import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
// The REAL scheduling authority at the TEST layer only (the accepted
// PROJECT-021 precedent): the renderer-core package itself stays
// scheduling-free with the scheduler injected through the session.
import { schedule } from '@genoffice/project-scheduling'
import {
  commitTaskEditThroughSession,
  createRendererSession,
  createViewState,
  projectDocumentView,
  redoRendererCommand,
  reduceViewState,
  undoRendererCommand,
} from '../src/index.js'
import { makeDocument, makeTask } from './fixtures.js'

/** A two-leaf document on the standard calendar: a (480 min from Mon 09:00),
 * b (960 min from Mon 09:00 — finishes Tue 17:00). */
const documentWithLeaves = (): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({
        id: 'a',
        name: 'Design',
        duration: asWorkingMinutes(480),
        start: asISODateTime('2026-08-03T09:00:00.000Z'),
      }),
      makeTask({ id: 'b', name: 'Build', duration: asWorkingMinutes(960) }),
    ],
  })

const gesture = (
  document: ProjectDocument,
  taskId: string,
  field: 'taskName' | 'duration' | 'start' | 'finish',
  draft: string,
) => {
  const session = createRendererSession(document, { schedule })
  let state = createViewState(document, session.schedule)
  // The user gesture chain: focus the row, activate the editor, type, commit.
  state = reduceViewState(
    state,
    { type: 'selectTask', taskId: asTaskId(taskId) },
    { document, schedule: session.schedule },
  )
  state = reduceViewState(
    state,
    { type: 'beginTaskEdit', taskId: asTaskId(taskId), field },
    { document, schedule: session.schedule },
  )
  state = reduceViewState(
    state,
    { type: 'updateTaskEditDraft', draft },
    { document, schedule: session.schedule },
  )
  return { session, state }
}

describe('PROJECT-023 edit flow — the full command/session pipeline', () => {
  it('a name edit: gesture → intent → RenameTask → engine → document → scheduler → projection refresh', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'a', 'taskName', 'Design v2')

    const outcome = commitTaskEditThroughSession(session, state)

    // The semantic command went through the engine and the session.
    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.result?.accepted).toBe(true)
    expect(outcome.entry?.command).toEqual({
      type: 'RenameTask',
      taskId: asTaskId('a'),
      name: 'Design v2',
    })
    expect(outcome.entry?.commandId).toBe('c1')
    // The document changed and was re-scheduled through the injected runner.
    expect(outcome.session.document.tasks[0]!.name).toBe('Design v2')
    expect(outcome.session.schedule).toEqual(schedule(outcome.session.document))
    expect(outcome.session.revision).toBe(1)
    // The editor ended; the selection SURVIVED the accepted mutation.
    expect('editing' in outcome.state).toBe(false)
    expect(outcome.state.tasks.taskIds).toEqual([asTaskId('a')])
    expect(outcome.state.tasks.focusId).toBe(asTaskId('a'))
    // The projection refresh shows the edited name — the hosts render this.
    const projection = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      outcome.state,
    )
    expect(projection.rows[0]!.name).toBe('Design v2')
    // The input session/state are never mutated.
    expect(document.tasks[0]!.name).toBe('Design')
    expect(state.editing?.draft).toBe('Design v2')
  })

  it('a duration edit moves the DERIVED finish through the real scheduler', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'b', 'duration', '1440')

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.result?.accepted).toBe(true)
    expect(outcome.entry?.command).toEqual({
      type: 'SetTaskDuration',
      taskId: asTaskId('b'),
      duration: asWorkingMinutes(1440),
    })
    // 1440 working minutes from Mon 09:00 on the standard 8-hour calendar
    // lands at Wed 17:00 (3 × 480-minute days) — the scheduling authority
    // derived the new finish.
    const row = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      outcome.state,
    ).rows[1]!
    expect(row.schedule?.scheduledFinish).toBe(asISODateTime('2026-08-05T17:00:00.000Z'))
    expect(row.duration).toBe(asWorkingMinutes(1440))
  })

  it('a start edit PINS the candidate start through the scheduler', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'b', 'start', '2026-08-04T09:00:00.000Z')

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.entry?.command).toEqual({
      type: 'SetTaskStart',
      taskId: asTaskId('b'),
      start: '2026-08-04T09:00:00.000Z',
    })
    expect(outcome.result?.accepted).toBe(true)
    const row = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      outcome.state,
    ).rows[1]!
    expect(row.schedule?.scheduledStart).toBe(asISODateTime('2026-08-04T09:00:00.000Z'))
  })

  it('a finish edit pins the STORED echo; the derived schedule stays honest', () => {
    const document = documentWithLeaves()
    const derivedFinishBefore = schedule(document).taskSchedules[asTaskId('b')]!.scheduledFinish
    const { session, state } = gesture(document, 'b', 'finish', '2026-08-10T17:00:00.000Z')

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.entry?.command).toEqual({
      type: 'SetTaskFinish',
      taskId: asTaskId('b'),
      finish: '2026-08-10T17:00:00.000Z',
    })
    expect(outcome.session.document.tasks[1]!.finish).toBe(
      asISODateTime('2026-08-10T17:00:00.000Z'),
    )
    // The stored finish is an interchange echo, not a scheduling input: the
    // derived finish is unchanged (start + duration rule it).
    expect(outcome.session.schedule?.taskSchedules[asTaskId('b')]!.scheduledFinish).toBe(
      derivedFinishBefore,
    )
  })

  it('an unparseable duration draft dispatches NO command and ends the editor', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'a', 'duration', '8h')

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit).toEqual({ kind: 'invalid', reason: 'unparseableDuration' })
    expect(outcome.result).toBeUndefined() // nothing reached the engine
    expect(outcome.session).toBe(session) // same reference — no journal entry
    expect(outcome.session.past).toHaveLength(0)
    expect('editing' in outcome.state).toBe(false)
    // Selection preserved after the rejected-at-translation edit.
    expect(outcome.state.tasks.taskIds).toEqual([asTaskId('a')])
  })

  it('an ENGINE-rejected edit surfaces the diagnostics and preserves everything', () => {
    const document = documentWithLeaves()
    // "-480" parses to a number; the ENGINE owns the semantic rejection.
    const { session, state } = gesture(document, 'a', 'duration', '-480')

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.result?.accepted).toBe(false)
    expect(outcome.result?.diagnostics.some((d) => d.code === 'INVALID_DURATION')).toBe(true)
    // Rejected: the session reference is unchanged, nothing journaled.
    expect(outcome.session).toBe(session)
    expect(outcome.session.past).toHaveLength(0)
    expect(outcome.session.document).toBe(document)
    // The editor still ends (the MS revert-cell behavior) and the selection
    // is preserved after the rejected mutation.
    expect('editing' in outcome.state).toBe(false)
    expect(outcome.state.tasks.taskIds).toEqual([asTaskId('a')])
    expect(outcome.state.tasks.focusId).toBe(asTaskId('a'))
  })

  it('an unchanged draft is a noChange: no command, no journal entry', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'a', 'taskName', 'Design')

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit).toEqual({ kind: 'noChange' })
    expect(outcome.session).toBe(session)
    expect('editing' in outcome.state).toBe(false)
  })

  it('committing with no active edit is a pure no-op (same references)', () => {
    const document = documentWithLeaves()
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.commit).toEqual({ kind: 'none' })
    expect(outcome.session).toBe(session)
    expect(outcome.state).toBe(state)
  })
})

describe('PROJECT-023 edit flow — undo/redo over committed edits', () => {
  it('undo restores the exact pre-edit document AND schedule; redo reapplies', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'a', 'duration', '960')
    const outcome = commitTaskEditThroughSession(session, state)
    expect(outcome.session.document.tasks[0]!.duration).toBe(asWorkingMinutes(960))

    const undo = undoRendererCommand(outcome.session)
    expect(undo.applied).toBe(true)
    expect(undo.session.document.tasks[0]!.duration).toBe(asWorkingMinutes(480))
    // Byte-identical restoration of the prior document AND schedule.
    expect(undo.session.document).toBe(document)
    expect(undo.session.schedule).toBe(session.schedule)

    const redo = redoRendererCommand(undo.session)
    expect(redo.applied).toBe(true)
    expect(redo.session.document.tasks[0]!.duration).toBe(asWorkingMinutes(960))
    expect(redo.session.schedule).toEqual(schedule(redo.session.document))
  })

  it('the view state reconciles cleanly against the undone document (selection preserved)', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'a', 'taskName', 'Design v2')
    const outcome = commitTaskEditThroughSession(session, state)
    expect(outcome.state.tasks.taskIds).toEqual([asTaskId('a')])

    const undo = undoRendererCommand(outcome.session)
    const reconciled = reduceViewState(
      outcome.state,
      { type: 'endTaskEdit' },
      { document: undo.session.document, schedule: undo.session.schedule },
    )
    // The task survived the undo: selection and focus are preserved.
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect(reconciled.tasks.focusId).toBe(asTaskId('a'))
    // And the projection over the undone document shows the original name.
    expect(
      projectDocumentView(undo.session.document, undo.session.schedule, reconciled).rows[0]!.name,
    ).toBe('Design')
  })
})

describe('PROJECT-023 edit flow — selection preservation under document replacement', () => {
  it('a multi-selection survives a committed edit of one selected row', () => {
    const document = documentWithLeaves()
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTasks', taskIds: [asTaskId('a'), asTaskId('b')] },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'beginTaskEdit', taskId: asTaskId('a'), field: 'taskName' },
      { document, schedule: session.schedule },
    )
    // beginTaskEdit selects the edited row; extend back to b for a range.
    state = reduceViewState(
      state,
      { type: 'moveTaskFocus', direction: 'last', extend: true },
      { document, schedule: session.schedule },
    )
    expect(state.tasks.taskIds).toEqual([asTaskId('a'), asTaskId('b')])

    // The active edit (target `a`, draft 'Design v2') is committed while the
    // LIVE selection is the extended range [a, b] — the commit reconciles
    // the selection against the new document and BOTH rows survive.
    state = reduceViewState(
      state,
      { type: 'updateTaskEditDraft', draft: 'Design v2' },
      { document, schedule: session.schedule },
    )
    const outcome = commitTaskEditThroughSession(session, state)
    expect(outcome.result?.accepted).toBe(true)
    expect(outcome.session.document.tasks[0]!.name).toBe('Design v2')
    // The live multi-selection [a, b] SURVIVES the committed edit of one of
    // its rows (reconciled against the new document — both tasks live).
    expect(outcome.state.tasks.taskIds).toEqual([asTaskId('a'), asTaskId('b')])
    expect(outcome.state.tasks.anchorId).toBe(asTaskId('a'))
    expect(outcome.state.tasks.focusId).toBe(asTaskId('b'))
  })

  it('an edit whose task is deleted before commit deterministically invalidates (missingTask)', () => {
    const document = documentWithLeaves()
    const { session, state } = gesture(document, 'a', 'taskName', 'X')

    // The document is replaced (e.g. an external delete) — the editing
    // state still points at the dead task; the COMMIT translation is the
    // last line of defense.
    const smaller = makeDocument({ tasks: [makeTask({ id: 'b' })] })
    const outcome = commitTaskEditThroughSession(
      { ...session, document: smaller, schedule: schedule(smaller) },
      state,
    )

    expect(outcome.commit).toEqual({ kind: 'invalid', reason: 'missingTask' })
    expect(outcome.session.document).toBe(smaller) // untouched
    expect('editing' in outcome.state).toBe(false)
    expect(outcome.state.tasks.taskIds).toEqual([]) // dead references pruned
  })

  it('the whole gesture chain is deterministic (3× byte-identical sessions and states)', () => {
    // One document built ONCE (the fixture uid counter advances per makeTask
    // call — building per run would leak different uids into the snapshots;
    // the documents are immutable so reusing one input is exactly the same
    // gesture three times).
    const document = documentWithLeaves()
    const run = (): string => {
      const { session, state } = gesture(document, 'a', 'duration', '960')
      const outcome = commitTaskEditThroughSession(session, state)
      return JSON.stringify({
        session: {
          document: outcome.session.document,
          schedule: outcome.session.schedule,
          past: outcome.session.past,
          revision: outcome.session.revision,
          commandSeq: outcome.session.commandSeq,
        },
        state: outcome.state,
      })
    }
    expect(run()).toBe(run())
    expect(run()).toBe(run())
  })
})
