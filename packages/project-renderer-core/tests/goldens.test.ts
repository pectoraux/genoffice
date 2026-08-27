/**
 * PROJECT-023 — the golden scenario battery (S01–S12).
 *
 * Each golden drives one acceptance scenario end-to-end through the real
 * machinery — intents through the reducer, edits through the commit flow,
 * commands through the session and the REAL scheduling authority — and
 * asserts the complete observable tuple: the selection (taskIds, anchorId,
 * focusId), the active edit target where applicable, the canonical
 * ProjectDocument, the DerivedSchedule, and the projection state the hosts
 * render. Pure inputs, no wall clock, no randomness — S12 proves the whole
 * battery deterministic (3× byte-identical).
 */
import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  applyRendererCommand,
  buildDeleteSelectionCommands,
  commitTaskEditThroughSession,
  createRendererSession,
  createViewState,
  projectDocumentView,
  reconcileViewState,
  redoRendererCommand,
  reduceViewState,
  undoRendererCommand,
} from '../src/index.js'
import type { ProjectViewState } from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

/**
 * The golden document: root summary over two leaves linked by an FS
 * dependency (a1 → a2). Real-scheduler dates on the standard Mon–Fri
 * 09:00–17:00 calendar: a1 Mon 09:00→17:00; a2 (after a1) Tue 09:00 →
 * Wed 17:00; root rolls both up. Built ONCE per scenario call — the fixture
 * uid counter makes per-call rebuilds differ, so determinism scenarios reuse
 * one instance.
 */
function goldenDocument(): ProjectDocument {
  return makeDocument({
    tasks: [
      makeTask({ id: 'root', uid: 1, outlineLevel: 1, summary: true, wbs: '1' }),
      makeTask({
        id: 'a1',
        uid: 2,
        parentTaskId: asTaskId('root'),
        outlineLevel: 2,
        name: 'Draft',
        duration: asWorkingMinutes(480),
        wbs: '1.1',
      }),
      makeTask({
        id: 'a2',
        uid: 3,
        parentTaskId: asTaskId('root'),
        outlineLevel: 2,
        name: 'Review',
        duration: asWorkingMinutes(960),
        wbs: '1.2',
      }),
    ],
    dependencies: [makeDependency('d1', 'a1', 'a2')],
  })
}

type Fixture = {
  document: ProjectDocument
  session: ReturnType<typeof createRendererSession>
  state: ProjectViewState
}

function goldenFixture(): Fixture {
  const document = goldenDocument()
  const session = createRendererSession(document, { schedule })
  const state = createViewState(document, session.schedule)
  return { document, session, state }
}

const ids = (...values: string[]) => values.map((value) => asTaskId(value))

describe('PROJECT-023 goldens — selection', () => {
  it('S01 — single selection', () => {
    const { document, session, state } = goldenFixture()
    const next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )

    expect(next.tasks).toEqual({
      taskIds: ids('a1'),
      anchorId: asTaskId('a1'),
      focusId: asTaskId('a1'),
    })
    expect('editing' in next).toBe(false)
    // The document and schedule are untouched by a pure view intent.
    expect(session.document).toBe(document)
    expect(session.schedule).toEqual(schedule(document))
    // The projection reflects the selection on exactly the focused row.
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows.map((row) => [row.taskId, row.selected, row.focused])).toEqual([
      [asTaskId('root'), false, false],
      [asTaskId('a1'), true, true],
      [asTaskId('a2'), false, false],
    ])
  })

  it('S02 — multi-selection (toggle)', () => {
    const { document, session, state } = goldenFixture()
    let next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    next = reduceViewState(
      next,
      { type: 'selectTask', taskId: asTaskId('a2'), mode: 'toggle' },
      { document, schedule: session.schedule },
    )
    next = reduceViewState(
      next,
      { type: 'selectTask', taskId: asTaskId('root'), mode: 'toggle' },
      { document, schedule: session.schedule },
    )

    // First-occurrence order; anchor/focus on the most recently toggled-in row.
    expect(next.tasks).toEqual({
      taskIds: ids('a1', 'a2', 'root'),
      anchorId: asTaskId('root'),
      focusId: asTaskId('root'),
    })
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows.map((row) => row.selected)).toEqual([true, true, true])
    expect(projection.rows.map((row) => row.focused)).toEqual([true, false, false])
    expect(session.document).toBe(document)
  })

  it('S03 — range selection (shift-extend over the canonical outline order)', () => {
    const { document, session, state } = goldenFixture()
    let next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    next = reduceViewState(
      next,
      { type: 'selectTask', taskId: asTaskId('root'), mode: 'extend' },
      { document, schedule: session.schedule },
    )

    // The outline-order range [root..a1] (both directions accepted).
    expect(next.tasks).toEqual({
      taskIds: ids('root', 'a1'),
      anchorId: asTaskId('a1'),
      focusId: asTaskId('root'),
    })
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows.map((row) => [row.taskId, row.selected])).toEqual([
      [asTaskId('root'), true],
      [asTaskId('a1'), true],
      [asTaskId('a2'), false],
    ])
    expect(projection.rows[0]!.focused).toBe(true)
  })

  it('S04 — collapsed hierarchy selection (hiding is not deselecting)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'root', uid: 1, outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({
          id: 'a',
          uid: 2,
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          summary: true,
          wbs: '1.1',
        }),
        makeTask({
          id: 'a1',
          uid: 3,
          parentTaskId: asTaskId('a'),
          outlineLevel: 3,
          duration: asWorkingMinutes(480),
          wbs: '1.1.1',
        }),
        makeTask({
          id: 'b',
          uid: 4,
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          duration: asWorkingMinutes(480),
          wbs: '1.2',
        }),
      ],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)

    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )

    // Hidden: not a row. Selected: still in the state — the documented policy.
    let projection = projectDocumentView(document, session.schedule, state)
    expect(projection.rows.map((row) => row.taskId)).toEqual(ids('root', 'a', 'b'))
    expect(state.tasks).toEqual({
      taskIds: ids('a1'),
      anchorId: asTaskId('a1'),
      focusId: asTaskId('a1'),
    })

    // Expand: the child re-projects as the selected, focused row.
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )
    projection = projectDocumentView(document, session.schedule, state)
    expect(projection.rows.map((row) => [row.taskId, row.selected])).toEqual([
      [asTaskId('root'), false],
      [asTaskId('a'), false],
      [asTaskId('a1'), true],
      [asTaskId('b'), false],
    ])
    expect(projection.rows[2]!.focused).toBe(true)
    expect(session.document).toBe(document)
  })

  it('S05 — keyboard range (bootstrap, walk, shift-extend)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'root', uid: 1, outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({
          id: 'a',
          uid: 2,
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          summary: true,
          wbs: '1.1',
        }),
        makeTask({
          id: 'a1',
          uid: 3,
          parentTaskId: asTaskId('a'),
          outlineLevel: 3,
          duration: asWorkingMinutes(480),
          wbs: '1.1.1',
        }),
        makeTask({
          id: 'a2',
          uid: 4,
          parentTaskId: asTaskId('a'),
          outlineLevel: 3,
          duration: asWorkingMinutes(480),
          wbs: '1.1.2',
        }),
        makeTask({
          id: 'b',
          uid: 5,
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          duration: asWorkingMinutes(480),
          wbs: '1.2',
        }),
      ],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)

    // Bootstrap from nothing: the first visible row.
    state = reduceViewState(
      state,
      { type: 'moveTaskFocus', direction: 'down' },
      { document, schedule: session.schedule },
    )
    expect(state.tasks).toEqual({
      taskIds: ids('root'),
      anchorId: asTaskId('root'),
      focusId: asTaskId('root'),
    })
    // Walk to `a`.
    state = reduceViewState(
      state,
      { type: 'moveTaskFocus', direction: 'down' },
      { document, schedule: session.schedule },
    )
    expect(state.tasks.focusId).toBe(asTaskId('a'))
    // Shift+down: the outline range from the anchor (a) to a1.
    state = reduceViewState(
      state,
      { type: 'moveTaskFocus', direction: 'down', extend: true },
      { document, schedule: session.schedule },
    )
    expect(state.tasks).toEqual({
      taskIds: ids('a', 'a1'),
      anchorId: asTaskId('a'),
      focusId: asTaskId('a1'),
    })
    // Shift+down again extends to a2.
    state = reduceViewState(
      state,
      { type: 'moveTaskFocus', direction: 'down', extend: true },
      { document, schedule: session.schedule },
    )
    expect(state.tasks).toEqual({
      taskIds: ids('a', 'a1', 'a2'),
      anchorId: asTaskId('a'),
      focusId: asTaskId('a2'),
    })
    const projection = projectDocumentView(document, session.schedule, state)
    expect(projection.rows.map((row) => row.selected)).toEqual([false, true, true, true, false])
    expect(projection.rows[3]!.focused).toBe(true)
  })
})

describe('PROJECT-023 goldens — editing', () => {
  it('S06 — rename through the full pipeline (identity stable, schedule recomputed)', () => {
    const { document, session, state } = goldenFixture()
    let viewState = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    viewState = reduceViewState(
      viewState,
      { type: 'beginTaskEdit', taskId: asTaskId('a1'), field: 'taskName' },
      { document, schedule: session.schedule },
    )
    expect(viewState.editing).toEqual({ taskId: asTaskId('a1'), field: 'taskName', draft: 'Draft' })
    viewState = reduceViewState(
      viewState,
      { type: 'updateTaskEditDraft', draft: 'Wireframes' },
      { document, schedule: session.schedule },
    )

    const outcome = commitTaskEditThroughSession(session, viewState)

    // The semantic command: RenameTask, journaled, accepted.
    expect(outcome.entry?.command).toEqual({
      type: 'RenameTask',
      taskId: asTaskId('a1'),
      name: 'Wireframes',
    })
    const renamed = outcome.session.document.tasks[1]!
    expect(renamed.name).toBe('Wireframes')
    // Stable identity and structure: id, uid, hierarchy, dependencies.
    expect(renamed.id).toBe(asTaskId('a1'))
    expect(renamed.uid).toBe(2)
    expect(renamed.parentTaskId).toBe(asTaskId('root'))
    expect(renamed.outlineLevel).toBe(2)
    expect(renamed.wbs).toBe('1.1')
    expect(outcome.session.document.dependencies).toEqual(document.dependencies)
    expect(outcome.session.document.baselines).toEqual(document.baselines)
    // The schedule was recomputed for the new document through the authority.
    expect(outcome.session.schedule).toEqual(schedule(outcome.session.document))
    expect(outcome.session.revision).toBe(1)
    // The editor ended; the selection survived the accepted mutation.
    expect('editing' in outcome.state).toBe(false)
    expect(outcome.state.tasks).toEqual({
      taskIds: ids('a1'),
      anchorId: asTaskId('a1'),
      focusId: asTaskId('a1'),
    })
    // The projection refresh carries the new name and the selected row.
    const projection = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      outcome.state,
    )
    expect(projection.rows[1]!.name).toBe('Wireframes')
    expect(projection.rows[1]!.selected).toBe(true)
  })

  it('S07 — percent completion (0 / partial / 100 / invalid, engine authority)', () => {
    const { session } = goldenFixture()

    // The documented percent-complete path: direct command construction
    // through the session (the PROJECT-021 rule).
    const at = (percent: number) => ({
      type: 'SetPercentComplete' as const,
      taskId: asTaskId('a1'),
      percentComplete: percent,
    })

    // 0 → accepted.
    const zero = applyRendererCommand(session, at(0))
    expect(zero.result.accepted).toBe(true)
    expect(zero.session.document.tasks[1]!.percentComplete).toBe(0)
    expect(zero.session.schedule).toEqual(schedule(zero.session.document))

    // partial → accepted.
    const partial = applyRendererCommand(zero.session, at(60))
    expect(partial.result.accepted).toBe(true)
    expect(partial.session.document.tasks[1]!.percentComplete).toBe(60)

    // 100 → accepted.
    const full = applyRendererCommand(partial.session, at(100))
    expect(full.result.accepted).toBe(true)
    expect(full.session.document.tasks[1]!.percentComplete).toBe(100)
    expect(full.session.revision).toBe(3)

    // invalid (out of range) → rejected, document unchanged (same reference).
    const invalid = applyRendererCommand(full.session, at(150))
    expect(invalid.result.accepted).toBe(false)
    expect(invalid.result.diagnostics[0]!.code).toBe('INVALID_PERCENT_COMPLETE')
    expect(invalid.session).toBe(full.session)

    // A summary's percent is derived — the engine rejects a direct set.
    const summary = applyRendererCommand(full.session, {
      type: 'SetPercentComplete',
      taskId: asTaskId('root'),
      percentComplete: 40,
    })
    expect(summary.result.accepted).toBe(false)
    expect(summary.session).toBe(full.session)
  })

  it('S08 — rejected edit (engine INVALID_DURATION leaves everything unchanged)', () => {
    const { document, session, state } = goldenFixture()
    let viewState = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    viewState = reduceViewState(
      viewState,
      { type: 'beginTaskEdit', taskId: asTaskId('a1'), field: 'duration' },
      { document, schedule: session.schedule },
    )
    viewState = reduceViewState(
      viewState,
      { type: 'updateTaskEditDraft', draft: '-5' },
      { document, schedule: session.schedule },
    )

    const outcome = commitTaskEditThroughSession(session, viewState)

    // The draft PARSED (canonical decimal) and became a command the ENGINE
    // rejected — the single validation authority.
    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.result?.accepted).toBe(false)
    expect(outcome.result?.diagnostics[0]!.code).toBe('INVALID_DURATION')
    // Nothing journaled; the session is the SAME reference.
    expect(outcome.entry).toBeUndefined()
    expect(outcome.session).toBe(session)
    expect(outcome.session.document).toBe(document)
    expect(outcome.session.revision).toBe(0)
    // The editor still ended (the revert-cell behavior); the selection is
    // deterministic and preserved.
    expect('editing' in outcome.state).toBe(false)
    expect(outcome.state.tasks).toEqual({
      taskIds: ids('a1'),
      anchorId: asTaskId('a1'),
      focusId: asTaskId('a1'),
    })
    const projection = projectDocumentView(document, session.schedule, outcome.state)
    expect(projection.rows[1]!.selected).toBe(true)
    expect(projection.rows[1]!.duration).toBe(asWorkingMinutes(480))
  })

  it('S09 — edit + reschedule (duration moves the derived finish)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          uid: 1,
          name: 'Design',
          duration: asWorkingMinutes(480),
          start: asISODateTime('2026-08-03T09:00:00.000Z'),
          wbs: '1',
        }),
        makeTask({ id: 'b', uid: 2, name: 'Build', duration: asWorkingMinutes(960), wbs: '2' }),
      ],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('b') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'beginTaskEdit', taskId: asTaskId('b'), field: 'duration' },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'updateTaskEditDraft', draft: '1440' },
      { document, schedule: session.schedule },
    )

    const outcome = commitTaskEditThroughSession(session, state)

    expect(outcome.entry?.command).toEqual({
      type: 'SetTaskDuration',
      taskId: asTaskId('b'),
      duration: asWorkingMinutes(1440),
    })
    expect(outcome.result?.accepted).toBe(true)
    // 1440 working minutes from Mon 09:00 on the standard 8-hour calendar:
    // the scheduling authority derives Wed 17:00 (3 × 480-minute days).
    const row = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      outcome.state,
    ).rows[1]!
    expect(row.schedule?.scheduledFinish).toBe(asISODateTime('2026-08-05T17:00:00.000Z'))
    expect(row.duration).toBe(asWorkingMinutes(1440))
    expect(outcome.session.schedule).toEqual(schedule(outcome.session.document))
    expect(outcome.state.tasks).toEqual({
      taskIds: ids('b'),
      anchorId: asTaskId('b'),
      focusId: asTaskId('b'),
    })
  })
})

describe('PROJECT-023 goldens — session', () => {
  it('S10 — undo/redo restores document + schedule + selection state', () => {
    const { document, session, state } = goldenFixture()
    let viewState = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    viewState = reduceViewState(
      viewState,
      { type: 'beginTaskEdit', taskId: asTaskId('a1'), field: 'taskName' },
      { document, schedule: session.schedule },
    )
    viewState = reduceViewState(
      viewState,
      { type: 'updateTaskEditDraft', draft: 'Wireframes' },
      { document, schedule: session.schedule },
    )
    const committed = commitTaskEditThroughSession(session, viewState)
    expect(committed.result?.accepted).toBe(true)

    // Undo restores the EXACT pre-edit document AND its derived schedule.
    const undone = undoRendererCommand(committed.session)
    expect(undone.applied).toBe(true)
    expect(undone.session.document).toEqual(document)
    expect(undone.session.schedule).toEqual(schedule(document))
    // The host reconciles the view state against the undone document: the
    // task still exists, so the selection survives verbatim.
    const undoneState = reconcileViewState(committed.state, undone.session.document)
    expect(undoneState.tasks).toEqual({
      taskIds: ids('a1'),
      anchorId: asTaskId('a1'),
      focusId: asTaskId('a1'),
    })

    // Redo restores the accepted post-edit state.
    const redone = redoRendererCommand(undone.session)
    expect(redone.applied).toBe(true)
    expect(redone.session.document).toEqual(committed.session.document)
    expect(redone.session.schedule).toEqual(committed.session.schedule)
    expect(redone.session.revision).toBe(3)
  })

  it('S11 — delete-selected-task reconciliation', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'root1', uid: 1, outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({
          id: 'a1',
          uid: 2,
          parentTaskId: asTaskId('root1'),
          outlineLevel: 2,
          duration: asWorkingMinutes(480),
          wbs: '1.1',
        }),
        makeTask({
          id: 'a2',
          uid: 3,
          parentTaskId: asTaskId('root1'),
          outlineLevel: 2,
          duration: asWorkingMinutes(480),
          wbs: '1.2',
        }),
        makeTask({
          id: 'root2',
          uid: 4,
          outlineLevel: 1,
          duration: asWorkingMinutes(480),
          wbs: '2',
        }),
      ],
      dependencies: [makeDependency('d1', 'a1', 'a2')],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a2') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('root2'), mode: 'toggle' },
      { document, schedule: session.schedule },
    )
    expect(state.tasks.taskIds).toEqual(ids('a2', 'root2'))

    // The multi-selection delete: deterministic reverse outline order.
    const commands = buildDeleteSelectionCommands(document, state.tasks.taskIds)
    expect(commands.map((command) => command.type)).toEqual(['DeleteTask', 'DeleteTask'])
    expect(
      commands.map((command) => (command.type === 'DeleteTask' ? command.taskId : undefined)),
    ).toEqual([asTaskId('root2'), asTaskId('a2')])

    let next = session
    for (const command of commands) {
      const outcome = applyRendererCommand(next, command)
      expect(outcome.result.accepted).toBe(true)
      next = outcome.session
    }
    // Both deleted; a1 and root1 survive with identity intact.
    expect(next.document.tasks.map((task) => task.id)).toEqual(ids('root1', 'a1'))
    // The dangling dependency was removed with its endpoints by the engine.
    expect(next.document.dependencies).toEqual([])
    expect(next.schedule).toEqual(schedule(next.document))

    // Host reconciliation: dead TaskIds pruned, anchor/focus reconciled to
    // nothing (no surviving members) — deterministic.
    const reconciled = reconcileViewState(state, next.document)
    expect(reconciled.tasks).toEqual({ taskIds: [] })
    const projection = projectDocumentView(next.document, next.schedule, reconciled)
    expect(projection.rows.map((row) => [row.taskId, row.selected])).toEqual([
      [asTaskId('root1'), false],
      [asTaskId('a1'), false],
    ])
  })

  it('S12 — deterministic repeated interaction sequence (3× byte-identical)', () => {
    // Built ONCE: the fixture uid counter would make per-run rebuilds differ.
    const document = goldenDocument()

    const run = () => {
      const session = createRendererSession(document, { schedule })
      let state = createViewState(document, session.schedule)
      const ctx = { document, schedule: session.schedule }

      state = reduceViewState(state, { type: 'moveTaskFocus', direction: 'down' }, ctx)
      state = reduceViewState(state, { type: 'moveTaskFocus', direction: 'down' }, ctx)
      state = reduceViewState(
        state,
        { type: 'moveTaskFocus', direction: 'down', extend: true },
        ctx,
      )
      state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('root') }, ctx)
      state = reduceViewState(state, { type: 'moveTaskFocus', direction: 'down' }, ctx)
      state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('root') }, ctx)
      state = reduceViewState(
        state,
        { type: 'beginTaskEdit', taskId: asTaskId('a2'), field: 'duration' },
        ctx,
      )
      state = reduceViewState(state, { type: 'updateTaskEditDraft', draft: '1440' }, ctx)
      const committed = commitTaskEditThroughSession(session, state)
      const undone = undoRendererCommand(committed.session)
      const redone = redoRendererCommand(undone.session)
      return {
        state: JSON.stringify(committed.state),
        undoneState: JSON.stringify(reconcileViewState(committed.state, undone.session.document)),
        document: JSON.stringify(redone.session.document),
        schedule: JSON.stringify(redone.session.schedule),
        projection: JSON.stringify(
          projectDocumentView(
            redone.session.document,
            redone.session.schedule,
            reconcileViewState(committed.state, redone.session.document),
          ),
        ),
      }
    }

    const first = run()
    expect(run()).toEqual(first)
    expect(run()).toEqual(first)
  })
})
