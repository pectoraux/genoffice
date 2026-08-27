import { describe, expect, it } from 'vitest'
import {
  asBaselineId,
  asISODateTime,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type { ProjectCommand, ProjectDocument } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  applyRendererCommand,
  canRedoRendererCommand,
  canUndoRendererCommand,
  createRendererSession,
  redoRendererCommand,
  rendererSessionJournal,
  undoRendererCommand,
} from '../src/index.js'
import { makeDocument, makeTask, outlineDocument } from './fixtures.js'

const documentWithWork = (): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({ id: 't1', uid: 1, duration: asWorkingMinutes(480) }),
      makeTask({ id: 't2', uid: 2, duration: asWorkingMinutes(960) }),
    ],
  })

const rename = (taskId: string, name: string): ProjectCommand => ({
  type: 'RenameTask',
  taskId: asTaskId(taskId),
  name,
})

describe('PROJECT-021 session — creation', () => {
  it('creates a session without a schedule when no runner is wired', () => {
    const session = createRendererSession(outlineDocument())
    expect(session.revision).toBe(0)
    expect(session.past).toEqual([])
    expect(session.future).toEqual([])
    expect('schedule' in session).toBe(false)
    expect('scheduleRunner' in session).toBe(false)
  })

  it('computes the initial schedule through the injected runner (real scheduling package)', () => {
    const document = documentWithWork()
    const session = createRendererSession(document, { schedule })
    expect(session.schedule).toEqual(schedule(document))
  })
})

describe('PROJECT-021 session — applyCommand', () => {
  it('journals an accepted command, reschedules, and never mutates the input session', () => {
    const document = documentWithWork()
    const session = createRendererSession(document, { schedule })
    const before = JSON.stringify(session)
    const outcome = applyRendererCommand(session, rename('t1', 'Renamed'))
    expect(outcome.result.accepted).toBe(true)
    expect(outcome.entry?.commandId).toBe('c1')
    expect(outcome.session.revision).toBe(1)
    expect(outcome.session.commandSeq).toBe(1)
    expect(outcome.session.document.tasks[0]!.name).toBe('Renamed')
    expect(outcome.session.schedule).toEqual(schedule(outcome.session.document))
    expect(JSON.stringify(session)).toBe(before) // immutability
  })

  it('returns the SAME session reference with diagnostics for a rejected command', () => {
    const session = createRendererSession(documentWithWork())
    const outcome = applyRendererCommand(session, rename('ghost', 'X'))
    expect(outcome.result.accepted).toBe(false)
    expect(outcome.result.diagnostics.some((d) => d.code === 'MISSING_TASK')).toBe(true)
    expect(outcome.session).toBe(session)
    expect(outcome.entry).toBeUndefined()
  })

  it('rejected commands never consume a command id', () => {
    const session = createRendererSession(documentWithWork())
    const rejected = applyRendererCommand(session, rename('ghost', 'X'))
    expect(rejected.result.commandId).toBe('c1')
    const accepted = applyRendererCommand(session, rename('t1', 'A'))
    expect(accepted.entry!.commandId).toBe('c1')
  })

  it('re-schedules after a scheduling-relevant mutation: the schedule tracks the live document', () => {
    const document = documentWithWork() // t1: 1 day, t2: 2 days
    const session = createRendererSession(document, { schedule })
    const first = session.schedule!
    const outcome = applyRendererCommand(session, {
      type: 'SetPercentComplete',
      taskId: asTaskId('t1'),
      percentComplete: 50,
    })
    expect(outcome.session.schedule!.taskSchedules[asTaskId('t1')]!.percentComplete).toBe(50)
    expect(outcome.session.schedule!.taskSchedules[asTaskId('t1')]!.actualDuration).toBe(
      asWorkingMinutes(240),
    )
    expect(outcome.session.schedule).not.toEqual(first)
  })

  it('a session without a runner applies commands and simply carries no schedule', () => {
    const session = createRendererSession(documentWithWork())
    const outcome = applyRendererCommand(session, rename('t1', 'Renamed'))
    expect(outcome.result.accepted).toBe(true)
    expect('schedule' in outcome.session).toBe(false)
  })
})

describe('PROJECT-021 session — snapshot undo/redo', () => {
  it('undo restores the EXACT prior document and derived schedule (byte-identical)', () => {
    const document = documentWithWork()
    const session = createRendererSession(document, { schedule })
    const documentBefore = JSON.stringify(session.document)
    const scheduleBefore = JSON.stringify(session.schedule)
    const stepped = applyRendererCommand(
      applyRendererCommand(session, rename('t1', 'One')).session,
      {
        type: 'SetPercentComplete',
        taskId: asTaskId('t2'),
        percentComplete: 75,
      },
    ).session
    expect(canUndoRendererCommand(stepped)).toBe(true)
    const undone = undoRendererCommand(stepped)
    expect(undone.applied).toBe(true)
    expect(undone.entry!.commandId).toBe('c2')
    expect(JSON.stringify(undone.session.document)).not.toBe(JSON.stringify(stepped.document))
    const undoneAgain = undoRendererCommand(undone.session)
    expect(undoneAgain.applied).toBe(true)
    expect(JSON.stringify(undoneAgain.session.document)).toBe(documentBefore)
    expect(JSON.stringify(undoneAgain.session.schedule)).toBe(scheduleBefore)
    expect(canUndoRendererCommand(undoneAgain.session)).toBe(false)
  })

  it('redo restores the after snapshot and a new command clears the future', () => {
    const session = createRendererSession(documentWithWork(), { schedule })
    const applied = applyRendererCommand(session, rename('t1', 'One')).session
    const undone = undoRendererCommand(applied)
    expect(canRedoRendererCommand(undone.session)).toBe(true)
    const redone = redoRendererCommand(undone.session)
    expect(redone.applied).toBe(true)
    expect(redone.session.document.tasks[0]!.name).toBe('One')
    expect(canRedoRendererCommand(redone.session)).toBe(false)
    // Undo, then apply a NEW command: the redo future is cleared (journal semantics).
    const undoneAgain = undoRendererCommand(redone.session).session
    const diverged = applyRendererCommand(undoneAgain, rename('t1', 'Two')).session
    expect(diverged.future).toEqual([])
    expect(canRedoRendererCommand(diverged)).toBe(false)
  })

  it('undo/redo are deterministic no-ops on empty stacks (same reference)', () => {
    const session = createRendererSession(documentWithWork())
    expect(undoRendererCommand(session).session).toBe(session)
    expect(redoRendererCommand(session).session).toBe(session)
    expect(undoRendererCommand(session).applied).toBe(false)
    expect(redoRendererCommand(session).applied).toBe(false)
  })

  it('undoes the deliberately non-invertible OutdentTask via the snapshot (engine-documented host-layer duty)', () => {
    const document = outlineDocument() // root, a, a1, b
    const session = createRendererSession(document, { schedule })
    const outcome = applyRendererCommand(session, {
      type: 'OutdentTask',
      taskId: asTaskId('a'),
    })
    expect(outcome.result.accepted).toBe(true)
    expect(outcome.result.inverse).toBeUndefined() // the engine provides no inverse
    // Canonical outdent semantics: `a` (with its subtree) becomes the next
    // ROOT sibling after `root`; `b` remains root's child, so DFS order is
    // root, b, a, a1.
    expect(outcome.session.document.tasks.map((task) => task.id)).toEqual([
      asTaskId('root'),
      asTaskId('b'),
      asTaskId('a'),
      asTaskId('a1'),
    ])
    expect(
      outcome.session.document.tasks.find((task) => task.id === asTaskId('a'))!.parentTaskId,
    ).toBeUndefined()
    const undone = undoRendererCommand(outcome.session)
    expect(JSON.stringify(undone.session.document)).toBe(JSON.stringify(document))
  })

  it('undoes the non-invertible CreateBaseline via the snapshot', () => {
    const document = documentWithWork()
    const session = createRendererSession(document, { schedule })
    const outcome = applyRendererCommand(session, {
      type: 'CreateBaseline',
      baseline: {
        id: asBaselineId('b1'),
        name: 'Baseline',
        capturedAt: asISODateTime('2026-08-03T09:00:00.000Z'),
        taskSnapshots: {},
      },
    })
    expect(outcome.result.accepted).toBe(true)
    expect(outcome.result.inverse).toBeUndefined()
    expect(outcome.session.document.baselines).toHaveLength(1)
    const undone = undoRendererCommand(outcome.session)
    expect(undone.session.document.baselines).toHaveLength(0)
    expect(JSON.stringify(undone.session.document)).toBe(JSON.stringify(document))
  })

  it('exposes the journal as plain engine JournalEntry values', () => {
    const session = createRendererSession(documentWithWork(), { schedule })
    const applied = applyRendererCommand(session, rename('t1', 'One')).session
    const journal = rendererSessionJournal(applied)
    expect(journal).toHaveLength(1)
    expect(journal[0]!.commandId).toBe('c1')
    expect(journal[0]!.command).toEqual(rename('t1', 'One'))
    expect(journal[0]!.result.accepted).toBe(true)
  })
})

describe('PROJECT-021 session — end-to-end determinism (real scheduling package)', () => {
  it('the same scripted session (commands + undo + redo) is byte-identical across 3 runs', () => {
    const run = (): string => {
      let session = createRendererSession(documentWithWork(), { schedule })
      session = applyRendererCommand(session, rename('t1', 'One')).session
      session = applyRendererCommand(session, {
        type: 'SetPercentComplete',
        taskId: asTaskId('t2'),
        percentComplete: 75,
      }).session
      session = applyRendererCommand(session, buildCreateFromDocument(session.document)).session
      session = undoRendererCommand(session).session
      session = redoRendererCommand(session).session
      session = applyRendererCommand(session, rename('t2', 'Two')).session
      return JSON.stringify({
        document: session.document,
        schedule: session.schedule,
        revision: session.revision,
        commandSeq: session.commandSeq,
        past: session.past.map((entry) => entry.commandId),
        future: session.future.map((entry) => entry.commandId),
      })
    }
    expect(run()).toBe(run())
    expect(run()).toBe(run())
  })
})

/** A CreateTask built the way the renderer core builds one (deterministic identity). */
function buildCreateFromDocument(document: ProjectDocument): ProjectCommand {
  const uid = document.tasks.reduce((max, task) => Math.max(max, task.uid), 0) + 1
  return {
    type: 'CreateTask',
    task: {
      ...makeTask({ id: 't-new', uid, name: 'New Task' }),
      id: asTaskId('t-new'),
    },
  }
}
