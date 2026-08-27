import { describe, expect, it } from 'vitest'
// Importing the scheduling package registers the canonical leveler (irrelevant
// here) AND gives this suite the REAL scheduling authority for the
// scheduling-effect tests below (the accepted PROJECT-013 test-layer import
// precedent: the engine package itself never depends on the scheduler).
import { schedule } from '@genoffice/project-scheduling'
import { applyProjectCommand, canonicalizeDocument } from '../src/index.js'
import { asISODateTime, asTaskId } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { MONDAY, MONDAY_FINISH, TUESDAY_FINISH, makeDocument, makeTask, wm } from './fixtures.js'

const parent = (id: string) => asTaskId(id)

const leafDocument = (): ProjectDocument =>
  canonicalizeDocument(
    makeDocument({
      tasks: [makeTask({ id: 'a', duration: wm(480), start: asISODateTime(MONDAY) })],
    }),
  )

/** A canonical two-level outline (summary + one leaf) — the summary's
 * duration/finish are derived roll-ups of its subtree. */
const summaryDocument = (): ProjectDocument =>
  canonicalizeDocument(
    makeDocument({
      tasks: [
        makeTask({ id: 's', summary: true, duration: wm(0) }),
        makeTask({ id: 'c', parentTaskId: parent('s'), outlineLevel: 2 }),
      ],
    }),
  )

const snapshot = (value: unknown): string => JSON.stringify(value)

// ===========================================================================
// PROJECT-023 SetTaskDuration
// ===========================================================================

describe('PROJECT-023 SetTaskDuration — accept', () => {
  it('sets the leaf duration and returns the previous-duration inverse', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(960),
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.tasks[0].duration).toBe(wm(960))
    expect(exec.result.affectedTaskIds).toEqual([parent('a')])
    expect(exec.result.inverse).toEqual({
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(480),
    })
  })

  it('undo via the inverse restores the previous duration', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(960),
    })
    const undone = applyProjectCommand(exec.document, exec.result.inverse!)
    expect(undone.document.tasks[0].duration).toBe(wm(480))
    // The inverse round-trip restores the input byte-identically.
    expect(snapshot(undone.document)).toBe(snapshot(document))
  })

  it('accepts zero duration (the milestone domain)', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(0),
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.tasks[0].duration).toBe(wm(0))
  })

  it('never mutates the input document and is deterministic', () => {
    const document = leafDocument()
    const before = snapshot(document)
    const first = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(600),
    })
    const second = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(600),
    })
    expect(snapshot(document)).toBe(before)
    expect(snapshot(first.document)).toBe(snapshot(second.document))
    expect(first.document).not.toBe(document)
  })
})

describe('PROJECT-023 SetTaskDuration — reject', () => {
  it('rejects a missing task', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('zzz'),
      duration: wm(480),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'MISSING_TASK')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a summary task (the duration is a derived roll-up)', () => {
    const document = summaryDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('s'),
      duration: wm(480),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'SUMMARY_DURATION_NOT_SETTABLE')).toBe(
      true,
    )
    expect(exec.document).toBe(document)
  })

  it('rejects values outside the non-negative integer working-minute domain', () => {
    const document = leafDocument()
    for (const value of [-1, -480, 0.5, 480.25, Number.NaN, Infinity, -Infinity]) {
      const exec = applyProjectCommand(document, {
        type: 'SetTaskDuration',
        taskId: parent('a'),
        duration: wm(value),
      })
      expect(exec.result.accepted, `duration ${value}`).toBe(false)
      expect(exec.result.diagnostics.some((d) => d.code === 'INVALID_DURATION')).toBe(true)
      expect(exec.document).toBe(document)
    }
  })
})

describe('PROJECT-023 SetTaskDuration — scheduling effect (real scheduler)', () => {
  it('moving the leaf duration moves the derived finish through the scheduling authority', () => {
    const document = leafDocument()
    const before = schedule(document)
    expect(before.taskSchedules[parent('a')].scheduledFinish).toBe(asISODateTime(MONDAY_FINISH))

    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(960),
    })
    const after = schedule(exec.document)
    // 960 working minutes from the Monday 09:00 start with the standard
    // 8-hour calendar lands at Tuesday 17:00 — the finish moved exactly as
    // the scheduler (the sole scheduling authority) derives it.
    expect(after.taskSchedules[parent('a')].scheduledFinish).toBe(asISODateTime(TUESDAY_FINISH))
    expect(after.taskSchedules[parent('a')].scheduledStart).toBe(asISODateTime(MONDAY))
  })

  it('the engine itself computes no dates — only the scheduler derives them', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskDuration',
      taskId: parent('a'),
      duration: wm(960),
    })
    // The engine stored the semantic input; every other task field is
    // byte-identical (no derived date was touched by the command).
    const before = document.tasks[0]
    const after = exec.document.tasks[0]
    expect(after.start).toBe(before.start)
    expect(after.finish).toBe(before.finish)
    expect(after.name).toBe(before.name)
    expect(after.percentComplete).toBe(before.percentComplete)
  })
})

// ===========================================================================
// PROJECT-023 SetTaskFinish
// ===========================================================================

describe('PROJECT-023 SetTaskFinish — accept', () => {
  it('sets the stored finish and returns the previous-finish inverse', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            duration: wm(480),
            start: asISODateTime(MONDAY),
            finish: asISODateTime(MONDAY_FINISH),
          }),
        ],
      }),
    )
    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('a'),
      finish: TUESDAY_FINISH,
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.tasks[0].finish).toBe(asISODateTime(TUESDAY_FINISH))
    expect(exec.result.affectedTaskIds).toEqual([parent('a')])
    expect(exec.result.inverse).toEqual({
      type: 'SetTaskFinish',
      taskId: parent('a'),
      finish: MONDAY_FINISH,
    })
  })

  it('emits no inverse when the task had no previous finish (undo needs a snapshot)', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('a'),
      finish: TUESDAY_FINISH,
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.tasks[0].finish).toBe(asISODateTime(TUESDAY_FINISH))
    expect(exec.result.inverse).toBeUndefined()
  })

  it('undo via the inverse restores the previous finish', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            duration: wm(480),
            finish: asISODateTime(MONDAY_FINISH),
          }),
        ],
      }),
    )
    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('a'),
      finish: TUESDAY_FINISH,
    })
    const undone = applyProjectCommand(exec.document, exec.result.inverse!)
    expect(undone.document.tasks[0].finish).toBe(asISODateTime(MONDAY_FINISH))
    expect(snapshot(undone.document)).toBe(snapshot(document))
  })
})

describe('PROJECT-023 SetTaskFinish — reject', () => {
  it('rejects a missing task', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('zzz'),
      finish: TUESDAY_FINISH,
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'MISSING_TASK')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a summary task (the finish is a derived roll-up)', () => {
    const document = summaryDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('s'),
      finish: TUESDAY_FINISH,
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'SUMMARY_FINISH_NOT_SETTABLE')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a malformed date', () => {
    const document = leafDocument()
    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('a'),
      finish: 'not-a-date',
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'INVALID_DATE')).toBe(true)
    expect(exec.document).toBe(document)
  })
})

describe('PROJECT-023 SetTaskFinish — honest inert scheduling semantics', () => {
  it('pins the STORED finish; the DERIVED schedule does not move', () => {
    const document = leafDocument()
    const before = schedule(document)
    const derivedFinishBefore = before.taskSchedules[parent('a')].scheduledFinish

    const exec = applyProjectCommand(document, {
      type: 'SetTaskFinish',
      taskId: parent('a'),
      finish: TUESDAY_FINISH,
    })
    expect(exec.document.tasks[0].finish).toBe(asISODateTime(TUESDAY_FINISH))

    // The stored finish is an interchange echo, NOT a scheduling input: the
    // scheduler derives the finish from start + duration and never reads the
    // stored field, so the derived schedule is byte-identical.
    const after = schedule(exec.document)
    expect(after.taskSchedules[parent('a')].scheduledFinish).toBe(derivedFinishBefore)
    expect(snapshot(after)).toBe(snapshot(before))
  })
})
