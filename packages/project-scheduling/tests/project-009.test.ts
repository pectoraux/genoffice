import { describe, expect, it } from 'vitest'
import { captureBaseline, compareBaseline, schedule, baselineIdOf } from '../src/index.js'
import { applyProjectCommand } from '@genoffice/project-engine'
import type {
  Baseline,
  BaselineComparison,
  BaselineVariance,
  DerivedSchedule,
  ProjectDocument,
  TaskSchedule,
} from '@genoffice/project-contracts'
import { asISODateTime } from '@genoffice/project-contracts'
import {
  MONDAY,
  MONDAY_FINISH,
  TUESDAY,
  TUESDAY_FINISH,
  WEDNESDAY,
  WEDNESDAY_FINISH,
  holiday,
  iso,
  makeCalendar,
  makeDependency,
  makeDocument,
  makeTask,
  parseDocument,
  taskId,
  wm,
  makeBaseline,
  asBaselineId,
} from './fixtures.js'
import type { Task } from '@genoffice/project-contracts'

const day = (minutes: number) => wm(minutes)

const resultOf = (document: ProjectDocument): DerivedSchedule => {
  const result = schedule(document)
  expect(result.diagnostics).toEqual([])
  return result
}

const scheduleOf = (result: DerivedSchedule, id: string): TaskSchedule => {
  const entry = result.taskSchedules[taskId(id)]
  if (!entry) throw new Error(`missing schedule for ${id}`)
  return entry
}

const varianceOf = (comparison: BaselineComparison, id: string): BaselineVariance => {
  const entry = comparison.variances[taskId(id)]
  if (!entry) throw new Error(`missing variance for ${id}`)
  return entry
}

const withStatusDate = (document: ProjectDocument, statusDate: string): ProjectDocument => ({
  ...document,
  properties: { ...document.properties, statusDate: asISODateTime(statusDate) },
})

const summary = (id: string, parentTaskId?: string, outlineLevel = 1): Task =>
  makeTask({
    id,
    summary: true,
    duration: wm(0),
    parentTaskId: parentTaskId === undefined ? undefined : taskId(parentTaskId),
    outlineLevel,
  })

describe('PROJECT-009 baseline capture', () => {
  it('captureBaseline defaults capturedAt to ProjectProperties.statusDate (never Date.now())', () => {
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] }),
      TUESDAY,
    )
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'))
    expect(baseline?.capturedAt).toBe(TUESDAY)
    expect(baseline?.id).toEqual(asBaselineId('b1'))
  })

  it('captureBaseline honors an explicit capturedAt override over statusDate', () => {
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] }),
      TUESDAY,
    )
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(WEDNESDAY),
    })
    expect(baseline?.capturedAt).toBe(WEDNESDAY)
  })

  it('captureBaseline returns undefined when no capturedAt and no statusDate are available', () => {
    // No statusDate on the project and no explicit capturedAt: a baseline
    // cannot be captured without a deterministic instant. The helper returns
    // undefined so the caller surfaces a clean diagnostic instead of silently
    // inventing a wall-clock timestamp.
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'))
    expect(baseline).toBeUndefined()
  })

  it('captureBaseline selection=all snapshots every task with its scheduled state', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(240) })],
    })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
      selection: { kind: 'all' },
    })!
    expect(Object.keys(baseline.taskSnapshots).sort()).toEqual(['a', 'b'])
    expect(baseline.taskSnapshots.a.start).toBe(MONDAY)
    expect(baseline.taskSnapshots.a.finish).toBe(MONDAY_FINISH)
    expect(baseline.taskSnapshots.a.duration).toBe(480)
    expect(baseline.taskSnapshots.b.duration).toBe(240)
  })

  it('captureBaseline selection=leaves omits summary tasks', () => {
    const document = makeDocument({
      tasks: [
        summary('s'),
        makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'b', duration: day(240), parentTaskId: taskId('s'), outlineLevel: 2 }),
      ],
    })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
      selection: { kind: 'leaves' },
    })!
    expect(Object.keys(baseline.taskSnapshots).sort()).toEqual(['a', 'b'])
    expect(baseline.taskSnapshots.s).toBeUndefined()
  })

  it('captureBaseline selection=tasks captures only the listed tasks', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
    })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
      selection: { kind: 'tasks', taskIds: [taskId('a'), taskId('c')] },
    })!
    expect(Object.keys(baseline.taskSnapshots).sort()).toEqual(['a', 'c'])
  })

  it('captureBaseline omits unschedulable tasks (no schedule entry)', () => {
    // A task whose schedule produced no entry is silently omitted rather than
    // carrying an empty snapshot.
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
      selection: { kind: 'tasks', taskIds: [taskId('a'), taskId('ghost')] },
    })!
    expect(Object.keys(baseline.taskSnapshots)).toEqual(['a'])
  })
})

describe('PROJECT-009 baseline comparison (variance sign convention)', () => {
  it('start variance is positive when the current start slips past the baseline', () => {
    // Baseline: a starts Monday. Current: a constrained to start Tuesday.
    const baseline = makeBaseline('b1', MONDAY, {
      a: { start: MONDAY, finish: MONDAY_FINISH, duration: 480 },
    })
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ],
    })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBe(480) // +1 working day slipped
    expect(v.finishVariance).toBe(480)
    expect(v.durationVariance).toBe(0)
  })

  it('finish variance is negative when the current finish moves earlier (ahead of plan)', () => {
    // Baseline: a finishes Tuesday 17:00 (SNET Tuesday). Current: a is ASAP,
    // so it finishes Monday 17:00 — one working day ahead.
    const baseline = makeBaseline('b1', MONDAY, {
      a: { start: TUESDAY, finish: TUESDAY_FINISH, duration: 480 },
    })
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBe(-480)
    expect(v.finishVariance).toBe(-480)
    expect(v.durationVariance).toBe(0)
  })

  it('duration variance is positive when the current task is longer than planned', () => {
    const baseline = makeBaseline('b1', MONDAY, {
      a: { start: MONDAY, finish: MONDAY_FINISH, duration: 480 },
    })
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(720) })] })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.durationVariance).toBe(240) // 720 - 480
    expect(v.finishVariance).toBe(240) // finishes one half-day later
    expect(v.startVariance).toBe(0) // starts at the same instant
  })

  it('start/finish variance are undefined when the baseline snapshot lacks the date', () => {
    // Baseline captured a duration-only snapshot (no start/finish). The
    // comparison cannot compute start/finish variance, so they are undefined
    // (not zero). Duration variance is still defined.
    const baseline = makeBaseline('b1', MONDAY, { a: { duration: 480 } })
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBeUndefined()
    expect(v.finishVariance).toBeUndefined()
    expect(v.durationVariance).toBe(0)
    expect(v.baselineStart).toBeUndefined()
    expect(v.baselineFinish).toBeUndefined()
  })

  it('compareBaseline omits tasks that the baseline did not snapshot', () => {
    const baseline = makeBaseline('b1', MONDAY, { a: { duration: 480 } })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(240) })],
    })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    expect(Object.keys(comparison.variances).sort()).toEqual(['a'])
    expect(comparison.variances[taskId('b')]).toBeUndefined()
  })

  it('compareBaseline omits tasks that no longer exist in the document (dangling snapshot)', () => {
    const baseline = makeBaseline('b1', MONDAY, {
      a: { duration: 480 },
      ghost: { duration: 240 },
    })
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    expect(Object.keys(comparison.variances).sort()).toEqual(['a'])
  })
})

describe('PROJECT-009 baseline golden scenarios', () => {
  it('B01: baseline == current schedule produces zero variance on every field', () => {
    const document = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const comparison = compareBaseline(document, result, baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBe(0)
    expect(v.finishVariance).toBe(0)
    expect(v.durationVariance).toBe(0)
    expect(v.baselineStart).toBe(MONDAY)
    expect(v.baselineFinish).toBe(MONDAY_FINISH)
    expect(v.baselineDuration).toBe(480)
  })

  it('B02: a slipped task produces positive start and finish variance', () => {
    const original = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const originalSchedule = resultOf(original)
    const baseline = captureBaseline(original, originalSchedule, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const slipped = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ],
    })
    const comparison = compareBaseline(slipped, resultOf(slipped), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBe(480)
    expect(v.finishVariance).toBe(480)
    expect(v.durationVariance).toBe(0)
  })

  it('B03: an ahead-of-plan task produces negative start and finish variance', () => {
    const planned = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ],
    })
    const plannedSchedule = resultOf(planned)
    const baseline = captureBaseline(planned, plannedSchedule, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const ahead = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const comparison = compareBaseline(ahead, resultOf(ahead), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBe(-480)
    expect(v.finishVariance).toBe(-480)
    expect(v.durationVariance).toBe(0)
  })

  it('B04: a longer duration produces positive duration and finish variance, zero start variance', () => {
    const original = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const baseline = captureBaseline(original, resultOf(original), baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const grown = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(960) })] })
    const comparison = compareBaseline(grown, resultOf(grown), baseline)
    const v = varianceOf(comparison, 'a')
    expect(v.startVariance).toBe(0)
    expect(v.durationVariance).toBe(480) // 960 - 480
    expect(v.finishVariance).toBe(480) // finishes one day later
  })

  it('B05: a milestone baseline variance compares the single scheduled instant', () => {
    const planned = makeDocument({
      tasks: [
        makeTask({
          id: 'm',
          duration: day(0),
          milestone: true,
          constraintType: 'mustStartOn',
          constraintDate: iso(TUESDAY),
        }),
      ],
    })
    const baseline = captureBaseline(planned, resultOf(planned), baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const moved = makeDocument({
      tasks: [
        makeTask({
          id: 'm',
          duration: day(0),
          milestone: true,
          constraintType: 'mustStartOn',
          constraintDate: iso(WEDNESDAY),
        }),
      ],
    })
    const comparison = compareBaseline(moved, resultOf(moved), baseline)
    const v = varianceOf(comparison, 'm')
    expect(v.startVariance).toBe(480) // Tue -> Wed = 1 working day
    expect(v.finishVariance).toBe(480) // milestone: start == finish
    expect(v.durationVariance).toBe(0)
    expect(v.baselineStart).toBe(TUESDAY)
    expect(v.baselineFinish).toBe(TUESDAY)
  })

  it('B06: a summary baseline variance reflects the rolled-up child changes', () => {
    // Children are FS-linked (a -> b) so growing b actually pushes the summary
    // finish later. Without the link both children start Monday and the
    // summary finish never moves.
    const original = makeDocument({
      tasks: [
        summary('s'),
        makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'b', duration: day(240), parentTaskId: taskId('s'), outlineLevel: 2 }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
    })
    const originalSchedule = resultOf(original)
    const baseline = captureBaseline(original, originalSchedule, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    // a: Mon 09:00-17:00. b (240 min after a): Tue 09:00-13:00. Summary rolls
    // up to Mon 09:00 .. Tue 13:00 (working span = 480 Mon + 240 Tue = 720).
    expect(baseline.taskSnapshots.s.start).toBe(MONDAY)
    expect(baseline.taskSnapshots.s.finish).toBe('2026-08-04T13:00:00.000Z')
    // Grow b to a full day so the summary now finishes Tuesday 17:00.
    const grown = makeDocument({
      tasks: [
        summary('s'),
        makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'b', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
    })
    const comparison = compareBaseline(grown, resultOf(grown), baseline)
    const s = varianceOf(comparison, 's')
    // Summary finish slipped from Tue 13:00 to Tue 17:00 = 240 working minutes.
    expect(s.finishVariance).toBe(240)
    expect(s.startVariance).toBe(0) // summary start unchanged (a still starts Monday)
    expect(s.durationVariance).toBe(240) // summary span grew 720 -> 960
  })

  it('B07: an FS chain baseline variance isolates the slippage downstream', () => {
    const original = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
    })
    const baseline = captureBaseline(original, resultOf(original), baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    // Add a one-day lag to the a->b link so b and c slip by a day; a is fixed.
    const slipped = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      dependencies: [
        makeDependency('d1', 'a', 'b', 'FS', 480),
        makeDependency('d2', 'b', 'c', 'FS'),
      ],
    })
    const comparison = compareBaseline(slipped, resultOf(slipped), baseline)
    expect(varianceOf(comparison, 'a').finishVariance).toBe(0)
    expect(varianceOf(comparison, 'b').startVariance).toBe(480)
    expect(varianceOf(comparison, 'c').startVariance).toBe(480)
    expect(varianceOf(comparison, 'c').finishVariance).toBe(480)
  })

  it('B08: multiple independent baselines on the same task produce independent variance', () => {
    // Baseline B1 captured at the original Monday schedule. Then the task
    // slips to Tuesday; baseline B2 is captured at the slipped schedule.
    // Comparing the (still-slipped) current schedule against B1 yields +480
    // finish variance; against B2 yields zero variance.
    const original = makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] })
    const b1 = captureBaseline(original, resultOf(original), baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const slipped = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ],
    })
    const slippedSchedule = resultOf(slipped)
    const b2 = captureBaseline(slipped, slippedSchedule, baselineIdOf('b2'), {
      capturedAt: iso(TUESDAY),
    })!
    const againstB1 = compareBaseline(slipped, slippedSchedule, b1)
    const againstB2 = compareBaseline(slipped, slippedSchedule, b2)
    expect(varianceOf(againstB1, 'a').finishVariance).toBe(480)
    expect(varianceOf(againstB2, 'a').finishVariance).toBe(0)
    // Both baselines coexist independently in the document.
    const documentWithBoth: ProjectDocument = {
      ...slipped,
      baselines: [b1, b2],
    }
    expect(documentWithBoth.baselines.map((b) => b.id)).toEqual([
      asBaselineId('b1'),
      asBaselineId('b2'),
    ])
  })

  it('B09: captureBaseline is deterministic and never invokes Date.now()', () => {
    // The capturedAt comes from statusDate only; two captures of the same
    // serialized document produce byte-identical baseline bytes.
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] }),
      TUESDAY,
    )
    const serialized = JSON.stringify(document)
    const result = resultOf(parseDocument(serialized))
    const first = captureBaseline(parseDocument(serialized), result, baselineIdOf('b1'))!
    const second = captureBaseline(parseDocument(serialized), result, baselineIdOf('b1'))!
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.capturedAt).toBe(TUESDAY) // statusDate, not wall-clock
  })

  it('B10: captureBaseline selection=tasks isolates variance to the tracked subset', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
    })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
      selection: { kind: 'tasks', taskIds: [taskId('a'), taskId('c')] },
    })!
    // After capture, only a and c are tracked. b is omitted from variance.
    const comparison = compareBaseline(document, result, baseline)
    expect(Object.keys(comparison.variances).sort()).toEqual(['a', 'c'])
  })

  it('B11: captureBaseline selection=leaves skips summary snapshots but still compares leaves', () => {
    const document = makeDocument({
      tasks: [
        summary('s'),
        makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'b', duration: day(240), parentTaskId: taskId('s'), outlineLevel: 2 }),
      ],
    })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
      selection: { kind: 'leaves' },
    })!
    expect(Object.keys(baseline.taskSnapshots).sort()).toEqual(['a', 'b'])
    const comparison = compareBaseline(document, result, baseline)
    expect(Object.keys(comparison.variances).sort()).toEqual(['a', 'b'])
  })

  it('B12: a pruned snapshot (after task deletion) is omitted from the comparison', () => {
    // Simulate the post-DeleteTask state: baseline b1 snapshotted a and b, but
    // a's snapshot was pruned by the engine on deletion. Comparing the current
    // (b-only) document against the pruned baseline yields variance only for b.
    const baseline: Baseline = {
      ...makeBaseline('b1', MONDAY, { b: { duration: 480 } }),
    }
    const document = makeDocument({ tasks: [makeTask({ id: 'b', duration: day(480) })] })
    const comparison = compareBaseline(document, resultOf(document), baseline)
    expect(Object.keys(comparison.variances).sort()).toEqual(['b'])
  })

  it('B13: variance is calendar-aware (a holiday reduces the working-minute span)', () => {
    // Baseline: a finishes Monday 17:00. Current: a is pinned (MFO) to finish
    // Wednesday 17:00 in BOTH calendars. The working-minute finishVariance
    // from Mon 17:00 to Wed 17:00 differs by calendar:
    //  - standard calendar: Tuesday (480) + Wednesday (480) = 960.
    //  - Tuesday-holiday calendar: Tuesday (0) + Wednesday (480) = 480.
    // The holiday removes Tuesday's 480 working minutes from the span.
    const baseline = makeBaseline('b1', MONDAY, {
      a: { start: MONDAY, finish: MONDAY_FINISH, duration: 480 },
    })
    const withHoliday = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'mustFinishOn',
          constraintDate: iso(WEDNESDAY_FINISH),
        }),
      ],
      calendars: [makeCalendar('standard', { exceptions: [holiday('2026-08-04')] })],
    })
    const holidayResult = resultOf(withHoliday)
    expect(scheduleOf(holidayResult, 'a').scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(
      varianceOf(compareBaseline(withHoliday, holidayResult, baseline), 'a').finishVariance,
    ).toBe(480)

    // The same pinned finish on the standard calendar yields 960 because
    // Tuesday is a normal working day.
    const standard = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'mustFinishOn',
          constraintDate: iso(WEDNESDAY_FINISH),
        }),
      ],
    })
    const standardResult = resultOf(standard)
    expect(scheduleOf(standardResult, 'a').scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(
      varianceOf(compareBaseline(standard, standardResult, baseline), 'a').finishVariance,
    ).toBe(960)
  })

  it('B14: repeated and reordered comparison produces byte-identical variance bytes', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(240), constraintType: 'asLateAsPossible' }),
        makeTask({ id: 'm', duration: day(0), milestone: true }),
      ],
      dependencies: [
        makeDependency('d1', 'a', 'b', 'FS', 60),
        makeDependency('d2', 'b', 'm', 'FS'),
      ],
    })
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'), {
      capturedAt: iso(MONDAY),
    })!
    const first = compareBaseline(document, result, baseline)
    const second = compareBaseline(document, result, baseline)
    const third = compareBaseline(document, result, baseline)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(JSON.stringify(second)).toBe(JSON.stringify(third))
    // Reordering the baseline snapshot keys must not change the comparison
    // bytes (the comparator sorts snapshot keys deterministically).
    const reorderedSnapshots: Baseline['taskSnapshots'] = {}
    for (const key of Object.keys(baseline.taskSnapshots).reverse()) {
      reorderedSnapshots[key] = baseline.taskSnapshots[key]
    }
    const reorderedBaseline: Baseline = { ...baseline, taskSnapshots: reorderedSnapshots }
    const reordered = compareBaseline(document, result, reorderedBaseline)
    expect(JSON.stringify(reordered)).toBe(JSON.stringify(first))
  })
})

describe('PROJECT-009 CreateBaseline + captureBaseline integration', () => {
  it('capturing then issuing CreateBaseline stores a snapshot the comparison can read back', () => {
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] }),
      MONDAY,
    )
    const result = resultOf(document)
    const baseline = captureBaseline(document, result, baselineIdOf('b1'))!
    // Issue the command via the engine to store the captured baseline.
    const stored = applyProjectCommand(
      { ...document, tasks: [...document.tasks] },
      { type: 'CreateBaseline', baseline },
    )
    expect(stored.result.accepted).toBe(true)
    expect(stored.document.baselines.length).toBe(1)
    expect(stored.document.baselines[0]).toEqual(baseline)
    // The comparison reads the stored baseline back and matches a direct compare.
    const direct = compareBaseline(document, result, baseline)
    const fromStored = compareBaseline(document, result, stored.document.baselines[0])
    expect(JSON.stringify(fromStored)).toBe(JSON.stringify(direct))
  })
})
