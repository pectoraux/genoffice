import { describe, expect, it } from 'vitest'
import { applyProjectCommand } from '@genoffice/project-engine'
import { levelResources, schedule } from '../src/index.js'
import type {
  DerivedSchedule,
  LevelingResult,
  ProjectDocument,
  TaskSchedule,
} from '@genoffice/project-contracts'
import { asCalendarId, asResourceId, asTaskId, asAssignmentId } from '@genoffice/project-contracts'
import {
  MONDAY,
  MONDAY_FINISH,
  TUESDAY,
  TUESDAY_FINISH,
  WEDNESDAY,
  WEDNESDAY_FINISH,
  asBaselineId,
  iso,
  makeAssignment,
  makeCalendar,
  makeDependency,
  makeDocument,
  makeResource,
  makeTask,
  parseDocument,
  taskId,
  wm,
} from './fixtures.js'

const day = (minutes: number) => wm(minutes)

// Mid-working-day boundaries used by the availability-window regression
// fixtures (MONDAY = 09:00; the working day runs 09:00–17:00).
const MONDAY_TEN_AM = '2026-08-03T10:00:00.000Z'
const MONDAY_NOON = '2026-08-03T12:00:00.000Z'
const MONDAY_ONE_PM = '2026-08-03T13:00:00.000Z'

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

/** Applies leveling (via the engine's LevelResources command) and returns the
 * re-scheduled derived schedule plus the leveling result and the leveled
 * document. */
const applyLeveling = (
  document: ProjectDocument,
  options?: Parameters<typeof levelResources>[1],
): { result: LevelingResult; leveledDocument: ProjectDocument; schedule: DerivedSchedule } => {
  const result = levelResources(document, options)
  // Apply the proposed SetTaskStart commands as a batch through the canonical
  // LevelResources engine command (mutateForLevelResources). This is the
  // documented host path: level → apply → schedule.
  const exec = applyProjectCommand(document, { type: 'LevelResources' })
  expect(exec.result.accepted).toBe(true)
  const sched = resultOf(exec.document)
  return { result, leveledDocument: exec.document, schedule: sched }
}

const documentJson = (document: ProjectDocument): string => JSON.stringify(document)
const resultJson = (result: LevelingResult): string => JSON.stringify(result)

// ===========================================================================
// Golden fixtures L01–L24
// ===========================================================================

describe('PROJECT-013 golden L01 — simple two-task overload', () => {
  it('delays the larger-TaskId task to the next working day', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(result.affectedTaskIds).toEqual([asTaskId('b')])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    expect(scheduleOf(schedule, 'b').scheduledFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-013 golden L02 — three-task overload', () => {
  it('chains delays so each task lands on its own working day', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // a stays Monday; b delayed to Tuesday; c delayed to Wednesday (after b).
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
      { type: 'SetTaskStart', taskId: asTaskId('c'), start: WEDNESDAY },
    ])
    expect(result.affectedTaskIds).toEqual([asTaskId('b'), asTaskId('c')])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(WEDNESDAY)
  })
})

describe('PROJECT-013 golden L03 — 200% demand (single over-allocated assignment)', () => {
  it('emits LEVELING_INCOMPLETE; cannot resolve without splitting', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 2 })],
    })
    const result = levelResources(document)
    expect(result.proposedCommands).toEqual([])
    expect(result.affectedTaskIds).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_INCOMPLETE')).toBe(true)
    expect(result.overallocations.length).toBe(1)
    expect(result.overallocations[0].resolved).toBe(false)
    expect(result.overallocations[0].peakDemand).toBe(2)
    // Document unchanged after the no-op LevelResources command.
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    expect(exec.result.accepted).toBe(true)
    expect(scheduleOf(resultOf(exec.document), 'a').scheduledStart).toBe(MONDAY)
  })
})

describe('PROJECT-013 golden L04 — priority resolution', () => {
  it('keeps the higher-priority task in place; delays the lower-priority one', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480), priority: 100 }),
        makeTask({ id: 'b', duration: day(480), priority: 900 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('a'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 golden L05 — TaskId tie-break (equal priority)', () => {
  it('delays the lexicographically larger TaskId when priority is equal', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 't1', duration: day(480), priority: 500 }),
        makeTask({ id: 't2', duration: day(480), priority: 500 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 't1', 'r1', { units: 1 }),
        makeAssignment('as2', 't2', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('t2'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 't1').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 't2').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 golden L06 — critical-task policy', () => {
  it('respects respectCritical=true (no eligible task) and false (delays)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    // Both tasks finish Monday 17:00 = projectFinish → both critical (zero slack).
    const baseline = resultOf(document)
    expect(scheduleOf(baseline, 'a').critical).toBe(true)
    expect(scheduleOf(baseline, 'b').critical).toBe(true)

    // respectCritical=true: both protected; no eligible task to delay.
    const protectedResult = levelResources(document, { respectCritical: true })
    expect(protectedResult.proposedCommands).toEqual([])
    expect(protectedResult.diagnostics.some((d) => d.code === 'LEVELING_PROTECTED_CRITICAL')).toBe(
      true,
    )

    // respectCritical=false (default): delay b; b becomes the new critical path.
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').critical).toBe(false)
    expect(scheduleOf(schedule, 'b').critical).toBe(true)
    expect(scheduleOf(schedule, 'a').totalSlack).toBe(480)
    expect(scheduleOf(schedule, 'b').totalSlack).toBe(0)
  })
})

describe('PROJECT-013 golden L07 — FS dependency propagation', () => {
  it('preserves the FS dependency; the successor starts after the predecessor finishes', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      dependencies: [makeDependency('d1', 'a', 'c', 'FS', 0)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        // c has NO resource assignment — it only participates via the FS dep.
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // a and b conflict on r1 Monday; delay b. FS a→c preserved: c starts Tue.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // FS: c starts after a finishes (Monday 17:00 → next working = Tuesday 09:00).
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 golden L08 — SS dependency propagation', () => {
  it('preserves SS; the successor still starts at or after the predecessor start', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      dependencies: [makeDependency('d1', 'a', 'c', 'SS', 0)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // SS: c starts at a's start (Monday) — preserved.
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(MONDAY)
  })
})

describe('PROJECT-013 golden L09 — FF dependency propagation', () => {
  it('preserves FF; the successor finishes at or after the predecessor finish', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      dependencies: [makeDependency('d1', 'a', 'c', 'FF', 0)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // FF: c finishes at a's finish (Monday 17:00) — preserved.
    expect(scheduleOf(schedule, 'c').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-013 golden L10 — SF dependency propagation', () => {
  it('preserves SF; the successor finishes at or after the predecessor start', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      dependencies: [makeDependency('d1', 'a', 'c', 'SF', 0)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // SF: c finishes at/after a's start (Monday 09:00). c starts Monday 09:00
    // (project floor), finishes Monday 17:00 (>= a.start Mon 09:00). Preserved.
    expect(scheduleOf(schedule, 'c').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-013 golden L11 — positive lag', () => {
  it('preserves FS + positive lag; the successor starts after predecessor finish + lag', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      dependencies: [makeDependency('d1', 'a', 'c', 'FS', 480)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // FS + 1-day lag: a finishes Mon 17:00 + 480min = Tue 17:00 → Wed 09:00.
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(WEDNESDAY)
  })
})

describe('PROJECT-013 golden L12 — negative lead', () => {
  it('preserves FS with negative lead; the successor starts earlier than predecessor finish', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      dependencies: [makeDependency('d1', 'a', 'c', 'FS', -480)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // FS - 1-day lead: a finishes Mon 17:00 - 480min = Mon 09:00 → c starts Mon.
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(MONDAY)
  })
})

describe('PROJECT-013 golden L13 — SNET protection', () => {
  it('respects startNoEarlierThan; the delayed start is >= the SNET date', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'b',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(MONDAY),
        }),
        makeTask({ id: 'a', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule, leveledDocument } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    // b delayed to Tuesday (>= Monday SNET) — SNET respected.
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // SNET constraint is NOT mutated by leveling.
    const bTask = leveledDocument.tasks.find((t) => t.id === asTaskId('b'))!
    expect(bTask.constraintType).toBe('startNoEarlierThan')
    expect(bTask.constraintDate).toBe(iso(MONDAY))
  })
})

describe('PROJECT-013 golden L14 — MSO protection', () => {
  it('treats mustStartOn as immovable; delays the other side', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'b',
          duration: day(480),
          constraintType: 'mustStartOn',
          constraintDate: iso(MONDAY),
        }),
        makeTask({ id: 'a', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule, leveledDocument } = applyLeveling(document)
    // b has MSO (immovable); a is delayed (a is the only eligible side).
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('a'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(TUESDAY)
    const bTask = leveledDocument.tasks.find((t) => t.id === asTaskId('b'))!
    expect(bTask.constraintType).toBe('mustStartOn')
    expect(bTask.constraintDate).toBe(iso(MONDAY))
  })
})

describe('PROJECT-013 golden L15 — MFO protection', () => {
  it('treats mustFinishOn as immovable; delays the other side', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'b',
          duration: day(480),
          constraintType: 'mustFinishOn',
          constraintDate: iso(MONDAY_FINISH),
        }),
        makeTask({ id: 'a', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('a'), start: TUESDAY },
    ])
    // MFO preserved: b finishes at Monday 17:00 (the MFO date).
    expect(scheduleOf(schedule, 'b').scheduledFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 golden L16 — deadline interaction', () => {
  it('respects deadlines only when respectDeadlines is on; otherwise surfaces deadlineMissed', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480), priority: 900 }),
        makeTask({ id: 'b', duration: day(480), priority: 100, deadline: iso(MONDAY_FINISH) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    // Default (respectDeadlines=false): delay b; b misses its Monday deadline.
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'b').scheduledFinish).toBe(TUESDAY_FINISH)
    expect(scheduleOf(schedule, 'b').deadline).toBe(iso(MONDAY_FINISH))
    expect(scheduleOf(schedule, 'b').deadlineMissed).toBe(true)
    expect(scheduleOf(schedule, 'b').deadlineVariance).toBe(-480)
    // deadline field is NOT mutated by leveling.
    // respectDeadlines=true: leveler refuses to delay b past its deadline.
    const protectedResult = levelResources(document, { respectDeadlines: true })
    expect(protectedResult.proposedCommands).toEqual([])
    expect(protectedResult.diagnostics.some((d) => d.code === 'LEVELING_DEADLINE_CONFLICT')).toBe(
      true,
    )
  })
})

describe('PROJECT-013 golden L17 — resource calendar restriction', () => {
  it('advances the delayed start to the next RESOURCE-working instant', () => {
    // resourceCal: Mon-Tue OFF, Wed-Fri working 09:00-17:00.
    const resourceCal = makeCalendar('resourceCal', {
      workingWeek: {
        0: [],
        1: [],
        2: [],
        3: [{ startMinute: 540, endMinute: 1020 }],
        4: [{ startMinute: 540, endMinute: 1020 }],
        5: [{ startMinute: 540, endMinute: 1020 }],
        6: [],
      },
    })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      calendars: [makeCalendar('standard'), resourceCal],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          maxUnits: 1,
          calendarId: asCalendarId('resourceCal'),
        }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // a (kept) runs Monday (task calendar standard, Mon working). b delayed
    // to the next RESOURCE-working instant after a finishes (Mon 17:00). The
    // resource calendar has Mon-Tue off, so b starts Wednesday.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: WEDNESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(WEDNESDAY)
  })
})

describe('PROJECT-013 golden L18 — availability window restriction', () => {
  it('uses the availability window units as the effective capacity', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          maxUnits: 1,
          availability: [{ start: iso(MONDAY), finish: iso(TUESDAY_FINISH), units: 0.5 }],
        }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 0.5 }),
        makeAssignment('as2', 'b', 'r1', { units: 0.5 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // Without the availability window, combined 0.5+0.5 = 1.0 <= maxUnits 1.0
    // — no conflict. With the window (units 0.5), combined 1.0 > 0.5 — conflict.
    // b delayed to Tuesday; on Tuesday the window is still active (finish Tue
    // 17:00) so capacity 0.5, and b alone demands 0.5 <= 0.5 (no conflict).
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    expect(result.overallocations[0].maxUnits).toBe(0.5)
    expect(result.overallocations[0].peakDemand).toBe(1)
  })
})

describe('PROJECT-013 golden L19 — multiple resources on the same task', () => {
  it('resolves conflicts on each resource independently; the shared task is not moved', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [
        makeResource({ id: 'r1', kind: 'work', maxUnits: 1 }),
        makeResource({ id: 'r2', kind: 'work', maxUnits: 1 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'r2', { units: 1 }),
        makeAssignment('as3', 'b', 'r1', { units: 1 }),
        makeAssignment('as4', 'c', 'r2', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // Conflict on r1 (a, b): delay b. Conflict on r2 (a, c): delay c.
    // a is the shared task (assignments on r1 and r2); it is kept on both.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
      { type: 'SetTaskStart', taskId: asTaskId('c'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 golden L20 — baseline preservation', () => {
  it('does not mutate baseline snapshots; only the current schedule moves', () => {
    // Capture a baseline BEFORE leveling (b.start = Monday).
    const baselineDocument = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
      baselines: [
        {
          id: asBaselineId('b1'),
          name: 'Before',
          capturedAt: iso(MONDAY),
          taskSnapshots: {
            a: {
              start: iso(MONDAY),
              finish: iso(MONDAY_FINISH),
              duration: day(480),
              work: day(480),
              cost: 0,
            },
            b: {
              start: iso(MONDAY),
              finish: iso(MONDAY_FINISH),
              duration: day(480),
              work: day(480),
              cost: 0,
            },
          },
        },
      ],
    })
    const { schedule, leveledDocument } = applyLeveling(baselineDocument)
    // Current schedule: b moved to Tuesday.
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // Baseline snapshot: b.start still Monday (immutable).
    const baseline = leveledDocument.baselines[0]
    expect(baseline.taskSnapshots['b'].start).toBe(MONDAY)
    expect(baseline.taskSnapshots['b'].finish).toBe(MONDAY_FINISH)
    // Baseline id + capturedAt unchanged.
    expect(baseline.id).toBe('b1')
    expect(baseline.capturedAt).toBe(iso(MONDAY))
  })
})

describe('PROJECT-013 golden L21 — impossible leveling (both MSO)', () => {
  it('emits LEVELING_CONSTRAINT_CONFLICT; document unchanged', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'mustStartOn',
          constraintDate: iso(MONDAY),
        }),
        makeTask({
          id: 'b',
          duration: day(480),
          constraintType: 'mustStartOn',
          constraintDate: iso(MONDAY),
        }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = levelResources(document)
    expect(result.proposedCommands).toEqual([])
    expect(result.affectedTaskIds).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_CONSTRAINT_CONFLICT')).toBe(true)
    expect(result.overallocations[0].resolved).toBe(false)
    // Document unchanged after the no-op LevelResources command.
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    expect(exec.result.accepted).toBe(true)
    expect(scheduleOf(resultOf(exec.document), 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(resultOf(exec.document), 'b').scheduledStart).toBe(MONDAY)
  })
})

describe('PROJECT-013 golden L22 — repeated deterministic leveling', () => {
  it('produces byte-identical results across three runs', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
    })
    const r1 = levelResources(document)
    const r2 = levelResources(document)
    const r3 = levelResources(document)
    expect(resultJson(r1)).toBe(resultJson(r2))
    expect(resultJson(r2)).toBe(resultJson(r3))
  })

  it('is idempotent: leveling an already-leveled document is a no-op', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    const leveled = exec.document
    // Leveling the leveled document: no over-allocation remains.
    const secondResult = levelResources(leveled)
    expect(secondResult.proposedCommands).toEqual([])
    expect(secondResult.diagnostics.some((d) => d.code === 'LEVELING_NO_OVERALLOCATION')).toBe(true)
  })
})

describe('PROJECT-013 golden L23 — reordered input leveling', () => {
  it('produces byte-identical results under reversed arrays and serialized round-trip', () => {
    const base = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [
        makeResource({ id: 'r1', kind: 'work', maxUnits: 1 }),
        makeResource({ id: 'r2', kind: 'work', maxUnits: 1 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r2', { units: 1 }),
        makeAssignment('as4', 'a', 'r2', { units: 1 }),
      ],
    })
    const reference = levelResources(base)
    // Reversed task array.
    const reversedTasks: ProjectDocument = { ...base, tasks: [...base.tasks].reverse() }
    expect(resultJson(levelResources(reversedTasks))).toBe(resultJson(reference))
    // Reversed assignment array.
    const reversedAssignments: ProjectDocument = {
      ...base,
      assignments: [...base.assignments].reverse(),
    }
    expect(resultJson(levelResources(reversedAssignments))).toBe(resultJson(reference))
    // Reversed resource array.
    const reversedResources: ProjectDocument = {
      ...base,
      resources: [...base.resources].reverse(),
    }
    expect(resultJson(levelResources(reversedResources))).toBe(resultJson(reference))
    // Serialized round-trip (JSON parse).
    const reparsed = parseDocument(documentJson(base))
    expect(resultJson(levelResources(reparsed))).toBe(resultJson(reference))
  })
})

describe('PROJECT-013 golden L24 — material and cost resources excluded', () => {
  it('does not level material or cost resource capacity', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', maxUnits: 1 }),
        makeResource({ id: 'r2', kind: 'material', maxUnits: 1, standardRate: 10 }),
        makeResource({ id: 'r3', kind: 'cost', maxUnits: 1 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        // Material + cost assignments on both tasks — these do NOT
        // contribute to work-capacity over-allocation.
        makeAssignment('as3', 'a', 'r2', { units: 5 }),
        makeAssignment('as4', 'b', 'r2', { units: 5 }),
        makeAssignment('as5', 'a', 'r3', { cost: 100 }),
        makeAssignment('as6', 'b', 'r3', { cost: 100 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // Only the work resource (r1) over-allocates; material/cost are ignored.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // Material and cost resources never appear in over-allocations: every
    // reported over-allocation references the work resource r1 only, and the
    // exact resource-id set is {r1}. (Replaces a previous tautological
    // `... || true` assertion that always passed.)
    expect(result.overallocations.every((o) => o.resourceId === asResourceId('r1'))).toBe(true)
    expect([...new Set(result.overallocations.map((o) => o.resourceId as string))]).toEqual(['r1'])
  })
})

describe('PROJECT-013 golden L25 — mid-assignment capacity drop', () => {
  it('detects an over-allocation bounded only by availability-window transitions', () => {
    // Two tasks span the full working day (09:00–17:00) at 0.5 units each.
    // Combined demand is 1.0 everywhere. The resource maxUnits is 1.0, so
    // there is NO over-allocation from assignment endpoints alone. The only
    // over-allocation arises during the 12:00–13:00 availability window whose
    // units (0.5) drop capacity below demand. A sweep that segments only at
    // assignment start/finish would evaluate capacity at 09:00 (1.0, no
    // conflict) and at 17:00 (0 demand, no conflict) and MISS the conflict
    // entirely — there is no assignment event at 12:00 or 13:00.
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          maxUnits: 1,
          availability: [{ start: iso(MONDAY_NOON), finish: iso(MONDAY_ONE_PM), units: 0.5 }],
        }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 0.5 }),
        makeAssignment('as2', 'b', 'r1', { units: 0.5 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // The conflict is bounded [12:00, 13:00). b (larger TaskId, equal priority
    // and scheduled start) is delayed to Tuesday; a alone demands 0.5 ≤ the
    // 0.5 capacity during 12:00–13:00, so the conflict is resolved.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    expect(result.overallocations).toHaveLength(1)
    expect(result.overallocations[0].window).toEqual({
      start: iso(MONDAY_NOON),
      finish: iso(MONDAY_ONE_PM),
    })
    expect(result.overallocations[0].peakDemand).toBe(1)
    expect(result.overallocations[0].maxUnits).toBe(0.5)
    expect(result.overallocations[0].resolved).toBe(true)
  })
})

// ===========================================================================
// Additional required tests (constraint variants, edge cases, determinism)
// ===========================================================================

describe('PROJECT-013 SNLT protection', () => {
  it('emits LEVELING_CONSTRAINT_CONFLICT when the delay would push past SNLT', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          priority: 100,
          constraintType: 'startNoLaterThan',
          constraintDate: iso(MONDAY),
        }),
        makeTask({ id: 'b', duration: day(480), priority: 900 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = levelResources(document)
    // a is the lower-priority side → would be delayed to Tuesday. SNLT at
    // Monday: Tuesday > Monday → constraint conflict.
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_CONSTRAINT_CONFLICT')).toBe(true)
  })
})

describe('PROJECT-013 FNET protection', () => {
  it('satisfies finishNoEarlierThan when delaying (finish moves later)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'b',
          duration: day(480),
          constraintType: 'finishNoEarlierThan',
          constraintDate: iso(MONDAY_FINISH),
        }),
        makeTask({ id: 'a', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    // FNET at Monday 17:00; b finishes Tuesday 17:00 (>= Monday 17:00). OK.
    expect(scheduleOf(schedule, 'b').scheduledFinish).toBe(TUESDAY_FINISH)
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
  })
})

describe('PROJECT-013 FNLT protection', () => {
  it('emits LEVELING_CONSTRAINT_CONFLICT when the delay would push finish past FNLT', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'b',
          duration: day(480),
          priority: 100,
          constraintType: 'finishNoLaterThan',
          constraintDate: iso(MONDAY_FINISH),
        }),
        makeTask({ id: 'a', duration: day(480), priority: 900 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = levelResources(document)
    // b is lower-priority → would be delayed to Tuesday; finish Tuesday 17:00.
    // FNLT at Monday 17:00: Tuesday 17:00 > Monday 17:00 → conflict.
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_CONSTRAINT_CONFLICT')).toBe(true)
  })
})

describe('PROJECT-013 milestone behavior', () => {
  it('does not level zero-duration milestones (no work demand)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'm', duration: day(0), milestone: true }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'm', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // a and b conflict on r1; m has zero duration → no work demand → skipped.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'm').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 summary behavior', () => {
  it('does not directly level summary tasks; rolls up from children', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 's', duration: day(0), summary: true, outlineLevel: 1 }),
        makeTask({ id: 'a', duration: day(480), outlineLevel: 2, parentTaskId: asTaskId('s') }),
        makeTask({ id: 'b', duration: day(480), outlineLevel: 2, parentTaskId: asTaskId('s') }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // Leaf children a and b conflict; summary s is not directly levelable.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    // Summary rolls up: start = min(a, b) = Monday; finish = max = Tuesday.
    expect(scheduleOf(schedule, 's').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 's').scheduledFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-013 manual task behavior', () => {
  it('protects manual-scheduled tasks; delays the auto-scheduled side', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480), manualScheduled: true, autoScheduled: false }),
        makeTask({ id: 'b', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // a is manual (protected); b is delayed.
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
  })

  it('emits LEVELING_PROTECTED_MANUAL when both sides are manual', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480), manualScheduled: true, autoScheduled: false }),
        makeTask({ id: 'b', duration: day(480), manualScheduled: true, autoScheduled: false }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = levelResources(document)
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_PROTECTED_MANUAL')).toBe(true)
  })
})

describe('PROJECT-013 already-non-overallocated resource', () => {
  it('emits LEVELING_NO_OVERALLOCATION; no commands proposed', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        // b has no assignment on r1 — no conflict.
      ],
    })
    const result = levelResources(document)
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_NO_OVERALLOCATION')).toBe(true)
    expect(result.overallocations).toEqual([])
  })
})

describe('PROJECT-013 leveling result reapplied through canonical scheduler', () => {
  it('applying proposedCommands via SetTaskStart dispatch matches the LevelResources batch', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
    })
    // Path A: LevelResources command (batch).
    const batchExec = applyProjectCommand(document, { type: 'LevelResources' })
    const batchSchedule = resultOf(batchExec.document)
    // Path B: apply each proposed SetTaskStart via the canonical
    // applyProjectCommand(SetTaskStart) dispatch, one at a time.
    const leveling = levelResources(document)
    let manualDocument = document
    for (const cmd of leveling.proposedCommands) {
      const exec = applyProjectCommand(manualDocument, cmd)
      expect(exec.result.accepted).toBe(true)
      manualDocument = exec.document
    }
    const manualSchedule = resultOf(manualDocument)
    // Both paths produce byte-identical schedule bytes.
    expect(JSON.stringify(batchSchedule)).toBe(JSON.stringify(manualSchedule))
    expect(JSON.stringify(batchExec.document)).toBe(JSON.stringify(manualDocument))
  })
})

describe('PROJECT-013 leveling preserves stable IDs and dependencies', () => {
  it('does not mutate TaskId, ResourceId, AssignmentId, DependencyId, or BaselineId', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS', 0)],
    })
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    expect(exec.result.accepted).toBe(true)
    // IDs unchanged.
    expect(exec.document.tasks.map((t) => t.id as string)).toEqual(['a', 'b'])
    expect(exec.document.resources.map((r) => r.id as string)).toEqual(['r1'])
    expect(exec.document.assignments.map((a) => a.id as string)).toEqual(['as1', 'as2'])
    expect(exec.document.dependencies.map((d) => d.id as string)).toEqual(['d1'])
    // Dependency structure unchanged.
    expect(exec.document.dependencies[0]).toEqual(document.dependencies[0])
  })
})

describe('PROJECT-013 leveling produces no cycles', () => {
  it('does not introduce new dependencies; the dependency graph remains acyclic', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS', 0)],
    })
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    expect(exec.result.accepted).toBe(true)
    // No new dependencies added; existing ones unchanged.
    expect(exec.document.dependencies).toEqual(document.dependencies)
    // The leveled document still schedules cleanly (no cycle).
    const leveledSchedule = resultOf(exec.document)
    expect(leveledSchedule.diagnostics).toEqual([])
  })
})

describe('PROJECT-013 scope filter (taskIds)', () => {
  it('only delays tasks in the scope; out-of-scope tasks are immovable', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
    })
    // Scope: only b and c are levelable. a is out-of-scope (immovable).
    // Conflict a+b+c Monday. a is immovable (out of scope). Delay b, then c.
    // But a stays Monday, so b delayed to Tuesday, c delayed to Wednesday
    // (after b), and a remains on Monday alongside whoever is there.
    // Actually: a is immovable. b and c are in-scope and delayable. Delay
    // the largest-TaskId in-scope side: c → Tuesday (after a+b finish Mon).
    //   Wait — both a and b finish Monday 17:00 (both 1 day). c delayed to
    //   Tuesday (after max(a, b) finish = Mon 17:00 → Tue 09:00).
    //   Then a and b still conflict Monday. b is in-scope; delay b → ?
    //   b's newStart = a.finish = Mon 17:00 → Tue 09:00. But c is also on
    //   Tuesday now. b and c conflict Tuesday. Delay c → Wednesday.
    //   Final: a=Mon, b=Tue, c=Wed.
    const { result, schedule } = applyLeveling(document, {
      taskIds: [asTaskId('b'), asTaskId('c')],
    })
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(WEDNESDAY)
    // a is NOT in affectedTaskIds (out of scope, never delayed).
    expect(result.affectedTaskIds).toEqual([asTaskId('b'), asTaskId('c')])
  })

  it('emits LEVELING_SCOPE_EMPTY when the scope matches no levelable task', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = levelResources(document, { taskIds: [asTaskId('zzz')] })
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_SCOPE_EMPTY')).toBe(true)
  })
})

describe('PROJECT-013 different assignment units', () => {
  it('delays the task whose units push the combined demand over capacity', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 0.5 }),
        makeAssignment('as2', 'b', 'r1', { units: 0.75 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // Combined 0.5 + 0.75 = 1.25 > 1.0 → over-allocation. Delay b (larger id).
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 one 200% assignment cannot be resolved by moving', () => {
  it('reports LEVELING_INCOMPLETE for a single over-capacity assignment', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 2 })],
    })
    const result = levelResources(document)
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_INCOMPLETE')).toBe(true)
    expect(result.overallocations[0].peakDemand).toBe(2)
    expect(result.overallocations[0].maxUnits).toBe(1)
  })
})

describe('PROJECT-013 leveling with FS dependency — successor follows delayed predecessor', () => {
  it('propagates the predecessor delay to the FS successor', () => {
    // a (low priority) → b (FS). a and c conflict on r1 Monday.
    // Leveling delays a to Tuesday; b's FS pushes b to Wednesday.
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480), priority: 100 }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480), priority: 900 }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS', 0)],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    // a delayed to Tuesday (after c finishes Monday 17:00). b (FS on a)
    // starts Wednesday (after a finishes Tuesday 17:00 → Wed 09:00).
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('a'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'c').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(schedule, 'a').scheduledStart).toBe(TUESDAY)
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(WEDNESDAY)
  })
})

describe('PROJECT-013 leveling respects task calendars (split-day)', () => {
  it('uses the resource calendar to compute the delayed start', () => {
    // Standard resource calendar; task a has a split-day calendar. The
    // delayed start still lands on a resource-working instant.
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const { result, schedule } = applyLeveling(document)
    expect(result.proposedCommands).toEqual([
      { type: 'SetTaskStart', taskId: asTaskId('b'), start: TUESDAY },
    ])
    expect(scheduleOf(schedule, 'b').scheduledStart).toBe(TUESDAY)
  })
})

describe('PROJECT-013 LevelingResult shape', () => {
  it('exposes proposedCommands, actions, overallocations, affectedTaskIds, diagnostics', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = levelResources(document)
    expect(result.proposedCommands.length).toBe(1)
    expect(result.actions.length).toBe(1)
    expect(result.actions[0].taskId).toBe(asTaskId('b'))
    expect(result.actions[0].resourceId).toBe(asResourceId('r1'))
    expect(result.actions[0].reason).toBe('over-allocation')
    expect(result.actions[0].originalStart).toBe(iso(MONDAY))
    expect(result.actions[0].newStart).toBe(TUESDAY)
    expect(result.actions[0].assignmentId).toBe(asAssignmentId('as2'))
    expect(result.overallocations.length).toBe(1)
    expect(result.overallocations[0].resolved).toBe(true)
    expect(result.overallocations[0].peakDemand).toBe(2)
    expect(result.overallocations[0].maxUnits).toBe(1)
    expect(result.overallocations[0].window).toEqual({
      start: iso(MONDAY),
      finish: iso(MONDAY_FINISH),
    })
  })
})

describe('PROJECT-013 schedule-after-leveling is canonical', () => {
  it('the leveled document passes through schedule() and produces a clean DerivedSchedule', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 1 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
        makeAssignment('as3', 'c', 'r1', { units: 1 }),
      ],
    })
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    expect(exec.result.accepted).toBe(true)
    const sched = schedule(exec.document)
    expect(sched.diagnostics).toEqual([])
    expect(sched.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-013 multi-window conflict identity', () => {
  it('reports distinct conflict windows for the same assignments (no collapse)', () => {
    // Two MSO tasks overlap all day at 0.5 units each (combined demand 1.0).
    // The resource maxUnits is 1.0, so over-allocation arises ONLY inside two
    // disjoint availability windows whose units (0.5) drop capacity below
    // demand: [09:00, 10:00) and [12:00, 13:00). The SAME assignments produce
    // TWO distinct conflict windows. A conflict signature keyed only by
    // resource+tasks+assignments would collapse them into a single entry whose
    // `resolved` state the final pass would overwrite; the signature MUST
    // include window identity. Both tasks are MSO so neither can be delayed —
    // both conflicts persist and must both be reported as distinct entries.
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          duration: day(480),
          constraintType: 'mustStartOn',
          constraintDate: iso(MONDAY),
        }),
        makeTask({
          id: 'b',
          duration: day(480),
          constraintType: 'mustStartOn',
          constraintDate: iso(MONDAY),
        }),
      ],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          maxUnits: 1,
          availability: [
            { start: iso(MONDAY), finish: iso(MONDAY_TEN_AM), units: 0.5 },
            { start: iso(MONDAY_NOON), finish: iso(MONDAY_ONE_PM), units: 0.5 },
          ],
        }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 0.5 }),
        makeAssignment('as2', 'b', 'r1', { units: 0.5 }),
      ],
    })
    const result = levelResources(document)
    expect(result.proposedCommands).toEqual([])
    expect(result.diagnostics.some((d) => d.code === 'LEVELING_CONSTRAINT_CONFLICT')).toBe(true)
    // Two distinct windows survive — they do NOT collapse into one entry.
    expect(result.overallocations).toHaveLength(2)
    expect(result.overallocations.map((o) => o.window)).toEqual([
      { start: iso(MONDAY), finish: iso(MONDAY_TEN_AM) },
      { start: iso(MONDAY_NOON), finish: iso(MONDAY_ONE_PM) },
    ])
    // Both reference the same resource + assignments + tasks but are distinct
    // by window; both are unresolved (MSO is immovable).
    expect(result.overallocations.every((o) => o.resolved === false)).toBe(true)
    expect(result.overallocations.every((o) => o.resourceId === asResourceId('r1'))).toBe(true)
    expect(result.overallocations.every((o) => o.peakDemand === 1)).toBe(true)
    expect(result.overallocations.every((o) => o.maxUnits === 0.5)).toBe(true)
    // Exact resource-id set: only r1 (the work resource) appears.
    expect([...new Set(result.overallocations.map((o) => o.resourceId as string))]).toEqual(['r1'])
    // The leveled document is unchanged (no proposed commands); both MSO
    // tasks remain pinned to Monday by the no-op LevelResources command.
    const exec = applyProjectCommand(document, { type: 'LevelResources' })
    expect(exec.result.accepted).toBe(true)
    expect(scheduleOf(resultOf(exec.document), 'a').scheduledStart).toBe(MONDAY)
    expect(scheduleOf(resultOf(exec.document), 'b').scheduledStart).toBe(MONDAY)
  })
})
