import { describe, expect, it } from 'vitest'
import { resolveCalendar, schedule, resolveResourceCalendarId } from '../src/index.js'
import { applyProjectCommand } from '@genoffice/project-engine'
import type {
  AssignmentSchedule,
  DerivedSchedule,
  ProjectDocument,
  TaskSchedule,
} from '@genoffice/project-contracts'
import { asAssignmentId, asCalendarId, asResourceId, asTaskId } from '@genoffice/project-contracts'
import {
  MONDAY,
  MONDAY_FINISH,
  TUESDAY,
  TUESDAY_FINISH,
  iso,
  makeAssignment,
  makeCalendar,
  makeDocument,
  makeResource,
  makeTask,
  parseDocument,
  taskId,
  wm,
} from './fixtures.js'

const day = (minutes: number) => wm(minutes)

// PROJECT-011 calendar-precedence fixtures use an 08:00 project start so the
// task calendar's 08:00-16:00 opening is observable (the default MONDAY start
// is 09:00, which is already inside the task calendar window and would mask
// the precedence signal).
const MONDAY_08 = '2026-08-03T08:00:00.000Z'
const MONDAY_08_FINISH = '2026-08-03T16:00:00.000Z'

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

const assignmentScheduleOf = (result: DerivedSchedule, id: string): AssignmentSchedule => {
  const entry = result.assignmentSchedules?.[asAssignmentId(id)]
  if (!entry) throw new Error(`missing assignment schedule for ${id}`)
  return entry
}

const documentJson = (document: ProjectDocument): string => JSON.stringify(document)
const scheduleJson = (result: DerivedSchedule): string => JSON.stringify(result)

// ===========================================================================
// Golden fixtures W01-W20
// ===========================================================================

describe('PROJECT-011 golden W01 — one work resource', () => {
  it('derives assignment work = duration × units and standard-rate cost', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          name: 'Engineer',
          kind: 'work',
          maxUnits: 1,
          standardRate: 50,
          overtimeRate: 75,
          costPerUse: 0,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const task = scheduleOf(result, 'a')
    expect(task.scheduledStart).toBe(MONDAY)
    expect(task.scheduledFinish).toBe(MONDAY_FINISH)
    expect(task.work).toEqual(wm(480))
    expect(task.actualWork).toEqual(wm(0))
    expect(task.remainingWork).toEqual(wm(480))
    expect(task.cost).toBe(400)
    expect(task.actualCost).toBe(0)
    expect(task.remainingCost).toBe(400)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.assignmentId).toEqual(asAssignmentId('as1'))
    expect(assignment.taskId).toEqual(asTaskId('a'))
    expect(assignment.resourceId).toEqual(asResourceId('r1'))
    expect(assignment.resourceType).toBe('work')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('standard'))
    expect(assignment.maxUnits).toBe(1)
    expect(assignment.units).toBe(1)
    expect(assignment.work).toEqual(wm(480))
    expect(assignment.actualWork).toEqual(wm(0))
    expect(assignment.remainingWork).toEqual(wm(480))
    expect(assignment.cost).toBe(400)
    expect(assignment.actualCost).toBe(0)
    expect(assignment.remainingCost).toBe(400)
    expect(assignment.standardRate).toBe(50)
    expect(assignment.overtimeRate).toBe(75)
    expect(assignment.costPerUse).toBe(0)
  })
})

describe('PROJECT-011 golden W02 — half-unit assignment', () => {
  it('derives work = duration × 0.5 for a 50% assignment', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 0.5 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.work).toEqual(wm(240))
    expect(assignment.actualWork).toEqual(wm(0))
    expect(assignment.remainingWork).toEqual(wm(240))
    expect(assignment.cost).toBe(200)
    const task = scheduleOf(result, 'a')
    expect(task.work).toEqual(wm(240))
    expect(task.cost).toBe(200)
  })
})

describe('PROJECT-011 golden W03 — two resources', () => {
  it('aggregates work and cost from two work-resource assignments', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'r2', kind: 'work', standardRate: 60 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'r2', { units: 1 }),
      ],
    })
    const result = resultOf(document)
    const as1 = assignmentScheduleOf(result, 'as1')
    const as2 = assignmentScheduleOf(result, 'as2')
    expect(as1.work).toEqual(wm(480))
    expect(as1.cost).toBe(400)
    expect(as2.work).toEqual(wm(480))
    expect(as2.cost).toBe(480)
    const task = scheduleOf(result, 'a')
    expect(task.work).toEqual(wm(960))
    expect(task.cost).toBe(880)
  })
})

describe('PROJECT-011 golden W04 — unequal units', () => {
  it('derives work proportionally from unequal assignment units', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'r2', kind: 'work', standardRate: 50 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'r2', { units: 0.5 }),
      ],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as2').work).toEqual(wm(240))
    const task = scheduleOf(result, 'a')
    expect(task.work).toEqual(wm(720))
    expect(task.cost).toBe(600)
  })
})

describe('PROJECT-011 golden W05 — shared resource across tasks', () => {
  it('the same resource on two tasks produces independent assignment schedules', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'b', 'r1')],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as2').work).toEqual(wm(480))
    expect(scheduleOf(result, 'a').work).toEqual(wm(480))
    expect(scheduleOf(result, 'b').work).toEqual(wm(480))
    expect(scheduleOf(result, 'a').cost).toBe(400)
    expect(scheduleOf(result, 'b').cost).toBe(400)
  })
})

describe('PROJECT-011 golden W06 — standard-rate cost', () => {
  it('cost = (work/60) × standardRate with no overtime or cost-per-use', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          standardRate: 50,
          overtimeRate: 75,
          costPerUse: 0,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.standardRate).toBe(50)
    expect(assignment.cost).toBe(400)
    expect(assignment.actualCost).toBe(0)
    expect(assignment.remainingCost).toBe(400)
  })
})

describe('PROJECT-011 golden W07 — overtime-rate cost (deferred)', () => {
  it('overtimeRate is echoed but overtime cost is 0 (no overtimeWork input in the frozen contract)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          standardRate: 50,
          overtimeRate: 75,
          costPerUse: 0,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    // The resource carries an overtime rate, but the frozen Assignment contract
    // has no overtimeWork input. Overtime cost is therefore structurally 0 in
    // PROJECT-011 (a documented deferred limitation, not a guess). The full
    // cost is the standard-rate cost only.
    expect(assignment.overtimeRate).toBe(75)
    expect(assignment.cost).toBe(400)
    expect(assignment.standardRate).toBe(50)
    // Task scheduling is unaffected by rates.
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-011 golden W08 — cost-per-use', () => {
  it('cost = standardRateCost + costPerUse (flat per-assignment)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50, costPerUse: 100 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.costPerUse).toBe(100)
    expect(assignment.cost).toBe(500)
    expect(assignment.actualCost).toBe(0)
    expect(assignment.remainingCost).toBe(500)
  })
})

describe('PROJECT-011 golden W09 — material cost', () => {
  it('material: no work; cost = units × standardRate + costPerUse', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'm1',
          name: 'Concrete',
          kind: 'material',
          standardRate: 100,
          costPerUse: 0,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'm1', { units: 10 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resourceType).toBe('material')
    expect(assignment.work).toEqual(wm(0))
    expect(assignment.actualWork).toEqual(wm(0))
    expect(assignment.remainingWork).toEqual(wm(0))
    expect(assignment.cost).toBe(1000)
    expect(assignment.actualCost).toBe(0)
    expect(assignment.remainingCost).toBe(1000)
    expect(scheduleOf(result, 'a').work).toEqual(wm(0))
    expect(scheduleOf(result, 'a').cost).toBe(1000)
  })
})

describe('PROJECT-011 golden W10 — cost resource', () => {
  it('cost resource: assignment.cost is the authoritative input; no work', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'c1', name: 'Travel', kind: 'cost', costPerUse: 0 })],
      assignments: [makeAssignment('as1', 'a', 'c1', { units: 1, cost: 500 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resourceType).toBe('cost')
    expect(assignment.work).toEqual(wm(0))
    expect(assignment.actualWork).toEqual(wm(0))
    expect(assignment.remainingWork).toEqual(wm(0))
    expect(assignment.cost).toBe(500)
    expect(assignment.actualCost).toBe(0)
    expect(assignment.remainingCost).toBe(500)
    expect(scheduleOf(result, 'a').work).toEqual(wm(0))
    expect(scheduleOf(result, 'a').cost).toBe(500)
  })
})

describe('PROJECT-011 golden W11 — actual/remaining work', () => {
  it('50% progress derives actualWork = work/2 and remainingWork = work/2', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 50 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.work).toEqual(wm(480))
    expect(assignment.actualWork).toEqual(wm(240))
    expect(assignment.remainingWork).toEqual(wm(240))
    expect(assignment.cost).toBe(400)
    expect(assignment.actualCost).toBe(200)
    expect(assignment.remainingCost).toBe(200)
    const task = scheduleOf(result, 'a')
    expect(task.work).toEqual(wm(480))
    expect(task.actualWork).toEqual(wm(240))
    expect(task.remainingWork).toEqual(wm(240))
    expect(task.cost).toBe(400)
    expect(task.actualCost).toBe(200)
    expect(task.remainingCost).toBe(200)
  })
})

describe('PROJECT-011 golden W12 — progress + work', () => {
  it('25% progress derives proportional actual/remaining work and cost', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 25 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50, costPerUse: 100 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.work).toEqual(wm(480))
    expect(assignment.actualWork).toEqual(wm(120))
    expect(assignment.remainingWork).toEqual(wm(360))
    expect(assignment.cost).toBe(500)
    expect(assignment.actualCost).toBe(125)
    expect(assignment.remainingCost).toBe(375)
  })
})

describe('PROJECT-011 golden W13 — resource calendar', () => {
  it('work uses task duration (task calendar); resource calendar is an input only', () => {
    const resourceCalendar = makeCalendar('rescal', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 480, endMinute: 960 }],
        2: [{ startMinute: 480, endMinute: 960 }],
        3: [{ startMinute: 480, endMinute: 960 }],
        4: [{ startMinute: 480, endMinute: 960 }],
        5: [{ startMinute: 480, endMinute: 960 }],
        6: [],
      },
    })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      calendars: [makeCalendar('standard'), resourceCalendar],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50, calendarId: resourceCalendar.id }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    // Task scheduling uses the TASK calendar (default = standard, 09:00-17:00).
    const task = scheduleOf(result, 'a')
    expect(task.resolvedCalendarId).toEqual(asCalendarId('standard'))
    expect(task.scheduledStart).toBe(MONDAY)
    expect(task.scheduledFinish).toBe(MONDAY_FINISH)
    // Work uses the task's scheduled duration × units (resource calendar does
    // not change the work formula in PROJECT-011).
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('rescal'))
    expect(assignment.work).toEqual(wm(480))
    expect(assignment.cost).toBe(400)
  })
})

describe('PROJECT-011 golden W14 — task/resource calendar distinction', () => {
  it('task calendar governs task scheduling; resource calendar is independent', () => {
    const taskcal = makeCalendar('taskcal', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 480, endMinute: 960 }],
        2: [{ startMinute: 480, endMinute: 960 }],
        3: [{ startMinute: 480, endMinute: 960 }],
        4: [{ startMinute: 480, endMinute: 960 }],
        5: [{ startMinute: 480, endMinute: 960 }],
        6: [],
      },
    })
    const rescal = makeCalendar('rescal', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 600, endMinute: 1200 }],
        2: [{ startMinute: 600, endMinute: 1200 }],
        3: [{ startMinute: 600, endMinute: 1200 }],
        4: [{ startMinute: 600, endMinute: 1200 }],
        5: [{ startMinute: 600, endMinute: 1200 }],
        6: [],
      },
    })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), calendarId: taskcal.id })],
      calendars: [makeCalendar('standard'), taskcal, rescal],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50, calendarId: rescal.id }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
      startDate: iso(MONDAY_08),
    })
    const result = resultOf(document)
    const task = scheduleOf(result, 'a')
    expect(task.resolvedCalendarId).toEqual(asCalendarId('taskcal'))
    expect(task.scheduledStart).toBe(MONDAY_08)
    expect(task.scheduledFinish).toBe(MONDAY_08_FINISH)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('rescal'))
    expect(assignment.work).toEqual(wm(480))
    expect(assignment.cost).toBe(400)
  })
})

describe('PROJECT-011 golden W15 — summary work', () => {
  it('summary work = sum of children derived work', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 's', summary: true, outlineLevel: 1 }),
        makeTask({ id: 'a', duration: day(480), parentTaskId: asTaskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'b', duration: day(480), parentTaskId: asTaskId('s'), outlineLevel: 2 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').work).toEqual(wm(480))
    expect(scheduleOf(result, 'b').work).toEqual(wm(480))
    expect(scheduleOf(result, 's').work).toEqual(wm(960))
  })
})

describe('PROJECT-011 golden W16 — summary cost', () => {
  it('summary cost = sum of children derived cost', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 's', summary: true, outlineLevel: 1 }),
        makeTask({ id: 'a', duration: day(480), parentTaskId: asTaskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'b', duration: day(480), parentTaskId: asTaskId('s'), outlineLevel: 2 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').cost).toBe(400)
    expect(scheduleOf(result, 'b').cost).toBe(400)
    expect(scheduleOf(result, 's').cost).toBe(800)
    expect(scheduleOf(result, 's').actualCost).toBe(0)
    expect(scheduleOf(result, 's').remainingCost).toBe(800)
  })
})

describe('PROJECT-011 golden W17 — nested summary roll-up', () => {
  it('nested summaries roll up work/cost recursively (deepest first)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'root', summary: true, outlineLevel: 1 }),
        makeTask({ id: 's', summary: true, parentTaskId: asTaskId('root'), outlineLevel: 2 }),
        makeTask({ id: 'a', duration: day(480), parentTaskId: asTaskId('s'), outlineLevel: 3 }),
        makeTask({ id: 'b', duration: day(480), parentTaskId: asTaskId('s'), outlineLevel: 3 }),
      ],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').work).toEqual(wm(480))
    expect(scheduleOf(result, 'b').work).toEqual(wm(480))
    expect(scheduleOf(result, 's').work).toEqual(wm(960))
    expect(scheduleOf(result, 'root').work).toEqual(wm(960))
    expect(scheduleOf(result, 's').cost).toBe(800)
    expect(scheduleOf(result, 'root').cost).toBe(800)
  })
})

describe('PROJECT-011 golden W18 — assignment removal', () => {
  it('unassigning a resource recomputes task work/cost to zero', () => {
    const base = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const before = resultOf(base)
    expect(scheduleOf(before, 'a').work).toEqual(wm(480))
    expect(scheduleOf(before, 'a').cost).toBe(400)
    const removed = applyProjectCommand(base, {
      type: 'UnassignResource',
      assignmentId: asAssignmentId('as1'),
    })
    expect(removed.result.accepted).toBe(true)
    const after = resultOf(removed.document)
    expect(after.assignmentSchedules ?? {}).toEqual({})
    expect(scheduleOf(after, 'a').work).toEqual(wm(0))
    expect(scheduleOf(after, 'a').cost).toBe(0)
    expect(scheduleOf(after, 'a').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-011 golden W19 — task deletion', () => {
  it('deleting a task removes its assignment work/cost; siblings retain theirs', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const before = resultOf(document)
    expect(scheduleOf(before, 'a').work).toEqual(wm(480))
    expect(scheduleOf(before, 'b').work).toEqual(wm(480))
    const deleted = applyProjectCommand(document, { type: 'DeleteTask', taskId: asTaskId('a') })
    expect(deleted.result.accepted).toBe(true)
    const after = resultOf(deleted.document)
    expect(Object.keys(after.assignmentSchedules ?? {})).toEqual(['as2'])
    expect(scheduleOf(after, 'b').work).toEqual(wm(480))
    expect(scheduleOf(after, 'b').cost).toBe(400)
  })
})

describe('PROJECT-011 golden W20 — deterministic reordered-input result', () => {
  it('reordered resources AND assignments produce byte-identical schedule bytes', () => {
    const r1 = makeResource({ id: 'r1', kind: 'work', standardRate: 50 })
    const r2 = makeResource({ id: 'r2', kind: 'material', standardRate: 100 })
    const as1 = makeAssignment('as1', 'a', 'r1', { units: 1 })
    const as2 = makeAssignment('as2', 'a', 'r2', { units: 10 })
    const build = (resources: (typeof r1)[], assignments: (typeof as1)[]): ProjectDocument =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources,
        assignments,
      })
    const a = scheduleJson(resultOf(build([r1, r2], [as1, as2])))
    const b = scheduleJson(resultOf(build([r2, r1], [as2, as1])))
    expect(a).toBe(b)
  })
})

// ===========================================================================
// Additional required tests (21-32) — semantic coverage
// ===========================================================================

describe('PROJECT-011 required 3 — 100% work-resource assignment', () => {
  it('units 1.0 (100%) derives work equal to task duration', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(scheduleOf(result, 'a').work).toEqual(wm(480))
  })
})

describe('PROJECT-011 required 7-8 — task/assignment work aggregation', () => {
  it('7. task work = sum of assignment work across all assignments on the task', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'r2', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'r3', kind: 'work', standardRate: 50 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'r2', { units: 0.5 }),
        makeAssignment('as3', 'a', 'r3', { units: 0.25 }),
      ],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as2').work).toEqual(wm(240))
    expect(assignmentScheduleOf(result, 'as3').work).toEqual(wm(120))
    expect(scheduleOf(result, 'a').work).toEqual(wm(840))
  })

  it('8. assignment work derives from task.duration × assignment.units deterministically', () => {
    // 200% aggregate: two resources at 1.0 each = 200% demand. Over-allocation
    // is a data condition; the engine does NOT level (PROJECT-013 territory).
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'r2', kind: 'work', standardRate: 50 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'r2', { units: 1 }),
      ],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as2').work).toEqual(wm(480))
    expect(scheduleOf(result, 'a').work).toEqual(wm(960))
    // Dates are unchanged — no leveling.
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-011 required 9-10 — actual/remaining work derivation', () => {
  it('9. remaining work = work − actualWork at 0% progress', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').actualWork).toEqual(wm(0))
    expect(assignmentScheduleOf(result, 'as1').remainingWork).toEqual(wm(480))
  })

  it('10. actual work = round(work × percentComplete / 100) at 100% progress', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 100 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50, costPerUse: 100 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as1').actualWork).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as1').remainingWork).toEqual(wm(0))
    expect(assignmentScheduleOf(result, 'as1').cost).toBe(500)
    expect(assignmentScheduleOf(result, 'as1').actualCost).toBe(500)
    expect(assignmentScheduleOf(result, 'as1').remainingCost).toBe(0)
  })
})

describe('PROJECT-011 required 11-13 — progress extremes', () => {
  it('11. 0% progress: actualWork = 0, remainingWork = work', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.actualWork).toEqual(wm(0))
    expect(assignment.remainingWork).toEqual(wm(480))
    expect(assignment.actualCost).toBe(0)
    expect(assignment.remainingCost).toBe(400)
  })

  it('12. partial progress (33%): actualWork = round(work × 33/100)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 33 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.work).toEqual(wm(480))
    expect(assignment.actualWork).toEqual(wm(158))
    expect(assignment.remainingWork).toEqual(wm(322))
    expect(assignment.actualCost).toBe(132)
    expect(assignment.remainingCost).toBe(268)
  })

  it('13. 100% progress: actualWork = work, remainingWork = 0', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 100 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.actualWork).toEqual(wm(480))
    expect(assignment.remainingWork).toEqual(wm(0))
    expect(assignment.actualCost).toBe(400)
    expect(assignment.remainingCost).toBe(0)
  })
})

describe('PROJECT-011 required 19 — cost resource contributes zero work', () => {
  it('a cost resource assignment contributes cost but zero work to the task', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'c1', kind: 'cost', costPerUse: 0 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'c1', { units: 1, cost: 500 }),
      ],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as2').work).toEqual(wm(0))
    expect(assignmentScheduleOf(result, 'as2').cost).toBe(500)
    const task = scheduleOf(result, 'a')
    expect(task.work).toEqual(wm(480))
    expect(task.cost).toBe(900)
  })
})

describe('PROJECT-011 required 20 — multiple assignments aggregate deterministically', () => {
  it('three assignments with distinct units aggregate deterministically', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 50 })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
        makeResource({ id: 'r2', kind: 'work', standardRate: 60 }),
        makeResource({ id: 'm1', kind: 'material', standardRate: 100 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'r2', { units: 0.5 }),
        makeAssignment('as3', 'a', 'm1', { units: 5 }),
      ],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').work).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as2').work).toEqual(wm(240))
    expect(assignmentScheduleOf(result, 'as3').work).toEqual(wm(0))
    const task = scheduleOf(result, 'a')
    // work = 480 + 240 + 0 = 720; material contributes 0 work.
    expect(task.work).toEqual(wm(720))
    expect(task.actualWork).toEqual(wm(360))
    expect(task.remainingWork).toEqual(wm(360))
    // cost = 400 (r1) + 120 (r2: 240/60 × 60... wait, 240/60=4, 4×60=240) + 500 (m1: 5×100)
    // cost = 400 + 240 + 500 = 1140
    expect(task.cost).toBe(1140)
    expect(task.actualCost).toBe(570)
    expect(task.remainingCost).toBe(570)
  })
})

describe('PROJECT-011 required 23 — status-date interaction', () => {
  it('statusDate affects derived status but does NOT override percentComplete for work', () => {
    // percentComplete = 0; statusDate is AFTER the scheduled finish. Per
    // PROJECT-008, the derived status is `inProgress` (work should have begun).
    // But actualWork is derived purely from percentComplete (= 0), so
    // actualWork = 0. The status date does NOT silently invent progress.
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const withStatus: ProjectDocument = {
      ...document,
      properties: {
        ...document.properties,
        statusDate: iso(TUESDAY_FINISH),
      },
    }
    const result = resultOf(withStatus)
    const task = scheduleOf(result, 'a')
    // Status is derived from percentComplete + statusDate (PROJECT-008).
    expect(task.status).toBe('inProgress')
    // But actualWork is derived purely from percentComplete (= 0).
    expect(task.actualWork).toEqual(wm(0))
    expect(task.remainingWork).toEqual(wm(480))
    expect(assignmentScheduleOf(result, 'as1').actualWork).toEqual(wm(0))
  })

  it('statusDate before start: status is notStarted; actualWork still 0', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const beforeStart: ProjectDocument = {
      ...document,
      properties: {
        ...document.properties,
        statusDate: iso('2026-08-01T09:00:00.000Z'),
      },
    }
    const result = resultOf(beforeStart)
    const task = scheduleOf(result, 'a')
    expect(task.status).toBe('notStarted')
    expect(task.actualWork).toEqual(wm(0))
  })
})

describe('PROJECT-011 required 27 — task deletion removes assignment work/cost', () => {
  it('after deletion the sibling task retains full work/cost; deleted task is gone', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', standardRate: 50 })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r1', { units: 1 }),
      ],
    })
    const deleted = applyProjectCommand(document, { type: 'DeleteTask', taskId: asTaskId('a') })
    expect(deleted.result.accepted).toBe(true)
    const after = resultOf(deleted.document)
    expect(after.taskSchedules[taskId('a')]).toBeUndefined()
    expect(scheduleOf(after, 'b').work).toEqual(wm(480))
    expect(scheduleOf(after, 'b').cost).toBe(400)
  })
})

describe('PROJECT-011 required 29-30 — determinism (reorder)', () => {
  it('29. reordered assignment array produces byte-identical schedule bytes', () => {
    const as1 = makeAssignment('as1', 'a', 'r1', { units: 1 })
    const as2 = makeAssignment('as2', 'a', 'r2', { units: 0.5 })
    const build = (assignments: (typeof as1)[]): ProjectDocument =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources: [
          makeResource({ id: 'r1', kind: 'work', standardRate: 50 }),
          makeResource({ id: 'r2', kind: 'work', standardRate: 60 }),
        ],
        assignments,
      })
    const a = scheduleJson(resultOf(build([as1, as2])))
    const b = scheduleJson(resultOf(build([as2, as1])))
    expect(a).toBe(b)
  })

  it('30. reordered resource array produces byte-identical schedule bytes', () => {
    const r1 = makeResource({ id: 'r1', kind: 'work', standardRate: 50 })
    const r2 = makeResource({ id: 'r2', kind: 'work', standardRate: 60 })
    const build = (resources: (typeof r1)[]): ProjectDocument =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources,
        assignments: [
          makeAssignment('as1', 'a', 'r1', { units: 1 }),
          makeAssignment('as2', 'a', 'r2', { units: 0.5 }),
        ],
      })
    const a = scheduleJson(resultOf(build([r1, r2])))
    const b = scheduleJson(resultOf(build([r2, r1])))
    expect(a).toBe(b)
  })
})

describe('PROJECT-011 required 31 — repeated schedule determinism', () => {
  it('three consecutive schedule() calls produce byte-identical output', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 50 })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50, costPerUse: 100 }),
        makeResource({ id: 'm1', kind: 'material', standardRate: 100 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'm1', { units: 5 }),
      ],
    })
    const first = scheduleJson(resultOf(document))
    const second = scheduleJson(resultOf(document))
    const third = scheduleJson(resultOf(document))
    expect(first).toBe(second)
    expect(second).toBe(third)
  })
})

describe('PROJECT-011 required 32 — serialized round-trip determinism', () => {
  it('the same serialized document + options produce byte-identical schedules', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480), percentComplete: 25 }),
        makeTask({ id: 'b', duration: day(480) }),
      ],
      resources: [
        makeResource({ id: 'r1', kind: 'work', maxUnits: 1, standardRate: 50 }),
        makeResource({ id: 'r2', kind: 'material', standardRate: 100 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'b', 'r2', { units: 10 }),
      ],
    })
    const run1 = scheduleJson(resultOf(document))
    const run2 = scheduleJson(resultOf(parseDocument(documentJson(document))))
    expect(run1).toBe(run2)
  })
})

// Cross-run byte equality across a serialized round-trip, mirroring the
// PROJECT-006 determinism contract and extending it to work/cost derivation.
describe('PROJECT-011 determinism — work/cost serialized round-trip byte equality', () => {
  it('derived work/cost survive JSON serialization round-trips unchanged', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 50 })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', standardRate: 50, costPerUse: 100 }),
        makeResource({ id: 'c1', kind: 'cost', costPerUse: 0 }),
      ],
      assignments: [
        makeAssignment('as1', 'a', 'r1', { units: 1 }),
        makeAssignment('as2', 'a', 'c1', { units: 1, cost: 500 }),
      ],
    })
    const run1 = scheduleJson(resultOf(document))
    const run2 = scheduleJson(resultOf(parseDocument(documentJson(document))))
    expect(run1).toBe(run2)
  })
})

// Suppress unused-import lint for TUESDAY which documents the calendar
// arithmetic contract but is only used in the status-date test.
void TUESDAY
void resolveCalendar
void resolveResourceCalendarId
