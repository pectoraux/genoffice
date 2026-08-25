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

// PROJECT-010 calendar-precedence fixtures use an 08:00 project start so the
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
// Golden fixtures R01-R15
// ===========================================================================

describe('PROJECT-010 golden R01 — single work resource', () => {
  it('schedules the task and exposes the work resource assignment schedule', () => {
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
    expect(task.earlyStart).toBe(MONDAY)
    expect(task.earlyFinish).toBe(MONDAY_FINISH)
    expect(task.resolvedCalendarId).toEqual(asCalendarId('standard'))
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.assignmentId).toEqual(asAssignmentId('as1'))
    expect(assignment.taskId).toEqual(asTaskId('a'))
    expect(assignment.resourceId).toEqual(asResourceId('r1'))
    expect(assignment.resourceType).toBe('work')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('standard'))
    expect(assignment.maxUnits).toBe(1)
    expect(assignment.units).toBe(1)
  })
})

describe('PROJECT-010 golden R02 — work resource with limited units', () => {
  it('echoes maxUnits 0.5 without changing task dates (no work calc yet)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work', maxUnits: 0.5 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 0.5 })],
    })
    const result = resultOf(document)
    const task = scheduleOf(result, 'a')
    expect(task.scheduledStart).toBe(MONDAY)
    expect(task.scheduledFinish).toBe(MONDAY_FINISH)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.maxUnits).toBe(0.5)
    expect(assignment.units).toBe(0.5)
  })
})

describe('PROJECT-010 golden R03 — work resource with standard/overtime rates', () => {
  it('exposes the resource rates on the Resource; assignment schedule echoes type/units', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          kind: 'work',
          maxUnits: 1,
          standardRate: 50,
          overtimeRate: 75,
          costPerUse: 100,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const result = resultOf(document)
    const resource = document.resources[0]
    expect(resource.standardRate).toBe(50)
    expect(resource.overtimeRate).toBe(75)
    expect(resource.costPerUse).toBe(100)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resourceType).toBe('work')
    expect(assignment.units).toBe(1)
    // Task scheduling is unaffected by rates in PROJECT-010.
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-010 golden R04 — material resource', () => {
  it('schedules with a material resource assignment (no work-capacity semantics)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'm1',
          name: 'Concrete',
          kind: 'material',
          maxUnits: 0,
          standardRate: 100,
          costPerUse: 0,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'm1', { units: 1 })],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resourceType).toBe('material')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('standard'))
  })
})

describe('PROJECT-010 golden R05 — cost resource', () => {
  it('schedules with a cost resource (pure cost category, never work-capacity)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'c1',
          name: 'Travel',
          kind: 'cost',
          maxUnits: 0,
          costPerUse: 500,
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'c1', { units: 1 })],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resourceType).toBe('cost')
    // A cost resource never carries work capacity; maxUnits is echoed but
    // downstream must not treat it as a capacity bound.
    expect(assignment.maxUnits).toBe(0)
  })
})

describe('PROJECT-010 golden R06 — resource-specific calendar', () => {
  it('exposes the resource calendar id on the assignment schedule without moving task dates', () => {
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
      resources: [makeResource({ id: 'r1', calendarId: resourceCalendar.id })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    const result = resultOf(document)
    // Task scheduling uses the TASK calendar (default = standard, 09:00-17:00),
    // NOT the resource calendar. The resource calendar is a separate input.
    const task = scheduleOf(result, 'a')
    expect(task.resolvedCalendarId).toEqual(asCalendarId('standard'))
    expect(task.scheduledStart).toBe(MONDAY)
    expect(task.scheduledFinish).toBe(MONDAY_FINISH)
    // The assignment schedule exposes the resource's resolved calendar.
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('rescal'))
  })
})

describe('PROJECT-010 golden R07 — inherited resource calendar', () => {
  it('resolves the resource calendar through its inheritance chain', () => {
    const base = makeCalendar('base')
    const child = makeCalendar('child', {
      baseCalendarId: asCalendarId('base'),
      // Child overrides Monday to 08:00-16:00; the rest is inherited from base.
      workingWeek: {
        0: [],
        1: [{ startMinute: 480, endMinute: 960 }],
        2: [{ startMinute: 540, endMinute: 1020 }],
        3: [{ startMinute: 540, endMinute: 1020 }],
        4: [{ startMinute: 540, endMinute: 1020 }],
        5: [{ startMinute: 540, endMinute: 1020 }],
        6: [],
      },
    })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      calendars: [makeCalendar('standard'), base, child],
      resources: [makeResource({ id: 'r1', calendarId: child.id })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    const result = resultOf(document)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('child'))
    // The resolved calendar merges child over base: Monday is 08:00-16:00.
    const resolved = resolveCalendar({ calendars: document.calendars }, child.id)
    expect(resolved.workingWeek[1]).toEqual([{ startMinute: 480, endMinute: 960 }])
    // Tuesday is inherited from base (09:00-17:00).
    expect(resolved.workingWeek[2]).toEqual([{ startMinute: 540, endMinute: 1020 }])
  })
})

describe('PROJECT-010 golden R08 — resource calendar exception', () => {
  it('the resource calendar carries an exception without affecting task scheduling', () => {
    const resourceCalendar = makeCalendar('rescal', {
      exceptions: [{ date: '2026-08-03', periods: [{ startMinute: 600, endMinute: 1080 }] }],
    })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      calendars: [makeCalendar('standard'), resourceCalendar],
      resources: [makeResource({ id: 'r1', calendarId: resourceCalendar.id })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    const result = resultOf(document)
    // Task scheduling uses the task calendar (standard); the exception on the
    // resource calendar does not move the task.
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('rescal'))
  })
})

describe('PROJECT-010 golden R09 — multiple resources on one task', () => {
  it('exposes both assignment schedules sorted by AssignmentId', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work' }),
        makeResource({ id: 'm1', kind: 'material', standardRate: 100 }),
      ],
      assignments: [makeAssignment('as2', 'a', 'm1'), makeAssignment('as1', 'a', 'r1')],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    const ids = Object.keys(result.assignmentSchedules ?? {}).sort()
    expect(ids).toEqual(['as1', 'as2'])
    expect(assignmentScheduleOf(result, 'as1').resourceType).toBe('work')
    expect(assignmentScheduleOf(result, 'as2').resourceType).toBe('material')
  })
})

describe('PROJECT-010 golden R10 — one resource across multiple tasks', () => {
  it('exposes one assignment schedule per task, all pointing at the same resource', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work' })],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'b', 'r1')],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(result, 'b').scheduledFinish).toBe(MONDAY_FINISH)
    expect(assignmentScheduleOf(result, 'as1').resourceId).toEqual(asResourceId('r1'))
    expect(assignmentScheduleOf(result, 'as2').resourceId).toEqual(asResourceId('r1'))
  })
})

describe('PROJECT-010 golden R11 — assignment mutation (assign then unassign)', () => {
  it('assigning adds an assignment schedule; unassigning removes it; task dates stable', () => {
    const base = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1' })],
    })
    // No assignments initially.
    let result = resultOf(base)
    expect(result.assignmentSchedules ?? {}).toEqual({})
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    // Assign a resource.
    const assigned = applyProjectCommand(base, {
      type: 'AssignResource',
      assignment: makeAssignment('as1', 'a', 'r1'),
    })
    expect(assigned.result.accepted).toBe(true)
    result = resultOf(assigned.document)
    expect(Object.keys(result.assignmentSchedules ?? {})).toEqual(['as1'])
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    // Unassign.
    const unassigned = applyProjectCommand(assigned.document, {
      type: 'UnassignResource',
      assignmentId: asAssignmentId('as1'),
    })
    expect(unassigned.result.accepted).toBe(true)
    result = resultOf(unassigned.document)
    expect(result.assignmentSchedules ?? {}).toEqual({})
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-010 golden R12 — task deletion with assignments', () => {
  it("deleting a task removes only that task's assignments", () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'b', 'r1')],
    })
    const before = resultOf(document)
    expect(Object.keys(before.assignmentSchedules ?? {}).sort()).toEqual(['as1', 'as2'])
    const deleted = applyProjectCommand(document, { type: 'DeleteTask', taskId: asTaskId('a') })
    expect(deleted.result.accepted).toBe(true)
    const after = resultOf(deleted.document)
    expect(Object.keys(after.assignmentSchedules ?? {})).toEqual(['as2'])
    expect(scheduleOf(after, 'b').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-010 golden R13 — reordered-resource deterministic schedule', () => {
  it('produces byte-identical schedule bytes regardless of resource array order', () => {
    const r1 = makeResource({ id: 'r1', kind: 'work' })
    const r2 = makeResource({ id: 'r2', kind: 'material' })
    const base = (resources: (typeof r1)[]) =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources,
        assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'a', 'r2')],
      })
    const docA = base([r1, r2])
    const docB = base([r2, r1])
    const jsonA = scheduleJson(resultOf(docA))
    const jsonB = scheduleJson(resultOf(docB))
    expect(jsonA).toBe(jsonB)
  })
})

describe('PROJECT-010 golden R14 — reordered-assignment deterministic schedule', () => {
  it('produces byte-identical schedule bytes regardless of assignment array order', () => {
    const as1 = makeAssignment('as1', 'a', 'r1')
    const as2 = makeAssignment('as2', 'a', 'r2')
    const base = (assignments: (typeof as1)[]) =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources: [makeResource({ id: 'r1' }), makeResource({ id: 'r2' })],
        assignments,
      })
    const docA = base([as1, as2])
    const docB = base([as2, as1])
    const jsonA = scheduleJson(resultOf(docA))
    const jsonB = scheduleJson(resultOf(docB))
    expect(jsonA).toBe(jsonB)
  })
})

describe('PROJECT-010 golden R15 — project/task/resource calendar precedence', () => {
  it('task calendar governs task scheduling; resource calendar is a separate input', () => {
    // Three calendars: project default ('project'), task calendar ('taskcal'),
    // resource calendar ('rescal'). Each has a distinct Monday window so the
    // precedence is observable: the task uses 'taskcal' (08:00-16:00), not the
    // resource calendar or the project default.
    const project = makeCalendar('project', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 540, endMinute: 1020 }],
        2: [{ startMinute: 540, endMinute: 1020 }],
        3: [{ startMinute: 540, endMinute: 1020 }],
        4: [{ startMinute: 540, endMinute: 1020 }],
        5: [{ startMinute: 540, endMinute: 1020 }],
        6: [],
      },
    })
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
      calendars: [project, taskcal, rescal],
      resources: [makeResource({ id: 'r1', calendarId: rescal.id })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
      startDate: iso(MONDAY_08),
    })
    // Override the default calendar to 'project' (makeDocument defaults to the
    // first calendar; we want the explicit 'project' calendar as default).
    const fixed: ProjectDocument = {
      ...document,
      properties: { ...document.properties, defaultCalendarId: project.id },
    }
    const result = resultOf(fixed)
    const task = scheduleOf(result, 'a')
    // Task scheduling uses the TASK calendar ('taskcal', 08:00-16:00), so the
    // task starts at 08:00 and finishes at 16:00 on Monday — NOT 09:00-17:00
    // (project default) and NOT 10:00-20:00 (resource calendar).
    expect(task.resolvedCalendarId).toEqual(asCalendarId('taskcal'))
    expect(task.scheduledStart).toBe(MONDAY_08)
    expect(task.scheduledFinish).toBe(MONDAY_08_FINISH)
    // The assignment schedule exposes the RESOURCE calendar ('rescal'),
    // independent of the task calendar.
    const assignment = assignmentScheduleOf(result, 'as1')
    expect(assignment.resolvedCalendarId).toEqual(asCalendarId('rescal'))
  })
})

// ===========================================================================
// Required tests 14, 17, 24-30 (scheduling semantics)
// ===========================================================================

describe('PROJECT-010 required 14 — resource calendar inheritance resolution', () => {
  it('resolveCalendar merges a child resource calendar over its base', () => {
    const base = makeCalendar('base', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 540, endMinute: 1020 }],
        2: [{ startMinute: 540, endMinute: 1020 }],
        3: [{ startMinute: 540, endMinute: 1020 }],
        4: [{ startMinute: 540, endMinute: 1020 }],
        5: [{ startMinute: 540, endMinute: 1020 }],
        6: [],
      },
    })
    const child = makeCalendar('child', {
      baseCalendarId: base.id,
      workingWeek: {
        0: [],
        1: [{ startMinute: 480, endMinute: 960 }],
        2: [{ startMinute: 540, endMinute: 1020 }],
        3: [{ startMinute: 540, endMinute: 1020 }],
        4: [{ startMinute: 540, endMinute: 1020 }],
        5: [{ startMinute: 540, endMinute: 1020 }],
        6: [],
      },
    })
    const resolved = resolveCalendar(
      { calendars: [makeCalendar('standard'), base, child] },
      child.id,
    )
    expect(resolved.workingWeek[1]).toEqual([{ startMinute: 480, endMinute: 960 }])
    expect(resolved.workingWeek[2]).toEqual([{ startMinute: 540, endMinute: 1020 }])
  })
})

describe('PROJECT-010 required 17 — overlapping availability determinism', () => {
  it('overlapping availability windows resolve deterministically (sorted by start)', () => {
    const start = iso('2026-08-03T09:00:00.000Z')
    const finish = iso('2026-09-03T09:00:00.000Z')
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          availability: [
            { start, finish, units: 1 },
            { start, finish, units: 0.5 },
          ],
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    // The document is valid (overlaps are accepted, not silently merged).
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    // Reordering the availability windows produces the same schedule bytes.
    const reordered = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({
          id: 'r1',
          availability: [
            { start, finish, units: 0.5 },
            { start, finish, units: 1 },
          ],
        }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    expect(scheduleJson(resultOf(reordered))).toBe(scheduleJson(result))
  })
})

describe('PROJECT-010 required 24 — resource calendar deterministic resolution', () => {
  it('resolveResourceCalendarId returns the same id for repeated calls', () => {
    const rescal = makeCalendar('rescal')
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      calendars: [makeCalendar('standard'), rescal],
      resources: [makeResource({ id: 'r1', calendarId: rescal.id })],
    })
    const id1 = resolveResourceCalendarId(document, asResourceId('r1'))
    const id2 = resolveResourceCalendarId(document, asResourceId('r1'))
    expect(id1).toEqual(asCalendarId('rescal'))
    expect(id1).toBe(id2)
    // A resource with no calendar falls back to the project default.
    const document2 = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r2' })],
    })
    expect(resolveResourceCalendarId(document2, asResourceId('r2'))).toEqual(
      document2.properties.defaultCalendarId,
    )
    // A missing resource returns undefined (no silent fallback to a wrong calendar).
    expect(resolveResourceCalendarId(document2, asResourceId('ghost'))).toBeUndefined()
  })
})

describe('PROJECT-010 required 25 — task/project/resource calendar precedence', () => {
  it('task calendar wins for task scheduling; resource calendar is independent', () => {
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
    const rescal = makeCalendar('rescal')
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480), calendarId: taskcal.id })],
      calendars: [makeCalendar('standard'), taskcal, rescal],
      resources: [makeResource({ id: 'r1', calendarId: rescal.id })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
      startDate: iso(MONDAY_08),
    })
    const result = resultOf(document)
    const task = scheduleOf(result, 'a')
    expect(task.resolvedCalendarId).toEqual(asCalendarId('taskcal'))
    expect(task.scheduledStart).toBe(MONDAY_08)
    expect(task.scheduledFinish).toBe(MONDAY_08_FINISH)
    expect(assignmentScheduleOf(result, 'as1').resolvedCalendarId).toEqual(asCalendarId('rescal'))
  })

  it('falls back to project default when no task calendar is set', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    const result = resultOf(document)
    expect(scheduleOf(result, 'a').resolvedCalendarId).toEqual(
      document.properties.defaultCalendarId,
    )
    expect(assignmentScheduleOf(result, 'as1').resolvedCalendarId).toEqual(
      document.properties.defaultCalendarId,
    )
  })
})

describe('PROJECT-010 required 26-27 — multiple resources / one resource multiple tasks', () => {
  it('26. multiple resources on one task each get an assignment schedule', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work' }),
        makeResource({ id: 'r2', kind: 'work' }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'a', 'r2')],
    })
    const result = resultOf(document)
    expect(Object.keys(result.assignmentSchedules ?? {}).sort()).toEqual(['as1', 'as2'])
    expect(assignmentScheduleOf(result, 'as1').resourceId).toEqual(asResourceId('r1'))
    expect(assignmentScheduleOf(result, 'as2').resourceId).toEqual(asResourceId('r2'))
  })

  it('27. one resource assigned to multiple tasks appears in multiple assignment schedules', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [makeResource({ id: 'r1', kind: 'work' })],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'b', 'r1')],
    })
    const result = resultOf(document)
    expect(assignmentScheduleOf(result, 'as1').taskId).toEqual(asTaskId('a'))
    expect(assignmentScheduleOf(result, 'as2').taskId).toEqual(asTaskId('b'))
    expect(assignmentScheduleOf(result, 'as1').resourceId).toEqual(asResourceId('r1'))
    expect(assignmentScheduleOf(result, 'as2').resourceId).toEqual(asResourceId('r1'))
    // Both tasks schedule independently.
    expect(scheduleOf(result, 'a').scheduledFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(result, 'b').scheduledFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-010 required 28-30 — determinism (repeat / reorder)', () => {
  it('28. repeated resource resolution produces byte-identical schedule bytes', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work' }),
        makeResource({ id: 'r2', kind: 'material' }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'a', 'r2')],
    })
    const first = scheduleJson(resultOf(document))
    const second = scheduleJson(resultOf(parseDocument(documentJson(document))))
    expect(first).toBe(second)
    // Three runs in a row.
    const third = scheduleJson(resultOf(document))
    expect(third).toBe(first)
  })

  it('29. reordered resource array produces byte-identical schedule bytes', () => {
    const r1 = makeResource({ id: 'r1' })
    const r2 = makeResource({ id: 'r2' })
    const build = (resources: (typeof r1)[]) =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources,
        assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'a', 'r2')],
      })
    const a = scheduleJson(resultOf(build([r1, r2])))
    const b = scheduleJson(resultOf(build([r2, r1])))
    expect(a).toBe(b)
  })

  it('30. reordered assignment array produces byte-identical schedule bytes', () => {
    const as1 = makeAssignment('as1', 'a', 'r1')
    const as2 = makeAssignment('as2', 'a', 'r2')
    const build = (assignments: (typeof as1)[]) =>
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480) })],
        resources: [makeResource({ id: 'r1' }), makeResource({ id: 'r2' })],
        assignments,
      })
    const a = scheduleJson(resultOf(build([as1, as2])))
    const b = scheduleJson(resultOf(build([as2, as1])))
    expect(a).toBe(b)
  })
})

// Cross-run byte equality across a serialized round-trip, mirroring the
// PROJECT-006 determinism contract.
describe('PROJECT-010 determinism — serialized round-trip byte equality', () => {
  it('the same serialized document + options produce byte-identical schedules', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', duration: day(480) }), makeTask({ id: 'b', duration: day(480) })],
      resources: [
        makeResource({ id: 'r1', kind: 'work', maxUnits: 1 }),
        makeResource({ id: 'r2', kind: 'material', standardRate: 100 }),
      ],
      assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'b', 'r2')],
    })
    const run1 = scheduleJson(resultOf(document))
    const run2 = scheduleJson(resultOf(parseDocument(documentJson(document))))
    expect(run1).toBe(run2)
  })
})

// Suppress unused-import lint for TUESDAY/TUESDAY_FINISH which document the
// calendar arithmetic contract but are not asserted in every golden above.
void TUESDAY
void TUESDAY_FINISH
