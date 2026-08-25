import { describe, expect, it } from 'vitest'
import { schedule } from '../src/index.js'
import type { DerivedSchedule, ProjectDocument, TaskSchedule } from '@genoffice/project-contracts'
import { asCalendarId } from '@genoffice/project-contracts'
import {
  FRIDAY_FINISH,
  MONDAY,
  MONDAY_FINISH,
  NEXT_MONDAY,
  THURSDAY_FINISH,
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
  standardWeek,
  taskId,
  wm,
} from './fixtures.js'
import type { Calendar, CalendarPeriod, Task } from '@genoffice/project-contracts'

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

const splitDay = (): CalendarPeriod[] => [
  { startMinute: 540, endMinute: 720 },
  { startMinute: 780, endMinute: 1020 },
]

const day = (minutes: number) => wm(minutes)

describe('scheduling golden scenarios', () => {
  it('G01: schedules a single ASAP task on the project start', () => {
    const result = resultOf(makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480) })] }))
    const a = scheduleOf(result, 'a')
    expect(a.earlyStart).toBe(MONDAY)
    expect(a.earlyFinish).toBe(MONDAY_FINISH)
    expect(a.lateStart).toBe(MONDAY)
    expect(a.lateFinish).toBe(MONDAY_FINISH)
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(a.totalSlack).toBe(0)
    expect(a.freeSlack).toBe(0)
    expect(a.critical).toBe(true)
    expect(result.projectStart).toBe(MONDAY)
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })

  it('G02: carries an FS chain across a weekend boundary', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(2400) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe(FRIDAY_FINISH)
    expect(b.scheduledStart).toBe(NEXT_MONDAY)
    expect(b.scheduledFinish).toBe('2026-08-10T17:00:00.000Z')
    expect(a.critical).toBe(true)
    expect(b.critical).toBe(true)
    expect(a.totalSlack).toBe(0)
    expect(b.totalSlack).toBe(0)
    expect(result.projectFinish).toBe('2026-08-10T17:00:00.000Z')
  })

  it('G03: applies a positive FS lag as whole working days', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS', 480)],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    // The lag consumes all of Tuesday; work resumes Wednesday 09:00.
    expect(b.scheduledStart).toBe(WEDNESDAY)
    expect(b.scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(a.totalSlack).toBe(0)
    expect(b.totalSlack).toBe(0)
  })

  it('G04: applies a negative FS lag (lead) deterministically', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS', -240)],
      }),
    )
    const b = scheduleOf(result, 'b')
    expect(b.scheduledStart).toBe('2026-08-03T13:00:00.000Z')
    expect(b.scheduledFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(scheduleOf(result, 'a').totalSlack).toBe(0)
    expect(b.totalSlack).toBe(0)
  })

  it('G05: honors SS with lag and marks the driving start critical', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(240) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'SS', 60)],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(b.scheduledStart).toBe('2026-08-03T10:00:00.000Z')
    expect(b.scheduledFinish).toBe('2026-08-04T10:00:00.000Z')
    expect(a.totalSlack).toBe(0)
    expect(b.totalSlack).toBe(0)
    expect(a.critical).toBe(true)
    expect(b.critical).toBe(true)
  })

  it('G06: honors FF with lag so the successor finishes after the predecessor', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(240) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FF', 60)],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(b.scheduledStart).toBe('2026-08-03T14:00:00.000Z')
    expect(b.scheduledFinish).toBe('2026-08-04T10:00:00.000Z')
    expect(a.totalSlack).toBe(0)
    expect(b.totalSlack).toBe(0)
  })

  it('G07: honors SF so the successor finishes no earlier than the predecessor start', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(120) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'SF')],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(b.scheduledStart).toBe(MONDAY)
    expect(b.scheduledFinish).toBe('2026-08-03T11:00:00.000Z')
    // The SF relationship is satisfied (b finishes after a starts).
    expect(b.scheduledFinish! >= a.scheduledStart!).toBe(true)
    expect(a.totalSlack).toBe(0)
    expect(a.critical).toBe(true)
    expect(b.totalSlack).toBe(360)
    expect(b.critical).toBe(false)
  })

  it('G08: chains zero-duration milestones without collapsing the schedule', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'm', duration: day(0), milestone: true }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'm', 'FS'), makeDependency('d2', 'm', 'b', 'FS')],
      }),
    )
    const m = scheduleOf(result, 'm')
    expect(m.scheduledStart).toBe(TUESDAY)
    expect(m.scheduledFinish).toBe(TUESDAY)
    expect(m.duration).toBe(0)
    expect(scheduleOf(result, 'b').scheduledStart).toBe(TUESDAY)
    expect(scheduleOf(result, 'b').scheduledFinish).toBe(TUESDAY_FINISH)
    expect(m.critical).toBe(true)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
  })

  it('G09: skips a holiday exception between FS tasks', () => {
    const calendar = makeCalendar('standard', { exceptions: [holiday('2026-08-04')] })
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
        calendars: [calendar],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(b.scheduledStart).toBe(WEDNESDAY)
    expect(b.scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(a.critical).toBe(true)
    expect(b.critical).toBe(true)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('G10: resolves an inherited calendar with weekday and exception overrides', () => {
    const base = makeCalendar('base')
    const child = makeCalendar('child', {
      baseCalendarId: asCalendarId('base'),
      workingWeek: { ...standardWeek(), 3: [] },
      exceptions: [holiday('2026-08-06')],
    })
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(1440), calendarId: asCalendarId('child') })],
        calendars: [base, child],
      }),
    )
    const a = scheduleOf(result, 'a')
    // Mon + Tue work, Wed is off, Thu is a holiday, Friday completes the task.
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe(FRIDAY_FINISH)
    expect(a.duration).toBe(1440)
    expect(a.totalSlack).toBe(0)
    expect(result.projectFinish).toBe(FRIDAY_FINISH)
  })

  it('G11: uses a task-level split-day calendar', () => {
    const split: Calendar = {
      ...makeCalendar('split'),
      workingWeek: {
        0: [],
        1: splitDay(),
        2: splitDay(),
        3: splitDay(),
        4: splitDay(),
        5: splitDay(),
        6: [],
      },
    }
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480), calendarId: asCalendarId('split') })],
        calendars: [split],
      }),
    )
    const a = scheduleOf(result, 'a')
    // 420 working minutes per day (09-12 and 13-17).
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe('2026-08-04T10:00:00.000Z')
    expect(a.totalSlack).toBe(0)
  })

  it('G12: SNET pushes the early start but leaves slack intact', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({
            id: 'd',
            duration: day(240),
            constraintType: 'startNoEarlierThan',
            constraintDate: iso(TUESDAY),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    const d = scheduleOf(result, 'd')
    expect(d.earlyStart).toBe(TUESDAY)
    expect(d.earlyFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.lateStart).toBe('2026-08-05T13:00:00.000Z')
    expect(d.lateFinish).toBe(WEDNESDAY_FINISH)
    expect(d.totalSlack).toBe(720)
    expect(d.freeSlack).toBe(720)
    expect(d.critical).toBe(false)
    for (const id of ['a', 'b', 'c']) expect(scheduleOf(result, id).critical).toBe(true)
  })

  it('G13: MSO pins start and late dates so the same task becomes critical (MSO differs from SNET)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'mustStartOn',
          constraintDate: iso(TUESDAY),
        }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
    })
    const result = resultOf(document)
    const d = scheduleOf(result, 'd')
    // Identical early placement to the SNET variant...
    expect(d.earlyStart).toBe(TUESDAY)
    expect(d.earlyFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledStart).toBe(TUESDAY)
    // ...but MSO pins the late window to the constraint, so slack is zero.
    expect(d.lateStart).toBe(TUESDAY)
    expect(d.lateFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.totalSlack).toBe(0)
    expect(d.critical).toBe(true)
  })

  it('G14: SNLT pulls the scheduled start to the late window bounded by the constraint', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({
            id: 'd',
            duration: day(240),
            constraintType: 'startNoLaterThan',
            constraintDate: iso(TUESDAY),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    const d = scheduleOf(result, 'd')
    expect(d.earlyStart).toBe(MONDAY)
    expect(d.earlyFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(d.lateStart).toBe(TUESDAY)
    expect(d.lateFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledStart).toBe(TUESDAY)
    expect(d.scheduledFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.totalSlack).toBe(480)
    expect(d.critical).toBe(false)
  })

  it('G15: FNET pushes the finish while preserving slack', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({
            id: 'd',
            duration: day(240),
            constraintType: 'finishNoEarlierThan',
            constraintDate: iso(TUESDAY_FINISH),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    const d = scheduleOf(result, 'd')
    expect(d.scheduledStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledFinish).toBe(TUESDAY_FINISH)
    expect(d.totalSlack).toBe(480)
    expect(d.critical).toBe(false)
  })

  it('G16: MFO pins the finish and makes the task critical (MFO differs from FNET)', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({
            id: 'd',
            duration: day(240),
            constraintType: 'mustFinishOn',
            constraintDate: iso(TUESDAY_FINISH),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    const d = scheduleOf(result, 'd')
    expect(d.scheduledStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledFinish).toBe(TUESDAY_FINISH)
    expect(d.lateFinish).toBe(TUESDAY_FINISH)
    expect(d.lateStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.totalSlack).toBe(0)
    expect(d.critical).toBe(true)
  })

  it('G17: ALAP schedules the task at its late window', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(240), constraintType: 'asLateAsPossible' }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    const d = scheduleOf(result, 'd')
    expect(d.earlyStart).toBe(MONDAY)
    expect(d.lateStart).toBe('2026-08-05T13:00:00.000Z')
    expect(d.lateFinish).toBe(WEDNESDAY_FINISH)
    expect(d.scheduledStart).toBe('2026-08-05T13:00:00.000Z')
    expect(d.scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(d.totalSlack).toBe(1200)
    expect(d.critical).toBe(false)
  })

  it('G18: rolls nested summaries up from their children', () => {
    const summary = (id: string, parentTaskId?: string, outlineLevel = 1): Task =>
      makeTask({
        id,
        summary: true,
        duration: wm(0),
        parentTaskId: parentTaskId === undefined ? undefined : taskId(parentTaskId),
        outlineLevel,
      })
    const result = resultOf(
      makeDocument({
        tasks: [
          summary('s1'),
          makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s1'), outlineLevel: 2 }),
          summary('s2', taskId('s1'), 2),
          makeTask({
            id: 'b',
            duration: day(240),
            parentTaskId: taskId('s2'),
            outlineLevel: 3,
            constraintType: 'startNoEarlierThan',
            constraintDate: iso(TUESDAY),
          }),
          makeTask({
            id: 'c',
            duration: day(480),
            parentTaskId: taskId('s2'),
            outlineLevel: 3,
            constraintType: 'startNoEarlierThan',
            constraintDate: iso(WEDNESDAY),
          }),
        ],
        dependencies: [makeDependency('d1', 'b', 'c', 'FS')],
      }),
    )
    const s1 = scheduleOf(result, 's1')
    const s2 = scheduleOf(result, 's2')
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    const c = scheduleOf(result, 'c')
    expect(s2.scheduledStart).toBe(TUESDAY)
    expect(s2.scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(s2.duration).toBe(960)
    expect(s1.scheduledStart).toBe(MONDAY)
    expect(s1.scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(s1.duration).toBe(1440)
    expect(s2.critical).toBe(true)
    expect(s1.critical).toBe(true)
    expect(c.critical).toBe(true)
    expect(b.totalSlack).toBe(240)
    expect(a.totalSlack).toBe(960)
    expect(a.critical).toBe(false)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('G19: computes the critical path with total and free slack across branches', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
          makeTask({ id: 'p', duration: day(240) }),
          makeTask({ id: 'p1', duration: day(480) }),
          makeTask({ id: 's', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'c', 'd', 'FS'),
          makeDependency('d4', 'p', 's', 'FS'),
          makeDependency('d5', 'p1', 's', 'FS'),
        ],
      }),
    )
    expect(result.projectFinish).toBe(THURSDAY_FINISH)
    for (const id of ['a', 'b', 'c', 'd']) {
      const entry = scheduleOf(result, id)
      expect(entry.critical).toBe(true)
      expect(entry.totalSlack).toBe(0)
      expect(entry.freeSlack).toBe(0)
    }
    const s = scheduleOf(result, 's')
    expect(s.totalSlack).toBe(960)
    expect(s.freeSlack).toBe(960)
    const p1 = scheduleOf(result, 'p1')
    expect(p1.totalSlack).toBe(960)
    expect(p1.freeSlack).toBe(0)
    const p = scheduleOf(result, 'p')
    expect(p.totalSlack).toBe(1200)
    expect(p.freeSlack).toBe(240)
    expect(p.critical).toBe(false)
  })

  it('G20: produces byte-identical DerivedSchedule output for repeated and reordered runs', () => {
    const standard = makeCalendar('standard')
    const split: Calendar = {
      ...makeCalendar('split'),
      workingWeek: {
        0: [],
        1: splitDay(),
        2: splitDay(),
        3: splitDay(),
        4: splitDay(),
        5: splitDay(),
        6: [],
      },
    }
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(240), calendarId: asCalendarId('split') }),
        makeTask({ id: 'm', duration: day(0), milestone: true }),
        makeTask({
          id: 'c',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
        makeTask({ id: 'd', duration: day(240), constraintType: 'asLateAsPossible' }),
        makeTask({ id: 's', summary: true, duration: wm(0) }),
        makeTask({ id: 'e', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'f', duration: day(120), parentTaskId: taskId('s'), outlineLevel: 2 }),
      ],
      dependencies: [
        makeDependency('d1', 'a', 'b', 'FS', 60),
        makeDependency('d2', 'b', 'm', 'SS', -30),
        makeDependency('d3', 'a', 'c', 'FF', 120),
        makeDependency('d4', 'm', 'e', 'SF'),
        makeDependency('d5', 'e', 'f', 'FS'),
      ],
      calendars: [standard, split],
    })

    const serialized = JSON.stringify(document)
    const first = schedule(parseDocument(serialized))
    const second = schedule(parseDocument(serialized))
    const third = schedule(parseDocument(serialized))
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(JSON.stringify(second)).toBe(JSON.stringify(third))

    // Reordering the input arrays must not change the output bytes either.
    const shuffled = parseDocument(serialized)
    shuffled.tasks = [...shuffled.tasks].reverse()
    shuffled.dependencies = [...shuffled.dependencies].reverse()
    shuffled.calendars = [...shuffled.calendars].reverse()
    const reordered = schedule(shuffled)
    expect(JSON.stringify(reordered)).toBe(JSON.stringify(first))

    // Sanity: the document actually schedules without diagnostics.
    expect(first.diagnostics).toEqual([])
    expect(Object.keys(first.taskSchedules).sort()).toEqual([
      'a',
      'b',
      'c',
      'd',
      'e',
      'f',
      'm',
      's',
    ])
  })
})

describe('scheduling input rejection', () => {
  const expectRejection = (document: ProjectDocument, code: string) => {
    const result = schedule(document)
    expect(result.taskSchedules).toEqual({})
    expect(result.diagnostics.length).toBeGreaterThan(0)
    expect(result.diagnostics.some((diagnostic) => diagnostic.code === code)).toBe(true)
    expect(result.diagnostics.every((diagnostic) => diagnostic.severity === 'error')).toBe(true)
  }

  it('rejects a dependency cycle', () => {
    expectRejection(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'a', 'FS')],
      }),
      'DEPENDENCY_CYCLE',
    )
  })

  it('rejects a self dependency', () => {
    expectRejection(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        dependencies: [makeDependency('d1', 'a', 'a', 'FS')],
      }),
      'SELF_DEPENDENCY',
    )
  })

  it('rejects a dependency referencing a missing task', () => {
    expectRejection(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        dependencies: [makeDependency('d1', 'a', 'ghost', 'FS')],
      }),
      'MISSING_TASK_REFERENCE',
    )
  })

  it('rejects a parent-child cycle', () => {
    expectRejection(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', parentTaskId: taskId('b') }),
          makeTask({ id: 'b', parentTaskId: taskId('a') }),
        ],
      }),
      'PARENT_CYCLE',
    )
  })

  it('rejects a calendar inheritance cycle', () => {
    const first = makeCalendar('first', { baseCalendarId: asCalendarId('second') })
    const second = makeCalendar('second', { baseCalendarId: asCalendarId('first') })
    expectRejection(
      makeDocument({ tasks: [makeTask({ id: 'a' })], calendars: [first, second] }),
      'CALENDAR_CYCLE',
    )
  })

  it('rejects malformed calendar periods', () => {
    const broken = makeCalendar('broken', {
      workingWeek: { ...standardWeek(), 1: [{ startMinute: 1020, endMinute: 540 }] },
    })
    expectRejection(
      makeDocument({ tasks: [makeTask({ id: 'a' })], calendars: [broken] }),
      'CALENDAR_PERIOD_MALFORMED',
    )
  })

  it('rejects duplicate task ids', () => {
    expectRejection(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'a' })],
      }),
      'DUPLICATE_TASK_ID',
    )
  })

  it('rejects a dependency between a summary and its own descendant', () => {
    expectRejection(
      makeDocument({
        tasks: [
          makeTask({ id: 's', summary: true, duration: wm(0) }),
          makeTask({ id: 'child', parentTaskId: taskId('s'), outlineLevel: 2 }),
        ],
        dependencies: [makeDependency('d1', 'child', 's', 'FS')],
      }),
      'SUMMARY_DEPENDENCY',
    )
  })
})
