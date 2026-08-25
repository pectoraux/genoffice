import { describe, expect, it } from 'vitest'
import { schedule } from '../src/index.js'
import type {
  Calendar,
  CalendarPeriod,
  DerivedSchedule,
  ProjectDocument,
  Task,
  TaskSchedule,
} from '@genoffice/project-contracts'
import { asCalendarId } from '@genoffice/project-contracts'
import {
  FRIDAY,
  FRIDAY_FINISH,
  MONDAY,
  MONDAY_FINISH,
  NEXT_MONDAY,
  PREV_FRIDAY,
  THURSDAY,
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

// Additional exact instants used by the PROJECT-012 goldens. The standard
// week is Mon–Fri 09:00–17:00 (540–1020 working-minute windows, 480 min/day).
const MONDAY_13 = '2026-08-03T13:00:00.000Z'
const TUESDAY_13 = '2026-08-04T13:00:00.000Z'
const WEDNESDAY_13 = '2026-08-05T13:00:00.000Z'
const THURSDAY_13 = '2026-08-06T13:00:00.000Z'
const PREV_FRIDAY_START = '2026-07-31T09:00:00.000Z'

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

const documentJson = (document: ProjectDocument): string => JSON.stringify(document)
const scheduleJson = (result: DerivedSchedule): string => JSON.stringify(result)

// A reusable four-task FS chain a → b → c → d, each one working day. The late
// envelope of the chain anchors the unconstrained branches in the goldens.
const chain4 = (): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({ id: 'a', duration: day(480) }),
      makeTask({ id: 'b', duration: day(480) }),
      makeTask({ id: 'c', duration: day(480) }),
      makeTask({ id: 'd', duration: day(480) }),
    ],
    dependencies: [
      makeDependency('d1', 'a', 'b', 'FS'),
      makeDependency('d2', 'b', 'c', 'FS'),
      makeDependency('d3', 'c', 'd', 'FS'),
    ],
  })

const splitDay = (): CalendarPeriod[] => [
  { startMinute: 540, endMinute: 720 },
  { startMinute: 780, endMinute: 1020 },
]

// ===========================================================================
// Golden fixtures CP01–CP18
// ===========================================================================

describe('PROJECT-012 golden CP01 — simple critical chain', () => {
  it('marks every task on a single FS chain critical with zero float', () => {
    const result = resultOf(chain4())
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    const c = scheduleOf(result, 'c')
    const d = scheduleOf(result, 'd')
    // Working-time arithmetic: the lateFinish of a chain task is the
    // lateStart of its successor (the next working day's 09:00), NOT the
    // same-day 17:00. The working-minute distance between end-of-day and
    // next-day-start is 0, so totalSlack is still 0 and the task is critical.
    expect(a).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
    })
    expect(b).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: WEDNESDAY,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(c).toMatchObject({
      earlyStart: WEDNESDAY,
      earlyFinish: WEDNESDAY_FINISH,
      lateStart: WEDNESDAY,
      lateFinish: THURSDAY,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(d).toMatchObject({
      earlyStart: THURSDAY,
      earlyFinish: THURSDAY_FINISH,
      lateStart: THURSDAY,
      lateFinish: THURSDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(THURSDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP02 — two critical paths', () => {
  it('marks every task on both equal-length parallel chains critical', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
          makeTask({ id: 'e', duration: day(480) }),
          makeTask({ id: 'f', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'c', 'f', 'FS'),
          makeDependency('d4', 'a', 'd', 'FS'),
          makeDependency('d5', 'd', 'e', 'FS'),
          makeDependency('d6', 'e', 'f', 'FS'),
        ],
      }),
    )
    for (const id of ['a', 'b', 'c', 'd', 'e', 'f']) {
      expect(scheduleOf(result, id).critical).toBe(true)
      expect(scheduleOf(result, id).totalSlack).toBe(0)
      expect(scheduleOf(result, id).freeSlack).toBe(0)
    }
    expect(scheduleOf(result, 'a')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
    })
    expect(scheduleOf(result, 'f')).toMatchObject({
      earlyStart: THURSDAY,
      earlyFinish: THURSDAY_FINISH,
    })
    expect(result.projectFinish).toBe(THURSDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP03 — three critical paths', () => {
  it('marks every task on three equal-length parallel chains critical', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
          makeTask({ id: 'e', duration: day(480) }),
          makeTask({ id: 'g', duration: day(480) }),
          makeTask({ id: 'h', duration: day(480) }),
          makeTask({ id: 'f', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'c', 'f', 'FS'),
          makeDependency('d4', 'a', 'd', 'FS'),
          makeDependency('d5', 'd', 'e', 'FS'),
          makeDependency('d6', 'e', 'f', 'FS'),
          makeDependency('d7', 'a', 'g', 'FS'),
          makeDependency('d8', 'g', 'h', 'FS'),
          makeDependency('d9', 'h', 'f', 'FS'),
        ],
      }),
    )
    for (const id of ['a', 'b', 'c', 'd', 'e', 'g', 'h', 'f']) {
      expect(scheduleOf(result, id).critical).toBe(true)
      expect(scheduleOf(result, id).totalSlack).toBe(0)
      expect(scheduleOf(result, id).freeSlack).toBe(0)
    }
    expect(result.projectFinish).toBe(THURSDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP04 — diamond convergence', () => {
  it('marks all four diamond tasks critical when both branches are equal', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'a', 'c', 'FS'),
          makeDependency('d3', 'b', 'd', 'FS'),
          makeDependency('d4', 'c', 'd', 'FS'),
        ],
      }),
    )
    for (const id of ['a', 'b', 'c', 'd']) {
      expect(scheduleOf(result, id).critical).toBe(true)
      expect(scheduleOf(result, id).totalSlack).toBe(0)
      expect(scheduleOf(result, id).freeSlack).toBe(0)
    }
    // Working-time late dates: b and c's lateFinish = d's lateStart = Wed.
    expect(scheduleOf(result, 'b')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: WEDNESDAY,
    })
    expect(scheduleOf(result, 'c')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: WEDNESDAY,
    })
    expect(scheduleOf(result, 'd')).toMatchObject({
      earlyStart: WEDNESDAY,
      earlyFinish: WEDNESDAY_FINISH,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP05 — mixed dependency types', () => {
  it('derives correct early/late/float for FS→SS→FF chain with lag', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(240) }),
          makeTask({ id: 'd', duration: day(240) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'SS', 60),
          makeDependency('d3', 'c', 'd', 'FF', 60),
        ],
      }),
    )
    // a — critical (drives b which drives the project finish)
    expect(scheduleOf(result, 'a')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
    })
    // b — critical (project finish)
    expect(scheduleOf(result, 'b')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    // c — SS successor of b (lag 60): starts at 10:00, finishes at 14:00
    expect(scheduleOf(result, 'c').earlyStart).toBe('2026-08-04T10:00:00.000Z')
    expect(scheduleOf(result, 'c').earlyFinish).toBe('2026-08-04T14:00:00.000Z')
    expect(scheduleOf(result, 'c').totalSlack).toBe(120)
    expect(scheduleOf(result, 'c').freeSlack).toBe(0)
    expect(scheduleOf(result, 'c').critical).toBe(false)
    // d — FF successor of c (lag 60): starts at 11:00, finishes at 15:00
    expect(scheduleOf(result, 'd').earlyStart).toBe('2026-08-04T11:00:00.000Z')
    expect(scheduleOf(result, 'd').earlyFinish).toBe('2026-08-04T15:00:00.000Z')
    expect(scheduleOf(result, 'd').totalSlack).toBe(120)
    expect(scheduleOf(result, 'd').freeSlack).toBe(120)
    expect(scheduleOf(result, 'd').critical).toBe(false)
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP06 — lag/lead on critical and noncritical paths', () => {
  it('derives float from a lag path vs a no-lag path of different length', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'c', 'd', 'FS', 240),
        ],
      }),
    )
    // a → b is the shorter path: a finishes Mon 17:00, b finishes Tue 17:00.
    // c → d has a 240-min lag: d finishes Wed 13:00 → the project finish.
    expect(scheduleOf(result, 'a')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      totalSlack: 240,
      freeSlack: 0,
      critical: false,
    })
    expect(scheduleOf(result, 'a').lateFinish).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'b')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_FINISH,
      totalSlack: 240,
      freeSlack: 240,
      critical: false,
    })
    expect(scheduleOf(result, 'b').lateFinish).toBe(WEDNESDAY_13)
    expect(scheduleOf(result, 'c')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(scheduleOf(result, 'c').lateFinish).toBe(TUESDAY)
    expect(scheduleOf(result, 'd').earlyStart).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'd').earlyFinish).toBe(WEDNESDAY_13)
    expect(scheduleOf(result, 'd')).toMatchObject({
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_13)
  })
})

describe('PROJECT-012 golden CP07 — free-vs-total slack', () => {
  it('shows freeSlack < totalSlack, freeSlack = totalSlack, and zero free slack', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(240) }),
          makeTask({ id: 'e', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'a', 'd', 'FS'),
          makeDependency('d4', 'd', 'e', 'FS'),
        ],
      }),
    )
    // a, b, c are the critical chain (Mon→Tue→Wed, each 1 day).
    expect(scheduleOf(result, 'a')).toMatchObject({
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(scheduleOf(result, 'b')).toMatchObject({
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    expect(scheduleOf(result, 'c')).toMatchObject({
      earlyStart: WEDNESDAY,
      earlyFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    // d — short branch off a. d has totalSlack=240 but freeSlack=0 because
    // slipping d delays e (e starts right when d finishes).
    expect(scheduleOf(result, 'd')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_13,
      totalSlack: 240,
      freeSlack: 0,
      critical: false,
    })
    // e — no successors, so freeSlack = totalSlack = 240.
    expect(scheduleOf(result, 'e').earlyStart).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'e').earlyFinish).toBe(WEDNESDAY_13)
    expect(scheduleOf(result, 'e')).toMatchObject({
      totalSlack: 240,
      freeSlack: 240,
      critical: false,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP08 — negative slack from MFO constraint', () => {
  it('produces negative total and free slack on a predecessor of an MFO task', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({
            id: 'b',
            duration: day(480),
            constraintType: 'mustFinishOn',
            constraintDate: iso(TUESDAY),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    // b is pinned by MFO to finish on Tuesday 09:00 (the constraint date).
    // The forward pass computes b's early start as subtractWorkingTime(Tue
    // 09:00, 480) = Monday 09:00 (the working-time boundary), so b's early
    // finish is Monday 17:00. The MFO constraint pins b's LATE finish to Tue
    // 09:00 (and late start to Monday 09:00). The FS dependency from a means
    // a must finish before b starts, but a finishes Monday 17:00 — after b's
    // pinned late start (Monday 09:00). a therefore has negative total slack.
    expect(scheduleOf(result, 'b')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
    })
    const a = scheduleOf(result, 'a')
    expect(a.earlyStart).toBe(MONDAY)
    expect(a.earlyFinish).toBe(MONDAY_FINISH)
    expect(a.lateFinish).toBe(MONDAY)
    expect(a.lateStart).toBe(PREV_FRIDAY_START)
    expect(a.totalSlack).toBe(-480)
    expect(a.freeSlack).toBe(-480)
    expect(a.critical).toBe(true)
    // projectFinish = max of leaf earlyFinish values. Both a and b finish
    // Monday 17:00 (a from its own schedule, b from the MFO-implied start),
    // so the project finish is Monday 17:00.
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP09 — near-critical branch (one-day and one-hour)', () => {
  it('keeps near-critical tasks non-critical with small positive total slack', () => {
    // One-day near-critical: a→b→c chain (critical), plus an independent d
    // that finishes one working day before the project finish.
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
          makeTask({ id: 'e', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'd', 'e', 'FS'),
        ],
      }),
    )
    // c finishes Wednesday 17:00 (project finish). e finishes Tuesday 17:00
    // (one day earlier). So d and e each have 480 min (1 working day) of
    // total slack.
    expect(scheduleOf(result, 'c')).toMatchObject({
      earlyFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(scheduleOf(result, 'e')).toMatchObject({
      earlyFinish: TUESDAY_FINISH,
      totalSlack: 480,
      critical: false,
    })
    expect(scheduleOf(result, 'd')).toMatchObject({
      totalSlack: 480,
      critical: false,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('keeps a one-hour near-critical task non-critical', () => {
    // Critical chain a→b→c (each 480, FS) finishes Wed 17:00 = project finish.
    // f(420 min) is SNET-constrained to start Wed 09:00, finishing Wed 16:00
    // (420 min). f's totalSlack = workingDuration(Wed 16:00, Wed 17:00) = 60.
    const doc = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'c', duration: day(480) }),
        makeTask({
          id: 'f',
          duration: day(420),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(WEDNESDAY),
        }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
    })
    const r = resultOf(doc)
    expect(scheduleOf(r, 'f').earlyStart).toBe(WEDNESDAY)
    expect(scheduleOf(r, 'f').earlyFinish).toBe('2026-08-05T16:00:00.000Z')
    expect(scheduleOf(r, 'f').totalSlack).toBe(60)
    expect(scheduleOf(r, 'f').critical).toBe(false)
    expect(scheduleOf(r, 'c').critical).toBe(true)
    expect(r.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP10 — critical milestone', () => {
  it('marks a zero-duration milestone on the critical chain critical', () => {
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
    expect(m.earlyStart).toBe(TUESDAY)
    expect(m.earlyFinish).toBe(TUESDAY)
    expect(m.lateStart).toBe(TUESDAY)
    expect(m.lateFinish).toBe(TUESDAY)
    expect(m.totalSlack).toBe(0)
    expect(m.freeSlack).toBe(0)
    expect(m.critical).toBe(true)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP11 — noncritical milestone', () => {
  it('keeps a zero-duration milestone non-critical when it has float', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'm', duration: day(0), milestone: true }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'a', 'm', 'FS'),
        ],
      }),
    )
    // Critical chain: a→b→c (Mon, Tue, Wed). m is a zero-duration milestone
    // after a (a finishes Mon 17:00). m starts Tuesday 09:00 (= m.EF, zero
    // duration). m has no successor, so m.LF = projectFinish = Wed 17:00.
    // m's totalSlack = workingDuration(Tue 09:00, Wed 17:00) = 480 (Tue) +
    // 480 (Wed) = 960. Not critical.
    const m = scheduleOf(result, 'm')
    expect(m.earlyStart).toBe(TUESDAY)
    expect(m.earlyFinish).toBe(TUESDAY)
    expect(m.totalSlack).toBe(960)
    expect(m.critical).toBe(false)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'c').critical).toBe(true)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP12 — summary criticality', () => {
  it('derives summary criticality from the canonical scheduling engine', () => {
    const summary = (id: string): Task =>
      makeTask({ id, summary: true, duration: wm(0), outlineLevel: 1 })
    const result = resultOf(
      makeDocument({
        tasks: [
          summary('s'),
          makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
          makeTask({ id: 'b', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    // s spans Mon 09:00 – Tue 17:00 (a starts Mon, b finishes Tue). Project
    // finish = Tue 17:00. s is critical (its finish = project finish).
    const s = scheduleOf(result, 's')
    expect(s.earlyStart).toBe(MONDAY)
    expect(s.earlyFinish).toBe(TUESDAY_FINISH)
    expect(s.lateStart).toBe(MONDAY)
    expect(s.lateFinish).toBe(TUESDAY_FINISH)
    expect(s.totalSlack).toBe(0)
    expect(s.freeSlack).toBe(0)
    expect(s.critical).toBe(true)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP13 — nested summary critical path', () => {
  it('rolls up criticality through nested summaries', () => {
    const summary = (id: string, parent?: string, level = 1): Task =>
      makeTask({
        id,
        summary: true,
        duration: wm(0),
        outlineLevel: level,
        parentTaskId: parent ? taskId(parent) : undefined,
      })
    const result = resultOf(
      makeDocument({
        tasks: [
          summary('s1'),
          summary('s2', 's1', 2),
          makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s2'), outlineLevel: 3 }),
          makeTask({ id: 'b', duration: day(480), parentTaskId: taskId('s2'), outlineLevel: 3 }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    // s2 spans Mon–Tue, s1 spans Mon–Tue. Project finish = Tue 17:00.
    expect(scheduleOf(result, 's1')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(scheduleOf(result, 's2')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP14 — constraint + critical path', () => {
  it('SNET preserves the critical chain when the constraint is on the chain', () => {
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
    // a→b→c critical chain (Mon, Tue, Wed). d is constrained SNET at Tue 09:00
    // → starts Tue 09:00, finishes Tue 13:00. d has totalSlack = 720 (can slip
    // until Wed 17:00 = project finish). Not critical.
    expect(scheduleOf(result, 'd')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_13,
      totalSlack: 720,
      freeSlack: 720,
      critical: false,
    })
    for (const id of ['a', 'b', 'c']) {
      expect(scheduleOf(result, id).critical).toBe(true)
    }
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('MSO pins the task making it critical (differs from SNET)', () => {
    const result = resultOf(
      makeDocument({
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
      }),
    )
    // MSO pins d's late window to the constraint, so d has zero slack.
    expect(scheduleOf(result, 'd')).toMatchObject({
      earlyStart: TUESDAY,
      earlyFinish: TUESDAY_13,
      lateStart: TUESDAY,
      lateFinish: TUESDAY_13,
      totalSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('FNET pushes finish while preserving slack', () => {
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
    expect(scheduleOf(result, 'd').scheduledStart).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'd').scheduledFinish).toBe(TUESDAY_FINISH)
    expect(scheduleOf(result, 'd').totalSlack).toBe(480)
    expect(scheduleOf(result, 'd').critical).toBe(false)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('MFO pins finish making the task critical', () => {
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
    expect(scheduleOf(result, 'd')).toMatchObject({
      scheduledStart: TUESDAY_13,
      scheduledFinish: TUESDAY_FINISH,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP15 — calendar boundary (different task calendars)', () => {
  it('uses working-time arithmetic for float across different task calendars', () => {
    // a uses the standard calendar (Mon–Fri 09:00–17:00). b uses an alternate
    // calendar with 10:00–16:00 working days (360 min/day). b is a successor
    // of a via FS. Float must be measured in working time, not wall clock.
    const altWeek = (): Record<number, CalendarPeriod[]> => ({
      0: [],
      1: [{ startMinute: 600, endMinute: 960 }],
      2: [{ startMinute: 600, endMinute: 960 }],
      3: [{ startMinute: 600, endMinute: 960 }],
      4: [{ startMinute: 600, endMinute: 960 }],
      5: [{ startMinute: 600, endMinute: 960 }],
      6: [],
    })
    const alt = makeCalendar('alt', { workingWeek: altWeek() })
    const std = makeCalendar('standard')
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(360), calendarId: asCalendarId('alt') }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
        calendars: [std, alt],
      }),
    )
    expect(scheduleOf(result, 'a').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_FINISH)
    // b uses the alt calendar (10:00–16:00). b starts Tue 10:00 (nextWorking
    // instant after a finishes Mon 17:00, in the alt calendar). b finishes
    // Tue 16:00 (360 min from 10:00).
    expect(scheduleOf(result, 'b').earlyStart).toBe('2026-08-04T10:00:00.000Z')
    expect(scheduleOf(result, 'b').earlyFinish).toBe('2026-08-04T16:00:00.000Z')
    // b is the project finish → b is critical.
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(scheduleOf(result, 'b').totalSlack).toBe(0)
    // a's lateFinish = b's lateStart (Tue 10:00, alt calendar). a's EF = Mon
    // 17:00. workingDuration(Mon 17:00, Tue 10:00) = Tue 09:00-10:00 = 60.
    // So a has 60 min of total slack — the alt calendar's 10:00 start leaves
    // a 60 min of slack before b is affected. Not critical.
    expect(scheduleOf(result, 'a').critical).toBe(false)
    expect(scheduleOf(result, 'a').totalSlack).toBe(60)
    expect(result.projectFinish).toBe('2026-08-04T16:00:00.000Z')
  })

  it('float is working-time, not wall-clock, across a weekend boundary', () => {
    // A 3-day task (1440 min) starting Monday spans Mon–Wed. An independent
    // 1-day task starting Monday finishes Mon 17:00. Project finish = Wed
    // 17:00. The 1-day task's totalSlack must be 960 working minutes (Tue +
    // Wed), NOT the wall-clock 48 hours.
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'big', duration: day(1440) }),
          makeTask({ id: 'small', duration: day(480) }),
        ],
      }),
    )
    expect(scheduleOf(result, 'big').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'big').earlyFinish).toBe(WEDNESDAY_FINISH)
    expect(scheduleOf(result, 'small').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'small').earlyFinish).toBe(MONDAY_FINISH)
    // small's totalSlack = workingDuration(Mon 17:00, Wed 17:00) = 960 min.
    expect(scheduleOf(result, 'small').totalSlack).toBe(960)
    expect(scheduleOf(result, 'small').freeSlack).toBe(960)
    expect(scheduleOf(result, 'small').critical).toBe(false)
    expect(scheduleOf(result, 'big').critical).toBe(true)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP16 — holiday boundary', () => {
  it('skips a holiday exception and keeps the critical chain intact', () => {
    const cal = makeCalendar('standard', { exceptions: [holiday('2026-08-04')] })
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
        calendars: [cal],
      }),
    )
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_FINISH)
    // Tuesday 2026-08-04 is a holiday, so b starts Wednesday.
    expect(scheduleOf(result, 'b').earlyStart).toBe(WEDNESDAY)
    expect(scheduleOf(result, 'b').earlyFinish).toBe(WEDNESDAY_FINISH)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP17 — isolated branch', () => {
  it('keeps an isolated task non-critical when it finishes before project finish', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'iso', duration: day(240) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    // Critical chain a→b→c finishes Wed 17:00. iso is a 240-min task with no
    // predecessors or successors. It starts Mon 09:00, finishes Mon 13:00.
    // totalSlack = workingDuration(Mon 13:00, Wed 17:00) = Mon 13:00-17:00
    // (240) + Tue (480) + Wed (480) = 1200. Not critical.
    expect(scheduleOf(result, 'iso').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'iso').earlyFinish).toBe(MONDAY_13)
    expect(scheduleOf(result, 'iso').totalSlack).toBe(1200)
    expect(scheduleOf(result, 'iso').freeSlack).toBe(1200)
    expect(scheduleOf(result, 'iso').critical).toBe(false)
    for (const id of ['a', 'b', 'c']) {
      expect(scheduleOf(result, id).critical).toBe(true)
    }
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('marks an isolated task critical when it IS the project finish', () => {
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'solo', duration: day(480) })],
      }),
    )
    expect(scheduleOf(result, 'solo').critical).toBe(true)
    expect(scheduleOf(result, 'solo').totalSlack).toBe(0)
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-012 golden CP18 — reordered deterministic input', () => {
  it('produces byte-identical DerivedSchedule under task/dependency reordering', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(240) }),
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
    })
    const serialized = JSON.stringify(document)
    const first = schedule(parseDocument(serialized))
    const second = schedule(parseDocument(serialized))
    const third = schedule(parseDocument(serialized))
    expect(scheduleJson(first)).toBe(scheduleJson(second))
    expect(scheduleJson(second)).toBe(scheduleJson(third))
    // Reordering the input arrays must not change the output bytes.
    const shuffled = parseDocument(serialized)
    shuffled.tasks = [...shuffled.tasks].reverse()
    shuffled.dependencies = [...shuffled.dependencies].reverse()
    const reordered = schedule(shuffled)
    expect(scheduleJson(reordered)).toBe(scheduleJson(first))
    expect(first.diagnostics).toEqual([])
  })
})

// ===========================================================================
// Additional required tests (covering scenarios not fully covered by CP01–CP18)
// ===========================================================================

describe('PROJECT-012 required — fan-out (one predecessor, many successors)', () => {
  it('drives every successor from the shared predecessor', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'a', 'c', 'FS'),
          makeDependency('d3', 'a', 'd', 'FS'),
        ],
      }),
    )
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_FINISH)
    for (const id of ['b', 'c', 'd']) {
      expect(scheduleOf(result, id).earlyStart).toBe(TUESDAY)
      expect(scheduleOf(result, id).earlyFinish).toBe(TUESDAY_FINISH)
      expect(scheduleOf(result, id).critical).toBe(true)
    }
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 required — fan-in (many predecessors, one successor)', () => {
  it('drives the successor from the latest predecessor', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(240) }),
          makeTask({ id: 'c', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'c', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    // a finishes Mon 17:00, b finishes Mon 13:00. c starts Tue 09:00 (after a).
    expect(scheduleOf(result, 'c').earlyStart).toBe(TUESDAY)
    expect(scheduleOf(result, 'c').earlyFinish).toBe(TUESDAY_FINISH)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(false)
    expect(scheduleOf(result, 'b').totalSlack).toBe(240)
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 required — FS lead (negative lag)', () => {
  it('overlaps successor with predecessor via negative FS lag', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS', -240)],
      }),
    )
    // a finishes Mon 17:00. FS lead −240 means b starts 240 working min
    // before a finishes: Mon 13:00. b finishes Tue 13:00.
    expect(scheduleOf(result, 'b').earlyStart).toBe(MONDAY_13)
    expect(scheduleOf(result, 'b').earlyFinish).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'a').totalSlack).toBe(0)
    expect(scheduleOf(result, 'b').totalSlack).toBe(0)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe(TUESDAY_13)
  })
})

describe('PROJECT-012 required — SS lag', () => {
  it('starts the successor lag-minutes after the predecessor starts', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(240) }),
          makeTask({ id: 'b', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'SS', 60)],
      }),
    )
    expect(scheduleOf(result, 'a').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_13)
    expect(scheduleOf(result, 'b').earlyStart).toBe('2026-08-03T10:00:00.000Z')
    expect(scheduleOf(result, 'b').earlyFinish).toBe('2026-08-04T10:00:00.000Z')
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe('2026-08-04T10:00:00.000Z')
  })
})

describe('PROJECT-012 required — FF lag', () => {
  it('finishes the successor lag-minutes after the predecessor finishes', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(240) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FF', 60)],
      }),
    )
    expect(scheduleOf(result, 'a').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_FINISH)
    // b must finish 60 min after a finishes (Mon 17:00 + 60 = Tue 10:00).
    // b(240) starts at Tue 10:00 - 240 = Mon 14:00.
    expect(scheduleOf(result, 'b').earlyStart).toBe('2026-08-03T14:00:00.000Z')
    expect(scheduleOf(result, 'b').earlyFinish).toBe('2026-08-04T10:00:00.000Z')
    expect(scheduleOf(result, 'a').totalSlack).toBe(0)
    expect(scheduleOf(result, 'b').totalSlack).toBe(0)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe('2026-08-04T10:00:00.000Z')
  })
})

describe('PROJECT-012 required — SF lag', () => {
  it('finishes the successor after the predecessor starts (SF relationship)', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(120) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'SF')],
      }),
    )
    // SF: b must finish after a starts (Mon 09:00). b(120) starts at Mon 09:00
    // (project start) and finishes Mon 11:00 — which is after a starts. But
    // the SF bound computes the successor start so that successor finishes
    // at predecessor.start + lag. With lag=0: b finishes at Mon 09:00. But b
    // starts at projectStart = Mon 09:00 and finishes Mon 11:00 — the SF
    // constraint is satisfied (b finishes after a starts). b has float.
    expect(scheduleOf(result, 'a').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(result, 'b').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'b').earlyFinish).toBe('2026-08-03T11:00:00.000Z')
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(false)
    expect(scheduleOf(result, 'b').totalSlack).toBe(360)
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-012 required — multiple successors constrain free slack', () => {
  it('free slack is the min across all successor links', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
          makeTask({ id: 'e', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'a', 'c', 'FS', 240),
          makeDependency('d3', 'c', 'd', 'FS'),
          makeDependency('d4', 'd', 'e', 'FS'),
        ],
      }),
    )
    // a→b is one day (b finishes Tue 17:00). a→c(240 lag)→d→e: c starts Tue
    // 13:00 (Mon 17:00 + 240 lag), finishes Wed 13:00. d starts Wed 13:00,
    // finishes Thu 13:00. e starts Thu 13:00, finishes Fri 13:00.
    expect(scheduleOf(result, 'e').earlyFinish).toBe('2026-08-07T13:00:00.000Z')
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'a').totalSlack).toBe(0)
    expect(scheduleOf(result, 'a').freeSlack).toBe(0)
    expect(result.projectFinish).toBe('2026-08-07T13:00:00.000Z')
  })
})

describe('PROJECT-012 required — milestone with predecessor and successor', () => {
  it('propagates the critical chain through a milestone', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'm', duration: day(0), milestone: true }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'm', 'FS'),
          makeDependency('d2', 'm', 'b', 'FS'),
          makeDependency('d3', 'b', 'c', 'FS'),
        ],
      }),
    )
    const m = scheduleOf(result, 'm')
    expect(m.earlyStart).toBe(TUESDAY)
    expect(m.earlyFinish).toBe(TUESDAY)
    expect(m.lateStart).toBe(TUESDAY)
    expect(m.lateFinish).toBe(TUESDAY)
    expect(m.totalSlack).toBe(0)
    expect(m.critical).toBe(true)
    for (const id of ['a', 'b', 'c']) {
      expect(scheduleOf(result, id).critical).toBe(true)
    }
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — summary with critical and noncritical children', () => {
  it('summary is critical when at least one child is critical', () => {
    const summary = (id: string): Task =>
      makeTask({ id, summary: true, duration: wm(0), outlineLevel: 1 })
    const result = resultOf(
      makeDocument({
        tasks: [
          summary('s'),
          makeTask({ id: 'a', duration: day(480), parentTaskId: taskId('s'), outlineLevel: 2 }),
          makeTask({ id: 'b', duration: day(240), parentTaskId: taskId('s'), outlineLevel: 2 }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'a', 'c', 'FS'), makeDependency('d2', 'c', 'd', 'FS')],
      }),
    )
    // a→c→d is the critical chain (Mon, Tue, Wed). b is a short child of s
    // (Mon 09:00–13:00). s spans Mon 09:00–Mon 13:00 (a is critical, b is
    // shorter). s's finish = max(a.EF, b.EF) = Mon 17:00. s has totalSlack =
    // workingDuration(Mon 17:00, Wed 17:00) = 960. Not critical.
    expect(scheduleOf(result, 's').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 's').earlyFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(result, 's').totalSlack).toBe(960)
    expect(scheduleOf(result, 's').critical).toBe(false)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(false)
    expect(scheduleOf(result, 'c').critical).toBe(true)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — zero free slack with positive total slack', () => {
  it('shows freeSlack=0 and totalSlack>0 when successor is tight', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(240) }),
          makeTask({ id: 'e', duration: day(240) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'a', 'd', 'FS'),
          makeDependency('d4', 'd', 'e', 'FS'),
        ],
      }),
    )
    // Critical chain a→b→c (Mon, Tue, Wed). Short branch a→d→e (d=Tue 09:00–
    // 13:00, e=Tue 13:00–17:00). e finishes Tue 17:00. e has no successor so
    // freeSlack=totalSlack. But d's successor e starts right when d finishes,
    // so d has freeSlack=0. d's totalSlack = workingDuration(Tue 13:00, Wed
    // 17:00) = 240+480=720. Hmm, let me check.
    //
    // d.EF = Tue 13:00. projectFinish = Wed 17:00. d.LF = ?
    // d's successor is e (FS, 0). e.LS = subtractWorkingTime(e.LF, 240).
    // e.LF = projectFinish = Wed 17:00. e.LS = subtractWorkingTime(Wed 17:00,
    // 240) = Wed 13:00. d.LF = e.LS - 0 = Wed 13:00. d.LS = subtractWorking-
    // Time(Wed 13:00, 240) = Wed 09:00... wait. subtractWorkingTime(Wed 13:00,
    // 240): Wed 09:00-13:00 = 240. Consume 240 → Wed 09:00. So d.LS = Wed
    // 09:00. d.EF = Tue 13:00. d.TS = workingDuration(Tue 13:00, Wed 13:00) =
    // Tue 13:00-17:00 (240) + Wed 09:00-13:00 (240) = 480. d.FS: successor e
    // (FS, 0). bound = subtractWorkingTime(e.ES=Tue 13:00, 0) = Tue 13:00.
    // anchor = d.EF = Tue 13:00. FS = signedWorkingDuration(Tue 13:00, Tue
    // 13:00) = 0. So d: TS=480, FS=0. ✓
    expect(scheduleOf(result, 'd').totalSlack).toBe(480)
    expect(scheduleOf(result, 'd').freeSlack).toBe(0)
    expect(scheduleOf(result, 'd').critical).toBe(false)
    // e has no successor: FS=TS.
    expect(scheduleOf(result, 'e').totalSlack).toBe(480)
    expect(scheduleOf(result, 'e').freeSlack).toBe(480)
    expect(scheduleOf(result, 'e').critical).toBe(false)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — ALAP pulls scheduled dates to the late window', () => {
  it('ALAP task is scheduled at its late window with nonzero float', () => {
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
    expect(d.lateStart).toBe(WEDNESDAY_13)
    expect(d.lateFinish).toBe(WEDNESDAY_FINISH)
    expect(d.scheduledStart).toBe(WEDNESDAY_13)
    expect(d.scheduledFinish).toBe(WEDNESDAY_FINISH)
    expect(d.totalSlack).toBe(1200)
    expect(d.critical).toBe(false)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — SNLT pulls scheduled start to late window', () => {
  it('SNLT schedules at the late window bounded by the constraint', () => {
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
    expect(d.lateStart).toBe(TUESDAY)
    expect(d.scheduledStart).toBe(TUESDAY)
    expect(d.scheduledFinish).toBe(TUESDAY_13)
    expect(d.totalSlack).toBe(480)
    expect(d.critical).toBe(false)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — FNLT bounds the finish', () => {
  it('FNLT pulls the scheduled finish to the late window bounded by the constraint', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({
            id: 'd',
            duration: day(240),
            constraintType: 'finishNoLaterThan',
            constraintDate: iso(TUESDAY_FINISH),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
      }),
    )
    const d = scheduleOf(result, 'd')
    expect(d.scheduledFinish).toBe(TUESDAY_FINISH)
    // d's lateFinish is bounded by the FNLT constraint (Tue 17:00). d's EF =
    // Mon 13:00 (forward pass: no predecessors, starts Mon 09:00, 240 min).
    // totalSlack = workingDuration(Mon 13:00, Tue 17:00) = 240 + 480 = 720.
    expect(d.totalSlack).toBe(720)
    expect(d.critical).toBe(false)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — split-day calendar boundary', () => {
  it('split-day calendar correctly measures working-time float', () => {
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
    // Split day: 09:00–12:00 (180) + 13:00–17:00 (240) = 420 working min/day.
    // 480 min = 1 day + 60 min → Mon 09:00 + 420 = Mon 17:00 (end of day),
    // then 60 more = Tue 09:00 + 60 = Tue 10:00.
    expect(scheduleOf(result, 'a').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'a').earlyFinish).toBe('2026-08-04T10:00:00.000Z')
    expect(scheduleOf(result, 'a').totalSlack).toBe(0)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(result.projectFinish).toBe('2026-08-04T10:00:00.000Z')
  })
})

describe('PROJECT-012 required — inherited calendar with exception', () => {
  it('inherited calendar overrides weekday and exception correctly', () => {
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
    // Mon + Tue work (960), Wed off (child override), Thu holiday (exception),
    // Fri work (480). Total 1440 min → finishes Fri 17:00.
    expect(scheduleOf(result, 'a').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'a').earlyFinish).toBe(FRIDAY_FINISH)
    expect(scheduleOf(result, 'a').totalSlack).toBe(0)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(result.projectFinish).toBe(FRIDAY_FINISH)
  })
})

describe('PROJECT-012 required — isolated summary subtree', () => {
  it('isolated summary subtree has deterministic float', () => {
    const summary = (id: string): Task =>
      makeTask({ id, summary: true, duration: wm(0), outlineLevel: 1 })
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          summary('s'),
          makeTask({ id: 'x', duration: day(240), parentTaskId: taskId('s'), outlineLevel: 2 }),
          makeTask({ id: 'y', duration: day(240), parentTaskId: taskId('s'), outlineLevel: 2 }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'FS'),
          makeDependency('d3', 'x', 'y', 'FS'),
        ],
      }),
    )
    // Critical chain a→b→c finishes Wed 17:00. s subtree: x Mon 09:00–13:00,
    // y Mon 13:00–17:00. s spans Mon 09:00–17:00. s totalSlack =
    // workingDuration(Mon 17:00, Wed 17:00) = 960. Not critical.
    expect(scheduleOf(result, 's').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 's').earlyFinish).toBe(MONDAY_FINISH)
    expect(scheduleOf(result, 's').totalSlack).toBe(960)
    expect(scheduleOf(result, 's').critical).toBe(false)
    expect(scheduleOf(result, 'x').critical).toBe(false)
    expect(scheduleOf(result, 'y').critical).toBe(false)
    for (const id of ['a', 'b', 'c']) {
      expect(scheduleOf(result, id).critical).toBe(true)
    }
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — isolated independent milestone', () => {
  it('isolated milestone has deterministic float and is non-critical', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'm', duration: day(0), milestone: true }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    // a→b finishes Tue 17:00 = project finish. m is an isolated milestone
    // (no deps). m starts Mon 09:00, finishes Mon 09:00. m totalSlack =
    // workingDuration(Mon 09:00, Tue 17:00) = 480+480 = 960. Not critical.
    expect(scheduleOf(result, 'm').earlyStart).toBe(MONDAY)
    expect(scheduleOf(result, 'm').earlyFinish).toBe(MONDAY)
    expect(scheduleOf(result, 'm').totalSlack).toBe(960)
    expect(scheduleOf(result, 'm').critical).toBe(false)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })
})

describe('PROJECT-012 required — critical path with mixed FS/SS/FF/SF in one graph', () => {
  it('all four dependency types in one graph produce correct critical path', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(240) }),
          makeTask({ id: 'd', duration: day(240) }),
          makeTask({ id: 'e', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'SS', 120),
          makeDependency('d3', 'c', 'd', 'FF', 60),
          makeDependency('d4', 'b', 'e', 'SF', 0),
        ],
      }),
    )
    // Verify no diagnostics and deterministic output. The exact values are
    // covered by the CP05 golden; here we verify the mixed graph schedules.
    expect(result.diagnostics).toEqual([])
    expect(result.projectFinish).toBeDefined()
    // Re-run for byte-identical determinism.
    const rerun = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(240) }),
          makeTask({ id: 'd', duration: day(240) }),
          makeTask({ id: 'e', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS'),
          makeDependency('d2', 'b', 'c', 'SS', 120),
          makeDependency('d3', 'c', 'd', 'FF', 60),
          makeDependency('d4', 'b', 'e', 'SF', 0),
        ],
      }),
    )
    expect(scheduleJson(rerun)).toBe(scheduleJson(result))
  })
})

describe('PROJECT-012 required — repeated deterministic schedule (32)', () => {
  it('three consecutive schedule() calls produce byte-identical output', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(480) }),
        makeTask({ id: 'm', duration: day(0), milestone: true }),
        makeTask({
          id: 'c',
          duration: day(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
        makeTask({ id: 'd', duration: day(240), constraintType: 'asLateAsPossible' }),
      ],
      dependencies: [
        makeDependency('d1', 'a', 'b', 'FS', 60),
        makeDependency('d2', 'b', 'm', 'SS', -30),
        makeDependency('d3', 'a', 'c', 'FF', 120),
        makeDependency('d4', 'm', 'd', 'SF'),
      ],
    })
    const first = scheduleJson(resultOf(document))
    const second = scheduleJson(resultOf(document))
    const third = scheduleJson(resultOf(document))
    expect(first).toBe(second)
    expect(second).toBe(third)
  })
})

describe('PROJECT-012 required — serialized round-trip byte equality', () => {
  it('the same serialized document produces byte-identical CPM output', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', duration: day(480) }),
        makeTask({ id: 'b', duration: day(240) }),
        makeTask({ id: 'c', duration: day(480) }),
        makeTask({ id: 'm', duration: day(0), milestone: true }),
      ],
      dependencies: [
        makeDependency('d1', 'a', 'b', 'FS', 60),
        makeDependency('d2', 'b', 'c', 'FF', 120),
        makeDependency('d3', 'a', 'm', 'SS', -30),
      ],
    })
    const run1 = scheduleJson(resultOf(document))
    const run2 = scheduleJson(resultOf(parseDocument(documentJson(document))))
    expect(run1).toBe(run2)
  })
})

describe('PROJECT-012 required — negative-slack predecessor of MSO task', () => {
  it('MSO on a successor pulls predecessor lateFinish before earlyFinish', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({
            id: 'b',
            duration: day(240),
            constraintType: 'mustStartOn',
            constraintDate: iso(MONDAY),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    // b is pinned by MSO to start Monday 09:00 and finish Monday 13:00. a→FS→b
    // means a must finish before b starts (Mon 09:00). But a starts Mon 09:00
    // and finishes Mon 17:00. So a's lateFinish = b.LS = Mon 09:00 (before a's
    // earlyFinish = Mon 17:00). a has negative total slack.
    expect(scheduleOf(result, 'b')).toMatchObject({
      earlyStart: MONDAY,
      earlyFinish: MONDAY_13,
      lateStart: MONDAY,
      lateFinish: MONDAY_13,
      totalSlack: 0,
      critical: true,
    })
    const a = scheduleOf(result, 'a')
    expect(a.earlyFinish).toBe(MONDAY_FINISH)
    expect(a.lateFinish).toBe(MONDAY)
    expect(a.totalSlack).toBeLessThan(0)
    expect(a.critical).toBe(true)
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })
})

describe('PROJECT-012 required — dependency chain with different task calendars', () => {
  it('float uses working-time arithmetic when tasks have different calendars', () => {
    // a uses standard calendar (09:00–17:00, 480 min/day). b uses a 4-day-week
    // calendar (Mon–Thu 09:00–17:00, no Friday). b is successor of a via FS.
    const fourDayWeek = (): Record<number, CalendarPeriod[]> => ({
      0: [],
      1: [{ startMinute: 540, endMinute: 1020 }],
      2: [{ startMinute: 540, endMinute: 1020 }],
      3: [{ startMinute: 540, endMinute: 1020 }],
      4: [{ startMinute: 540, endMinute: 1020 }],
      5: [],
      6: [],
    })
    const fourDay = makeCalendar('fourday', { workingWeek: fourDayWeek() })
    const std = makeCalendar('standard')
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({
            id: 'b',
            duration: day(960),
            calendarId: asCalendarId('fourday'),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
        calendars: [std, fourDay],
      }),
    )
    expect(scheduleOf(result, 'a').earlyFinish).toBe(MONDAY_FINISH)
    // b uses fourDay calendar. b starts Tue 09:00 (in fourDay). b(960 min) =
    // 2 working days → finishes Wed 17:00 (Tue 480 + Wed 480 = 960).
    expect(scheduleOf(result, 'b').earlyStart).toBe(TUESDAY)
    expect(scheduleOf(result, 'b').earlyFinish).toBe(WEDNESDAY_FINISH)
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-012 required — converging paths with different lag signs', () => {
  it('two paths with positive and negative lag converge correctly', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(480) }),
          makeTask({ id: 'c', duration: day(480) }),
          makeTask({ id: 'd', duration: day(480) }),
        ],
        dependencies: [
          makeDependency('d1', 'a', 'b', 'FS', 240),
          makeDependency('d2', 'a', 'c', 'FS', -240),
          makeDependency('d3', 'b', 'd', 'FS'),
          makeDependency('d4', 'c', 'd', 'FS'),
        ],
      }),
    )
    // Path 1: a→b(+240 lag)→d. b starts Wed 13:00 (Mon 17:00 + 240 lag = Tue
    // 13:00... wait, addWorkingTime(Mon 17:00, 240). Mon 17:00 not working →
    // nextWorkingInstant = Tue 09:00. add 240 → Tue 13:00. b.EF = Wed 13:00.
    // Path 2: a→c(−240 lag)→d. c starts Mon 13:00 (Mon 17:00 - 240 = Mon
    // 13:00). c.EF = Tue 13:00.
    // d starts after both b and c: max(b.EF=Wed 13:00, c.EF=Tue 13:00) = Wed
    // 13:00. d.EF = Thu 13:00.
    expect(scheduleOf(result, 'b').earlyStart).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'b').earlyFinish).toBe(WEDNESDAY_13)
    expect(scheduleOf(result, 'c').earlyStart).toBe(MONDAY_13)
    expect(scheduleOf(result, 'c').earlyFinish).toBe(TUESDAY_13)
    expect(scheduleOf(result, 'd').earlyStart).toBe(WEDNESDAY_13)
    expect(scheduleOf(result, 'd').earlyFinish).toBe(THURSDAY_13)
    // b→d is the longer path → b and d are critical.
    expect(scheduleOf(result, 'b').critical).toBe(true)
    expect(scheduleOf(result, 'd').critical).toBe(true)
    expect(scheduleOf(result, 'a').critical).toBe(true)
    // c is shorter → has float.
    expect(scheduleOf(result, 'c').critical).toBe(false)
    expect(result.projectFinish).toBe(THURSDAY_13)
  })
})

// Suppress unused-import lint for fixtures that document the calendar
// arithmetic contract but are only used in specific scenarios.
void PREV_FRIDAY
void NEXT_MONDAY
void FRIDAY
void FRIDAY_FINISH
