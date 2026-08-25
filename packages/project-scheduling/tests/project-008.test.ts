import { describe, expect, it } from 'vitest'
import { schedule } from '../src/index.js'
import type { DerivedSchedule, ProjectDocument, TaskSchedule } from '@genoffice/project-contracts'
import { asISODateTime } from '@genoffice/project-contracts'
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
  iso,
  makeDependency,
  makeDocument,
  makeTask,
  parseDocument,
  taskId,
  wm,
} from './fixtures.js'

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

const withStatusDate = (document: ProjectDocument, statusDate: string): ProjectDocument => ({
  ...document,
  properties: { ...document.properties, statusDate: asISODateTime(statusDate) },
})

// A three-task FS chain a -> b -> c whose project finish anchors the late
// window for an unconstrained fourth task. Reused by the constraint scenarios
// so they can be compared against the accepted PROJECT-006 goldens.
const chainWith = (extra: ReturnType<typeof makeTask>[]): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({ id: 'a', duration: day(480) }),
      makeTask({ id: 'b', duration: day(480) }),
      makeTask({ id: 'c', duration: day(480) }),
      ...extra,
    ],
    dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'c', 'FS')],
  })

describe('PROJECT-008 constraints / deadlines / progress', () => {
  // ----- The canonical eight constraint types -----

  it('C01 ASAP: schedules from the project start with no constraint effect', () => {
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480), constraintType: 'asSoonAsPossible' })],
      }),
    )
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
    // ASAP carries no constraintDate in the canonical document model.
    expect(a.deadline).toBeUndefined()
    expect(a.deadlineVariance).toBeUndefined()
    expect(a.deadlineMissed).toBeUndefined()
    expect(a.status).toBe('notStarted')
    expect(a.percentComplete).toBe(0)
    expect(a.actualDuration).toBe(0)
    expect(a.remainingDuration).toBe(480)
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })

  it('C02 ALAP: pulls the scheduled window to the late dates', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({ id: 'b', duration: day(240), constraintType: 'asLateAsPossible' }),
        ],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(a.totalSlack).toBe(0)
    expect(a.critical).toBe(true)
    // b's late window is bounded by the project finish (a's finish).
    expect(b.earlyStart).toBe(MONDAY)
    expect(b.earlyFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(b.lateStart).toBe('2026-08-03T13:00:00.000Z')
    expect(b.lateFinish).toBe(MONDAY_FINISH)
    expect(b.scheduledStart).toBe('2026-08-03T13:00:00.000Z')
    expect(b.scheduledFinish).toBe(MONDAY_FINISH)
    expect(b.totalSlack).toBe(240)
    expect(b.freeSlack).toBe(240)
    expect(b.critical).toBe(false)
  })

  it('C03 SNET: pushes the early start later but leaves slack intact', () => {
    const result = resultOf(
      chainWith([
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ]),
    )
    const d = scheduleOf(result, 'd')
    expect(d.earlyStart).toBe(TUESDAY)
    expect(d.earlyFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.lateStart).toBe('2026-08-05T13:00:00.000Z')
    expect(d.lateFinish).toBe(WEDNESDAY_FINISH)
    expect(d.scheduledStart).toBe(TUESDAY)
    expect(d.scheduledFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.totalSlack).toBe(720)
    expect(d.freeSlack).toBe(720)
    expect(d.critical).toBe(false)
    for (const id of ['a', 'b', 'c']) expect(scheduleOf(result, id).critical).toBe(true)
  })

  it('C04 SNLT: pulls the scheduled start to the late window bounded by the constraint', () => {
    const result = resultOf(
      chainWith([
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'startNoLaterThan',
          constraintDate: iso(TUESDAY),
        }),
      ]),
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

  it('C05 MSO: pins start and late dates so the task becomes critical (MSO differs from SNET)', () => {
    const result = resultOf(
      chainWith([
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'mustStartOn',
          constraintDate: iso(TUESDAY),
        }),
      ]),
    )
    const d = scheduleOf(result, 'd')
    expect(d.earlyStart).toBe(TUESDAY)
    expect(d.earlyFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledStart).toBe(TUESDAY)
    expect(d.lateStart).toBe(TUESDAY)
    expect(d.lateFinish).toBe('2026-08-04T13:00:00.000Z')
    expect(d.totalSlack).toBe(0)
    expect(d.critical).toBe(true)
  })

  it('C06 FNET: pushes the finish later while preserving slack', () => {
    const result = resultOf(
      chainWith([
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'finishNoEarlierThan',
          constraintDate: iso(TUESDAY_FINISH),
        }),
      ]),
    )
    const d = scheduleOf(result, 'd')
    expect(d.scheduledStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledFinish).toBe(TUESDAY_FINISH)
    expect(d.totalSlack).toBe(480)
    expect(d.critical).toBe(false)
  })

  it('C07 FNLT: pulls the scheduled finish to the late window bounded by the constraint', () => {
    const result = resultOf(
      chainWith([
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'finishNoLaterThan',
          constraintDate: iso(TUESDAY_FINISH),
        }),
      ]),
    )
    const d = scheduleOf(result, 'd')
    expect(d.earlyStart).toBe(MONDAY)
    expect(d.earlyFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(d.lateStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.lateFinish).toBe(TUESDAY_FINISH)
    expect(d.scheduledStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledFinish).toBe(TUESDAY_FINISH)
    expect(d.totalSlack).toBe(720)
    expect(d.critical).toBe(false)
  })

  it('C08 MFO: pins the finish and makes the task critical (MFO differs from FNET)', () => {
    const result = resultOf(
      chainWith([
        makeTask({
          id: 'd',
          duration: day(240),
          constraintType: 'mustFinishOn',
          constraintDate: iso(TUESDAY_FINISH),
        }),
      ]),
    )
    const d = scheduleOf(result, 'd')
    expect(d.scheduledStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.scheduledFinish).toBe(TUESDAY_FINISH)
    expect(d.lateFinish).toBe(TUESDAY_FINISH)
    expect(d.lateStart).toBe('2026-08-04T13:00:00.000Z')
    expect(d.totalSlack).toBe(0)
    expect(d.critical).toBe(true)
  })

  // ----- Conflicting dependency + constraint cases -----

  it('C09 conflicting start dependency + MSO: MSO honors its date even when a finish-to-start dependency would start the task later', () => {
    // Predecessor a finishes Monday 17:00; FS would start b on Tuesday 09:00.
    // MSO pins b's start to Monday 09:00, so b begins BEFORE a finishes. The
    // engine resolves this deterministically: MSO wins over the dependency,
    // and the resulting schedule surfaces the conflict — b's scheduled start
    // precedes its FS predecessor's finish — without inventing negative slack
    // (MSO pins both early and late windows to the constraint, so total slack
    // is zero and the task is critical).
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
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(b.earlyStart).toBe(MONDAY)
    expect(b.earlyFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(b.lateStart).toBe(MONDAY)
    expect(b.lateFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(b.scheduledStart).toBe(MONDAY)
    expect(b.scheduledFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(b.totalSlack).toBe(0)
    expect(b.critical).toBe(true)
    // The conflict: b starts before its FS predecessor a finishes.
    expect(new Date(b.scheduledStart!).getTime() < new Date(a.scheduledFinish!).getTime()).toBe(
      true,
    )
  })

  it('C10 conflicting finish dependency + MFO: MFO honors its finish date even when a finish-to-finish dependency would finish the task later', () => {
    // Predecessor a finishes Monday 17:00; FF would finish b at Tuesday 09:00.
    // MFO pins b's finish to Monday 13:00, so b finishes before a — a conflict
    // the engine resolves deterministically by honoring MFO.
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', duration: day(480) }),
          makeTask({
            id: 'b',
            duration: day(240),
            constraintType: 'mustFinishOn',
            constraintDate: iso('2026-08-03T13:00:00.000Z'),
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FF')],
      }),
    )
    const a = scheduleOf(result, 'a')
    const b = scheduleOf(result, 'b')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(b.scheduledStart).toBe(MONDAY)
    expect(b.scheduledFinish).toBe('2026-08-03T13:00:00.000Z')
    expect(b.totalSlack).toBe(0)
    expect(b.critical).toBe(true)
  })

  // ----- Deadlines -----

  it('C11 deadline after finish: positive variance, not missed', () => {
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480), deadline: iso(WEDNESDAY) })],
      }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(a.deadline).toBe(WEDNESDAY)
    expect(a.deadlineVariance).toBe(480)
    expect(a.deadlineMissed).toBe(false)
  })

  it('C12 deadline exactly at finish: zero variance, not missed', () => {
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480), deadline: iso(MONDAY_FINISH) })],
      }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.deadline).toBe(MONDAY_FINISH)
    expect(a.deadlineVariance).toBe(0)
    expect(a.deadlineMissed).toBe(false)
  })

  it('C13 missed deadline: negative variance, missed', () => {
    const result = resultOf(
      makeDocument({
        tasks: [makeTask({ id: 'a', duration: day(480), deadline: iso(MONDAY) })],
      }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(a.deadline).toBe(MONDAY)
    expect(a.deadlineVariance).toBe(-480)
    expect(a.deadlineMissed).toBe(true)
    // A deadline never moves the task: scheduled finish is unchanged.
    expect(a.scheduledStart).toBe(MONDAY)
  })

  it('C14 milestone constraint: MSO pins a zero-duration task', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({
            id: 'm',
            duration: day(0),
            milestone: true,
            constraintType: 'mustStartOn',
            constraintDate: iso(TUESDAY),
          }),
        ],
      }),
    )
    const m = scheduleOf(result, 'm')
    expect(m.earlyStart).toBe(TUESDAY)
    expect(m.earlyFinish).toBe(TUESDAY)
    expect(m.lateStart).toBe(TUESDAY)
    expect(m.lateFinish).toBe(TUESDAY)
    expect(m.scheduledStart).toBe(TUESDAY)
    expect(m.scheduledFinish).toBe(TUESDAY)
    expect(m.duration).toBe(0)
    expect(m.totalSlack).toBe(0)
    expect(m.critical).toBe(true)
    expect(m.status).toBe('notStarted')
    expect(m.percentComplete).toBe(0)
    expect(m.actualDuration).toBe(0)
    expect(m.remainingDuration).toBe(0)
  })

  it('C15 zero-duration deadline: a milestone deadline is missed when finish is after it', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({
            id: 'm',
            duration: day(0),
            milestone: true,
            constraintType: 'mustStartOn',
            constraintDate: iso(TUESDAY),
            deadline: iso(MONDAY),
          }),
        ],
      }),
    )
    const m = scheduleOf(result, 'm')
    expect(m.scheduledFinish).toBe(TUESDAY)
    expect(m.deadline).toBe(MONDAY)
    expect(m.deadlineVariance).toBe(-480)
    expect(m.deadlineMissed).toBe(true)
  })

  // ----- Progress model -----

  it('C16 percentComplete 0: not started, full remaining duration', () => {
    const result = resultOf(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })] }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.percentComplete).toBe(0)
    expect(a.status).toBe('notStarted')
    expect(a.actualDuration).toBe(0)
    expect(a.remainingDuration).toBe(480)
  })

  it('C17 partial percentComplete: in progress, split actual/remaining', () => {
    const result = resultOf(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 50 })] }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.percentComplete).toBe(50)
    expect(a.status).toBe('inProgress')
    expect(a.actualDuration).toBe(240)
    expect(a.remainingDuration).toBe(240)
  })

  it('C18 percentComplete 100: complete, no remaining duration even when finish is in the future', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            duration: day(480),
            percentComplete: 100,
            constraintType: 'startNoEarlierThan',
            constraintDate: iso(NEXT_MONDAY),
          }),
        ],
      }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.scheduledStart).toBe(NEXT_MONDAY)
    expect(a.scheduledFinish).toBe('2026-08-10T17:00:00.000Z')
    expect(a.percentComplete).toBe(100)
    expect(a.status).toBe('complete')
    expect(a.actualDuration).toBe(480)
    expect(a.remainingDuration).toBe(0)
  })

  it('C19 physicalPercentComplete: echoed for leaf tasks and independent of percentComplete', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            duration: day(480),
            percentComplete: 40,
            physicalPercentComplete: 60,
          }),
        ],
      }),
    )
    const a = scheduleOf(result, 'a')
    expect(a.percentComplete).toBe(40)
    expect(a.physicalPercentComplete).toBe(60)
    expect(a.status).toBe('inProgress')
    expect(a.actualDuration).toBe(192)
    expect(a.remainingDuration).toBe(288)
  })

  it('C20 invalid progress is rejected by the scheduler with diagnostics and no task schedules', () => {
    const reject = (document: ProjectDocument, code: string) => {
      const result = schedule(document)
      expect(result.taskSchedules).toEqual({})
      expect(result.diagnostics.some((d) => d.code === code)).toBe(true)
      expect(result.diagnostics.every((d) => d.severity === 'error')).toBe(true)
    }
    reject(
      makeDocument({ tasks: [makeTask({ id: 'a', percentComplete: 150 })] }),
      'INVALID_PERCENT_COMPLETE',
    )
    reject(
      makeDocument({ tasks: [makeTask({ id: 'a', percentComplete: -5 })] }),
      'INVALID_PERCENT_COMPLETE',
    )
    reject(
      makeDocument({ tasks: [makeTask({ id: 'a', physicalPercentComplete: 101 })] }),
      'INVALID_PERCENT_COMPLETE',
    )
    // A dated constraint without a constraintDate is rejected too.
    reject(
      makeDocument({
        tasks: [makeTask({ id: 'a', constraintType: 'mustStartOn' })],
      }),
      'MISSING_CONSTRAINT_DATE',
    )
    // ASAP must not carry a constraintDate.
    reject(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            constraintType: 'asSoonAsPossible',
            constraintDate: iso(MONDAY),
          }),
        ],
      }),
      'CONSTRAINT_DATE_NOT_ALLOWED',
    )
  })

  // ----- Status date semantics (deterministic, no wall-clock) -----

  it('C21 status date before task start: 0% task is not started', () => {
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })] }),
      '2026-08-02T09:00:00.000Z', // Sunday, before Monday start
    )
    const a = scheduleOf(resultOf(document), 'a')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.status).toBe('notStarted')
    expect(a.percentComplete).toBe(0)
  })

  it('C22 status date during the task window: 0% task is in progress', () => {
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(960), percentComplete: 0 })] }),
      '2026-08-04T12:00:00.000Z', // Tuesday midday, inside Mon 09:00 .. Tue 17:00
    )
    const a = scheduleOf(resultOf(document), 'a')
    expect(a.scheduledStart).toBe(MONDAY)
    expect(a.scheduledFinish).toBe(TUESDAY_FINISH)
    expect(a.status).toBe('inProgress')
    expect(a.percentComplete).toBe(0)
  })

  it('C23 status date after task finish: 0% task is in progress (overdue-incomplete)', () => {
    const document = withStatusDate(
      makeDocument({ tasks: [makeTask({ id: 'a', duration: day(480), percentComplete: 0 })] }),
      WEDNESDAY, // after Monday finish
    )
    const a = scheduleOf(resultOf(document), 'a')
    expect(a.scheduledFinish).toBe(MONDAY_FINISH)
    expect(a.status).toBe('inProgress')
    expect(a.percentComplete).toBe(0)
  })

  // ----- Determinism -----

  it('C24 repeated and reordered scheduling with constraints/deadlines/progress is byte-identical', () => {
    const summary = (id: string, parentTaskId?: string, outlineLevel = 1) =>
      makeTask({
        id,
        summary: true,
        duration: wm(0),
        parentTaskId: parentTaskId === undefined ? undefined : taskId(parentTaskId),
        outlineLevel,
      })
    const document = withStatusDate(
      makeDocument({
        tasks: [
          summary('s'),
          makeTask({
            id: 'a',
            duration: day(480),
            parentTaskId: taskId('s'),
            outlineLevel: 2,
            percentComplete: 100,
            deadline: iso(TUESDAY_FINISH),
          }),
          makeTask({
            id: 'b',
            duration: day(480),
            parentTaskId: taskId('s'),
            outlineLevel: 2,
            constraintType: 'startNoEarlierThan',
            constraintDate: iso(TUESDAY),
            percentComplete: 25,
            deadline: iso(WEDNESDAY_FINISH),
          }),
          makeTask({
            id: 'c',
            duration: day(0),
            milestone: true,
            constraintType: 'mustFinishOn',
            constraintDate: iso(WEDNESDAY_FINISH),
            deadline: iso(THURSDAY_FINISH),
          }),
          makeTask({
            id: 'd',
            duration: day(240),
            constraintType: 'asLateAsPossible',
            percentComplete: 50,
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
      TUESDAY,
    )

    const serialized = JSON.stringify(document)
    const first = schedule(parseDocument(serialized))
    const second = schedule(parseDocument(serialized))
    const third = schedule(parseDocument(serialized))
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(JSON.stringify(second)).toBe(JSON.stringify(third))
    expect(first.diagnostics).toEqual([])

    // Reordering input arrays must not change the output bytes either.
    const shuffled = parseDocument(serialized)
    shuffled.tasks = [...shuffled.tasks].reverse()
    shuffled.dependencies = [...shuffled.dependencies].reverse()
    const reordered = schedule(shuffled)
    expect(JSON.stringify(reordered)).toBe(JSON.stringify(first))

    // Golden snapshot of the derived progress/deadline state for the summary.
    const s = scheduleOf(first, 's')
    // Summary progress is duration-weighted: a=100% (480 done/480), b=25% (120
    // done/480). actual = 480 + 120 = 600; remaining = 0 + 360 = 360; total =
    // 960; percent = round(600/960*100) = 63 (rounded).
    expect(s.status).toBe('inProgress')
    expect(s.percentComplete).toBe(63)
    expect(s.actualDuration).toBe(600)
    expect(s.remainingDuration).toBe(360)
    expect(s.physicalPercentComplete).toBeUndefined()
    expect(scheduleOf(first, 'a').status).toBe('complete')
    expect(scheduleOf(first, 'b').status).toBe('inProgress')
    expect(scheduleOf(first, 'c').status).toBe('notStarted')
    expect(scheduleOf(first, 'd').status).toBe('inProgress')
    // Deadline variance present on the milestone (finish == deadline date).
    expect(scheduleOf(first, 'a').deadlineMissed).toBe(false)
  })

  // ----- Extra golden: summary progress aggregation precedence -----

  it('C25 summary progress rolls up duration-weighted and derives status from children', () => {
    const summary = (id: string) =>
      makeTask({ id, summary: true, duration: wm(0), outlineLevel: 1 })
    const result = resultOf(
      makeDocument({
        tasks: [
          summary('s'),
          makeTask({
            id: 'a',
            duration: day(480),
            parentTaskId: taskId('s'),
            outlineLevel: 2,
            percentComplete: 100,
          }),
          makeTask({
            id: 'b',
            duration: day(480),
            parentTaskId: taskId('s'),
            outlineLevel: 2,
            percentComplete: 0,
          }),
        ],
      }),
    )
    const s = scheduleOf(result, 's')
    // One child complete (480 done), one not started (0 done / 480 remaining).
    // actual=480, remaining=480, total=960, percent=round(480/960*100)=50.
    expect(s.percentComplete).toBe(50)
    expect(s.actualDuration).toBe(480)
    expect(s.remainingDuration).toBe(480)
    expect(s.status).toBe('inProgress')
    // A summary's stored percentComplete is never authoritative; the derived
    // value reflects the subtree only.
    expect(scheduleOf(result, 'a').status).toBe('complete')
    expect(scheduleOf(result, 'b').status).toBe('notStarted')
  })

  it('C26 milestone 100% is complete and incomplete milestone is not started (binary)', () => {
    const result = resultOf(
      makeDocument({
        tasks: [
          makeTask({ id: 'done', duration: day(0), milestone: true, percentComplete: 100 }),
          makeTask({ id: 'pending', duration: day(0), milestone: true, percentComplete: 0 }),
        ],
      }),
    )
    expect(scheduleOf(result, 'done').status).toBe('complete')
    expect(scheduleOf(result, 'done').actualDuration).toBe(0)
    expect(scheduleOf(result, 'done').remainingDuration).toBe(0)
    expect(scheduleOf(result, 'pending').status).toBe('notStarted')
    expect(scheduleOf(result, 'pending').actualDuration).toBe(0)
    expect(scheduleOf(result, 'pending').remainingDuration).toBe(0)
  })
})
