import { describe, expect, it } from 'vitest'
import type { DerivedSchedule, ProjectCommand, ProjectDocument } from '@genoffice/project-contracts'
import { applyProjectCommand, canonicalizeDocument } from '@genoffice/project-engine'
import { schedule } from '../src/index.js'
import {
  MONDAY,
  MONDAY_FINISH,
  TUESDAY,
  TUESDAY_FINISH,
  WEDNESDAY,
  WEDNESDAY_FINISH,
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

const runCommand = (document: ProjectDocument, command: ProjectCommand): ProjectDocument => {
  const execution = applyProjectCommand(document, command)
  if (!execution.result.accepted) {
    throw new Error(`command rejected: ${JSON.stringify(execution.result.diagnostics)}`)
  }
  return execution.document
}

const runAll = (document: ProjectDocument, commands: ProjectCommand[]): ProjectDocument =>
  commands.reduce((current, command) => runCommand(current, command), document)

interface GoldenExpectation {
  wbs?: string
  outlineLevel?: number
  summary?: boolean
  parentTaskId?: string | null
  scheduledStart?: string
  scheduledFinish?: string
  earlyStart?: string
  earlyFinish?: string
  lateStart?: string
  lateFinish?: string
  totalSlack?: number
  freeSlack?: number
  critical?: boolean
  duration?: number
}

const expectGolden = (
  document: ProjectDocument,
  result: DerivedSchedule,
  id: string,
  expected: GoldenExpectation,
): void => {
  const task = document.tasks.find((candidate) => (candidate.id as string) === id)
  if (!task) throw new Error(`missing task ${id}`)
  const entry = result.taskSchedules[taskId(id)]
  if (!entry) throw new Error(`missing schedule for ${id}`)
  expect(entry.taskId).toBe(taskId(id))
  if (expected.wbs !== undefined) expect(task.wbs).toBe(expected.wbs)
  if (expected.outlineLevel !== undefined) expect(task.outlineLevel).toBe(expected.outlineLevel)
  if (expected.summary !== undefined) expect(task.summary).toBe(expected.summary)
  if (expected.parentTaskId !== undefined) {
    expect(task.parentTaskId ?? null).toBe(
      expected.parentTaskId === null ? null : taskId(expected.parentTaskId),
    )
  }
  if (expected.scheduledStart !== undefined)
    expect(entry.scheduledStart).toBe(expected.scheduledStart)
  if (expected.scheduledFinish !== undefined) {
    expect(entry.scheduledFinish).toBe(expected.scheduledFinish)
  }
  if (expected.earlyStart !== undefined) expect(entry.earlyStart).toBe(expected.earlyStart)
  if (expected.earlyFinish !== undefined) expect(entry.earlyFinish).toBe(expected.earlyFinish)
  if (expected.lateStart !== undefined) expect(entry.lateStart).toBe(expected.lateStart)
  if (expected.lateFinish !== undefined) expect(entry.lateFinish).toBe(expected.lateFinish)
  if (expected.totalSlack !== undefined) expect(entry.totalSlack).toBe(expected.totalSlack)
  if (expected.freeSlack !== undefined) expect(entry.freeSlack).toBe(expected.freeSlack)
  if (expected.critical !== undefined) expect(entry.critical).toBe(expected.critical)
  if (expected.duration !== undefined) expect(entry.duration).toBe(expected.duration)
}

describe('PROJECT-007 hierarchy scheduling goldens', () => {
  it('P01: flat task list derives WBS 1..n and schedules each task from the project start', () => {
    const document = runAll(makeDocument(), [
      { type: 'CreateTask', task: makeTask({ id: 'a', uid: 1, duration: day(480) }) },
      { type: 'CreateTask', task: makeTask({ id: 'b', uid: 2, duration: day(480) }) },
      { type: 'CreateTask', task: makeTask({ id: 'c', uid: 3, duration: day(480) }) },
    ])
    const result = resultOf(document)
    for (const id of ['a', 'b', 'c']) {
      expectGolden(document, result, id, {
        wbs: id === 'a' ? '1' : id === 'b' ? '2' : '3',
        outlineLevel: 1,
        summary: false,
        parentTaskId: null,
        scheduledStart: MONDAY,
        scheduledFinish: MONDAY_FINISH,
        earlyStart: MONDAY,
        earlyFinish: MONDAY_FINISH,
        lateStart: MONDAY,
        lateFinish: MONDAY_FINISH,
        totalSlack: 0,
        freeSlack: 0,
        critical: true,
        duration: 480,
      })
    }
    expect(result.projectStart).toBe(MONDAY)
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })

  it('P02: one-level WBS rolls the summary up from its children', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'a',
            uid: 2,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'b',
            uid: 3,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    const result = resultOf(document)
    expectGolden(document, result, 'S', {
      wbs: '1',
      outlineLevel: 1,
      summary: true,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      earlyStart: MONDAY,
      earlyFinish: TUESDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'a', {
      wbs: '1.1',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'S',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(document, result, 'b', {
      wbs: '1.2',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'S',
      scheduledStart: TUESDAY,
      scheduledFinish: TUESDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })

  it('P03: two-level WBS derives nested codes and outline levels', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S1', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'S2',
            uid: 2,
            summary: true,
            duration: day(0),
            parentTaskId: taskId('S1'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'a',
            uid: 3,
            duration: day(480),
            parentTaskId: taskId('S2'),
            outlineLevel: 3,
          }),
          makeTask({
            id: 'b',
            uid: 4,
            duration: day(240),
            parentTaskId: taskId('S1'),
            outlineLevel: 2,
          }),
        ],
      }),
    )
    const result = resultOf(document)
    expectGolden(document, result, 'S1', {
      wbs: '1',
      outlineLevel: 1,
      summary: true,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 480,
    })
    expectGolden(document, result, 'S2', {
      wbs: '1.1',
      outlineLevel: 2,
      summary: true,
      parentTaskId: 'S1',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 480,
    })
    expectGolden(document, result, 'a', {
      wbs: '1.1.1',
      outlineLevel: 3,
      summary: false,
      parentTaskId: 'S2',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(document, result, 'b', {
      wbs: '1.2',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'S1',
      scheduledStart: MONDAY,
      scheduledFinish: '2026-08-03T13:00:00.000Z',
      earlyStart: MONDAY,
      earlyFinish: '2026-08-03T13:00:00.000Z',
      lateStart: '2026-08-03T13:00:00.000Z',
      lateFinish: MONDAY_FINISH,
      totalSlack: 240,
      freeSlack: 240,
      critical: false,
      duration: 240,
    })
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })

  it('P04: three-level nested WBS rolls summaries up level by level', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S1', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'S2',
            uid: 2,
            summary: true,
            duration: day(0),
            parentTaskId: taskId('S1'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'S3',
            uid: 3,
            summary: true,
            duration: day(0),
            parentTaskId: taskId('S2'),
            outlineLevel: 3,
          }),
          makeTask({
            id: 'a',
            uid: 4,
            duration: day(960),
            parentTaskId: taskId('S3'),
            outlineLevel: 4,
          }),
          makeTask({
            id: 'b',
            uid: 5,
            duration: day(240),
            parentTaskId: taskId('S2'),
            outlineLevel: 3,
          }),
          makeTask({
            id: 'c',
            uid: 6,
            duration: day(480),
            parentTaskId: taskId('S1'),
            outlineLevel: 2,
          }),
        ],
      }),
    )
    const result = resultOf(document)
    expectGolden(document, result, 'S1', {
      wbs: '1',
      outlineLevel: 1,
      summary: true,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'S2', {
      wbs: '1.1',
      outlineLevel: 2,
      summary: true,
      parentTaskId: 'S1',
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'S3', {
      wbs: '1.1.1',
      outlineLevel: 3,
      summary: true,
      parentTaskId: 'S2',
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'a', {
      wbs: '1.1.1.1',
      outlineLevel: 4,
      summary: false,
      parentTaskId: 'S3',
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      earlyStart: MONDAY,
      earlyFinish: TUESDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'b', {
      wbs: '1.1.2',
      outlineLevel: 3,
      summary: false,
      parentTaskId: 'S2',
      scheduledStart: MONDAY,
      scheduledFinish: '2026-08-03T13:00:00.000Z',
      earlyStart: MONDAY,
      earlyFinish: '2026-08-03T13:00:00.000Z',
      lateStart: '2026-08-04T13:00:00.000Z',
      lateFinish: TUESDAY_FINISH,
      totalSlack: 720,
      freeSlack: 720,
      critical: false,
      duration: 240,
    })
    expectGolden(document, result, 'c', {
      wbs: '1.2',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'S1',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 480,
      freeSlack: 480,
      critical: false,
      duration: 480,
    })
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })

  it('P05: sibling insertion extends the summary roll-up and renumbers WBS', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'a',
            uid: 2,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'b',
            uid: 3,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    const before = resultOf(base)
    expect(before.projectFinish).toBe(TUESDAY_FINISH)

    const document = runCommand(base, {
      type: 'CreateTask',
      task: makeTask({ id: 'c', uid: 4, duration: day(1440), parentTaskId: taskId('S') }),
    })
    const result = resultOf(document)
    expectGolden(document, result, 'c', {
      wbs: '1.3',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'S',
      scheduledStart: MONDAY,
      scheduledFinish: WEDNESDAY_FINISH,
      earlyStart: MONDAY,
      earlyFinish: WEDNESDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
      duration: 1440,
    })
    expectGolden(document, result, 'S', {
      wbs: '1',
      summary: true,
      scheduledStart: MONDAY,
      scheduledFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 1440,
    })
    expectGolden(document, result, 'b', {
      wbs: '1.2',
      scheduledStart: TUESDAY,
      scheduledFinish: TUESDAY_FINISH,
      lateStart: WEDNESDAY,
      lateFinish: WEDNESDAY_FINISH,
      totalSlack: 480,
      critical: false,
    })
    expectGolden(document, result, 'a', {
      wbs: '1.1',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: '2026-08-05T09:00:00.000Z',
      totalSlack: 480,
      freeSlack: 0,
      critical: false,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('P05b: child deletion shrinks the summary roll-up and removes referencing dependencies', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'a',
            uid: 2,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'b',
            uid: 3,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'c',
            uid: 4,
            duration: day(1440),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    const document = runCommand(base, { type: 'DeleteTask', taskId: taskId('b') })
    expect(document.dependencies).toHaveLength(0)
    const result = resultOf(document)
    expectGolden(document, result, 'a', {
      wbs: '1.1',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: WEDNESDAY,
      lateFinish: WEDNESDAY_FINISH,
      totalSlack: 960,
      freeSlack: 960,
      critical: false,
    })
    expectGolden(document, result, 'c', {
      wbs: '1.2',
      scheduledStart: MONDAY,
      scheduledFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(document, result, 'S', {
      wbs: '1',
      summary: true,
      scheduledStart: MONDAY,
      scheduledFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 1440,
    })
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('P06: indenting a task turns its preceding sibling into a summary and reschedules', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', uid: 1, duration: day(480) }),
          makeTask({ id: 'b', uid: 2, duration: day(480) }),
          makeTask({ id: 'c', uid: 3, duration: day(240) }),
        ],
      }),
    )
    const document = runCommand(base, {
      type: 'IndentTask',
      taskId: taskId('c'),
      parentTaskId: taskId('b'),
    })
    const result = resultOf(document)
    expectGolden(document, result, 'a', {
      wbs: '1',
      outlineLevel: 1,
      summary: false,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(document, result, 'b', {
      wbs: '2',
      outlineLevel: 1,
      summary: true,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: '2026-08-03T13:00:00.000Z',
      earlyStart: MONDAY,
      earlyFinish: '2026-08-03T13:00:00.000Z',
      lateStart: '2026-08-03T13:00:00.000Z',
      lateFinish: MONDAY_FINISH,
      totalSlack: 240,
      freeSlack: 240,
      critical: false,
      // Summary duration is derived by the scheduler from the child span,
      // even though the document-level duration field stays untouched.
      duration: 240,
    })
    expectGolden(document, result, 'c', {
      wbs: '2.1',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'b',
      scheduledStart: MONDAY,
      scheduledFinish: '2026-08-03T13:00:00.000Z',
      lateStart: '2026-08-03T13:00:00.000Z',
      lateFinish: MONDAY_FINISH,
      totalSlack: 240,
      critical: false,
    })
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })

  it('P07: outdenting a child keeps dependencies authoritative and reschedules', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'a',
            uid: 2,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'b',
            uid: 3,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS')],
      }),
    )
    const document = runCommand(base, { type: 'OutdentTask', taskId: taskId('b') })
    expect(document.dependencies).toEqual([makeDependency('d1', 'a', 'b', 'FS')])
    const result = resultOf(document)
    expectGolden(document, result, 'b', {
      wbs: '2',
      outlineLevel: 1,
      summary: false,
      parentTaskId: null,
      scheduledStart: TUESDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(document, result, 'S', {
      wbs: '1',
      summary: true,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 480,
    })
    expectGolden(document, result, 'a', {
      wbs: '1.1',
      parentTaskId: 'S',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })

  it('P07b: outdenting the last child reverts the summary to a leaf', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'only',
            uid: 2,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
        ],
      }),
    )
    const document = runCommand(base, { type: 'OutdentTask', taskId: taskId('only') })
    const result = resultOf(document)
    expectGolden(document, result, 'S', {
      wbs: '1',
      outlineLevel: 1,
      summary: false,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY,
      earlyStart: MONDAY,
      earlyFinish: MONDAY,
      totalSlack: 480,
      critical: false,
      duration: 0,
    })
    expectGolden(document, result, 'only', {
      wbs: '2',
      outlineLevel: 1,
      summary: false,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expect(result.projectFinish).toBe(MONDAY_FINISH)
  })

  it('P08: summaries with children chain across summaries through child dependencies', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S', uid: 1, summary: true, duration: day(0) }),
          makeTask({
            id: 'a',
            uid: 2,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({
            id: 'b',
            uid: 3,
            duration: day(480),
            parentTaskId: taskId('S'),
            outlineLevel: 2,
          }),
          makeTask({ id: 'T', uid: 4, summary: true, duration: day(0) }),
          makeTask({
            id: 'd',
            uid: 5,
            duration: day(480),
            parentTaskId: taskId('T'),
            outlineLevel: 2,
          }),
        ],
        dependencies: [makeDependency('d1', 'a', 'b', 'FS'), makeDependency('d2', 'b', 'd', 'FS')],
      }),
    )
    const result = resultOf(document)
    expectGolden(document, result, 'S', {
      wbs: '1',
      summary: true,
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'T', {
      wbs: '2',
      summary: true,
      scheduledStart: WEDNESDAY,
      scheduledFinish: WEDNESDAY_FINISH,
      lateStart: WEDNESDAY,
      lateFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 480,
    })
    expectGolden(document, result, 'd', {
      wbs: '2.1',
      parentTaskId: 'T',
      scheduledStart: WEDNESDAY,
      scheduledFinish: WEDNESDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    for (const id of ['a', 'b']) {
      expectGolden(document, result, id, {
        parentTaskId: 'S',
        totalSlack: 0,
        critical: true,
      })
    }
    expect(result.projectFinish).toBe(WEDNESDAY_FINISH)
  })

  it('P09: nested summaries built purely from indent commands match the declarative fixture', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'S1', uid: 1, duration: day(0) }),
          makeTask({ id: 'S2', uid: 2, duration: day(0) }),
          makeTask({ id: 'S3', uid: 3, duration: day(0) }),
          makeTask({ id: 'a', uid: 4, duration: day(960) }),
          makeTask({ id: 'b', uid: 5, duration: day(240) }),
          makeTask({ id: 'c', uid: 6, duration: day(480) }),
        ],
      }),
    )
    const document = runAll(base, [
      { type: 'IndentTask', taskId: taskId('a'), parentTaskId: taskId('S3') },
      { type: 'IndentTask', taskId: taskId('S3'), parentTaskId: taskId('S2') },
      { type: 'IndentTask', taskId: taskId('b'), parentTaskId: taskId('S2') },
      { type: 'IndentTask', taskId: taskId('S2'), parentTaskId: taskId('S1') },
      { type: 'IndentTask', taskId: taskId('c'), parentTaskId: taskId('S1') },
    ])
    // Identity is preserved across every hierarchy mutation.
    expect(document.tasks.map((task) => task.id)).toEqual([
      taskId('S1'),
      taskId('S2'),
      taskId('S3'),
      taskId('a'),
      taskId('b'),
      taskId('c'),
    ])
    const result = resultOf(document)
    expectGolden(document, result, 'S1', {
      wbs: '1',
      outlineLevel: 1,
      summary: true,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'S2', {
      wbs: '1.1',
      outlineLevel: 2,
      summary: true,
      parentTaskId: 'S1',
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'S3', {
      wbs: '1.1.1',
      outlineLevel: 3,
      summary: true,
      parentTaskId: 'S2',
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'a', {
      wbs: '1.1.1.1',
      outlineLevel: 4,
      summary: false,
      parentTaskId: 'S3',
      scheduledStart: MONDAY,
      scheduledFinish: TUESDAY_FINISH,
      totalSlack: 0,
      critical: true,
      duration: 960,
    })
    expectGolden(document, result, 'b', {
      wbs: '1.1.2',
      outlineLevel: 3,
      summary: false,
      parentTaskId: 'S2',
      scheduledStart: MONDAY,
      scheduledFinish: '2026-08-03T13:00:00.000Z',
      lateStart: '2026-08-04T13:00:00.000Z',
      lateFinish: TUESDAY_FINISH,
      totalSlack: 720,
      freeSlack: 720,
      critical: false,
    })
    expectGolden(document, result, 'c', {
      wbs: '1.2',
      outlineLevel: 2,
      summary: false,
      parentTaskId: 'S1',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: TUESDAY,
      lateFinish: TUESDAY_FINISH,
      totalSlack: 480,
      freeSlack: 480,
      critical: false,
    })
    expect(result.projectFinish).toBe(TUESDAY_FINISH)
  })

  it('P10: hierarchy mutation preserves dependency identity and schedule determinism', () => {
    const buildBase = (): ProjectDocument =>
      canonicalizeDocument(
        makeDocument({
          tasks: [
            makeTask({ id: 'x', uid: 1, duration: day(480) }),
            makeTask({ id: 'y', uid: 2, duration: day(480) }),
            makeTask({ id: 'S', uid: 3, summary: true, duration: day(0) }),
            makeTask({
              id: 'z',
              uid: 4,
              duration: day(240),
              parentTaskId: taskId('S'),
              outlineLevel: 2,
            }),
          ],
          dependencies: [makeDependency('dep-1', 'y', 'z', 'FS')],
        }),
      )

    // Indent y beneath x; the dependency keeps its identity and endpoints.
    const indented = runCommand(buildBase(), {
      type: 'IndentTask',
      taskId: taskId('y'),
      parentTaskId: taskId('x'),
    })
    expect(indented.dependencies).toEqual([makeDependency('dep-1', 'y', 'z', 'FS')])
    const indentedResult = resultOf(indented)
    expectGolden(indented, indentedResult, 'x', {
      wbs: '1',
      summary: true,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      earlyStart: MONDAY,
      earlyFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY,
      totalSlack: 0,
      freeSlack: 0,
      critical: true,
      duration: 480,
    })
    expectGolden(indented, indentedResult, 'y', {
      wbs: '1.1',
      parentTaskId: 'x',
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: MONDAY,
      lateFinish: TUESDAY,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(indented, indentedResult, 'S', {
      wbs: '2',
      summary: true,
      scheduledStart: TUESDAY,
      scheduledFinish: '2026-08-04T13:00:00.000Z',
      totalSlack: 0,
      critical: true,
      duration: 240,
    })
    expectGolden(indented, indentedResult, 'z', {
      wbs: '2.1',
      parentTaskId: 'S',
      scheduledStart: TUESDAY,
      scheduledFinish: '2026-08-04T13:00:00.000Z',
      totalSlack: 0,
      critical: true,
    })
    expect(indentedResult.projectFinish).toBe('2026-08-04T13:00:00.000Z')

    // Outdent y back to the root; the same dependency still drives z.
    const outdented = runCommand(indented, { type: 'OutdentTask', taskId: taskId('y') })
    expect(outdented.dependencies).toEqual([makeDependency('dep-1', 'y', 'z', 'FS')])
    const outdentedResult = resultOf(outdented)
    expectGolden(outdented, outdentedResult, 'x', {
      wbs: '1',
      summary: false,
      parentTaskId: null,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      lateStart: '2026-08-03T13:00:00.000Z',
      lateFinish: '2026-08-04T13:00:00.000Z',
      totalSlack: 240,
      critical: false,
      duration: 480,
    })
    expectGolden(outdented, outdentedResult, 'y', {
      wbs: '2',
      parentTaskId: null,
      summary: false,
      scheduledStart: MONDAY,
      scheduledFinish: MONDAY_FINISH,
      totalSlack: 0,
      critical: true,
    })
    expectGolden(outdented, outdentedResult, 'S', {
      wbs: '3',
      summary: true,
      scheduledStart: TUESDAY,
      scheduledFinish: '2026-08-04T13:00:00.000Z',
      totalSlack: 0,
      critical: true,
      duration: 240,
    })
    expect(outdentedResult.projectFinish).toBe('2026-08-04T13:00:00.000Z')

    // Re-scheduling the mutated document is byte-identical on repeat.
    expect(JSON.stringify(resultOf(outdented))).toBe(JSON.stringify(outdentedResult))
  })

  it('repeated identical hierarchy command sequences produce deterministic documents and schedules', () => {
    const commands: ProjectCommand[] = [
      { type: 'CreateTask', task: makeTask({ id: 's', uid: 1, duration: day(0) }) },
      {
        type: 'CreateTask',
        task: makeTask({ id: 'a', uid: 2, duration: day(480), parentTaskId: taskId('s') }),
      },
      {
        type: 'CreateTask',
        task: makeTask({ id: 'b', uid: 3, duration: day(240), parentTaskId: taskId('s') }),
      },
      { type: 'CreateTask', task: makeTask({ id: 'c', uid: 4, duration: day(480) }) },
      { type: 'RenameTask', taskId: taskId('a'), name: 'A revised' },
      { type: 'IndentTask', taskId: taskId('c'), parentTaskId: taskId('s') },
      { type: 'OutdentTask', taskId: taskId('b') },
      { type: 'DeleteTask', taskId: taskId('c') },
    ]
    const serialized = JSON.stringify(
      canonicalizeDocument(
        makeDocument({
          tasks: [makeTask({ id: 'seed', uid: 99, duration: day(120) })],
        }),
      ),
    )
    const first = runAll(parseDocument(serialized), commands)
    const second = runAll(parseDocument(serialized), commands)
    const third = runAll(parseDocument(serialized), commands)
    expect(JSON.stringify(second)).toBe(JSON.stringify(first))
    expect(JSON.stringify(third)).toBe(JSON.stringify(first))
    const firstSchedule = resultOf(first)
    expect(JSON.stringify(resultOf(second))).toBe(JSON.stringify(firstSchedule))
    expect(JSON.stringify(resultOf(third))).toBe(JSON.stringify(firstSchedule))
    // Final shape: seed + s>[a], b at root.
    const taskById = new Map(first.tasks.map((task) => [task.id as string, task]))
    expect(taskById.get('seed')!.wbs).toBe('1')
    expect(taskById.get('s')!.wbs).toBe('2')
    expect(taskById.get('a')!.wbs).toBe('2.1')
    expect(taskById.get('b')!.wbs).toBe('3')
    expect(taskById.get('a')!.name).toBe('A revised')
  })

  it('hierarchy mutations never relax the summary-dependency invariant', () => {
    // Indenting a task beneath a task it depends on would create a
    // summary/descendant dependency; the command must be rejected so the
    // PROJECT-006 invariant survives hierarchy editing.
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'p', uid: 1, duration: day(480) }),
          makeTask({ id: 'q', uid: 2, duration: day(480) }),
          makeTask({ id: 'r', uid: 3, duration: day(480) }),
        ],
        dependencies: [makeDependency('d1', 'q', 'r', 'FS')],
      }),
    )
    const execution = applyProjectCommand(base, {
      type: 'IndentTask',
      taskId: taskId('r'),
      parentTaskId: taskId('q'),
    })
    expect(execution.result.accepted).toBe(false)
    expect(execution.document).toBe(base)
    expect(
      execution.result.diagnostics.some((diagnostic) => diagnostic.code === 'SUMMARY_DEPENDENCY'),
    ).toBe(true)
  })
})
