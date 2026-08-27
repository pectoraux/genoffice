import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type { DerivedSchedule, ISODateTime, TaskSchedule } from '@genoffice/project-contracts'
import { createViewState, projectDocumentView, reduceViewState } from '../src/index.js'
import {
  makeAssignment,
  makeDocument,
  makeResource,
  makeTask,
  outlineDocument,
} from './fixtures.js'

const scheduleOf = (
  entries: Record<string, TaskSchedule>,
  diagnostics: DerivedSchedule['diagnostics'] = [],
  projectStart?: string,
  projectFinish?: string,
): DerivedSchedule =>
  ({
    taskSchedules: entries,
    diagnostics,
    ...(projectStart !== undefined ? { projectStart: asISODateTime(projectStart) } : {}),
    ...(projectFinish !== undefined ? { projectFinish: asISODateTime(projectFinish) } : {}),
  }) as unknown as DerivedSchedule

const taskSchedule = (taskId: string, overrides: Partial<TaskSchedule> = {}): TaskSchedule =>
  ({
    taskId: asTaskId(taskId),
    totalSlack: 0,
    freeSlack: 0,
    critical: false,
    duration: asWorkingMinutes(480),
    ...overrides,
  }) as TaskSchedule

describe('PROJECT-021 projection — rows and collapse', () => {
  it('projects every task as a row in canonical outline order', () => {
    const document = outlineDocument() // root, a, a1, b
    const projection = projectDocumentView(document, undefined, createViewState(document))
    expect(projection.rows.map((row) => row.taskId)).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('a1'),
      asTaskId('b'),
    ])
  })

  it('hides the entire subtree of a collapsed summary, keeping deeper ancestors visible', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('a') }, { document })
    const projection = projectDocumentView(document, undefined, state)
    expect(projection.rows.map((row) => row.taskId)).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('b'),
    ])
    expect(projection.rows[1]!.collapsed).toBe(true)
    expect(projection.rows[2]!.collapsed).toBe(false)
    // Collapsing the root hides everything below it.
    state = reduceViewState(state, { type: 'expandAll' }, { document })
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('root') },
      { document },
    )
    expect(projectDocumentView(document, undefined, state).rows.map((row) => row.taskId)).toEqual([
      asTaskId('root'),
    ])
  })

  it('echoes canonical task fields verbatim (identity, structure, labeling)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 't1',
          uid: 77,
          wbs: '1',
          name: 'Design',
          outlineLevel: 1,
          milestone: true,
          manualScheduled: true,
          priority: 900,
          percentComplete: 40,
          constraintType: 'mustStartOn',
          constraintDate: asISODateTime('2026-08-05T09:00:00.000Z'),
          deadline: asISODateTime('2026-08-10T09:00:00.000Z'),
        }),
      ],
    })
    const [row] = projectDocumentView(document, undefined, createViewState(document)).rows
    expect(row!.uid).toBe(77)
    expect(row!.wbs).toBe('1')
    expect(row!.name).toBe('Design')
    expect(row!.milestone).toBe(true)
    expect(row!.manualScheduled).toBe(true)
    expect(row!.priority).toBe(900)
    expect(row!.percentComplete).toBe(40)
    expect(row!.constraintType).toBe('mustStartOn')
    expect(row!.constraintDate).toBe('2026-08-05T09:00:00.000Z')
    expect(row!.deadline).toBe('2026-08-10T09:00:00.000Z')
    expect(row!.summary).toBe(false)
    expect(row!.resourceNames).toEqual([])
    expect('schedule' in row!).toBe(false)
  })

  it('projects resource names from the assignment array in document order, deduplicated', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      resources: [makeResource({ id: 'r1', name: 'Ada' }), makeResource({ id: 'r2', name: 'Bo' })],
      assignments: [
        makeAssignment('a1', 't1', 'r2'),
        makeAssignment('a2', 't1', 'r1'),
        makeAssignment('a3', 't1', 'r2'),
      ],
    })
    const [row] = projectDocumentView(document, undefined, createViewState(document)).rows
    expect(row!.resourceNames).toEqual(['Bo', 'Ada'])
  })
})

describe('PROJECT-021 projection — schedule join (no scheduling authority)', () => {
  it('joins the authoritative TaskSchedule objects BY REFERENCE (never copied or recomputed)', () => {
    const document = outlineDocument()
    const rootSchedule = taskSchedule('root', {
      critical: true,
      scheduledStart: asISODateTime('2026-08-03T09:00:00.000Z') as ISODateTime,
      scheduledFinish: asISODateTime('2026-08-04T09:00:00.000Z') as ISODateTime,
    })
    const aSchedule = taskSchedule('a', { totalSlack: 10, freeSlack: 5 })
    const schedule = scheduleOf({
      [asTaskId('root')]: rootSchedule,
      [asTaskId('a')]: aSchedule,
    })
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const rootRow = projection.rows.find((row) => row.taskId === asTaskId('root'))!
    const aRow = projection.rows.find((row) => row.taskId === asTaskId('a'))!
    expect(rootRow.schedule).toBe(rootSchedule) // reference identity — verbatim join
    expect(aRow.schedule).toBe(aSchedule)
    // Tasks without a schedule entry carry `undefined` — never an invented value.
    expect('schedule' in projection.rows.find((row) => row.taskId === asTaskId('a1'))!).toBe(false)
  })

  it('without a schedule the projection renders the document and invents nothing', () => {
    const document = outlineDocument()
    const projection = projectDocumentView(document, undefined, createViewState(document))
    expect(projection.hasSchedule).toBe(false)
    expect(projection.scheduleDiagnostics).toEqual([])
    expect(projection.rows.every((row) => !('schedule' in row))).toBe(true)
    expect(projection.projectStart).toBe('2026-08-03T09:00:00.000Z') // properties fallback
    expect('projectFinish' in projection).toBe(false)
  })

  it('prefers the derived schedule for the project window and echoes its diagnostics verbatim', () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      finishDate: '2026-08-30T09:00:00.000Z',
    })
    const diagnostics = [{ code: 'SCHED_X', severity: 'warning' as const, message: 'note' }]
    const schedule = scheduleOf(
      {},
      diagnostics,
      '2026-08-05T09:00:00.000Z',
      '2026-08-20T09:00:00.000Z',
    )
    const projection = projectDocumentView(document, schedule, createViewState(document))
    expect(projection.hasSchedule).toBe(true)
    expect(projection.projectStart).toBe('2026-08-05T09:00:00.000Z')
    expect(projection.projectFinish).toBe('2026-08-20T09:00:00.000Z')
    expect(projection.scheduleDiagnostics).toBe(diagnostics)
  })
})

describe('PROJECT-021 projection — active view resolution and discipline', () => {
  it('resolves the active canonical view/table/filter/group definitions', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      views: [{ id: 'v1' as never, name: 'Gantt', type: 'gantt' }],
      tables: [{ id: 'tb1' as never, name: 'Entry', columns: ['name'] }],
      filters: [],
      groups: [],
    })
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'setActiveView', viewId: 'v1' as never }, { document })
    state = reduceViewState(
      state,
      { type: 'setActiveTable', tableId: 'tb1' as never },
      { document },
    )
    const projection = projectDocumentView(document, undefined, state)
    expect(projection.activeView?.name).toBe('Gantt')
    expect(projection.activeTable?.columns).toEqual(['name'])
    expect('activeFilter' in projection).toBe(false)
  })

  it('is deterministic across 3 runs and never mutates its inputs', () => {
    const document = outlineDocument()
    const schedule = scheduleOf({ [asTaskId('root')]: taskSchedule('root', { critical: true }) })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      { document, schedule },
    )
    const documentBefore = JSON.stringify(document)
    const scheduleBefore = JSON.stringify(schedule)
    const first = JSON.stringify(projectDocumentView(document, schedule, state))
    expect(JSON.stringify(projectDocumentView(document, schedule, state))).toBe(first)
    expect(JSON.stringify(projectDocumentView(document, schedule, state))).toBe(first)
    expect(JSON.stringify(document)).toBe(documentBefore)
    expect(JSON.stringify(schedule)).toBe(scheduleBefore)
  })

  it('scales linearly in task count (2000-task projection stays correct at the ends)', () => {
    const tasks = Array.from({ length: 2000 }, (_, index) =>
      makeTask({ id: `task-${index}`, name: `Task ${index}` }),
    )
    const document = makeDocument({ tasks })
    const projection = projectDocumentView(document, undefined, createViewState(document))
    expect(projection.rows).toHaveLength(2000)
    expect(projection.rows[0]!.taskId).toBe(asTaskId('task-0'))
    expect(projection.rows[1999]!.taskId).toBe(asTaskId('task-1999'))
  })
})
