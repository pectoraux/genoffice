import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import {
  DEFAULT_TASK_GRID_COLUMNS,
  buildRowWindow,
  buildTaskGrid,
  createViewState,
  projectDocumentView,
  reduceViewState,
} from '../../src/index.js'
import {
  ganttDocument,
  ganttSchedule,
  makeDependency,
  makeDocument,
  makeTask,
  outlineDocument,
} from '../fixtures.js'

const context = (document = ganttDocument()) => ({ document })
const fullWindow = (count: number) => buildRowWindow(count, { firstRow: 0, visibleRows: count })
const asProjectTableId = (id: string) => id as never

const projectionOf = (
  document = ganttDocument(),
  schedule = ganttSchedule(),
  state = createViewState(document),
) => projectDocumentView(document, schedule, state)

describe('PROJECT-022 task grid — columns', () => {
  it('uses the documented DEFAULT column set when no table is active', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const grid = buildTaskGrid(document, projection, fullWindow(projection.rows.length))
    expect(grid.columns.map((column) => column.source)).toEqual([...DEFAULT_TASK_GRID_COLUMNS])
    expect(grid.columns.map((column) => column.field)).toEqual([
      'rowNumber',
      'taskName',
      'duration',
      'start',
      'finish',
      'predecessors',
      'resourceNames',
    ])
  })

  it('resolves columns from the active canonical table (the .gproj field-name convention)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      tables: [
        {
          id: asProjectTableId('tbl1'),
          name: 'Entry',
          columns: ['name', 'duration', 'start', 'finish'],
        },
      ],
    })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setActiveTable', tableId: asProjectTableId('tbl1') },
      context(document),
    )
    const projection = projectDocumentView(document, undefined, state)
    const grid = buildTaskGrid(document, projection, fullWindow(1))
    expect(grid.columns.map((column) => column.field)).toEqual([
      'taskName',
      'duration',
      'start',
      'finish',
    ])
  })

  it('maps unrecognized column names to unsupported cells (never a crash, never invented data)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      tables: [
        {
          id: asProjectTableId('tbl1'),
          name: 'Custom',
          columns: ['name', 'phase', 'costCenter'],
        },
      ],
    })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setActiveTable', tableId: asProjectTableId('tbl1') },
      context(document),
    )
    const projection = projectDocumentView(document, undefined, state)
    const grid = buildTaskGrid(document, projection, fullWindow(1))
    expect(grid.columns[1]).toEqual({ field: 'unsupported', source: 'phase' })
    expect(grid.columns[2]).toEqual({ field: 'unsupported', source: 'costCenter' })
    expect(grid.rows[0]!.cells[1]).toEqual({ kind: 'unsupported', source: 'phase' })
    // The recognized column still produces its structured cell:
    expect(grid.rows[0]!.cells[0]).toMatchObject({ kind: 'taskName', text: 't1' })
  })

  it('falls back to the default set when the active table declares no columns', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      tables: [{ id: asProjectTableId('tbl1'), name: 'Empty', columns: [] }],
    })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setActiveTable', tableId: asProjectTableId('tbl1') },
      context(document),
    )
    const projection = projectDocumentView(document, undefined, state)
    const grid = buildTaskGrid(document, projection, fullWindow(1))
    expect(grid.columns.map((column) => column.source)).toEqual([...DEFAULT_TASK_GRID_COLUMNS])
  })
})

describe('PROJECT-022 task grid — cells', () => {
  it('produces structured, unformatted cells (rowNumber, taskName, duration, start, finish, predecessors, resourceNames)', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const grid = buildTaskGrid(document, projection, fullWindow(projection.rows.length))
    // Row `a1` (document position 3, 50% complete, schedule 08-03 09:00→17:00):
    const a1 = grid.rows.find((row) => row.row.taskId === asTaskId('a1'))!
    const [rowNumber, taskName, duration, start, finish, predecessors, resources] = a1.cells
    expect(rowNumber).toEqual({ kind: 'rowNumber', value: 3 })
    expect(taskName).toEqual({
      kind: 'taskName',
      text: 'a1',
      outlineLevel: 3,
      summary: false,
      milestone: false,
      collapsed: false,
    })
    expect(duration).toEqual({ kind: 'duration', minutes: 480 })
    expect(start).toEqual({ kind: 'instant', iso: '2026-08-03T09:00:00.000Z' })
    expect(finish).toEqual({ kind: 'instant', iso: '2026-08-03T17:00:00.000Z' })
    expect(predecessors).toEqual({ kind: 'predecessors', links: [] })
    expect(resources).toEqual({ kind: 'resources', names: [] })
    // `a2` carries the FS predecessor link with the predecessor uid:
    const a2 = grid.rows.find((row) => row.row.taskId === asTaskId('a2'))!
    expect(a2.cells[5]).toEqual({
      kind: 'predecessors',
      links: [
        {
          dependencyId: 'd1' as never,
          predecessorTaskId: asTaskId('a1'),
          predecessorUid: 3,
          type: 'FS',
          lagMinutes: 0,
        },
      ],
    })
  })

  it('keeps start/finish cells EMPTY when the row has no schedule — dates are never invented', () => {
    const document = ganttDocument()
    const projection = projectDocumentView(document, undefined, createViewState(document))
    const grid = buildTaskGrid(document, projection, fullWindow(projection.rows.length))
    for (const row of grid.rows) {
      expect(row.cells[3]).toEqual({ kind: 'empty' }) // start
      expect(row.cells[4]).toEqual({ kind: 'empty' }) // finish
    }
  })

  it('prefers the schedule echo for duration/percent (rolled-up summary values) and falls back to the task field', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const grid = buildTaskGrid(document, projection, fullWindow(projection.rows.length))
    // Summary `root`: schedule.duration 4800 wins over the task field:
    const root = grid.rows.find((row) => row.row.taskId === asTaskId('root'))!
    expect(root.cells[2]).toEqual({ kind: 'duration', minutes: 4800 })
    // A task without a schedule entry: the canonical task field is the value.
    const schedule = ganttSchedule()
    delete schedule.taskSchedules[asTaskId('b')]
    const projectionWithoutB = projectDocumentView(document, schedule, createViewState(document))
    const gridWithoutB = buildTaskGrid(
      document,
      projectionWithoutB,
      fullWindow(projectionWithoutB.rows.length),
    )
    const b = gridWithoutB.rows.find((row) => row.row.taskId === asTaskId('b'))!
    expect(b.cells[2]).toEqual({ kind: 'duration', minutes: 1440 })
    expect(b.cells[3]).toEqual({ kind: 'empty' })
  })

  it('rowNumber is the CANONICAL document position — stable under collapse (MS Project ID behavior)', () => {
    const document = ganttDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      context(document),
    )
    const projection = projectionOf(document, ganttSchedule(), state)
    const grid = buildTaskGrid(document, projection, fullWindow(projection.rows.length))
    // Visible: root(1), a(2), b(5), m(6) — a1(3)/a2(4) hidden keep their numbers:
    expect(grid.rows.map((row) => row.cells[0])).toEqual([
      { kind: 'rowNumber', value: 1 },
      { kind: 'rowNumber', value: 2 },
      { kind: 'rowNumber', value: 5 },
      { kind: 'rowNumber', value: 6 },
    ])
  })

  it('lists predecessors in document order with type and lag', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' }), makeTask({ id: 't2' }), makeTask({ id: 't3' })],
      dependencies: [
        makeDependency('d1', 't1', 't3'),
        makeDependency('d2', 't2', 't3', 'FF', -120),
      ],
    })
    const projection = projectDocumentView(document, undefined, createViewState(document))
    const grid = buildTaskGrid(document, projection, fullWindow(3))
    const t3 = grid.rows.find((row) => row.row.taskId === asTaskId('t3'))!
    const defaultPredecessors = t3.cells[5] as { kind: 'predecessors'; links: unknown[] }
    expect(defaultPredecessors.links).toHaveLength(2)
    expect(defaultPredecessors.links[0]).toMatchObject({ type: 'FS', lagMinutes: 0 })
    expect(defaultPredecessors.links[1]).toMatchObject({ type: 'FF', lagMinutes: -120 })
  })

  it('joins projection rows BY REFERENCE (row identity is the projection row)', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const grid = buildTaskGrid(document, projection, fullWindow(projection.rows.length))
    expect(grid.rows[0]!.row).toBe(projection.rows[0])
    expect(grid.rows[3]!.row).toBe(projection.rows[3])
  })
})

describe('PROJECT-022 task grid — virtualization and discipline', () => {
  it('slices rows to the shared window with ABSOLUTE indices', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const window = buildRowWindow(6, { firstRow: 2, visibleRows: 2 })
    const grid = buildTaskGrid(document, projection, window)
    expect(grid.rowWindow).toBe(window)
    expect(grid.rows.map((row) => row.index)).toEqual([2, 3])
    expect(grid.rows.map((row) => row.row.taskId)).toEqual([asTaskId('a1'), asTaskId('a2')])
    expect(grid.rows[0]!.row).toBe(projection.rows[2])
  })

  it('renders nothing for an empty window', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const grid = buildTaskGrid(document, projection, { firstIndex: 0, lastIndex: -1 })
    expect(grid.rows).toEqual([])
    expect(grid.columns).toHaveLength(DEFAULT_TASK_GRID_COLUMNS.length)
  })

  it('is pure: 3× byte-identical and inputs never mutated', () => {
    const document = ganttDocument()
    const projection = projectionOf(document)
    const documentBefore = JSON.stringify(document)
    const projectionBefore = JSON.stringify(projection)
    const run = () =>
      JSON.stringify(buildTaskGrid(document, projection, fullWindow(projection.rows.length)))
    expect(run()).toBe(run())
    expect(run()).toBe(run())
    expect(JSON.stringify(document)).toBe(documentBefore)
    expect(JSON.stringify(projection)).toBe(projectionBefore)
  })

  it('scales linearly (2000-row grid stays correct at the ends)', () => {
    const tasks = Array.from({ length: 2000 }, (_, index) =>
      makeTask({ id: `task-${index}`, name: `Task ${index}` }),
    )
    const document = makeDocument({ tasks })
    const projection = projectDocumentView(document, undefined, createViewState(document))
    const grid = buildTaskGrid(document, projection, fullWindow(2000))
    expect(grid.rows).toHaveLength(2000)
    expect(grid.rows[0]!.row.taskId).toBe(asTaskId('task-0'))
    expect(grid.rows[1999]!.row.taskId).toBe(asTaskId('task-1999'))
    expect(grid.rows[1999]!.cells[0]).toEqual({ kind: 'rowNumber', value: 2000 })
  })

  it('supports the remaining structured fields (wbs, outlineLevel, priority, uid, percentComplete)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 't1',
          wbs: '1.2',
          outlineLevel: 2,
          priority: 700,
          uid: 42,
          percentComplete: 40,
        }),
      ],
      tables: [
        {
          id: asProjectTableId('tbl1'),
          name: 'All',
          columns: ['wbs', 'outlineLevel', 'priority', 'uid', 'percentComplete'],
        },
      ],
    })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setActiveTable', tableId: asProjectTableId('tbl1') },
      { document },
    )
    const projection = projectDocumentView(document, undefined, state)
    const grid = buildTaskGrid(document, projection, fullWindow(1))
    expect(grid.rows[0]!.cells).toEqual([
      { kind: 'text', text: '1.2' },
      { kind: 'number', value: 2 },
      { kind: 'number', value: 700 },
      { kind: 'number', value: 42 },
      { kind: 'percentComplete', value: 40 },
    ])
  })

  it('keeps the collapse flag on taskName cells (the expander state)', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('root') },
      { document },
    )
    const projection = projectDocumentView(document, undefined, state)
    const grid = buildTaskGrid(document, projection, fullWindow(1))
    expect(grid.rows[0]!.cells[1]).toMatchObject({ kind: 'taskName', collapsed: true })
  })
})
