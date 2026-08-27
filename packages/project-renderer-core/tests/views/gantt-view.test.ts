import { describe, expect, it } from 'vitest'
import { asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  buildGanttView,
  createViewState,
  hitTestGantt,
  projectDocumentView,
  reduceViewState,
} from '../../src/index.js'
import { ganttDocument, ganttSchedule, makeDocument, makeTask } from '../fixtures.js'

const VIEWPORT = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-31T00:00:00.000Z' }
const START_MS = Date.parse(VIEWPORT.start)
const SPAN_MS = Date.parse(VIEWPORT.finish) - START_MS
const f = (iso: string): number => (Date.parse(iso) - START_MS) / SPAN_MS

const viewOf = (
  document = ganttDocument(),
  schedule = ganttSchedule(),
  state = createViewState(document),
  layout = { firstRow: 0, visibleRows: 6 },
) => buildGanttView(document, projectDocumentView(document, schedule, state), state, layout)

describe('PROJECT-022 gantt view — the synchronized surface', () => {
  it('shares ONE row window object across both panes (synchronization by construction)', () => {
    const view = viewOf()
    expect(view.taskGrid.rowWindow).toBe(view.rowWindow)
    expect(view.timeline.rowWindow).toBe(view.rowWindow)
    // The pane rows address the SAME absolute indices:
    expect(view.taskGrid.rows.map((row) => row.index)).toEqual(
      view.timeline.rows.map((row) => row.index),
    )
  })

  it('clamps the layout inputs once, for both panes', () => {
    const view = viewOf(ganttDocument(), ganttSchedule(), createViewState(ganttDocument()), {
      firstRow: 99,
      visibleRows: -3,
    })
    // visibleRows <= 0 degenerates to the canonical empty window:
    expect(view.rowWindow).toEqual({ firstIndex: 0, lastIndex: -1 })
    expect(view.taskGrid.rows).toEqual([])
    expect(view.timeline.rows).toEqual([])
    expect(view.timeline.bars).toEqual([])
  })

  it('renders the full composition over the gantt fixture', () => {
    const view = viewOf()
    expect(view.taskGrid.columns).toHaveLength(7)
    expect(view.timeline.bars).toHaveLength(5)
    expect(view.timeline.milestones).toHaveLength(1)
    expect(view.timeline.links).toHaveLength(2)
  })

  it('handles the empty document (empty windows, default columns, no geometry)', () => {
    const document = makeDocument()
    const view = buildGanttView(
      document,
      projectDocumentView(document, undefined, createViewState(document)),
      createViewState(document),
      { firstRow: 0, visibleRows: 10 },
    )
    expect(view.rowWindow).toEqual({ firstIndex: 0, lastIndex: -1 })
    expect(view.taskGrid.rows).toEqual([])
    expect(view.timeline.rows).toEqual([])
    expect(view.timeline.bars).toEqual([])
    expect(view.timeline.links).toEqual([])
    expect(view.taskGrid.columns.map((column) => column.source)).toEqual([
      'rowNumber',
      'name',
      'duration',
      'start',
      'finish',
      'predecessors',
      'resourceNames',
    ])
  })

  it('renders the grid from a document WITHOUT a schedule — no invented geometry anywhere', () => {
    const document = ganttDocument()
    const state = createViewState(document)
    const view = buildGanttView(document, projectDocumentView(document, undefined, state), state, {
      firstRow: 0,
      visibleRows: 6,
    })
    expect(view.taskGrid.rows).toHaveLength(6)
    expect(view.timeline.bars).toEqual([])
    expect(view.timeline.milestones).toEqual([])
    expect(view.timeline.links).toEqual([])
  })

  it('uses the view-state viewport the state carries', () => {
    const document = ganttDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      {
        type: 'setViewport',
        start: VIEWPORT.start,
        finish: VIEWPORT.finish,
      },
      { document },
    )
    const view = viewOf(document, ganttSchedule(), state)
    expect(view.timeline.viewport).toEqual(VIEWPORT)
    // The a1 bar sits at the August-3 position over that viewport:
    const a1 = view.timeline.bars.find((bar) => bar.taskId === asTaskId('a1'))!
    expect(a1.startFraction).toBeCloseTo(f('2026-08-03T09:00:00.000Z'), 10)
  })

  it('joins the REAL scheduling authority end-to-end (test-layer injected scheduler)', () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      tasks: [
        makeTask({ id: 't1', duration: asWorkingMinutes(480) }),
        makeTask({ id: 't2', duration: asWorkingMinutes(480) }),
      ],
    })
    const derived = schedule(document)
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setViewport', start: VIEWPORT.start, finish: VIEWPORT.finish },
      { document },
    )
    const view = buildGanttView(document, projectDocumentView(document, derived, state), state, {
      firstRow: 0,
      visibleRows: 2,
    })
    // Two leaf bars at the REAL scheduled instants:
    expect(view.timeline.bars.map((bar) => bar.taskId)).toEqual([asTaskId('t1'), asTaskId('t2')])
    const t1 = view.timeline.bars[0]!
    expect(t1.startFraction).toBeCloseTo(
      f(derived.taskSchedules[asTaskId('t1')]!.scheduledStart!),
      10,
    )
    expect(t1.finishFraction).toBeCloseTo(
      f(derived.taskSchedules[asTaskId('t1')]!.scheduledFinish!),
      10,
    )
    // And the grid start cells carry the same real instants:
    const t1Row = view.taskGrid.rows.find((row) => row.row.taskId === asTaskId('t1'))!
    expect(t1Row.cells[3]).toEqual({
      kind: 'instant',
      iso: derived.taskSchedules[asTaskId('t1')]!.scheduledStart!,
    })
  })

  it('is pure: 3× byte-identical and inputs never mutated', () => {
    const document = ganttDocument()
    const gantt = ganttSchedule()
    const state = createViewState(document)
    const projection = projectDocumentView(document, gantt, state)
    const before = JSON.stringify({ document, gantt, state, projection })
    const run = () =>
      JSON.stringify(buildGanttView(document, projection, state, { firstRow: 0, visibleRows: 6 }))
    expect(run()).toBe(run())
    expect(run()).toBe(run())
    expect(JSON.stringify({ document, gantt, state, projection })).toBe(before)
  })
})

describe('PROJECT-022 gantt view — hit testing (geometry→entity inverse)', () => {
  it('hits bars by row and fraction containment', () => {
    const view = viewOf()
    const a1 = view.timeline.bars.find((bar) => bar.taskId === asTaskId('a1'))!
    const hit = hitTestGantt(view.timeline, {
      rowIndex: a1.rowIndex,
      fraction: (a1.startFraction + a1.finishFraction) / 2,
    })
    expect(hit).toEqual({ kind: 'bar', taskId: asTaskId('a1') })
  })

  it('hits milestones exactly, before bars at the same position (deterministic priority)', () => {
    // A flagged milestone WITH a real span: both a bar and a diamond exist on
    // the same row; the diamond sits at the bar's start.
    const document = makeDocument({
      startDate: '2026-08-01T00:00:00.000Z',
      tasks: [makeTask({ id: 'm2', milestone: true, duration: asWorkingMinutes(960) })],
    })
    const schedule2 = {
      taskSchedules: {
        [asTaskId('m2')]: {
          taskId: asTaskId('m2'),
          scheduledStart: '2026-08-03T09:00:00.000Z' as never,
          scheduledFinish: '2026-08-05T17:00:00.000Z' as never,
          duration: asWorkingMinutes(960),
          totalSlack: 0,
          freeSlack: 0,
          critical: false,
        },
      },
      diagnostics: [],
    }
    const state = createViewState(document)
    const view = buildGanttView(document, projectDocumentView(document, schedule2, state), state, {
      firstRow: 0,
      visibleRows: 1,
    })
    expect(view.timeline.bars).toHaveLength(1)
    expect(view.timeline.milestones).toHaveLength(1)
    const diamondAt = f('2026-08-03T09:00:00.000Z')
    // Exactly on the diamond → milestone wins over the bar:
    expect(hitTestGantt(view.timeline, { rowIndex: 0, fraction: diamondAt })).toEqual({
      kind: 'milestone',
      taskId: asTaskId('m2'),
    })
    // Mid-bar (no diamond there) → the bar:
    expect(
      hitTestGantt(view.timeline, {
        rowIndex: 0,
        fraction:
          (view.timeline.bars[0]!.startFraction + view.timeline.bars[0]!.finishFraction) / 2,
      }),
    ).toEqual({ kind: 'bar', taskId: asTaskId('m2') })
  })

  it('applies the host-supplied tolerance and misses cleanly', () => {
    const view = viewOf()
    const m = view.timeline.milestones[0]!
    // Just off the diamond: exact hit-test misses…
    expect(
      hitTestGantt(view.timeline, { rowIndex: m.rowIndex, fraction: m.fraction + 0.004 }),
    ).toBeUndefined()
    // …but a tolerance (hosts pass pixelTolerance/width) catches it:
    expect(
      hitTestGantt(view.timeline, { rowIndex: m.rowIndex, fraction: m.fraction + 0.004 }, 0.005),
    ).toEqual({ kind: 'milestone', taskId: m.taskId })
    // A fraction on row 0 beyond every row-0 bar misses even with tolerance
    // (the milestone lives on another row and is never hit from row 0):
    expect(hitTestGantt(view.timeline, { rowIndex: 0, fraction: 0.99 }, 0.01)).toBeUndefined()
  })
})
