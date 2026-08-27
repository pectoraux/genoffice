import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import {
  buildRowWindow,
  buildTimeline,
  buildTimeAxis,
  createViewState,
  projectDocumentView,
  reduceViewState,
} from '../../src/index.js'
import { ganttDocument, ganttSchedule } from '../fixtures.js'

const VIEWPORT = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-31T00:00:00.000Z' }

const timelineOf = (
  document = ganttDocument(),
  schedule = ganttSchedule(),
  state = createViewState(document),
  window = buildRowWindow(6, { firstRow: 0, visibleRows: 6 }),
) => buildTimeline(document, projectDocumentView(document, schedule, state), VIEWPORT, window)

describe('PROJECT-022 timeline — composition', () => {
  it('echoes the viewport by reference and carries the deterministic axis bands', () => {
    const timeline = timelineOf()
    expect(timeline.viewport).toBe(VIEWPORT)
    expect(timeline.bands).toEqual(buildTimeAxis(VIEWPORT, timeline.axisLevel))
    expect(timeline.axisLevel).toBe('day') // a 30-day window
    expect(timeline.bands[0]).toMatchObject({ start: '2026-08-01T00:00:00.000Z' })
    expect(timeline.bands).toHaveLength(30)
  })

  it('lists the in-window rows with ABSOLUTE indices, joined by reference', () => {
    const timeline = timelineOf()
    expect(timeline.rows.map((row) => row.index)).toEqual([0, 1, 2, 3, 4, 5])
    expect(timeline.rows[3]!.row.taskId).toBe(asTaskId('a2'))
    const window = buildRowWindow(6, { firstRow: 4, visibleRows: 2 })
    const sliced = timelineOf(
      ganttDocument(),
      ganttSchedule(),
      createViewState(ganttDocument()),
      window,
    )
    expect(sliced.rows.map((row) => row.index)).toEqual([4, 5])
    expect(sliced.rows.map((row) => row.row.taskId)).toEqual([asTaskId('b'), asTaskId('m')])
  })

  it('composes the three geometry surfaces over the SAME row window', () => {
    const timeline = timelineOf()
    expect(timeline.bars.map((bar) => bar.taskId)).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('a1'),
      asTaskId('a2'),
      asTaskId('b'),
    ])
    expect(timeline.milestones.map((milestone) => milestone.taskId)).toEqual([asTaskId('m')])
    expect(timeline.links.map((link) => link.type)).toEqual(['FS', 'SS'])
    for (const bar of timeline.bars) {
      expect(bar.rowIndex).toBeGreaterThanOrEqual(timeline.rowWindow.firstIndex)
      expect(bar.rowIndex).toBeLessThanOrEqual(timeline.rowWindow.lastIndex)
    }
  })

  it('reflects collapse through every composed surface', () => {
    const document = ganttDocument()
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('a') }, { document })
    const timeline = timelineOf(
      document,
      ganttSchedule(),
      state,
      buildRowWindow(4, { firstRow: 0, visibleRows: 4 }),
    )
    expect(timeline.rows.map((row) => row.row.taskId)).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('b'),
      asTaskId('m'),
    ])
    expect(timeline.bars.map((bar) => bar.taskId)).not.toContain(asTaskId('a1'))
    // The FS link a1→a2 is entirely inside the collapsed subtree — omitted:
    expect(timeline.links.map((link) => link.type)).toEqual(['SS'])
  })

  it('degenerates to an EMPTY model for an unparseable viewport (nothing invented)', () => {
    const document = ganttDocument()
    const projection = projectDocumentView(document, ganttSchedule(), createViewState(document))
    const timeline = buildTimeline(
      document,
      projection,
      { start: 'nope', finish: '2026-08-31T00:00:00.000Z' },
      buildRowWindow(6, { firstRow: 0, visibleRows: 6 }),
    )
    expect(timeline.bands).toEqual([])
    expect(timeline.rows).toEqual([])
    expect(timeline.bars).toEqual([])
    expect(timeline.milestones).toEqual([])
    expect(timeline.links).toEqual([])
  })

  it('is pure: 3× byte-identical and inputs never mutated', () => {
    const document = ganttDocument()
    const schedule = ganttSchedule()
    const state = createViewState(document)
    const before = JSON.stringify({ document, schedule, state })
    const run = () =>
      JSON.stringify(
        buildTimeline(
          document,
          projectDocumentView(document, schedule, state),
          VIEWPORT,
          buildRowWindow(6, { firstRow: 0, visibleRows: 6 }),
        ),
      )
    expect(run()).toBe(run())
    expect(run()).toBe(run())
    expect(JSON.stringify({ document, schedule, state })).toBe(before)
  })
})
