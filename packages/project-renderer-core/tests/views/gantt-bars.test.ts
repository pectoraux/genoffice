import { describe, expect, it } from 'vitest'
import { asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import {
  buildGanttBars,
  buildRowWindow,
  createViewState,
  projectDocumentView,
} from '../../src/index.js'
import {
  ganttDocument,
  ganttSchedule,
  makeDocument,
  makeScheduleEntry,
  makeTask,
} from '../fixtures.js'

const VIEWPORT = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-31T00:00:00.000Z' }
const START_MS = Date.parse(VIEWPORT.start)
const SPAN_MS = Date.parse(VIEWPORT.finish) - START_MS
/** The documented fraction definition, computed independently in the test. */
const f = (iso: string): number => (Date.parse(iso) - START_MS) / SPAN_MS
const fullWindow = () => buildRowWindow(6, { firstRow: 0, visibleRows: 6 })

const projectionOf = () => {
  const document = ganttDocument()
  const schedule = ganttSchedule()
  return {
    document,
    projection: projectDocumentView(document, schedule, createViewState(document)),
  }
}

describe('PROJECT-022 Gantt bars — geometry', () => {
  it('maps scheduled windows onto clamped viewport fractions in ascending row order', () => {
    const { projection } = projectionOf()
    const bars = buildGanttBars(projection, VIEWPORT, fullWindow())
    expect(bars.map((bar) => bar.taskId)).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('a1'),
      asTaskId('a2'),
      asTaskId('b'),
    ])
    const a1 = bars[2]!
    expect(a1.startFraction).toBeCloseTo(f('2026-08-03T09:00:00.000Z'), 10)
    expect(a1.finishFraction).toBeCloseTo(f('2026-08-03T17:00:00.000Z'), 10)
    expect(a1.rowIndex).toBe(2)
    expect(a1.kind).toBe('leaf')
    expect(a1.startsBefore).toBe(false)
    expect(a1.finishesAfter).toBe(false)
  })

  it('marks summary rows with the summary bar kind (same date source — the rolled-up window)', () => {
    const { projection } = projectionOf()
    const bars = buildGanttBars(projection, VIEWPORT, fullWindow())
    expect(bars[0]).toMatchObject({ taskId: asTaskId('root'), kind: 'summary', rowIndex: 0 })
    expect(bars[1]).toMatchObject({ taskId: asTaskId('a'), kind: 'summary', rowIndex: 1 })
    expect(bars[0]!.startFraction).toBeCloseTo(f('2026-08-03T09:00:00.000Z'), 10)
    expect(bars[0]!.finishFraction).toBeCloseTo(f('2026-08-12T17:00:00.000Z'), 10)
  })

  it('interpolates the progress point linearly over the raw span (50% → the midpoint)', () => {
    const { projection } = projectionOf()
    const a1 = buildGanttBars(projection, VIEWPORT, fullWindow())[2]!
    const expected =
      f('2026-08-03T09:00:00.000Z') +
      0.5 * (f('2026-08-03T17:00:00.000Z') - f('2026-08-03T09:00:00.000Z'))
    expect(a1.progressFraction).toBeCloseTo(expected, 10)
    // 25% rolled-up progress on the summary:
    const root = buildGanttBars(projection, VIEWPORT, fullWindow())[0]!
    const rootExpected =
      f('2026-08-03T09:00:00.000Z') +
      0.25 * (f('2026-08-12T17:00:00.000Z') - f('2026-08-03T09:00:00.000Z'))
    expect(root.progressFraction).toBeCloseTo(rootExpected, 10)
  })

  it('falls back to the task percentComplete when the schedule echo is absent', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: asWorkingMinutes(480), percentComplete: 80 })],
      startDate: '2026-08-01T00:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {
        [asTaskId('t1')]: makeScheduleEntry(
          't1',
          '2026-08-02T00:00:00.000Z',
          '2026-08-04T00:00:00.000Z',
        ),
      },
      diagnostics: [],
    }
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const bar = buildGanttBars(
      projection,
      VIEWPORT,
      buildRowWindow(1, { firstRow: 0, visibleRows: 1 }),
    )[0]!
    const expected =
      f('2026-08-02T00:00:00.000Z') +
      0.8 * (f('2026-08-04T00:00:00.000Z') - f('2026-08-02T00:00:00.000Z'))
    expect(bar.progressFraction).toBeCloseTo(expected, 10)
  })

  it('clips bars to the viewport with explicit flags', () => {
    // A viewport of Aug 5..Aug 8: root/a/a1/a2 start before it; root and b
    // finish after it; `a` and `a2` finish inside it.
    const viewport = { start: '2026-08-05T00:00:00.000Z', finish: '2026-08-08T00:00:00.000Z' }
    const { projection } = projectionOf()
    const bars = buildGanttBars(projection, viewport, fullWindow())
    const byTask = new Map(bars.map((bar) => [bar.taskId, bar]))
    const root = byTask.get(asTaskId('root'))!
    expect(root.startsBefore).toBe(true)
    expect(root.startFraction).toBe(0)
    expect(root.finishesAfter).toBe(true)
    expect(root.finishFraction).toBe(1)
    const a2 = byTask.get(asTaskId('a2'))!
    expect(a2.startsBefore).toBe(true)
    expect(a2.startFraction).toBe(0)
    expect(a2.finishesAfter).toBe(false)
    // 08-05T17:00 sits 17/72 into the three-day viewport:
    expect(a2.finishFraction).toBeCloseTo(17 / 72, 10)
    const b = byTask.get(asTaskId('b'))!
    expect(b.startsBefore).toBe(false) // entirely AFTER the viewport
    expect(b.finishesAfter).toBe(true)
    expect(b.startFraction).toBe(1)
    expect(b.finishFraction).toBe(1)
  })

  it('produces NO bar without a complete schedule span — dates are never invented', () => {
    const document = ganttDocument()
    const schedule = ganttSchedule()
    delete schedule.taskSchedules[asTaskId('a2')] // no schedule at all
    schedule.taskSchedules[asTaskId('b')] = makeScheduleEntry('b', '2026-08-10T09:00:00.000Z', '') // missing finish → build over a partial entry
    const partial = {
      ...schedule.taskSchedules[asTaskId('b')],
      scheduledFinish: undefined,
    } as never
    schedule.taskSchedules[asTaskId('b')] = partial
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const bars = buildGanttBars(projection, VIEWPORT, fullWindow())
    const byTask = new Map(bars.map((bar) => [bar.taskId, bar]))
    expect(byTask.has(asTaskId('a2'))).toBe(false)
    expect(byTask.has(asTaskId('b'))).toBe(false)
    expect(byTask.has(asTaskId('a1'))).toBe(true)
  })

  it('produces no bar for a zero-span schedule (that is milestone geometry)', () => {
    const { projection } = projectionOf()
    const bars = buildGanttBars(projection, VIEWPORT, fullWindow())
    expect(bars.map((bar) => bar.taskId)).not.toContain(asTaskId('m'))
  })

  it('produces nothing for an empty window or a degenerate viewport', () => {
    const { projection } = projectionOf()
    expect(buildGanttBars(projection, VIEWPORT, { firstIndex: 0, lastIndex: -1 })).toEqual([])
    expect(
      buildGanttBars(
        projection,
        { start: 'nope', finish: '2026-08-31T00:00:00.000Z' },
        fullWindow(),
      ),
    ).toEqual([])
    expect(
      buildGanttBars(
        projection,
        { start: '2026-08-31T00:00:00.000Z', finish: '2026-08-01T00:00:00.000Z' },
        fullWindow(),
      ),
    ).toEqual([])
  })

  it('slices to the virtualized window with absolute indices', () => {
    const { projection } = projectionOf()
    const window = buildRowWindow(6, { firstRow: 4, visibleRows: 2 })
    const bars = buildGanttBars(projection, VIEWPORT, window)
    expect(bars.map((bar) => bar.taskId)).toEqual([asTaskId('b')])
    expect(bars[0]!.rowIndex).toBe(4)
  })

  it('is pure: 3× byte-identical and inputs never mutated', () => {
    const { projection } = projectionOf()
    const before = JSON.stringify(projection)
    const run = () => JSON.stringify(buildGanttBars(projection, VIEWPORT, fullWindow()))
    expect(run()).toBe(run())
    expect(run()).toBe(run())
    expect(JSON.stringify(projection)).toBe(before)
  })
})
