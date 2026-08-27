import { describe, expect, it } from 'vitest'
import { asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import {
  buildGanttBars,
  buildMilestones,
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
const f = (iso: string): number => (Date.parse(iso) - START_MS) / SPAN_MS
const fullWindow = () => buildRowWindow(6, { firstRow: 0, visibleRows: 6 })

const projectionOf = (document = ganttDocument(), schedule = ganttSchedule()) =>
  projectDocumentView(document, schedule, createViewState(document))

describe('PROJECT-022 milestones — geometry', () => {
  it('renders the flagged milestone at its scheduled start instant', () => {
    const milestones = buildMilestones(projectionOf(), VIEWPORT, fullWindow())
    expect(milestones).toHaveLength(1)
    expect(milestones[0]).toMatchObject({
      taskId: asTaskId('m'),
      rowIndex: 5,
      instant: '2026-08-07T09:00:00.000Z',
      beforeViewport: false,
      afterViewport: false,
    })
    expect(milestones[0]!.fraction).toBeCloseTo(f('2026-08-07T09:00:00.000Z'), 10)
  })

  it('renders an UNFLAGGED zero-duration task as a milestone (zero-span schedule)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'z', duration: asWorkingMinutes(0) })],
      startDate: '2026-08-01T00:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {
        [asTaskId('z')]: makeScheduleEntry(
          'z',
          '2026-08-06T12:00:00.000Z',
          '2026-08-06T12:00:00.000Z',
        ),
      },
      diagnostics: [],
    }
    const milestones = buildMilestones(
      projectionOf(document, schedule),
      VIEWPORT,
      buildRowWindow(1, { firstRow: 0, visibleRows: 1 }),
    )
    expect(milestones).toHaveLength(1)
    expect(milestones[0]!.instant).toBe('2026-08-06T12:00:00.000Z')
  })

  it('renders a milestone-like schedule even for a task with a nonzero duration field', () => {
    // The schedule window collapses to zero — milestone geometry by span:
    const document = makeDocument({
      tasks: [makeTask({ id: 'z', duration: asWorkingMinutes(480) })],
      startDate: '2026-08-01T00:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {
        [asTaskId('z')]: makeScheduleEntry(
          'z',
          '2026-08-06T12:00:00.000Z',
          '2026-08-06T12:00:00.000Z',
        ),
      },
      diagnostics: [],
    }
    const milestones = buildMilestones(
      projectionOf(document, schedule),
      VIEWPORT,
      buildRowWindow(1, { firstRow: 0, visibleRows: 1 }),
    )
    expect(milestones).toHaveLength(1)
  })

  it('keeps BOTH bar and milestone for a flagged milestone with a real span (orthogonal rule)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'm2', milestone: true, duration: asWorkingMinutes(960) })],
      startDate: '2026-08-01T00:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {
        [asTaskId('m2')]: makeScheduleEntry(
          'm2',
          '2026-08-03T09:00:00.000Z',
          '2026-08-05T17:00:00.000Z',
        ),
      },
      diagnostics: [],
    }
    const projection = projectionOf(document, schedule)
    const window = buildRowWindow(1, { firstRow: 0, visibleRows: 1 })
    const bars = buildGanttBars(projection, VIEWPORT, window)
    const milestones = buildMilestones(projection, VIEWPORT, window)
    expect(bars.map((bar) => bar.taskId)).toEqual([asTaskId('m2')])
    expect(milestones.map((milestone) => milestone.taskId)).toEqual([asTaskId('m2')])
    expect(milestones[0]!.instant).toBe('2026-08-03T09:00:00.000Z') // the diamond sits at the start
  })

  it('produces NO milestone without a schedule — the flag alone invents nothing', () => {
    const projection = projectDocumentView(
      ganttDocument(),
      undefined,
      createViewState(ganttDocument()),
    )
    expect(buildMilestones(projection, VIEWPORT, fullWindow())).toEqual([])
  })

  it('flags milestones outside the viewport and clamps the fraction', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'm1', milestone: true }), makeTask({ id: 'm2', milestone: true })],
      startDate: '2026-08-01T00:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {
        [asTaskId('m1')]: makeScheduleEntry(
          'm1',
          '2026-07-01T00:00:00.000Z',
          '2026-07-01T00:00:00.000Z',
        ),
        [asTaskId('m2')]: makeScheduleEntry(
          'm2',
          '2026-09-30T00:00:00.000Z',
          '2026-09-30T00:00:00.000Z',
        ),
      },
      diagnostics: [],
    }
    const milestones = buildMilestones(
      projectionOf(document, schedule),
      VIEWPORT,
      buildRowWindow(2, { firstRow: 0, visibleRows: 2 }),
    )
    const byTask = new Map(milestones.map((milestone) => [milestone.taskId, milestone]))
    expect(byTask.get(asTaskId('m1'))).toMatchObject({
      fraction: 0,
      beforeViewport: true,
      afterViewport: false,
    })
    expect(byTask.get(asTaskId('m2'))).toMatchObject({
      fraction: 1,
      beforeViewport: false,
      afterViewport: true,
    })
  })

  it('slices to the virtualized window and is pure (3× byte-identical, inputs untouched)', () => {
    const projection = projectionOf()
    const window = buildRowWindow(6, { firstRow: 5, visibleRows: 1 })
    expect(buildMilestones(projection, VIEWPORT, window).map((m) => m.taskId)).toEqual([
      asTaskId('m'),
    ])
    expect(buildMilestones(projection, VIEWPORT, { firstIndex: 0, lastIndex: -1 })).toEqual([])
    const before = JSON.stringify(projection)
    const run = () => JSON.stringify(buildMilestones(projection, VIEWPORT, fullWindow()))
    expect(run()).toBe(run())
    expect(run()).toBe(run())
    expect(JSON.stringify(projection)).toBe(before)
  })
})
