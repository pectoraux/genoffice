import { describe, expect, it } from 'vitest'
import { asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import {
  buildDependencies,
  buildRowWindow,
  createViewState,
  projectDocumentView,
  reduceViewState,
} from '../../src/index.js'
import {
  ganttDocument,
  ganttSchedule,
  makeDependency,
  makeDocument,
  makeScheduleEntry,
  makeTask,
} from '../fixtures.js'

const VIEWPORT = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-31T00:00:00.000Z' }
const START_MS = Date.parse(VIEWPORT.start)
const SPAN_MS = Date.parse(VIEWPORT.finish) - START_MS
const f = (iso: string): number => (Date.parse(iso) - START_MS) / SPAN_MS
const fullWindow = (count: number) => buildRowWindow(count, { firstRow: 0, visibleRows: count })

/** A four-leaf chain t1..t4 with fully controlled schedules, so each
 * relationship type's edge anchors are unambiguous. */
function chainFixture() {
  const document = makeDocument({
    tasks: [
      makeTask({ id: 't1', uid: 1, duration: asWorkingMinutes(480) }),
      makeTask({ id: 't2', uid: 2, duration: asWorkingMinutes(480) }),
      makeTask({ id: 't3', uid: 3, duration: asWorkingMinutes(480) }),
      makeTask({ id: 't4', uid: 4, duration: asWorkingMinutes(480) }),
    ],
    dependencies: [
      makeDependency('fs', 't1', 't2'), // finish → start
      makeDependency('ss', 't2', 't3', 'SS'), // start → start
      makeDependency('ff', 't3', 't4', 'FF'), // finish → finish
      makeDependency('sf', 't1', 't4', 'SF'), // start → finish
    ],
    startDate: '2026-08-01T00:00:00.000Z',
  })
  const schedule = {
    taskSchedules: {
      [asTaskId('t1')]: makeScheduleEntry(
        't1',
        '2026-08-02T00:00:00.000Z',
        '2026-08-03T00:00:00.000Z',
      ),
      [asTaskId('t2')]: makeScheduleEntry(
        't2',
        '2026-08-04T00:00:00.000Z',
        '2026-08-05T00:00:00.000Z',
      ),
      [asTaskId('t3')]: makeScheduleEntry(
        't3',
        '2026-08-04T00:00:00.000Z',
        '2026-08-06T00:00:00.000Z',
      ),
      [asTaskId('t4')]: makeScheduleEntry(
        't4',
        '2026-08-05T00:00:00.000Z',
        '2026-08-06T00:00:00.000Z',
      ),
    },
    diagnostics: [],
  }
  return { document, schedule }
}

describe('PROJECT-022 dependency links — edge anchors and routes', () => {
  it('anchors each relationship type at the documented edges with the four-point elbow route', () => {
    const { document, schedule } = chainFixture()
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const links = buildDependencies(document, projection, VIEWPORT, fullWindow(4))
    const byId = new Map(links.map((link) => [link.dependencyId, link]))

    const fs = byId.get('fs' as never)!
    expect(fs.type).toBe('FS')
    expect(fs.from).toMatchObject({ taskId: asTaskId('t1'), rowIndex: 0, edge: 'finish' })
    expect(fs.from.fraction).toBeCloseTo(f('2026-08-03T00:00:00.000Z'), 10)
    expect(fs.to).toMatchObject({ taskId: asTaskId('t2'), rowIndex: 1, edge: 'start' })
    expect(fs.to.fraction).toBeCloseTo(f('2026-08-04T00:00:00.000Z'), 10)
    const mid = (fs.from.fraction + fs.to.fraction) / 2
    expect(fs.route).toEqual([
      { fraction: fs.from.fraction, rowIndex: 0 },
      { fraction: mid, rowIndex: 0 },
      { fraction: mid, rowIndex: 1 },
      { fraction: fs.to.fraction, rowIndex: 1 },
    ])

    expect(byId.get('ss' as never)!.from).toMatchObject({ edge: 'start' })
    expect(byId.get('ss' as never)!.to).toMatchObject({ edge: 'start' })
    expect(byId.get('ff' as never)!.from).toMatchObject({ edge: 'finish' })
    expect(byId.get('ff' as never)!.to).toMatchObject({ edge: 'finish' })
    // SF is the inverted anchor: predecessor START → successor FINISH.
    const sf = byId.get('sf' as never)!
    expect(sf.from).toMatchObject({ taskId: asTaskId('t1'), edge: 'start' })
    expect(sf.to).toMatchObject({ taskId: asTaskId('t4'), edge: 'finish' })
  })

  it('keeps the CANONICAL endpoint ids on the link even after ancestor resolution', () => {
    const document = ganttDocument()
    let state = createViewState(document)
    // Collapse `a`: a1/a2 hidden; d1 (a1→a2) resolves both ends to `a`.
    state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('a') }, { document })
    const projection = projectDocumentView(document, ganttSchedule(), state)
    const links = buildDependencies(document, projection, VIEWPORT, fullWindow(4))
    const d2 = links.find((link) => link.dependencyId === ('d2' as never))!
    // d2 = a2→b SS with `a` collapsed: from resolves to the ancestor `a`:
    expect(d2.from).toMatchObject({ taskId: asTaskId('a'), edge: 'start' })
    expect(d2.from.fraction).toBeCloseTo(f('2026-08-03T09:00:00.000Z'), 10)
    // …but the canonical endpoints stay on the link record:
    expect(d2.predecessorTaskId).toBe(asTaskId('a2'))
    expect(d2.successorTaskId).toBe(asTaskId('b'))
  })

  it('omits a link whose endpoints resolve to the SAME row (inside one collapsed subtree)', () => {
    const document = ganttDocument()
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('a') }, { document })
    const projection = projectDocumentView(document, ganttSchedule(), state)
    const links = buildDependencies(document, projection, VIEWPORT, fullWindow(4))
    expect(links.map((link) => link.dependencyId)).not.toContain('d1' as never)
  })

  it('omits links whose endpoint row is scrolled OUT of the virtualized window', () => {
    const { document, schedule } = chainFixture()
    const projection = projectDocumentView(document, schedule, createViewState(document))
    // A window that shows only t3 and t4: every link touching t1/t2 is
    // omitted; the fully in-window t3→t4 FF link survives.
    const window = buildRowWindow(4, { firstRow: 2, visibleRows: 2 })
    const links = buildDependencies(document, projection, VIEWPORT, window)
    expect(links.map((link) => link.type)).toEqual(['FF'])
    // Widening to t2..t4 adds the SS link (t2→t3); the links anchored at t1
    // (rows outside the window) stay omitted:
    const window2 = buildRowWindow(4, { firstRow: 1, visibleRows: 3 })
    const links2 = buildDependencies(document, projection, VIEWPORT, window2)
    expect(links2.map((link) => link.type)).toEqual(['SS', 'FF'])
  })

  it('attaches a collapsed endpoint to the NEAREST visible ancestor, not a farther one', () => {
    // root > a > (b > t1, t2): collapsing `b` hides only t1; the link t1→t2
    // must resolve t1 to `b` (the collapsed summary itself — nearest), never
    // to `a` or `root`, while t2 keeps its own row.
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'root', outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({
          id: 'a',
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          summary: true,
          wbs: '1.1',
        }),
        makeTask({
          id: 'b',
          parentTaskId: asTaskId('a'),
          outlineLevel: 3,
          summary: true,
          wbs: '1.1.1',
        }),
        makeTask({ id: 't1', parentTaskId: asTaskId('b'), outlineLevel: 4, wbs: '1.1.1.1' }),
        makeTask({ id: 't2', parentTaskId: asTaskId('a'), outlineLevel: 3, wbs: '1.1.2' }),
      ],
      dependencies: [makeDependency('d', 't1', 't2')],
      startDate: '2026-08-01T00:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {
        [asTaskId('root')]: makeScheduleEntry(
          'root',
          '2026-08-02T00:00:00.000Z',
          '2026-08-10T00:00:00.000Z',
        ),
        [asTaskId('a')]: makeScheduleEntry(
          'a',
          '2026-08-02T00:00:00.000Z',
          '2026-08-10T00:00:00.000Z',
        ),
        [asTaskId('b')]: makeScheduleEntry(
          'b',
          '2026-08-03T00:00:00.000Z',
          '2026-08-04T00:00:00.000Z',
        ),
        [asTaskId('t1')]: makeScheduleEntry(
          't1',
          '2026-08-03T00:00:00.000Z',
          '2026-08-04T00:00:00.000Z',
        ),
        [asTaskId('t2')]: makeScheduleEntry(
          't2',
          '2026-08-05T00:00:00.000Z',
          '2026-08-06T00:00:00.000Z',
        ),
      },
      diagnostics: [],
    }
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'toggleCollapse', taskId: asTaskId('b') }, { document })
    const projection = projectDocumentView(document, schedule, state)
    const links = buildDependencies(document, projection, VIEWPORT, fullWindow(4))
    expect(links).toHaveLength(1)
    expect(links[0]!.from).toMatchObject({ taskId: asTaskId('b'), edge: 'finish' })
    expect(links[0]!.from.fraction).toBeCloseTo(f('2026-08-04T00:00:00.000Z'), 10)
    expect(links[0]!.to).toMatchObject({ taskId: asTaskId('t2'), edge: 'start' })
    expect(links[0]!.predecessorTaskId).toBe(asTaskId('t1')) // canonical id preserved
  })

  it('omits links whose resolved row lacks the needed schedule instant — never an invented position', () => {
    const { document, schedule } = chainFixture()
    delete schedule.taskSchedules[asTaskId('t2')]
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const links = buildDependencies(document, projection, VIEWPORT, fullWindow(4))
    expect(links.map((link) => link.dependencyId)).not.toContain('fs' as never)
    expect(links.map((link) => link.dependencyId)).not.toContain('ss' as never)
    // t3→t4 and t1→t4 survive:
    expect(links).toHaveLength(2)
  })

  it('returns nothing for an empty window or a degenerate viewport', () => {
    const { document, schedule } = chainFixture()
    const projection = projectDocumentView(document, schedule, createViewState(document))
    expect(
      buildDependencies(document, projection, VIEWPORT, { firstIndex: 0, lastIndex: -1 }),
    ).toEqual([])
    expect(
      buildDependencies(
        document,
        projection,
        { start: 'bad', finish: '2026-08-31T00:00:00.000Z' },
        fullWindow(4),
      ),
    ).toEqual([])
  })

  it('is pure: 3× byte-identical, document order preserved, inputs untouched', () => {
    const { document, schedule } = chainFixture()
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const documentBefore = JSON.stringify(document)
    const run = () =>
      JSON.stringify(buildDependencies(document, projection, VIEWPORT, fullWindow(4)))
    expect(run()).toBe(run())
    expect(run()).toBe(run())
    expect(JSON.stringify(document)).toBe(documentBefore)
    const links = buildDependencies(document, projection, VIEWPORT, fullWindow(4))
    expect(links.map((link) => link.dependencyId)).toEqual([
      'fs' as never,
      'ss' as never,
      'ff' as never,
      'sf' as never,
    ])
  })
})
