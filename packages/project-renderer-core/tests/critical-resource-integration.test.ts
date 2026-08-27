/**
 * PROJECT-026 — critical-path / resource visualization integration with
 * the REAL scheduling authority (the no-second-engine evidence).
 *
 * The canonical allocation query is bound to the REAL scheduling package
 * exactly as hosts bind it (the `ScheduleRunner` /
 * `CalendarWorkingTimeQuery` injection precedents — the scheduling import
 * lives at the TEST layer only):
 *
 *     (document, schedule) => resourceAllocations(document, schedule)
 *
 * Every float value and every utilization band therefore IS the
 * authority's answer: the critical-path surface's floats equal the REAL
 * scheduler's own `TaskSchedule` values (the "visual layers match derived
 * schedule" acceptance evidence), the resource bands equal the REAL
 * allocation output clipped to the viewport (never re-derived), the
 * over-allocation windows equal the REAL leveler's conflict record, and
 * the degradation paths mirror the authority's own error boundaries. The
 * battery also proves the full pipeline end-to-end (session → projection →
 * `buildGanttView` with BOTH the calendar and the resource inputs), the
 * additive-surface geometry-neutrality contract, and the
 * absent-without-input degradation.
 */
import { describe, expect, it } from 'vitest'
import { asCalendarId, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { resolveCalendar, schedule, workingIntervals } from '@genoffice/project-scheduling'
import { levelResources, resourceAllocations } from '@genoffice/project-scheduling'
import type { CalendarWorkingTimeQuery, ResourceAllocationQuery } from '../src/index.js'
import {
  applyRendererCommand,
  buildGanttView,
  buildResourceUtilization,
  createRendererSession,
  createViewState,
  projectDocumentView,
  undoRendererCommand,
} from '../src/index.js'
import {
  makeAssignment,
  makeCalendar,
  makeDependency,
  makeDocument,
  makeResource,
  makeTask,
} from './fixtures.js'

/** The canonical bindings — the documented host-side adapters over the REAL
 * scheduling package (the scheduling import lives at the test layer only;
 * the renderer package under test stays scheduling-free). */
const canonicalAllocation: ResourceAllocationQuery = (document, derived) =>
  resourceAllocations(document, derived)
const canonicalWorkingTime: CalendarWorkingTimeQuery = (calendars, calendarId, start, finish) =>
  workingIntervals(resolveCalendar({ calendars: [...calendars] }, calendarId), start, finish)

const id = (value: string) => asTaskId(value)

/**
 * The critical-path pipeline fixture: a real FS chain `a → b → c → m`
 * (one working day each; `m` a zero-duration milestone) that defines the
 * project finish — every chain task is critical — plus an independent
 * branch task `d` (one day, linked FS into `m`) that carries real float.
 */
const chainDocument = (): ProjectDocument =>
  makeDocument({
    startDate: '2026-08-03T09:00:00.000Z',
    tasks: [
      makeTask({ id: 'a', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'b', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'c', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'd', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'm', milestone: true, duration: asWorkingMinutes(0) }),
    ],
    dependencies: [
      makeDependency('d1', 'a', 'b'),
      makeDependency('d2', 'b', 'c'),
      makeDependency('d3', 'c', 'm'),
      makeDependency('d4', 'd', 'm'),
    ],
  })

/** The over-allocation pipeline fixture: two same-day tasks on one 100%
 * work resource (100% + 60%). */
const overloadDocument = (): ProjectDocument =>
  makeDocument({
    startDate: '2026-08-03T09:00:00.000Z',
    tasks: [
      makeTask({ id: 't1', duration: asWorkingMinutes(480) }),
      makeTask({ id: 't2', duration: asWorkingMinutes(480) }),
    ],
    resources: [makeResource({ id: 'r1', name: 'Builder' })],
    assignments: [
      makeAssignment('a1', 't1', 'r1'),
      makeAssignment('a2', 't2', 'r1', { units: 0.6 }),
    ],
  })

describe('PROJECT-026 integration — floats ARE the canonical schedule', () => {
  it("echoes the REAL scheduler's criticality/float values verbatim (the chain + branch fixture)", () => {
    const document = chainDocument()
    const session = createRendererSession(document, { schedule })
    expect(session.schedule?.diagnostics).toEqual([])
    const state = createViewState(document, session.schedule)
    const projection = projectDocumentView(document, session.schedule, state)
    const view = buildGanttView(document, projection, state, { firstRow: 0, visibleRows: 10 })
    const floats = new Map(
      view.timeline.criticalPath?.floats.map((float) => [float.taskId as string, float]) ?? [],
    )
    // Every scheduled task carries a float entry.
    expect([...floats.keys()].sort()).toEqual(['a', 'b', 'c', 'd', 'm'])
    // The chain tasks + milestone are critical; the branch is not.
    expect(floats.get('a')?.critical).toBe(true)
    expect(floats.get('b')?.critical).toBe(true)
    expect(floats.get('c')?.critical).toBe(true)
    expect(floats.get('m')?.critical).toBe(true)
    expect(floats.get('d')?.critical).toBe(false)
    // The equality evidence: every echoed value equals the REAL scheduler's
    // own TaskSchedule object for that task — never recomputed.
    for (const float of floats.values()) {
      const authoritative = session.schedule?.taskSchedules[float.taskId]
      expect(authoritative).toBeDefined()
      expect(float.critical).toBe(authoritative?.critical)
      expect(float.totalSlack).toBe(authoritative?.totalSlack)
      expect(float.freeSlack).toBe(authoritative?.freeSlack)
    }
    // The branch carries two working days of total slack (Mon 17:00 →
    // Wed 17:00) — the authority's signed working-minute answer.
    expect(floats.get('d')?.totalSlack).toBe(960)
    // The slack bar spans the branch's scheduled finish to its canonical
    // late finish; the zero-slack chain tasks carry no slack geometry.
    expect(floats.get('d')?.slack).toBeDefined()
    expect(floats.get('a')?.slack).toBeUndefined()
    expect(floats.get('m')?.slack).toBeUndefined()
  })

  it('classifies the chain links critical and leaves the branch link out (both-endpoints convention over the real flags)', () => {
    const document = chainDocument()
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)
    const projection = projectDocumentView(document, session.schedule, state)
    const view = buildGanttView(document, projection, state, { firstRow: 0, visibleRows: 10 })
    // d1/d2/d3 connect critical endpoints; d4's predecessor `d` is not
    // critical — the both-endpoints projection, never a driving-path claim.
    expect(
      view.timeline.criticalPath?.criticalDependencyIds.map((value) => value as string),
    ).toEqual(['d1', 'd2', 'd3'])
  })

  it('re-projects the floats onto the new derived schedule through the session (command → schedule → surface)', () => {
    const document = chainDocument()
    let session = createRendererSession(document, { schedule })
    const build = () => {
      const state = createViewState(session.document, session.schedule)
      const projection = projectDocumentView(session.document, session.schedule, state)
      return buildGanttView(session.document, projection, state, { firstRow: 0, visibleRows: 10 })
    }
    const before = build()
    const beforeFloats = new Map(
      before.timeline.criticalPath?.floats.map((float) => [float.taskId as string, float]) ?? [],
    )
    expect(beforeFloats.get('d')?.totalSlack).toBe(960)
    // Double the branch task's duration: it now runs Mon 09:00 → Tue 17:00,
    // leaving one working day of float to the (unchanged) project finish.
    const outcome = applyRendererCommand(session, {
      type: 'SetTaskDuration',
      taskId: id('d'),
      duration: asWorkingMinutes(960),
    })
    expect(outcome.result.accepted).toBe(true)
    session = outcome.session
    const after = build()
    const afterFloats = new Map(
      after.timeline.criticalPath?.floats.map((float) => [float.taskId as string, float]) ?? [],
    )
    expect(afterFloats.get('d')?.totalSlack).toBe(480)
    // Every value still equals the NEW authority schedule verbatim.
    for (const float of afterFloats.values()) {
      const authoritative = session.schedule?.taskSchedules[float.taskId]
      expect(float.critical).toBe(authoritative?.critical)
      expect(float.totalSlack).toBe(authoritative?.totalSlack)
      expect(float.freeSlack).toBe(authoritative?.freeSlack)
    }
    // Undo restores the exact prior surface (snapshot schedule restore).
    const undone = undoRendererCommand(session)
    expect(undone.applied).toBe(true)
    session = undone.session
    const restored = build()
    expect(JSON.stringify(restored.timeline.criticalPath)).toBe(
      JSON.stringify(before.timeline.criticalPath),
    )
  })
})

describe('PROJECT-026 integration — resource bands ARE the canonical allocation', () => {
  it('equals the REAL resourceAllocations output clipped to the viewport (never re-derived)', () => {
    const document = overloadDocument()
    const derived = schedule(document)
    const viewport = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-04T00:00:00.000Z' }
    const surface = buildResourceUtilization(document, derived, canonicalAllocation, viewport)
    expect(surface.status).toBe('ok')
    // The authority's own output for the same document/schedule.
    const authoritative = resourceAllocations(document, derived)
    expect(authoritative).toHaveLength(1)
    const segments = authoritative[0]?.segments ?? []
    expect(segments.length).toBeGreaterThan(0)
    // The surface's bands are those segments clipped to the Monday viewport
    // — the one-day overload tiles into a single in-viewport band.
    expect(surface.resources).toEqual([
      {
        resourceId: asTaskId('r1') as never,
        name: 'Builder',
        bands: segments
          .filter((segment) => segment.finish > viewport.start && segment.start < viewport.finish)
          .map((segment) => ({
            resourceId: authoritative[0]?.resourceId,
            start: segment.start > viewport.start ? (segment.start as string) : viewport.start,
            finish: segment.finish < viewport.finish ? (segment.finish as string) : viewport.finish,
            demandUnits: segment.demandUnits,
            capacityUnits: segment.capacityUnits,
            overallocated: segment.overallocated,
            assignmentIds: segment.assignmentIds,
          })),
      },
    ])
    // The Monday band is over-allocated (1.6 > 1) with BOTH assignments.
    const band = surface.resources?.[0]?.bands?.[0]
    expect(band).toMatchObject({
      demandUnits: 1.6,
      capacityUnits: 1,
      overallocated: true,
    })
    expect(band?.assignmentIds.map((value) => value as string)).toEqual(['a1', 'a2'])
  })

  it("the over-allocation windows equal the REAL leveler's conflict record (the same sweep)", () => {
    const document = overloadDocument()
    const derived = schedule(document)
    const viewport = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-05T00:00:00.000Z' }
    const surface = buildResourceUtilization(document, derived, canonicalAllocation, viewport)
    const overallocatedBands = (surface.resources ?? []).flatMap((entry) =>
      (entry.bands ?? []).filter((band) => band.overallocated),
    )
    const conflicts = levelResources(document).overallocations
    expect(conflicts).toHaveLength(1)
    expect(overallocatedBands.map((band) => [band.start, band.finish])).toEqual(
      conflicts.map((conflict) => [
        conflict.window.start as string,
        conflict.window.finish as string,
      ]),
    )
    expect(overallocatedBands[0]?.demandUnits).toBe(conflicts[0]?.peakDemand)
    expect(overallocatedBands[0]?.capacityUnits).toBe(conflicts[0]?.maxUnits)
  })

  it('degrades gracefully when the SCHEDULER itself rejects a broken resource calendar (the stricter authority)', () => {
    // The scheduler's own validation is STRICTER than the allocation
    // authority's: a resource referencing a missing calendar degrades
    // `schedule()` itself (MISSING_CALENDAR, "Resource r1 references missing
    // calendar ghost") — so the session carries a degraded schedule, there
    // is no demand to allocate, and the whole visualization pipeline
    // degrades to EMPTY surfaces (never a crash, never invented values).
    // The renderer's coded-error boundary itself (the 025 diagnostic-echo
    // mirror) is covered by the unit battery with a synthetic coded query —
    // through the REAL binding the scheduler's validation always fires
    // first, a strictly stronger guarantee.
    const document = makeDocument({
      calendars: [makeCalendar('standard')],
      startDate: '2026-08-03T09:00:00.000Z',
      tasks: [makeTask({ id: 't1', duration: asWorkingMinutes(480) })],
      resources: [makeResource({ id: 'r1', calendarId: asCalendarId('ghost') })],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    const session = createRendererSession(document, { schedule })
    expect(session.schedule?.diagnostics).toEqual([
      {
        code: 'MISSING_CALENDAR',
        severity: 'error',
        message: 'Resource r1 references missing calendar ghost',
      },
    ])
    const state = createViewState(document, session.schedule)
    const projection = projectDocumentView(document, session.schedule, state)
    // The schedule's diagnostics are echoed on the projection for the
    // host's diagnostics surface.
    expect(projection.scheduleDiagnostics).toEqual(session.schedule?.diagnostics)
    const view = buildGanttView(
      document,
      projection,
      state,
      { firstRow: 0, visibleRows: 10 },
      undefined,
      { allocation: canonicalAllocation, schedule: session.schedule! },
    )
    expect(view.timeline.criticalPath).toEqual({ floats: [], criticalDependencyIds: [] })
    expect(view.timeline.resourceUtilization).toMatchObject({ status: 'ok', resources: [] })
  })

  it('an empty allocation over a degraded schedule yields an OK empty surface (never a crash)', () => {
    // A cyclic DEFAULT calendar degrades the schedule itself: the session
    // carries no task schedules, so there is no criticality and no demand.
    const cyclicA = makeCalendar('cycA', { baseCalendarId: asCalendarId('cycB') })
    const cyclicB = makeCalendar('cycB', { baseCalendarId: asCalendarId('cycA') })
    const document = makeDocument({
      calendars: [cyclicA, cyclicB],
      tasks: [makeTask({ id: 't1', duration: asWorkingMinutes(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    const session = createRendererSession(document, { schedule })
    expect(session.schedule?.diagnostics[0]?.code).toBe('CALENDAR_CYCLE')
    const state = createViewState(document, session.schedule)
    const projection = projectDocumentView(document, session.schedule, state)
    const view = buildGanttView(
      document,
      projection,
      state,
      { firstRow: 0, visibleRows: 10 },
      undefined,
      { allocation: canonicalAllocation, schedule: session.schedule! },
    )
    // No schedules → no floats, no critical links — but the surface exists
    // (a schedule was joined) and is EMPTY, never invented.
    expect(view.timeline.criticalPath).toEqual({ floats: [], criticalDependencyIds: [] })
    // No demand intervals → the allocation authority returns nothing → an
    // OK surface with no resource entries.
    expect(view.timeline.resourceUtilization).toMatchObject({ status: 'ok', resources: [] })
  })
})

describe('PROJECT-026 integration — the additive pipeline (calendar + resources together)', () => {
  it('threads both inputs through buildGanttView with byte-identical geometry (additivity)', () => {
    const document = overloadDocument()
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)
    const projection = projectDocumentView(document, session.schedule, state)
    const layout = { firstRow: 0, visibleRows: 10 }
    const resources = { allocation: canonicalAllocation, schedule: session.schedule! }
    const calendar = { workingTime: canonicalWorkingTime }
    const withBoth = buildGanttView(document, projection, state, layout, calendar, resources)
    const withNone = buildGanttView(document, projection, state, layout)
    // Both PROJECT-026 surfaces present alongside the PROJECT-025 surfaces.
    expect(withBoth.timeline.criticalPath).toBeDefined()
    expect(withBoth.timeline.resourceUtilization).toBeDefined()
    expect(withBoth.timeline.calendar).toBeDefined()
    expect(withBoth.timeline.rowCalendars).toBeDefined()
    // Absent without the inputs / without a schedule — never invented.
    expect(withNone.timeline.resourceUtilization).toBeUndefined()
    expect(withNone.timeline.calendar).toBeUndefined()
    expect(withNone.timeline.rowCalendars).toBeUndefined()
    // The critical-path surface joins from the projection's schedule alone.
    expect(withNone.timeline.criticalPath).toEqual(withBoth.timeline.criticalPath)
    // ADDITIVITY: every other surface is byte-identical with and without
    // the threaded inputs (projection never feeds geometry).
    const strip = (view: typeof withBoth) => {
      const timeline = { ...view.timeline }
      delete (timeline as Record<string, unknown>).criticalPath
      delete (timeline as Record<string, unknown>).resourceUtilization
      delete (timeline as Record<string, unknown>).calendar
      delete (timeline as Record<string, unknown>).rowCalendars
      return { ...view, timeline }
    }
    expect(strip(withBoth)).toEqual(strip(withNone))
  })

  it('follows the schedule through session commands and replays 3× byte-identical (end-to-end determinism)', () => {
    const document = overloadDocument()
    let session = createRendererSession(document, { schedule })
    const build = () => {
      const state = createViewState(session.document, session.schedule)
      const projection = projectDocumentView(session.document, session.schedule, state)
      return buildGanttView(
        session.document,
        projection,
        state,
        { firstRow: 0, visibleRows: 10 },
        { workingTime: canonicalWorkingTime },
        { allocation: canonicalAllocation, schedule: session.schedule! },
      )
    }
    const first = JSON.stringify(build())
    for (let i = 0; i < 2; i += 1) {
      expect(JSON.stringify(build())).toBe(first)
    }
    // A command moves the schedule; both surfaces follow the NEW schedule.
    const outcome = applyRendererCommand(session, {
      type: 'SetTaskDuration',
      taskId: id('t2'),
      duration: asWorkingMinutes(960),
    })
    expect(outcome.result.accepted).toBe(true)
    session = outcome.session
    const moved = build()
    const bands = moved.timeline.resourceUtilization?.resources?.[0]?.bands ?? []
    // t2 now runs Mon 09:00 → Tue 17:00; the demand overlaps t1 only on
    // Monday (1.6 > 1 over-allocated) and stands alone on Tuesday.
    const monday = bands.find((band) => band.start === '2026-08-03T09:00:00.000Z')
    const tuesday = bands.find((band) => band.start === '2026-08-04T09:00:00.000Z')
    expect(monday).toMatchObject({ demandUnits: 1.6, overallocated: true })
    expect(tuesday).toMatchObject({ demandUnits: 0.6, overallocated: false })
    const again = JSON.stringify(build())
    expect(again).toBe(JSON.stringify(build()))
    expect(again).not.toBe(first)
  })
})
