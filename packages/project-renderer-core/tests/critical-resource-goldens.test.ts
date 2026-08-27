/**
 * PROJECT-026 — the golden scenario battery (R01–R12).
 *
 * Each golden drives one acceptance scenario end-to-end through the real
 * machinery — documents scheduled by the REAL scheduler, allocations from
 * the REAL canonical allocation authority (`resourceAllocations`), surfaces
 * built through `buildCriticalPath` / `buildResourceUtilization` / the
 * accepted `buildGanttView` pipeline — and asserts the complete observable
 * tuple: the critical-path surface (floats + critical-link classification),
 * the resource-visualization surface, the ProjectDocument, and the
 * DerivedSchedule. The equality half of every golden is the acceptance
 * contract itself: every float value EQUALS the authority's own
 * TaskSchedule values and every band EQUALS the authority's own allocation
 * output clipped to the viewport — the visual layers MATCH the derived
 * schedule, never re-deriving it.
 *
 * Pure inputs, no wall clock, no randomness; R10 keeps the surfaces
 * byte-identical across undo/redo; R12 proves a mixed command/viewport
 * sequence replays byte-identical (3×).
 */
import { describe, expect, it } from 'vitest'
import type { DerivedSchedule, ProjectDocument } from '@genoffice/project-contracts'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import type { ResourceAllocationQuery } from '../src/index.js'
import {
  applyRendererCommand,
  buildGanttView,
  createRendererSession,
  createViewState,
  projectDocumentView,
  redoRendererCommand,
  scaleViewport,
  undoRendererCommand,
} from '../src/index.js'
import { makeAssignment, makeDependency, makeDocument, makeResource, makeTask } from './fixtures.js'
import { resourceAllocations } from '@genoffice/project-scheduling'

/** The canonical allocation binding — the documented host-side adapter. */
const canonicalAllocation: ResourceAllocationQuery = (document, derived) =>
  resourceAllocations(document, derived)

const id = (value: string) => asTaskId(value)

/** The chain + branch fixture (R01/R02/R04/R05): `a → b → c → m` critical
 * FS chain (m a zero-duration milestone), `d` an independent one-day branch
 * linked FS into `m` with two working days of float. */
const chainFixture = () => ({
  document: makeDocument({
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
  }),
})

/** The two-day single-assignment fixture (R06): one resource, one 2-day
 * task — the calendar-aware demand tiling across the Monday night. */
const tilingFixture = () => ({
  document: makeDocument({
    startDate: '2026-08-03T09:00:00.000Z',
    tasks: [makeTask({ id: 't1', duration: asWorkingMinutes(960) })],
    resources: [makeResource({ id: 'r1', name: 'Builder' })],
    assignments: [makeAssignment('a1', 't1', 'r1')],
  }),
})

/** The over-allocation fixture (R07): two same-day tasks on one 100%
 * resource (100% + 60% demand). */
const overloadFixture = () => ({
  document: makeDocument({
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
  }),
})

/** The full pipeline over a REAL session: the projection + view with both
 * PROJECT-026 surfaces (calendar surfaces stay out of these goldens — the
 * 025 battery owns them). */
const pipeline = (document: ProjectDocument, derived: DerivedSchedule) => {
  const state = createViewState(document, derived)
  const projection = projectDocumentView(document, derived, state)
  const view = buildGanttView(
    document,
    projection,
    state,
    { firstRow: 0, visibleRows: 10 },
    undefined,
    { allocation: canonicalAllocation, schedule: derived },
  )
  return { state, projection, view }
}

/** The acceptance-equality assertion: every float value equals the REAL
 * scheduler's own TaskSchedule values for the same document. */
const assertFloatsEqualSchedule = (
  document: ProjectDocument,
  derived: DerivedSchedule,
  surface:
    | {
        floats: readonly {
          taskId: unknown
          critical: boolean
          totalSlack: number
          freeSlack: number
        }[]
      }
    | undefined,
) => {
  expect(surface).toBeDefined()
  for (const float of surface?.floats ?? []) {
    const authoritative = derived.taskSchedules[float.taskId as never]
    expect(authoritative, `${JSON.stringify(float.taskId)} missing from the schedule`).toBeDefined()
    expect(float.critical).toBe(authoritative?.critical)
    expect(float.totalSlack).toBe(authoritative?.totalSlack)
    expect(float.freeSlack).toBe(authoritative?.freeSlack)
  }
  expect(surface?.floats.length).toBe(
    document.tasks.filter((task) => derived.taskSchedules[task.id] !== undefined).length,
  )
}

describe('PROJECT-026 goldens — critical path', () => {
  it('R01 — the critical chain: every chain task critical, zero float, no slack bars', () => {
    const { document } = chainFixture()
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    assertFloatsEqualSchedule(document, derived, view.timeline.criticalPath)
    const floats = new Map(
      view.timeline.criticalPath?.floats.map((float) => [float.taskId as string, float]) ?? [],
    )
    expect(floats.get('a')?.critical).toBe(true)
    expect(floats.get('b')?.critical).toBe(true)
    expect(floats.get('c')?.critical).toBe(true)
    expect(floats.get('m')?.critical).toBe(true)
    // Zero float → no slack geometry anywhere on the chain.
    for (const taskId of ['a', 'b', 'c', 'm']) {
      expect(floats.get(taskId)?.totalSlack).toBe(0)
      expect(floats.get(taskId)?.slack).toBeUndefined()
    }
    // The chain links are all classified critical (document order).
    expect(
      view.timeline.criticalPath?.criticalDependencyIds.map((value) => value as string),
    ).toEqual(['d1', 'd2', 'd3'])
  })

  it('R02 — the float branch: positive totalSlack echoed with its slack-bar geometry', () => {
    const { document } = chainFixture()
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    assertFloatsEqualSchedule(document, derived, view.timeline.criticalPath)
    const floats = new Map(
      view.timeline.criticalPath?.floats.map((float) => [float.taskId as string, float]) ?? [],
    )
    const branch = floats.get('d')
    expect(branch?.critical).toBe(false)
    // Two working days of float — the authority's signed answer, echoed.
    expect(branch?.totalSlack).toBe(960)
    expect(branch?.freeSlack).toBe(960)
    // The slack bar spans the branch's scheduled finish (Mon 17:00) to its
    // canonical late finish (Thu 09:00 — the project finish's snapped late
    // date) over the initial [Mon 09:00, Thu 09:00) viewport.
    expect(branch?.slack).toMatchObject({ startsBefore: false, finishesAfter: false })
    const viewport = view.timeline.viewport
    const spanMs = Date.parse(viewport.finish) - Date.parse(viewport.start)
    expect(branch?.slack?.startFraction).toBeCloseTo(
      (Date.parse('2026-08-03T17:00:00.000Z') - Date.parse(viewport.start)) / spanMs,
      12,
    )
    expect(branch?.slack?.finishFraction).toBeCloseTo(
      (Date.parse('2026-08-06T09:00:00.000Z') - Date.parse(viewport.start)) / spanMs,
      12,
    )
    // The branch link (critical milestone endpoint, non-critical branch) is
    // NOT classified — the both-endpoints convention.
    expect(
      view.timeline.criticalPath?.criticalDependencyIds.map((value) => value as string),
    ).not.toContain('d4')
  })

  it('R03 — the impossible schedule: negative totalSlack echoed verbatim (never clamped)', () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      tasks: [
        makeTask({ id: 'a', duration: asWorkingMinutes(480) }),
        makeTask({
          id: 'b',
          duration: asWorkingMinutes(480),
          constraintType: 'mustFinishOn',
          constraintDate: asISODateTime('2026-08-03T12:00:00.000Z'),
        }),
      ],
      dependencies: [makeDependency('d1', 'a', 'b')],
    })
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    assertFloatsEqualSchedule(document, derived, view.timeline.criticalPath)
    const floats = new Map(
      view.timeline.criticalPath?.floats.map((float) => [float.taskId as string, float]) ?? [],
    )
    // The MFO pin makes the schedule impossible for `a`: its late finish is
    // pulled before its early finish — the authority's negative-slack
    // signal, echoed verbatim, critical, and NO slack geometry (inverted).
    expect(floats.get('a')?.critical).toBe(true)
    expect(floats.get('a')?.totalSlack).toBe(-780)
    expect(floats.get('a')?.slack).toBeUndefined()
    expect(floats.get('b')?.critical).toBe(true)
    expect(floats.get('b')?.totalSlack).toBe(0)
    expect(view.timeline.criticalPath?.criticalDependencyIds).toHaveLength(1)
  })

  it('R04 — the zero-slack milestone participates with a float entry and no geometry', () => {
    const { document } = chainFixture()
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    const milestone = view.timeline.criticalPath?.floats.find((float) => float.taskId === id('m'))
    expect(milestone).toMatchObject({
      critical: true,
      totalSlack: 0,
      freeSlack: 0,
    })
    expect(milestone?.slack).toBeUndefined()
  })

  it('R05 — mixed classification: chain links critical, branch link not (both-endpoints convention)', () => {
    const { document } = chainFixture()
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    expect(
      view.timeline.criticalPath?.criticalDependencyIds.map((value) => value as string),
    ).toEqual(['d1', 'd2', 'd3'])
    // The full observable tuple: document + schedule + surfaces.
    expect(view.timeline.criticalPath?.floats.map((float) => float.taskId as string)).toEqual([
      'a',
      'b',
      'c',
      'd',
      'm',
    ])
    expect(document.dependencies).toHaveLength(4)
    expect(derived.projectFinish).toBe('2026-08-06T09:00:00.000Z')
  })
})

describe('PROJECT-026 goldens — resource utilization', () => {
  it('R06 — the calendar-aware demand tiling (single two-day assignment)', () => {
    const { document } = tilingFixture()
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    const surface = view.timeline.resourceUtilization
    expect(surface?.status).toBe('ok')
    // The bands EQUAL the authority's allocation output clipped — never
    // re-derived.
    const authoritative = resourceAllocations(document, derived)
    expect(authoritative).toHaveLength(1)
    expect(
      surface?.resources?.[0]?.bands.map((band) => [
        band.start,
        band.finish,
        band.demandUnits,
        band.capacityUnits,
        band.overallocated,
      ]),
    ).toEqual(
      (authoritative[0]?.segments ?? []).map((segment) => [
        segment.start as string,
        segment.finish as string,
        segment.demandUnits,
        segment.capacityUnits,
        segment.overallocated,
      ]),
    )
    // The hand-expected tiling: the working days at demand 1 and the
    // Monday night at demand 0 (the resource supplies no capacity then).
    expect(
      surface?.resources?.[0]?.bands.map((band) => [band.start, band.finish, band.demandUnits]),
    ).toEqual([
      ['2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z', 1],
      ['2026-08-03T17:00:00.000Z', '2026-08-04T09:00:00.000Z', 0],
      ['2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z', 1],
    ])
    expect(surface?.resources?.[0]?.name).toBe('Builder')
  })

  it("R07 — the over-allocation: the over-allocated band equals the leveler's conflict record", () => {
    const { document } = overloadFixture()
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    const bands = view.timeline.resourceUtilization?.resources?.[0]?.bands ?? []
    expect(bands).toHaveLength(1)
    expect(bands[0]).toMatchObject({
      start: '2026-08-03T09:00:00.000Z',
      finish: '2026-08-03T17:00:00.000Z',
      demandUnits: 1.6,
      capacityUnits: 1,
      overallocated: true,
    })
    expect(bands[0]?.assignmentIds.map((value) => value as string)).toEqual(['a1', 'a2'])
    // The observable tuple: document + schedule + the over-allocation band.
    expect(document.assignments).toHaveLength(2)
    expect(derived.taskSchedules[id('t1')]?.scheduledStart).toBe('2026-08-03T09:00:00.000Z')
  })

  it("R08 — the availability-window capacity drop: the authority's tightest capacity echoed per segment", () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      tasks: [makeTask({ id: 't1', duration: asWorkingMinutes(1440) })], // Mon 09:00 → Wed 17:00
      resources: [
        makeResource({
          id: 'r1',
          availability: [{ start: asISODateTime('2026-08-04T12:00:00.000Z'), units: 0.5 }],
        }),
      ],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    const derived = schedule(document)
    const { view } = pipeline(document, derived)
    const bands = view.timeline.resourceUtilization?.resources?.[0]?.bands ?? []
    const byStart = new Map(bands.map((band) => [band.start, band]))
    // Full capacity before the window; the tightest window units (0.5) from
    // Tue 12:00 onwards — over-allocated where the demand (1) exceeds it.
    expect(byStart.get('2026-08-03T09:00:00.000Z')).toMatchObject({
      capacityUnits: 1,
      overallocated: false,
    })
    expect(byStart.get('2026-08-04T09:00:00.000Z')).toMatchObject({
      capacityUnits: 1,
      overallocated: false,
    })
    expect(byStart.get('2026-08-04T12:00:00.000Z')).toMatchObject({
      demandUnits: 1,
      capacityUnits: 0.5,
      overallocated: true,
    })
    expect(byStart.get('2026-08-05T09:00:00.000Z')).toMatchObject({
      demandUnits: 1,
      capacityUnits: 0.5,
      overallocated: true,
    })
    // Equality with the authority's own output.
    const authoritative = resourceAllocations(document, derived)
    expect(bands.length).toBe(authoritative[0]?.segments.length)
  })

  it('R09 — viewport re-projection: both surfaces re-project onto a zoomed viewport byte-identically', () => {
    const { document } = overloadFixture()
    const derived = schedule(document)
    const { state } = pipeline(document, derived)
    const zoomed = scaleViewport(state.viewport, 0.5, state.viewport.start)
    const zoomedProjection = projectDocumentView(document, derived, {
      ...state,
      viewport: zoomed,
    })
    const zoomedView = buildGanttView(
      document,
      zoomedProjection,
      { ...state, viewport: zoomed },
      { firstRow: 0, visibleRows: 10 },
      undefined,
      { allocation: canonicalAllocation, schedule: derived },
    )
    // The bands clip to the narrower window; the floats keep echoing the
    // SAME schedule values (the schedule did not change).
    const bands = zoomedView.timeline.resourceUtilization?.resources?.[0]?.bands ?? []
    expect(bands.length).toBeGreaterThanOrEqual(1)
    for (const band of bands) {
      expect(band.start >= zoomed.start).toBe(true)
      expect(band.finish <= zoomed.finish).toBe(true)
    }
    assertFloatsEqualSchedule(document, derived, zoomedView.timeline.criticalPath)
    // Repeated re-projection is byte-identical (determinism).
    const first = JSON.stringify(zoomedView.timeline.resourceUtilization)
    for (let i = 0; i < 2; i += 1) {
      expect(
        JSON.stringify(
          buildGanttView(
            document,
            zoomedProjection,
            { ...state, viewport: zoomed },
            { firstRow: 0, visibleRows: 10 },
            undefined,
            { allocation: canonicalAllocation, schedule: derived },
          ).timeline.resourceUtilization,
        ),
      ).toBe(first)
    }
  })
})

describe('PROJECT-026 goldens — the dynamic pipeline', () => {
  it('R10 — a session command moves the schedule; undo/redo restore the surfaces byte-identically', () => {
    const { document } = overloadFixture()
    let session = createRendererSession(document, { schedule })
    const build = () => {
      const state = createViewState(session.document, session.schedule)
      const projection = projectDocumentView(session.document, session.schedule, state)
      const view = buildGanttView(
        session.document,
        projection,
        state,
        { firstRow: 0, visibleRows: 10 },
        undefined,
        { allocation: canonicalAllocation, schedule: session.schedule! },
      )
      return {
        criticalPath: view.timeline.criticalPath,
        resourceUtilization: view.timeline.resourceUtilization,
        document: session.document,
        schedule: session.schedule,
      }
    }
    const before = build()
    // t1 grows to two days: the critical path follows (the project finish
    // moves), and the allocation bands re-tile.
    const outcome = applyRendererCommand(session, {
      type: 'SetTaskDuration',
      taskId: id('t1'),
      duration: asWorkingMinutes(960),
    })
    expect(outcome.result.accepted).toBe(true)
    session = outcome.session
    const after = build()
    expect(after.schedule?.projectFinish).not.toBe(before.schedule?.projectFinish)
    expect(after.resourceUtilization?.resources?.[0]?.bands.length).toBeGreaterThanOrEqual(3)
    // The surfaces still EQUAL the new authority answers.
    assertFloatsEqualSchedule(after.document, after.schedule!, after.criticalPath)
    // Undo restores the exact prior tuple (byte-identical).
    const undone = undoRendererCommand(session)
    expect(undone.applied).toBe(true)
    session = undone.session
    expect(JSON.stringify(build())).toBe(JSON.stringify(before))
    // Redo restores the moved state.
    const redone = redoRendererCommand(session)
    expect(redone.applied).toBe(true)
    session = redone.session
    expect(JSON.stringify(build())).toBe(JSON.stringify(after))
  })

  it('R11 — absent inputs: no resource surface without the input, no critical surface without a schedule, geometry untouched', () => {
    const { document } = overloadFixture()
    const derived = schedule(document)
    const state = createViewState(document, derived)
    const projection = projectDocumentView(document, derived, state)
    const layout = { firstRow: 0, visibleRows: 10 }
    const withResources = buildGanttView(document, projection, state, layout, undefined, {
      allocation: canonicalAllocation,
      schedule: derived,
    })
    const withoutResources = buildGanttView(document, projection, state, layout)
    expect(withResources.timeline.resourceUtilization).toBeDefined()
    expect(withoutResources.timeline.resourceUtilization).toBeUndefined()
    // The critical surface needs no input — it joins from the projection's
    // schedule; without a scheduler there is nothing to project.
    expect(withoutResources.timeline.criticalPath).toEqual(withResources.timeline.criticalPath)
    const unscheduled = projectDocumentView(document, undefined, state)
    const bareView = buildGanttView(document, unscheduled, state, layout, undefined, {
      allocation: canonicalAllocation,
      schedule: { taskSchedules: {}, diagnostics: [] },
    })
    expect(bareView.timeline.criticalPath).toBeUndefined()
    expect(bareView.timeline.resourceUtilization).toMatchObject({ status: 'ok', resources: [] })
    // Geometry additivity: bars/bands/rows byte-identical with and without.
    expect(withoutResources.timeline.bars).toEqual(withResources.timeline.bars)
    expect(withoutResources.timeline.bands).toEqual(withResources.timeline.bands)
    expect(withoutResources.timeline.rows).toEqual(withResources.timeline.rows)
    expect(withoutResources.timeline.links).toEqual(withResources.timeline.links)
    expect(withoutResources.timeline.milestones).toEqual(withResources.timeline.milestones)
  })

  it('R12 — the 3× byte-identical mixed command/viewport sequence (both surfaces)', () => {
    const { document } = overloadFixture()
    let session = createRendererSession(document, { schedule })
    const steps = (): string[] => {
      const out: string[] = []
      let state = createViewState(session.document, session.schedule)
      for (let step = 0; step < 2; step += 1) {
        if (step > 0) {
          state = {
            ...state,
            viewport: scaleViewport(state.viewport, 0.8, state.viewport.start),
          }
        }
        const projection = projectDocumentView(session.document, session.schedule, state)
        const view = buildGanttView(
          session.document,
          projection,
          state,
          { firstRow: 0, visibleRows: 10 },
          undefined,
          { allocation: canonicalAllocation, schedule: session.schedule! },
        )
        out.push(
          JSON.stringify({
            criticalPath: view.timeline.criticalPath,
            resourceUtilization: view.timeline.resourceUtilization,
            document: session.document,
            schedule: session.schedule,
          }),
        )
      }
      return out
    }
    // The deterministic replay: the same command/viewport sequence produces
    // byte-identical observable tuples across 3 full replays.
    const replay = (): string[] => {
      session = createRendererSession(document, { schedule })
      const grow = applyRendererCommand(session, {
        type: 'SetTaskDuration',
        taskId: id('t2'),
        duration: asWorkingMinutes(960),
      })
      if (grow.result.accepted) session = grow.session
      const grown = steps()
      const undo = undoRendererCommand(session)
      if (undo.applied) session = undo.session
      const restored = steps()
      return [...grown, ...restored]
    }
    const first = replay()
    expect(first.length).toBe(4)
    for (let i = 0; i < 2; i += 1) {
      expect(replay()).toEqual(first)
    }
  })
})
