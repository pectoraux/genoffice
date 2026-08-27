import { describe, expect, it } from 'vitest'
import { levelResources, resourceAllocations, schedule } from '../src/index.js'
import type {
  DerivedSchedule,
  LevelingOverallocation,
  ProjectDocument,
} from '@genoffice/project-contracts'
import type { ResourceAllocation } from '../src/index.js'
import { asAssignmentId, asResourceId } from '@genoffice/project-contracts'
import {
  MONDAY,
  MONDAY_FINISH,
  TUESDAY,
  TUESDAY_FINISH,
  WEDNESDAY,
  WEDNESDAY_FINISH,
  iso,
  makeAssignment,
  makeDocument,
  makeResource,
  makeTask,
  parseDocument,
  taskId,
  wm,
} from './fixtures.js'

/**
 * PROJECT-026 — the canonical time-phased resource allocation (the canonical
 * resource/assignment result the renderer's resource visualization projects
 * through the injected `ResourceAllocationQuery`).
 *
 * The battery proves the canonical semantics: the demand tiling (calendar-
 * aware — zero demand where the resource is not working), the availability-
 * window capacity resolution (tightest coverage, segment-splitting), the
 * canonical over-allocation flag, the skip rules (material/cost/no-capacity
 * resources; milestone/summary/zero-duration/unscheduled tasks contribute no
 * demand), and the leveler's determinism contract (sorted output, invariance
 * under array reordering and serialization round-trips, purity). The
 * leveler cross-check proves the mirrored-sweep equivalence: the union of
 * consecutive over-allocated segments IS the leveler's own conflict record
 * — the same authority, never a second capacity engine.
 */

const overallocatedWindows = (allocations: readonly ResourceAllocation[]) =>
  allocations.flatMap((allocation) => {
    const windows: {
      resourceId: string
      start: string
      finish: string
      peak: number
      capacity: number
      assignmentIds: string[]
    }[] = []
    for (const segment of allocation.segments) {
      if (!segment.overallocated) continue
      const last = windows[windows.length - 1]
      if (last !== undefined && last.finish === segment.start) {
        last.finish = segment.finish
        last.peak = Math.max(last.peak, segment.demandUnits)
        last.capacity = Math.min(last.capacity, segment.capacityUnits)
        for (const id of segment.assignmentIds) {
          if (!last.assignmentIds.includes(id)) last.assignmentIds.push(id)
        }
        last.assignmentIds.sort()
      } else {
        windows.push({
          resourceId: allocation.resourceId as string,
          start: segment.start as string,
          finish: segment.finish as string,
          peak: segment.demandUnits,
          capacity: segment.capacityUnits,
          assignmentIds: [...(segment.assignmentIds as readonly string[])].sort(),
        })
      }
    }
    return windows
  })

const levelerSignature = (overallocation: LevelingOverallocation) => ({
  resourceId: overallocation.resourceId as string,
  start: overallocation.window.start as string,
  finish: overallocation.window.finish as string,
  peak: overallocation.peakDemand,
  capacity: overallocation.maxUnits,
  assignmentIds: [...(overallocation.assignmentIds as readonly string[])].sort(),
})

describe('PROJECT-026 canonical resource allocation — demand tiling', () => {
  it('tiles a single assignment into one demand segment (the standard day)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    const scheduleResult = schedule(document)
    const allocations = resourceAllocations(document, scheduleResult)
    expect(allocations).toHaveLength(1)
    expect(allocations[0]?.resourceId).toBe(asResourceId('r1'))
    expect(allocations[0]?.segments).toEqual([
      {
        start: MONDAY,
        finish: MONDAY_FINISH,
        demandUnits: 1,
        capacityUnits: 1,
        overallocated: false,
        assignmentIds: [asAssignmentId('a1')],
      },
    ])
  })

  it('zeroes demand on non-working segments (the calendar-aware tiling across a night)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(960) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    const scheduleResult = schedule(document)
    const allocations = resourceAllocations(document, scheduleResult)
    const segments = allocations[0]?.segments ?? []
    // The two-day task window Mon 09:00 → Tue 17:00 tiles into three
    // segments: the working days at full demand and the intervening night
    // at ZERO demand (the resource supplies no capacity while not working).
    expect(segments.map((segment) => [segment.start, segment.finish, segment.demandUnits])).toEqual(
      [
        [MONDAY, MONDAY_FINISH, 1],
        [MONDAY_FINISH, TUESDAY, 0],
        [TUESDAY, TUESDAY_FINISH, 1],
      ],
    )
    expect(segments.every((segment) => !segment.overallocated)).toBe(true)
    // Segments tile the assignment span contiguously.
    expect(segments[0]?.start).toBe(MONDAY)
    expect(segments[segments.length - 1]?.finish).toBe(TUESDAY_FINISH)
    for (let i = 1; i < segments.length; i += 1) {
      expect(segments[i]?.start).toBe(segments[i - 1]?.finish)
    }
  })

  it('marks over-allocation when combined units exceed capacity', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(480) }), makeTask({ id: 't2', duration: wm(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [
        makeAssignment('a1', 't1', 'r1'),
        makeAssignment('a2', 't2', 'r1', { units: 0.6 }),
      ],
    })
    const scheduleResult = schedule(document)
    const allocations = resourceAllocations(document, scheduleResult)
    expect(allocations[0]?.segments).toEqual([
      {
        start: MONDAY,
        finish: MONDAY_FINISH,
        demandUnits: 1.6,
        capacityUnits: 1,
        overallocated: true,
        assignmentIds: [asAssignmentId('a1'), asAssignmentId('a2')],
      },
    ])
  })

  it('never over-allocates disjoint assignments (sequential days)', () => {
    // t1 starts ASAP (Monday); t2 is pinned to Tuesday by a soft SNET
    // constraint — the simplest deterministic disjoint case.
    const document = makeDocument({
      tasks: [
        makeTask({ id: 't1', duration: wm(480) }),
        makeTask({
          id: 't2',
          duration: wm(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('a1', 't1', 'r1'), makeAssignment('a2', 't2', 'r1')],
    })
    const scheduleResult = schedule(document)
    const allocations = resourceAllocations(document, scheduleResult)
    const segments = allocations[0]?.segments ?? []
    expect(segments.length).toBeGreaterThanOrEqual(3)
    expect(segments.every((segment) => !segment.overallocated)).toBe(true)
    // The demand appears on Monday (a1) and Tuesday (a2), never both.
    expect(segments.find((segment) => segment.start === MONDAY)?.assignmentIds).toEqual([
      asAssignmentId('a1'),
    ])
    expect(segments.find((segment) => segment.start === TUESDAY)?.assignmentIds).toEqual([
      asAssignmentId('a2'),
    ])
  })

  it('splits segments at availability-window capacity changes (tightest coverage)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(1440) })], // Mon 09:00 → Wed 17:00
      resources: [
        makeResource({
          id: 'r1',
          availability: [{ start: iso('2026-08-04T12:00:00.000Z'), units: 0.5 }],
        }),
      ],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    const scheduleResult = schedule(document)
    const allocations = resourceAllocations(document, scheduleResult)
    const segments = allocations[0]?.segments ?? []
    const byStart = new Map(segments.map((segment) => [segment.start as string, segment]))
    // Monday: full capacity, demand 1 — fine.
    expect(byStart.get(MONDAY)?.capacityUnits).toBe(1)
    expect(byStart.get(MONDAY)?.overallocated).toBe(false)
    // Tuesday morning (09:00–12:00): still before the window — full capacity.
    expect(byStart.get(TUESDAY)?.capacityUnits).toBe(1)
    expect(byStart.get(TUESDAY)?.finish).toBe('2026-08-04T12:00:00.000Z')
    expect(byStart.get(TUESDAY)?.overallocated).toBe(false)
    // Tuesday afternoon (12:00–17:00): the open-ended window's tightest
    // units (0.5) drop the capacity mid-assignment — the segment splits at
    // the window start and the demand now exceeds capacity.
    const afternoon = byStart.get('2026-08-04T12:00:00.000Z')
    expect(afternoon?.capacityUnits).toBe(0.5)
    expect(afternoon?.demandUnits).toBe(1)
    expect(afternoon?.overallocated).toBe(true)
    // Wednesday: the open-ended window still covers — capacity stays 0.5.
    expect(byStart.get(WEDNESDAY)?.capacityUnits).toBe(0.5)
    expect(byStart.get(WEDNESDAY)?.overallocated).toBe(true)
    expect(byStart.get(WEDNESDAY)?.finish).toBe(WEDNESDAY_FINISH)
  })
})

describe('PROJECT-026 canonical resource allocation — skip rules', () => {
  it('skips material/cost/no-capacity/assignment-less resources', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(480) })],
      resources: [
        makeResource({ id: 'r_material', kind: 'material' }),
        makeResource({ id: 'r_cost', kind: 'cost' }),
        makeResource({ id: 'r_zero', maxUnits: 0 }),
        makeResource({ id: 'r_empty' }),
      ],
      assignments: [
        makeAssignment('a1', 't1', 'r_material'),
        makeAssignment('a2', 't1', 'r_cost'),
        makeAssignment('a3', 't1', 'r_zero'),
      ],
    })
    const scheduleResult = schedule(document)
    expect(resourceAllocations(document, scheduleResult)).toEqual([])
  })

  it('milestones, summary roll-ups, and zero-duration tasks contribute no demand', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'root',
          summary: true,
          duration: wm(960),
        }),
        makeTask({
          id: 'child',
          parentTaskId: taskId('root'),
          outlineLevel: 2,
          duration: wm(960),
        }),
        makeTask({ id: 'mile', milestone: true, duration: wm(0) }),
      ],
      resources: [makeResource({ id: 'r1' })],
      assignments: [
        makeAssignment('a1', 'root', 'r1'), // summary with children — no own demand
        makeAssignment('a2', 'mile', 'r1'), // milestone — no demand
      ],
    })
    const scheduleResult = schedule(document)
    expect(resourceAllocations(document, scheduleResult)).toEqual([])
    // The child's assignment WOULD contribute — proving the skips above are
    // the task-shape rules, not a broken pipeline.
    const withChild = makeDocument({
      tasks: document.tasks,
      resources: document.resources,
      assignments: [makeAssignment('a3', 'child', 'r1')],
    })
    expect(resourceAllocations(withChild, schedule(withChild))).toHaveLength(1)
  })

  it('a task without a scheduled window contributes no demand (never invented)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('a1', 't1', 'r1')],
    })
    // A schedule with no entry for the task at all.
    const emptySchedule: DerivedSchedule = { taskSchedules: {}, diagnostics: [] }
    expect(resourceAllocations(document, emptySchedule)).toEqual([])
    // A schedule entry present but missing its scheduled window.
    const windowless: DerivedSchedule = {
      taskSchedules: {
        [taskId('t1')]: {
          taskId: taskId('t1'),
          totalSlack: 0,
          freeSlack: 0,
          critical: false,
          duration: wm(480),
        },
      },
      diagnostics: [],
    }
    expect(resourceAllocations(document, windowless)).toEqual([])
  })

  it('returns [] for a document with no resources', () => {
    expect(resourceAllocations(makeDocument({}), schedule(makeDocument({})))).toEqual([])
  })
})

describe('PROJECT-026 canonical resource allocation — determinism and purity', () => {
  const overloaded = (): ProjectDocument =>
    makeDocument({
      tasks: [
        makeTask({ id: 't1', duration: wm(480) }),
        makeTask({ id: 't2', duration: wm(960) }),
        makeTask({ id: 't3', duration: wm(480) }),
      ],
      resources: [makeResource({ id: 'r2' }), makeResource({ id: 'r1' })],
      assignments: [
        makeAssignment('a1', 't1', 'r2'),
        makeAssignment('a2', 't2', 'r1'),
        makeAssignment('a3', 't2', 'r2', { units: 0.75 }),
        makeAssignment('a4', 't3', 'r1', { units: 0.5 }),
      ],
    })

  it('produces byte-identical output on 3× repeat', () => {
    const document = overloaded()
    const scheduleResult = schedule(document)
    const first = JSON.stringify(resourceAllocations(document, scheduleResult))
    for (let i = 0; i < 2; i += 1) {
      expect(JSON.stringify(resourceAllocations(document, scheduleResult))).toBe(first)
    }
  })

  it('is invariant under task/assignment/resource array reordering and serialization round-trips', () => {
    const document = overloaded()
    const scheduleResult = schedule(document)
    const expected = JSON.stringify(resourceAllocations(document, scheduleResult))
    const reordered: ProjectDocument = {
      ...document,
      tasks: [...document.tasks].reverse(),
      assignments: [...document.assignments].reverse(),
      resources: [...document.resources].reverse(),
    }
    expect(JSON.stringify(resourceAllocations(reordered, scheduleResult))).toBe(expected)
    const roundTripped = parseDocument(JSON.stringify(document))
    expect(JSON.stringify(resourceAllocations(roundTripped, scheduleResult))).toBe(expected)
  })

  it('sorts allocations by resourceId and assignment ids within segments', () => {
    const document = overloaded()
    const allocations = resourceAllocations(document, schedule(document))
    expect(allocations.map((allocation) => allocation.resourceId as string)).toEqual(['r1', 'r2'])
    for (const allocation of allocations) {
      for (const segment of allocation.segments) {
        const ids = segment.assignmentIds as readonly string[]
        expect([...ids]).toEqual([...ids].sort())
      }
    }
  })

  it('never mutates its inputs', () => {
    const document = overloaded()
    const scheduleResult = schedule(document)
    const documentBefore = JSON.stringify(document)
    const scheduleBefore = JSON.stringify(scheduleResult)
    resourceAllocations(document, scheduleResult)
    expect(JSON.stringify(document)).toBe(documentBefore)
    expect(JSON.stringify(scheduleResult)).toBe(scheduleBefore)
  })
})

describe('PROJECT-026 canonical resource allocation — the leveler cross-check (mirrored sweep)', () => {
  it("the union of consecutive over-allocated segments equals the leveler's resolvable conflict record", () => {
    // Two 1-day tasks on one 100% resource (100% + 60%): a resolvable
    // conflict — the leveler delays one task; the ORIGINAL conflict window
    // is what both authorities report.
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(480) }), makeTask({ id: 't2', duration: wm(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [
        makeAssignment('a1', 't1', 'r1'),
        makeAssignment('a2', 't2', 'r1', { units: 0.6 }),
      ],
    })
    const scheduleResult = schedule(document)
    const allocationWindows = overallocatedWindows(resourceAllocations(document, scheduleResult))
    const levelerWindows = levelResources(document).overallocations.map(levelerSignature)
    expect(allocationWindows).toEqual(levelerWindows)
    expect(levelerWindows).toEqual([
      {
        resourceId: 'r1',
        start: MONDAY,
        finish: MONDAY_FINISH,
        peak: 1.6,
        capacity: 1,
        assignmentIds: ['a1', 'a2'],
      },
    ])
  })

  it("equals the leveler's UNRESOLVABLE conflict record (200% single assignment)", () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1', duration: wm(480) })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('a1', 't1', 'r1', { units: 2 })],
    })
    const scheduleResult = schedule(document)
    const allocationWindows = overallocatedWindows(resourceAllocations(document, scheduleResult))
    const levelerWindows = levelResources(document).overallocations.map(levelerSignature)
    expect(allocationWindows).toEqual(levelerWindows)
    expect(levelerWindows).toEqual([
      {
        resourceId: 'r1',
        start: MONDAY,
        finish: MONDAY_FINISH,
        peak: 2,
        capacity: 1,
        assignmentIds: ['a1'],
      },
    ])
  })

  it('agrees with the leveler across multiple resources and conflict windows', () => {
    // Two independent resources, each with one cleanly resolvable conflict
    // (the delay lands on a day with spare capacity — no cascades): r1 is
    // over-allocated Monday, r2 Tuesday. Both authorities must report the
    // same two windows for the ORIGINAL schedule.
    const document = makeDocument({
      tasks: [
        makeTask({ id: 't1', duration: wm(480) }),
        makeTask({ id: 't2', duration: wm(480) }),
        makeTask({
          id: 't3',
          duration: wm(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
        makeTask({
          id: 't4',
          duration: wm(480),
          constraintType: 'startNoEarlierThan',
          constraintDate: iso(TUESDAY),
        }),
      ],
      resources: [makeResource({ id: 'r1' }), makeResource({ id: 'r2' })],
      assignments: [
        makeAssignment('a1', 't1', 'r1'),
        makeAssignment('a2', 't2', 'r1', { units: 0.6 }),
        makeAssignment('a3', 't3', 'r2'),
        makeAssignment('a4', 't4', 'r2', { units: 0.8 }),
      ],
    })
    const scheduleResult = schedule(document)
    const allocationWindows = overallocatedWindows(resourceAllocations(document, scheduleResult))
    const levelerWindows = levelResources(document).overallocations.map(levelerSignature)
    expect(levelerWindows).toEqual([
      {
        resourceId: 'r1',
        start: MONDAY,
        finish: MONDAY_FINISH,
        peak: 1.6,
        capacity: 1,
        assignmentIds: ['a1', 'a2'],
      },
      {
        resourceId: 'r2',
        start: TUESDAY,
        finish: TUESDAY_FINISH,
        peak: 1.8,
        capacity: 1,
        assignmentIds: ['a3', 'a4'],
      },
    ])
    expect(allocationWindows).toEqual(levelerWindows)
  })
})
