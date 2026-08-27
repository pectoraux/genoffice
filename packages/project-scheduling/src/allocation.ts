import type {
  AssignmentId,
  Calendar,
  DerivedSchedule,
  ISODateTime,
  ProjectDocument,
  ResourceId,
} from '@genoffice/project-contracts'
import type { CalendarBook } from './calendar.js'
import {
  allocationSegments,
  compareIds,
  demandIntervalsForResource,
  resourceCalendarFor,
} from './allocation-kernel.js'

// ===========================================================================
// PROJECT-026 — canonical time-phased resource allocation (kernel consumer).
//
// `resourceAllocations(document, schedule)` is the canonical resource/assignment
// RESULT the renderer's resource visualization projects (the injected
// `ResourceAllocationQuery` binding). It is a pure read-only projection of a
// GIVEN DerivedSchedule — it never schedules, never mutates the document, and
// never proposes commands (leveling stays `levelResources`; `schedule()` stays
// the sole authoritative scheduling operation, architecture-lock §6).
//
// SINGLE AUTHORITY (the PROJECT-026 correction): every demand/capacity
// semantic this projection exposes comes from the shared canonical
// allocation kernel (`allocation-kernel.ts`) — the ONE implementation of the
// demand-contribution rule, the demand-interval construction, the
// resource-calendar resolution, the availability-capacity resolution, the
// calendar-aware segmentation, and the over-allocation predicate. The
// accepted leveler (`leveling.ts`) consumes the SAME kernel for its conflict
// windows, so there are not two semantic authorities for resource capacity
// in this package — there is one kernel with two projections:
//
//     resourceAllocations() emits EVERY segment of the kernel's tiling
//     (including zero-demand non-working segments) so the projection
//     receives the authority's complete time-phased answer; the leveler
//     collapses the consecutive over-allocated runs of the same segments
//     into its conflict windows and leveling decisions.
//
// The scheduling architecture suite enforces the single-authority property:
// it fails if the kernel's primitives are defined anywhere else and fails if
// this module (or the leveler) stops consuming them.
//
// Determinism (the shared kernel contract): the same `(document, schedule)`
// pair always produces byte-identical output, and the output is invariant
// under task/assignment/resource array reordering and serialization
// round-trips — allocations are sorted by `resourceId`, segments ascending
// by start, and each segment's assignment ids sorted (locale-free code-unit
// comparison). Pure: inputs are never mutated.
// ===========================================================================

/** One maximal `[start, finish)` window of the allocation tiling where the
 * active-assignment set, the demand, and the effective capacity are all
 * constant. Segments tile the resource's assignment span contiguously:
 * every segment's `finish` equals its successor's `start`.
 *
 * - `demandUnits` is the combined units of the active assignments — the sum
 *   of `Assignment.units` over every assignment active throughout the
 *   segment. On segments where the resource's resolved calendar says the
 *   resource is NOT working, demand is zero: the resource supplies no work
 *   capacity there (the kernel's calendar-aware demand rule — a task window
 *   overlapping a resource's day off produces a zero-demand segment, never
 *   a false over-allocation).
 * - `capacityUnits` is the resource's effective capacity over the segment —
 *   the resource's capacity, or the tightest covering availability window's
 *   units when one covers the segment (the kernel's capacity resolution).
 * - `overallocated` is the canonical conflict predicate: demand exceeds
 *   capacity on this segment. This is the kernel's own classification —
 *   projections echo it, never re-derive it. */
export interface ResourceAllocationSegment {
  readonly start: ISODateTime
  readonly finish: ISODateTime
  readonly demandUnits: number
  readonly capacityUnits: number
  readonly overallocated: boolean
  /** The assignments active throughout the segment, sorted by id. */
  readonly assignmentIds: readonly AssignmentId[]
}

/** The time-phased allocation of ONE work resource: the segments tiling its
 * assignment span, ascending by start. */
export interface ResourceAllocation {
  readonly resourceId: ResourceId
  readonly segments: readonly ResourceAllocationSegment[]
}

/**
 * The canonical time-phased work-resource allocation of a GIVEN derived
 * schedule — the kernel's complete tiling per resource. For every work
 * resource that carries at least one contributing assignment, returns the
 * contiguous segment tiling of its assignment span with the demand, the
 * effective capacity, and the canonical over-allocation flag per segment.
 * Allocations are sorted by `resourceId` (locale-free code-unit comparison)
 * — the shared kernel's deterministic-output contract.
 *
 * Pure and deterministic: the same `(document, schedule)` pair always
 * produces byte-identical allocation bytes; neither input is mutated.
 * Material and cost resources (no work capacity) and resources with no
 * capacity at all produce no entry — the kernel's rules. A task without a
 * scheduled window contributes no demand (never invented).
 */
export function resourceAllocations(
  document: ProjectDocument,
  schedule: DerivedSchedule,
): readonly ResourceAllocation[] {
  const allocations: ResourceAllocation[] = []
  const book: CalendarBook = { calendars: document.calendars }
  const calendarCache = new Map<string, Calendar>()
  for (const resource of document.resources) {
    if (resource.kind !== 'work') continue
    if (resource.maxUnits <= 0) continue
    const intervals = demandIntervalsForResource(document, schedule, resource)
    if (intervals.length === 0) continue
    const resourceCal = resourceCalendarFor(document, book, calendarCache, resource)
    const segments = allocationSegments(resource, resourceCal, intervals)
    if (segments.length === 0) continue
    allocations.push({
      resourceId: resource.id,
      segments: segments.map((segment) => ({
        start: segment.start,
        finish: segment.finish,
        demandUnits: segment.demandUnits,
        capacityUnits: segment.capacityUnits,
        overallocated: segment.overallocated,
        assignmentIds: segment.active.map((interval) => interval.assignment.id).sort(compareIds),
      })),
    })
  }
  allocations.sort((a, b) => compareIds(a.resourceId as string, b.resourceId as string))
  return allocations
}
