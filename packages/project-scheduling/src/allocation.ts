import type {
  Assignment,
  AssignmentId,
  Calendar,
  DerivedSchedule,
  ISODateTime,
  ProjectDocument,
  Resource,
  ResourceId,
  Task,
} from '@genoffice/project-contracts'
import { asISODateTime } from '@genoffice/project-contracts'
import { type CalendarBook, isWorking, resolveCalendar, workingIntervals } from './calendar.js'
import { resolveResourceCalendarId } from './schedule.js'

// ===========================================================================
// PROJECT-026 — canonical time-phased resource allocation.
//
// `resourceAllocations(document, schedule)` is the canonical resource/assignment
// RESULT the renderer's resource visualization projects (the injected
// `ResourceAllocationQuery` binding). It is a pure read-only projection of a
// GIVEN DerivedSchedule — it never schedules, never mutates the document, and
// never proposes commands (leveling stays `levelResources`; `schedule()` stays
// the sole authoritative scheduling operation, architecture-lock §6).
//
// The sweep MIRRORS the accepted leveler's demand/capacity semantics exactly
// (the same rules `packages/project-scheduling/src/leveling.ts` applies
// internally): which assignments contribute demand, the availability-window
// capacity resolution, the resource-calendar-aware demand zeroing, and the
// segment boundary collection. Where the leveler collapses consecutive
// over-allocated segments into conflict windows, this function emits EVERY
// segment of the tiling (including zero-demand non-working segments) so the
// projection receives the authority's complete time-phased answer. The
// leveler itself is untouched; the equivalence is asserted by test (the
// union of consecutive over-allocated segments equals the leveler's
// `overallocations` windows).
//
// Determinism (the leveler's contract, mirrored): the same
// `(document, schedule)` pair always produces byte-identical output, and the
// output is invariant under task/assignment/resource array reordering and
// serialization round-trips — allocations are sorted by `resourceId`,
// segments ascending by start, and each segment's assignment ids sorted
// (locale-free code-unit comparison). Pure: inputs are never mutated.
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
 *   capacity there (the accepted leveler's calendar-aware demand rule — a
 *   task window overlapping a resource's day off produces a zero-demand
 *   segment, never a false over-allocation).
 * - `capacityUnits` is the resource's effective capacity over the segment —
 *   the resource's capacity, or the tightest covering availability window's
 *   units when one covers the segment (the accepted leveler's capacity
 *   resolution, mirrored).
 * - `overallocated` is the canonical conflict predicate: demand exceeds
 *   capacity on this segment. This is the authority's own classification —
 *   projections echo it, never re-derive it. */
export interface ResourceAllocationSegment {
  readonly start: ISODateTime
  readonly finish: ISODateTime
  readonly demandUnits: number
  readonly capacityUnits: number
  readonly overallocated: boolean
  /** The assignments active throughout the segment, in document order. */
  readonly assignmentIds: readonly AssignmentId[]
}

/** The time-phased allocation of ONE work resource: the segments tiling its
 * assignment span, ascending by start. */
export interface ResourceAllocation {
  readonly resourceId: ResourceId
  readonly segments: readonly ResourceAllocationSegment[]
}

/** The locale-free code-unit id comparison (the scheduler's established
 * deterministic ordering). */
const compareIds = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0)

/** One assignment's demand interval: the assignment record joined with its
 * task's canonical scheduled window and the assignment's units. */
interface DemandInterval {
  assignment: Assignment
  start: ISODateTime
  finish: ISODateTime
  units: number
}

/** True if the task is a leaf (not a summary with children). Mirrors the
 * leveler's `isLeaf`. */
const isLeaf = (document: ProjectDocument, task: Task): boolean => {
  if (!task.summary) return true
  return !document.tasks.some((child) => child.parentTaskId === task.id)
}

/**
 * True if the task CONTRIBUTES to work-resource demand — the leveler's
 * `contributesToDemand` rule, mirrored verbatim: a summary-with-children
 * rolls up from children (its own assignments are not double-counted);
 * milestones and zero-duration tasks have an empty window and consume no
 * resource capacity; everything else (manual, auto, leaf summaries)
 * consumes the resource while it runs.
 */
const contributesToDemand = (document: ProjectDocument, task: Task): boolean => {
  if (task.summary && !isLeaf(document, task)) return false
  if (task.milestone) return false
  if ((task.duration as number) <= 0) return false
  return true
}

/**
 * The resource's effective capacity at an instant — the leveler's
 * `effectiveMaxUnits` rule, mirrored verbatim: a covering availability
 * window overrides the resource's nominal capacity with the TIGHTEST
 * covering window's units (an over-allocation is never masked by a wider
 * window); open-ended windows are active from their start onwards; with no
 * covering window the nominal capacity applies; a resource with no capacity
 * at all is skipped earlier (mirroring the leveler's no-capacity rule).
 */
const effectiveCapacity = (resource: Resource, timestamp: ISODateTime): number => {
  if (resource.maxUnits <= 0) return 0
  const ts = new Date(timestamp).getTime()
  let effective = resource.maxUnits
  let covered = false
  for (const slot of resource.availability) {
    const startMs = new Date(slot.start).getTime()
    const finishMs =
      slot.finish !== undefined ? new Date(slot.finish).getTime() : Number.POSITIVE_INFINITY
    if (ts >= startMs && ts < finishMs) {
      covered = true
      // Take the tightest capacity across all covering windows.
      if (slot.units < effective) effective = slot.units
    }
  }
  return covered ? effective : resource.maxUnits
}

/** The demand intervals of one work resource: every assignment on the
 * resource whose task contributes demand and whose canonical scheduled
 * window is present in the GIVEN DerivedSchedule. */
const demandIntervalsForResource = (
  document: ProjectDocument,
  schedule: DerivedSchedule,
  resource: Resource,
): DemandInterval[] => {
  const intervals: DemandInterval[] = []
  for (const assignment of document.assignments) {
    if (assignment.resourceId !== resource.id) continue
    const task = document.tasks.find((candidate) => candidate.id === assignment.taskId)
    if (task === undefined) continue
    if (!contributesToDemand(document, task)) continue
    const taskSchedule = schedule.taskSchedules[task.id]
    if (taskSchedule === undefined) continue
    const start = taskSchedule.scheduledStart
    const finish = taskSchedule.scheduledFinish
    if (start === undefined || finish === undefined) continue
    intervals.push({ assignment, start, finish, units: assignment.units })
  }
  return intervals
}

/**
 * Builds the allocation tiling of one work resource: the maximal segments
 * over which the active-assignment set, the demand, and the capacity are
 * constant. SEGMENTATION (the leveler's correctness-critical rule,
 * mirrored): the sweep collects EVERY boundary timestamp at which demand,
 * capacity, or resource working status can change — assignment endpoints,
 * availability-window endpoints, and resource-calendar working-period
 * endpoints, all bounded to the assignment span — then evaluates each
 * consecutive pair once (midpoint evaluation). Between two boundaries all
 * three inputs are constant, so the tiling is both sufficient and complete.
 */
const allocationSegmentsForResource = (
  document: ProjectDocument,
  resource: Resource,
  intervals: readonly DemandInterval[],
): readonly ResourceAllocationSegment[] => {
  if (intervals.length === 0 || resource.maxUnits <= 0) return []
  const book: CalendarBook = { calendars: document.calendars }
  const calendarId =
    resolveResourceCalendarId(document, resource.id) ?? document.properties.defaultCalendarId
  const resourceCal: Calendar = resolveCalendar(book, calendarId)

  const startMs = intervals.reduce(
    (min, interval) => Math.min(min, new Date(interval.start).getTime()),
    Number.POSITIVE_INFINITY,
  )
  const finishMs = intervals.reduce(
    (max, interval) => Math.max(max, new Date(interval.finish).getTime()),
    Number.NEGATIVE_INFINITY,
  )
  const boundaries = new Set<number>()
  for (const interval of intervals) {
    boundaries.add(new Date(interval.start).getTime())
    boundaries.add(new Date(interval.finish).getTime())
  }
  for (const slot of resource.availability) {
    const slotStart = new Date(slot.start).getTime()
    if (slotStart >= startMs && slotStart <= finishMs) boundaries.add(slotStart)
    if (slot.finish !== undefined) {
      const slotFinish = new Date(slot.finish).getTime()
      if (slotFinish >= startMs && slotFinish <= finishMs) boundaries.add(slotFinish)
    }
  }
  // Resource-calendar working-period boundaries: each segment then lies
  // wholly inside or wholly outside a working period, so the working-status
  // check at the segment midpoint is representative of the whole segment.
  const spanStartIso = asISODateTime(new Date(startMs).toISOString())
  const spanFinishIso = asISODateTime(new Date(finishMs).toISOString())
  for (const working of workingIntervals(resourceCal, spanStartIso, spanFinishIso)) {
    boundaries.add(new Date(working.start).getTime())
    boundaries.add(new Date(working.finish).getTime())
  }
  const sorted = [...boundaries].sort((a, b) => a - b)
  if (sorted.length < 2) return []

  const isoAt = (ms: number): ISODateTime => asISODateTime(new Date(ms).toISOString())
  const segments: ResourceAllocationSegment[] = []
  for (let k = 0; k < sorted.length - 1; k += 1) {
    const t0 = sorted[k] as number
    const t1 = sorted[k + 1] as number
    if (t0 >= t1) continue
    // Active assignments on the open segment (t0, t1): a half-open interval
    // [start, finish) is active throughout (t0, t1) iff it was already
    // active at t0 and remains active past t1.
    const active = intervals.filter((interval) => {
      const s = new Date(interval.start).getTime()
      const f = new Date(interval.finish).getTime()
      return s <= t0 && f >= t1
    })
    const midpoint = isoAt(t0 + (t1 - t0) / 2)
    const working = isWorking(resourceCal, midpoint)
    // Calendar-aware demand: during non-working segments the resource
    // supplies no work capacity, so demand against capacity is zero.
    const demand = working ? active.reduce((sum, interval) => sum + interval.units, 0) : 0
    const capacity = effectiveCapacity(resource, midpoint)
    segments.push({
      start: isoAt(t0),
      finish: isoAt(t1),
      demandUnits: demand,
      capacityUnits: capacity,
      overallocated: demand > capacity,
      assignmentIds: active.map((interval) => interval.assignment.id).sort(compareIds),
    })
  }
  return segments
}

/**
 * The canonical time-phased work-resource allocation of a GIVEN derived
 * schedule. For every work resource that carries at least one contributing
 * assignment, returns the contiguous segment tiling of its assignment span
 * with the demand, the effective capacity, and the canonical
 * over-allocation flag per segment. Allocations are sorted by `resourceId`
 * (locale-free code-unit comparison) — the leveler's deterministic-output
 * contract, mirrored.
 *
 * Pure and deterministic: the same `(document, schedule)` pair always
 * produces byte-identical allocation bytes; neither input is mutated.
 * Material and cost resources (no work capacity) and resources with no
 * capacity at all produce no entry — mirroring the leveler's rules. A task
 * without a scheduled window contributes no demand (never invented).
 */
export function resourceAllocations(
  document: ProjectDocument,
  schedule: DerivedSchedule,
): readonly ResourceAllocation[] {
  const allocations: ResourceAllocation[] = []
  for (const resource of document.resources) {
    if (resource.kind !== 'work') continue
    if (resource.maxUnits <= 0) continue
    const intervals = demandIntervalsForResource(document, schedule, resource)
    if (intervals.length === 0) continue
    const segments = allocationSegmentsForResource(document, resource, intervals)
    if (segments.length === 0) continue
    allocations.push({ resourceId: resource.id, segments })
  }
  allocations.sort((a, b) => compareIds(a.resourceId as string, b.resourceId as string))
  return allocations
}
