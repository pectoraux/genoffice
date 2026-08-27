import type {
  Assignment,
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
// The shared canonical resource-allocation kernel — ONE authority.
//
// PROJECT-026 correction (Principal Architect review of PR #28): the
// resource-allocation semantics that `resourceAllocations()` (allocation.ts —
// the read-only time-phased tiling the renderer projects) and
// `levelResources()` (leveling.ts — conflict windows and leveling decisions)
// both need live in exactly ONE implementation — this kernel — consumed by
// both:
//
//                     allocation-kernel.ts (this module)
//                        /                            \
//             resourceAllocations()               levelResources()
//             (allocation.ts — the full           (leveling.ts — conflict
//              tiling projection)                  windows + delay decisions)
//
// The kernel owns the shared semantic concepts:
//   - the demand-contribution rule (which tasks consume resource capacity),
//   - the demand-interval construction (assignments joined with the canonical
//     scheduled windows, optionally filtered to a leveling date window),
//   - the resource-calendar resolution (resource.calendarId ?? default),
//   - the availability-capacity resolution (tightest covering window),
//   - the calendar-aware segmentation (the boundary sweep: assignment +
//     capacity-window + working-period endpoints, midpoint evaluation, zero
//     demand while the resource is not working),
//   - the over-allocation predicate (demand > capacity).
//
// Neither consumer may re-implement any of these rules: the scheduling
// architecture suite (tests/architecture.test.ts) fails if the kernel
// primitives are defined anywhere else in this package, and fails if either
// consumer stops consuming the kernel. A higher-level semantic concept has
// ONE canonical authority (the frozen architecture's core rule) — a
// cross-check test proving two implementations agree is not that
// architecture; a single implementation shared by both consumers is.
//
// The kernel is pure and deterministic (the leveler's PROJECT-013 contract,
// now the shared contract): no wall clock, no randomness, no locale-sensitive
// ordering; inputs are never mutated.
// ===========================================================================

/** The locale-free code-unit id comparison — the deterministic ordering
 * primitive shared by every allocation consumer (the scheduler's established
 * convention). */
export const compareIds = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0)

const isBefore = (a: ISODateTime, b: ISODateTime): boolean =>
  new Date(a).getTime() < new Date(b).getTime()

/** True if the task is a leaf (not a summary with children). */
export const isLeaf = (document: ProjectDocument, task: Task): boolean => {
  if (!task.summary) return true
  return !document.tasks.some((child) => child.parentTaskId === task.id)
}

/**
 * True if the task CONTRIBUTES to work-resource demand — the kernel's single
 * demand-contribution rule. A task contributes demand when it has a non-zero
 * duration and is not a summary-with-children (summaries roll up from
 * children; their own assignments, if any, are not double-counted). Manual
 * tasks, auto tasks, and leaf summaries all contribute demand — the resource
 * is consumed when the task runs regardless of whether the task is manually
 * scheduled. Whether a task can be DELAYED is a separate question the
 * leveler answers with `isProtected` (leveling policy, not allocation
 * semantics).
 *
 * Milestones and zero-duration tasks have an empty [start, finish) window,
 * so they contribute no work demand and are skipped.
 */
export const contributesToDemand = (document: ProjectDocument, task: Task): boolean => {
  if (task.summary && !isLeaf(document, task)) return false
  if (task.milestone) return false
  if ((task.duration as number) <= 0) return false
  return true
}

/**
 * Effective max-units capacity for a work resource at a given instant — the
 * kernel's single availability-capacity resolution. When an availability
 * window covers the instant, its `units` override `resource.maxUnits`
 * (MS Project semantics: availability windows define the resource's max
 * units over time). When no window covers the instant, the resource's
 * `maxUnits` is the capacity. When `maxUnits` is 0 the resource has no
 * capacity at any time.
 *
 * Windows with `finish === undefined` are open-ended (active from `start`
 * onwards). Overlapping windows are resolved by taking the MINIMUM units
 * (the tightest capacity) so an over-allocation is never masked by a wider
 * window.
 */
export const effectiveMaxUnits = (resource: Resource, timestamp: ISODateTime): number => {
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

/**
 * One assignment's demand interval — the kernel's single demand-interval
 * shape: the assignment record joined with its task, the task's canonical
 * scheduled window, and the assignment's units. `levelResources()` reads
 * `task` for its delay decisions; `resourceAllocations()` projects only the
 * assignment identity and the demand values.
 */
export interface ResourceDemandInterval {
  readonly assignment: Assignment
  readonly task: Task
  readonly start: ISODateTime
  readonly finish: ISODateTime
  readonly units: number
}

/** An optional `[start, finish)` demand window (the leveler's
 * `levelingDateWindow` scope): assignments whose scheduled window does not
 * overlap it are omitted from the demand set. `resourceAllocations()`
 * passes no window — it reports the complete tiling. */
export interface DemandWindow {
  readonly start?: ISODateTime
  readonly finish?: ISODateTime
}

/**
 * The demand intervals of one work resource — the kernel's single
 * demand-interval construction: every assignment on the resource whose task
 * contributes demand and whose canonical scheduled window is present in the
 * GIVEN DerivedSchedule. Material and cost resources are never work-capacity
 * and produce no intervals. A task without a scheduled window contributes no
 * demand (never invented). Intervals keep document order (the consumers'
 * deterministic-ordering contracts sort their own outputs).
 */
export const demandIntervalsForResource = (
  document: ProjectDocument,
  schedule: DerivedSchedule,
  resource: Resource,
  window?: DemandWindow,
): ResourceDemandInterval[] => {
  if (resource.kind !== 'work') return []
  const intervals: ResourceDemandInterval[] = []
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
    // Date-window filter (leveling scope): skip assignments whose scheduled
    // window does not overlap the window.
    if (window) {
      if (window.start && isBefore(finish, window.start)) continue
      if (window.finish && isBefore(window.finish, start)) continue
    }
    intervals.push({ assignment, task, start, finish, units: assignment.units })
  }
  return intervals
}

/**
 * Resolves the resource's calendar (resource.calendarId ?? defaultCalendarId)
 * — the kernel's single resource-calendar resolution, so the tiler and the
 * leveler always reason about the same working-time windows. `cache` is a
 * per-call memo (calendars are pure to resolve; the cache exists so a
 * many-resource sweep resolves each calendar once).
 */
export const resourceCalendarFor = (
  document: ProjectDocument,
  book: CalendarBook,
  cache: Map<string, Calendar>,
  resource: Resource,
): Calendar => {
  const calendarId =
    resolveResourceCalendarId(document, resource.id as ResourceId) ??
    document.properties.defaultCalendarId
  const key = calendarId as string
  let resolved = cache.get(key)
  if (!resolved) {
    resolved = resolveCalendar(book, calendarId)
    cache.set(key, resolved)
  }
  return resolved
}

/**
 * One maximal `[start, finish)` window of the kernel's allocation tiling
 * where the active-assignment set, the demand, and the effective capacity
 * are all constant, together with the kernel's own over-allocation
 * classification. Segments tile the resource's assignment span contiguously:
 * every segment's `finish` equals its successor's `start`.
 *
 * - `demandUnits` is the combined units of the active assignments. On
 *   segments where the resource's resolved calendar says the resource is NOT
 *   working, demand is zero: the resource supplies no work capacity there (a
 *   task window overlapping a resource's day off produces a zero-demand
 *   segment, never a false over-allocation).
 * - `capacityUnits` is the resource's effective capacity over the segment
 *   (the tightest covering availability window's units when one covers it).
 * - `overallocated` is the kernel's single conflict predicate: demand exceeds
 *   capacity on this segment. The leveler's conflict windows are exactly the
 *   maximal runs of consecutive over-allocated segments; the allocation
 *   projection echoes the flag per segment.
 * - `active` carries the active assignments (document order) — the leveler's
 *   conflict sides and the projection's assignment ids both derive from it.
 */
export interface AllocationKernelSegment {
  readonly start: ISODateTime
  readonly finish: ISODateTime
  readonly demandUnits: number
  readonly capacityUnits: number
  readonly overallocated: boolean
  readonly active: readonly ResourceDemandInterval[]
}

/**
 * Builds the allocation tiling of one work resource — the kernel's single
 * calendar-aware segmentation: the maximal segments over which the
 * active-assignment set, the demand, and the capacity are constant.
 *
 * SEGMENTATION (correctness-critical, the PROJECT-013 rule): the sweep
 * collects EVERY boundary timestamp at which demand, capacity, or resource
 * working status can change — assignment endpoints, availability-window
 * endpoints, and resource-calendar working-period endpoints, all bounded to
 * the assignment span — then evaluates each consecutive pair once (midpoint
 * evaluation). Between two boundaries all three inputs are constant, so the
 * tiling is both sufficient and complete.
 *
 * RESOURCE-CALENDAR-AWARE DEMAND (correctness-critical): over-allocation is
 * evaluated only where the resource can actually supply work capacity.
 * During a segment where the resource's calendar says it is NOT working, the
 * resource supplies no work capacity, so the demand against capacity on that
 * segment is zero — there is no over-allocation there. This clips conflict
 * windows to the resource's working periods: any open conflict is closed at
 * the working→non-working transition, so a reported conflict window never
 * spans a non-working interval. Omitting the resource calendar from the
 * segmentation would report a FALSE over-allocation whenever two task
 * windows overlap on a day the resource does not work.
 *
 * Omitting availability-window boundaries would miss over-allocations that
 * arise ONLY from a mid-assignment capacity drop (the assignment endpoints
 * alone do not bracket the conflict).
 *
 * The sweep covers the COMPLETE tiling (including zero-demand non-working
 * segments): `resourceAllocations()` projects every segment, while the
 * leveler collapses the consecutive over-allocated runs into its conflict
 * windows — the same segments, two projections of one authority.
 */
export const allocationSegments = (
  resource: Resource,
  resourceCal: Calendar,
  intervals: readonly ResourceDemandInterval[],
): readonly AllocationKernelSegment[] => {
  if (intervals.length === 0 || resource.maxUnits <= 0) return []
  // Bound the sweep to the assignment span. Availability windows and
  // resource working periods outside this span cannot bracket a conflict (no
  // assignment is active there).
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
  const segments: AllocationKernelSegment[] = []
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
    // Capacity is constant on (t0, t1); evaluate at the midpoint. On
    // non-working segments capacity is irrelevant (demand is already zero).
    const capacity = effectiveMaxUnits(resource, midpoint)
    segments.push({
      start: isoAt(t0),
      finish: isoAt(t1),
      demandUnits: demand,
      capacityUnits: capacity,
      overallocated: demand > capacity,
      active,
    })
  }
  return segments
}
