/**
 * PROJECT-026 — the resource-visualization projection layer.
 *
 * The authorized semantic path (verbatim from the work-item directive):
 *
 *     ProjectDocument + DerivedSchedule + canonical resource/assignment results
 *                            ↓
 *     project-renderer-core resource projection
 *                            ↓
 *     resource visualizations
 *
 * The renderer visualizes canonical resource allocation results ALREADY
 * PRODUCED by the Project domain; it never re-derives them. Concretely
 * there is NO second capacity engine here:
 *
 * - Every "how much demand does this resource face over this window"
 *   answer comes from the INJECTED `ResourceAllocationQuery` — a structural
 *   function type satisfied by binding `resourceAllocations` from
 *   `@genoffice/project-scheduling` (the accepted `ScheduleRunner` /
 *   `CalendarWorkingTimeQuery` injection precedents; this package stays
 *   statically scheduling-free, architecture-lock §3/§6). The canonical
 *   binding hosts write is:
 *
 *     (document, schedule) => resourceAllocations(document, schedule)
 *
 * - The authority's segments are joined VERBATIM into the band surface:
 *   the projection CLIPS each segment to the timeline viewport, drops the
 *   degenerate spans, keeps the document-order entries, and echoes the
 *   demand, the effective capacity, the contributing assignment ids, and
 *   the authority's OWN over-allocation flag. There is no demand
 *   aggregation, no capacity resolution, no over-allocation comparison,
 *   and no working-time evaluation anywhere in this module — without the
 *   authority's segments no capacity statement is possible here; the
 *   architecture discipline suite scans this module for exactly those
 *   markers.
 *
 * Degradation mirrors the scheduling authority exactly (the PROJECT-025
 * `buildCalendarSurface` boundary): an evaluator error that carries a
 * string code (structurally, the canonical `CalendarError` family the
 * scheduling package throws for unresolvable calendars) degrades the
 * surface to `status: 'unresolvable'` with the diagnostic echoed verbatim
 * and NO bands; an error without a code (a host-binding bug) is re-thrown.
 * A degenerate window (unparseable or empty span) has no bands and does
 * not consult the evaluator at all — never a crash, never invented values.
 *
 * Determinism: every function is pure — the same inputs always produce
 * byte-identical outputs (3×-repeat tested), no wall clock, no randomness,
 * no locale ordering; inputs are never mutated.
 */
import type {
  AssignmentId,
  DerivedSchedule,
  ImportDiagnostic,
  ProjectDocument,
  ResourceId,
} from '@genoffice/project-contracts'
import { parseInstant, formatInstant } from './state.js'
import type { TimelineViewport } from './state.js'

/**
 * One maximal window of the canonical allocation tiling: the authority's
 * answer for ONE work resource over a span where the demand, the capacity,
 * and the active-assignment set are constant. Structurally identical to
 * the scheduling package's `ResourceAllocationSegment` — the canonical
 * binding returns its output verbatim and this layer joins it without
 * re-deriving anything.
 */
export interface ResourceAllocationSegment {
  readonly start: string
  readonly finish: string
  /** The combined demand of the active assignments over the segment
   * (ZERO on segments where the resource's own calendar says it is not
   * working — the authority's calendar-aware demand rule). */
  readonly demandUnits: number
  /** The resource's effective capacity over the segment (the authority's
   * capacity-window resolution — echoed, never re-derived). */
  readonly capacityUnits: number
  /** The authority's OWN over-allocation classification for the segment. */
  readonly overallocated: boolean
  /** The assignments active throughout the segment (sorted ids). */
  readonly assignmentIds: readonly AssignmentId[]
}

/** The canonical allocation of ONE work resource: its segments, ascending
 * by start, tiling the resource's assignment span contiguously. */
export interface ResourceAllocation {
  readonly resourceId: ResourceId
  readonly segments: readonly ResourceAllocationSegment[]
}

/**
 * The injected canonical allocation evaluator. Hosts bind this to the
 * scheduling package's `resourceAllocations` (see the module header); the
 * renderer core NEVER implements it — implementing it here would be a
 * second capacity engine, which the PROJECT-026 directive forbids. The
 * query receives the canonical document and the CURRENT derived schedule
 * (the very object the renderer session holds), so the allocation always
 * matches the schedule the rest of the view projects. Errors: a thrown
 * error carrying a string `code` (the canonical `CalendarError` family)
 * degrades the surface to `unresolvable`; errors without a code propagate.
 */
export type ResourceAllocationQuery = (
  document: ProjectDocument,
  schedule: DerivedSchedule,
) => readonly ResourceAllocation[]

/** The PROJECT-026 resource-visualization inputs for the Gantt/timeline
 * builders: the injected canonical allocation query and the CURRENT
 * derived schedule it evaluates (the session's own schedule object — the
 * authority's answer must match the projected schedule, never a stale or
 * re-derived one). A per-RENDER input, like the layout and calendar
 * inputs — never persisted `ProjectViewState` (the PROJECT-021 boundary
 * between persisted interaction state and ephemeral host render choices).
 * Absent input → no resource surface is built (never invented). */
export interface ResourceViewInput {
  readonly allocation: ResourceAllocationQuery
  readonly schedule: DerivedSchedule
}

/** One utilization band: the authority's segment projected onto the
 * timeline viewport — clipped, degenerate spans dropped, values echoed
 * verbatim. Hosts map `demandUnits`/`capacityUnits` to their histogram
 * scale (the core never sees a pixel); `overallocated` is the authority's
 * own flag, never a renderer-side comparison. */
export interface ProjectResourceBand {
  readonly resourceId: ResourceId
  readonly start: string
  readonly finish: string
  readonly demandUnits: number
  readonly capacityUnits: number
  readonly overallocated: boolean
  readonly assignmentIds: readonly AssignmentId[]
}

/** One work resource's utilization projection: the resource's in-viewport
 * bands (ascending), with the resource's display name echoed from the
 * document when the id exists there (pure echo; absent otherwise). An
 * EMPTY band list is real information — the resource carries allocation,
 * but none of it intersects this viewport. */
export interface ProjectResourceUtilization {
  readonly resourceId: ResourceId
  readonly name?: string
  readonly bands: readonly ProjectResourceBand[]
}

/** `ok` — bands present (possibly empty when no allocation intersects the
 * viewport). `unresolvable` — the canonical evaluator rejected the
 * document/schedule pair (e.g. an unresolvable resource calendar); bands
 * are absent, never invented, and the diagnostic is echoed. */
export type ResourceSurfaceStatus = 'ok' | 'unresolvable'

/** The resource-visualization surface of ONE timeline build over ONE
 * viewport: every allocated work resource's utilization projection, in the
 * authority's (resource-id-sorted) entry order. */
export interface ProjectResourceViewSurface {
  readonly status: ResourceSurfaceStatus
  /** The per-resource utilization projections — present iff
   * `status === 'ok'`. */
  readonly resources?: readonly ProjectResourceUtilization[]
  /** The canonical evaluator's error echo, present iff
   * `status === 'unresolvable'` — the same code/message the scheduling
   * authority surfaces for the failure (an `ImportDiagnostic` shape, so
   * hosts can feed it into their diagnostics surface). */
  readonly diagnostic?: ImportDiagnostic
  /** The evaluated window, echoed verbatim. */
  readonly start: string
  readonly finish: string
}

/** The fallback diagnostic code for evaluator errors that carry no code of
 * their own. The canonical `CalendarError` family always carries a string
 * code, so this label only ever appears for non-canonical failures of a
 * host-supplied binding — which `buildResourceUtilization` re-throws —
 * making it a defensive guard, not an expected value. */
export const RESOURCE_ALLOCATION_FAILED = 'RESOURCE_ALLOCATION_FAILED'

/**
 * Builds the resource-visualization surface of one viewport. The evaluator
 * is consulted exactly ONCE; its segments are clipped to the viewport and
 * echoed verbatim (see the module header for the degradation contract:
 * degenerate window → no bands without consulting the evaluator; coded
 * evaluator error → `unresolvable` with the diagnostic echoed; uncoded
 * error → re-thrown). Pure and deterministic; never mutates its inputs.
 */
export function buildResourceUtilization(
  document: ProjectDocument,
  schedule: DerivedSchedule,
  allocation: ResourceAllocationQuery,
  window: TimelineViewport,
): ProjectResourceViewSurface {
  const base = { start: window.start, finish: window.finish }
  const startMs = parseInstant(window.start)
  const finishMs = parseInstant(window.finish)
  if (startMs === undefined || finishMs === undefined || finishMs <= startMs) {
    // A degenerate window is not a span — there is nothing to project and
    // the canonical evaluator is not consulted (never invented values).
    return { ...base, status: 'ok', resources: [] }
  }
  let entries: readonly ResourceAllocation[]
  try {
    entries = allocation(document, schedule)
  } catch (error) {
    const code = (error as { code?: unknown }).code
    if (typeof code !== 'string') throw error
    const message = error instanceof Error ? error.message : String(error)
    return {
      ...base,
      status: 'unresolvable',
      diagnostic: { code, severity: 'error', message },
    }
  }
  const names = new Map<string, string>()
  for (const resource of document.resources) names.set(resource.id as string, resource.name)
  return {
    ...base,
    status: 'ok',
    resources: entries.map((entry) => ({
      resourceId: entry.resourceId,
      ...(names.get(entry.resourceId as string) !== undefined
        ? { name: names.get(entry.resourceId as string) }
        : {}),
      bands: projectBands(entry, startMs, finishMs),
    })),
  }
}

/** Clips one resource's segments to the viewport: parse, drop the spans
 * that do not intersect, clip the rest, keep the authority's values
 * verbatim — pure projection, no aggregation, no merging, no
 * classification. */
function projectBands(
  entry: ResourceAllocation,
  viewportStartMs: number,
  viewportFinishMs: number,
): readonly ProjectResourceBand[] {
  const bands: ProjectResourceBand[] = []
  for (const segment of entry.segments) {
    const startMs = parseInstant(segment.start)
    const finishMs = parseInstant(segment.finish)
    if (startMs === undefined || finishMs === undefined || finishMs <= startMs) continue
    const clippedStart = Math.max(startMs, viewportStartMs)
    const clippedFinish = Math.min(finishMs, viewportFinishMs)
    if (clippedFinish <= clippedStart) continue
    bands.push({
      resourceId: entry.resourceId,
      start: formatInstant(clippedStart),
      finish: formatInstant(clippedFinish),
      demandUnits: segment.demandUnits,
      capacityUnits: segment.capacityUnits,
      overallocated: segment.overallocated,
      assignmentIds: segment.assignmentIds,
    })
  }
  return bands
}
