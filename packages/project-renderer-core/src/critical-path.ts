/**
 * PROJECT-026 — the critical-path projection layer.
 *
 * The authorized semantic path (verbatim from the work-item directive):
 *
 *     ProjectDocument + DerivedSchedule
 *                 ↓
 *     project-renderer-core critical-path projection
 *                 ↓
 *     critical-path visualizations
 *
 * The renderer visualizes canonical critical-path and float values ALREADY
 * PRODUCED by the scheduling authority; it never re-derives them. The
 * PROJECT-012 canonical clarification is explicit: "Critical-path and float
 * calculations belong exclusively in the scheduling engine. Renderer, host,
 * UI, and selector code MUST NOT compute authoritative critical path or
 * float. No second CPM engine is created." This module is that rule made
 * structural:
 *
 * - Every criticality/float value is a VERBATIM echo of the projection
 *   rows' by-reference `TaskSchedule` join: `critical`, `totalSlack`,
 *   `freeSlack` are read straight off the authority's object — there is no
 *   working-time arithmetic, no dependency-graph traversal, and no slack
 *   formula anywhere in this module. Without those primitives no CPM or
 *   float computation is possible; the architecture discipline suite scans
 *   this module for exactly those markers.
 *
 * - The slack/float BAR is pure interval GEOMETRY over two canonical
 *   instants — the task's current scheduled finish to its canonical late
 *   finish (`[scheduledFinish, lateFinish)`), clamped to the viewport with
 *   the same edge-flag convention the Gantt bars use. The bar exists iff
 *   the authority's OWN `totalSlack` is positive (there IS float to draw —
 *   the authority's answer, not a renderer-side calculation) and the late
 *   finish follows the current finish as a real span. Zero-slack critical
 *   tasks carry no float bar even when the authority's snapped late dates
 *   differ wall-clock (a zero-slack task's lateFinish can sit at the next
 *   working instant while the WORKING-TIME distance is zero — the bar
 *   visualizes FLOAT, and the authority says there is none); negative
 *   slack (the late finish precedes the current finish — the authority's
 *   impossible-schedule signal) produces no geometry either, and the
 *   echoed (possibly negative) values remain the authority's own — never
 *   clamped, never interpreted here.
 *
 * - A dependency link is classified critical iff BOTH of its canonical
 *   endpoint tasks carry `critical: true` — a pure two-boolean projection
 *   of the authority's task flags (the "every legitimately critical task
 *   on every critical chain" rule of PROJECT-012), NEVER a driving-path
 *   analysis: which links actually bind the chain is a slack analysis that
 *   stays with the scheduler. The convention is documented as exactly
 *   that: a projection of task criticality onto the link surface.
 *
 * Degradation mirrors the accepted projection rules: a row without a
 * schedule carries no float entry (dates are never invented,
 * architecture-lock §11), and a degenerate viewport yields an EMPTY
 * surface rather than invented values (the PROJECT-022 rule).
 *
 * Determinism: every function is pure — the same inputs always produce
 * byte-identical outputs (3×-repeat tested), no wall clock, no randomness,
 * no locale ordering; inputs are never mutated.
 */
import type { DependencyId, ProjectDocument, TaskId } from '@genoffice/project-contracts'
import type { ProjectViewProjection } from './projection.js'
import type { TimelineViewport } from './state.js'
import { parseInstant } from './state.js'
import { viewportSpanMs } from './views/gantt-bars.js'
import type { ProjectRowWindow } from './views/virtualization.js'
import { rowWindowIsEmpty } from './views/virtualization.js'

/** The float/slack geometry of one in-window row. Present iff the row's
 * schedule carries BOTH `scheduledFinish` and `lateFinish` as a REAL span
 * (`lateFinish > scheduledFinish`); fractions are clamped to [0, 1] and the
 * flags state when the span extends beyond the viewport on either side —
 * the Gantt-bar edge-flag convention. */
export interface ProjectTaskSlackGeometry {
  readonly startFraction: number
  readonly finishFraction: number
  readonly startsBefore: boolean
  readonly finishesAfter: boolean
}

/** One in-window task's canonical critical/float projection: the
 * scheduling authority's own values echoed verbatim (joined from the
 * projection row's by-reference `TaskSchedule`), plus the pure interval
 * geometry of the float bar. */
export interface ProjectTaskFloat {
  readonly taskId: TaskId
  /** Absolute index into the projection's visible rows. */
  readonly rowIndex: number
  /** Canonical echo: `TaskSchedule.critical` (critical := totalSlack <= 0). */
  readonly critical: boolean
  /** Canonical echo: `TaskSchedule.totalSlack` — SIGNED working minutes,
   * never clamped (negative slack is the authority's impossible-schedule
   * signal; the reporting/UI layer interprets it, never this module). */
  readonly totalSlack: number
  /** Canonical echo: `TaskSchedule.freeSlack` (working minutes). */
  readonly freeSlack: number
  /** The float-bar geometry — the window from the task's current scheduled
   * finish to its canonical late finish. Present iff the authority's
   * `totalSlack` is POSITIVE (there is float to draw — the authority's own
   * answer) and the late finish follows the current finish as a real span;
   * absent for zero-slack tasks (including the snapped-late-date case),
   * negative slack, and missing instants — geometry is never invented. */
  readonly slack?: ProjectTaskSlackGeometry
}

/** The critical-path surface (PROJECT-026): the canonical critical/float
 * projection of every in-window row, plus the critical-link
 * classification. Present on the timeline iff a schedule was joined into
 * the projection; EMPTY values are real information (nothing scheduled in
 * the window / nothing critical / no critical links). */
export interface ProjectCriticalPathSurface {
  /** Every in-window row that carries a schedule, in ascending row order
   * (rows without a schedule carry no float — never invented). */
  readonly floats: readonly ProjectTaskFloat[]
  /** The document dependencies (in canonical document order) whose BOTH
   * canonical endpoint tasks carry `critical: true` — the pure
   * both-endpoints projection convention (see the module header); never a
   * driving-path analysis. */
  readonly criticalDependencyIds: readonly DependencyId[]
}

/**
 * Builds the critical-path surface. Pure and deterministic; never mutates
 * its inputs. A degenerate viewport (unparseable or empty span) yields an
 * EMPTY surface rather than invented values. Every float value is the
 * authority's own echo; the slack geometry is pure two-instant interval
 * projection over the viewport.
 */
export function buildCriticalPath(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
): ProjectCriticalPathSurface {
  const span = viewportSpanMs(viewport)
  if (span === undefined || rowWindowIsEmpty(rowWindow)) {
    return { floats: [], criticalDependencyIds: [] }
  }

  const floats: ProjectTaskFloat[] = []
  for (
    let index = rowWindow.firstIndex;
    index <= rowWindow.lastIndex && index < projection.rows.length;
    index += 1
  ) {
    const row = projection.rows[index]
    if (row === undefined) continue
    const schedule = row.schedule
    if (schedule === undefined) continue
    const scheduledFinishMs = parseInstant(schedule.scheduledFinish ?? '')
    const lateFinishMs = parseInstant(schedule.lateFinish ?? '')
    // The float bar exists iff the authority says there IS float (a
    // positive totalSlack — its own answer, read verbatim) AND the two
    // canonical instants form a real span. Zero slack with snapped late
    // dates and negative slack both carry no geometry; the echoed values
    // below remain the authority's own either way.
    const slack =
      schedule.totalSlack > 0 &&
      scheduledFinishMs !== undefined &&
      lateFinishMs !== undefined &&
      lateFinishMs > scheduledFinishMs
        ? slackGeometry(scheduledFinishMs, lateFinishMs, span)
        : undefined
    floats.push({
      taskId: row.taskId,
      rowIndex: index,
      critical: schedule.critical,
      totalSlack: schedule.totalSlack,
      freeSlack: schedule.freeSlack,
      ...(slack !== undefined ? { slack } : {}),
    })
  }

  // The critical-link classification: a pure two-boolean projection of the
  // canonical endpoint tasks' own critical flags. The schedule join comes
  // from the projection rows (the authority's by-reference objects).
  const criticalByTask = new Map<TaskId, boolean>()
  for (const row of projection.rows) {
    criticalByTask.set(row.taskId, row.schedule?.critical ?? false)
  }
  const criticalDependencyIds: DependencyId[] = []
  for (const dependency of document.dependencies) {
    if (criticalByTask.get(dependency.predecessorId) !== true) continue
    if (criticalByTask.get(dependency.successorId) !== true) continue
    criticalDependencyIds.push(dependency.id)
  }

  return { floats, criticalDependencyIds }
}

/** The clamped float-bar geometry of `[startMs, finishMs)` over the
 * viewport span — the same clamping/flag convention the Gantt bars use. */
function slackGeometry(
  startMs: number,
  finishMs: number,
  span: { start: number; ms: number },
): ProjectTaskSlackGeometry {
  const rawStart = (startMs - span.start) / span.ms
  const rawFinish = (finishMs - span.start) / span.ms
  return {
    startFraction: clamp01(rawStart),
    finishFraction: clamp01(rawFinish),
    startsBefore: rawStart < 0,
    finishesAfter: rawFinish > 1,
  }
}

const clamp01 = (value: number): number =>
  Number.isFinite(value) ? Math.max(0, Math.min(1, value)) : 0
