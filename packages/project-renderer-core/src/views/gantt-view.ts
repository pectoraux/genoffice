/**
 * PROJECT-022 — the synchronized Gantt view composition.
 *
 * `buildGanttView` is the single entry both hosts call per render: it
 * computes the ONE shared row window from the host's logical layout inputs
 * and builds BOTH panes from it — the task grid (`./grid.js`) and the
 * timeline (`./timeline.js`, which composes bars, milestones, and
 * dependency links). Synchronization is by construction: the grid rows,
 * timeline rows, bars, milestones, and links all address rows by the same
 * absolute visible-row indices over the same window, so the two panes
 * cannot drift (the acceptance contract "synchronized virtualized core
 * surface operational").
 *
 * Layout inputs (`firstRow`, `visibleRows`, `overscan`) are HOST scroll
 * state passed as plain arguments — logical rows, never pixels, and never
 * persisted `ProjectViewState` (the PROJECT-021 boundary between
 * persisted interaction state and ephemeral host layout).
 *
 * `hitTestGantt` is the geometry→entity inverse for pointer interaction:
 * given a point in (absolute row index, viewport fraction) space it
 * returns the milestone or bar under it, milestones first (the smaller,
 * more precise target wins at overlapping positions). Link hit-testing
 * requires a pixel tolerance against the route polyline and stays a host
 * concern. What hosts DO with a hit (selection intents, editing) is
 * PROJECT-023/024 scope.
 *
 * PROJECT-025: the optional `calendar` parameter threads the injected
 * canonical working-time query (+ optional background calendar id) into
 * the timeline's calendar surfaces. It is a per-render input like the
 * layout inputs — never persisted view state — and is ADDITIVE: the
 * geometry surfaces are identical with or without it.
 *
 * PROJECT-026: the optional `resources` parameter threads the injected
 * canonical allocation query + the current derived schedule into the
 * timeline's resource-visualization surface (the critical-path surface
 * joins automatically from the projection's schedule — no input needed).
 * Per-render input, ADDITIVE geometry, exactly like `calendar`.
 */
import type { ProjectDocument, TaskId } from '@genoffice/project-contracts'
import type { CalendarViewInput } from '../calendar.js'
import type { ResourceViewInput } from '../resources.js'
import type { ProjectViewProjection } from '../projection.js'
import type { ProjectViewState } from '../state.js'
import { type ProjectTaskGrid, buildTaskGrid } from './grid.js'
import { type ProjectTimeline, buildTimeline } from './timeline.js'
import { type ProjectRowWindow, type RowWindowInput, buildRowWindow } from './virtualization.js'

/** The synchronized Gantt view (PROJECT-022): both panes over one shared
 * virtualized row window. */
export interface ProjectGanttView {
  readonly rowWindow: ProjectRowWindow
  readonly taskGrid: ProjectTaskGrid
  readonly timeline: ProjectTimeline
}

/** The host layout inputs for one render (logical rows, never pixels). */
export type GanttViewLayout = RowWindowInput

/**
 * Builds the synchronized Gantt view: one shared row window, both panes.
 * Pure and deterministic: the same `(document, projection, state, layout,
 * calendar?, resources?)` always produces the same view; inputs are never
 * mutated.
 */
export function buildGanttView(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  state: ProjectViewState,
  layout: GanttViewLayout,
  calendar?: CalendarViewInput,
  resources?: ResourceViewInput,
): ProjectGanttView {
  const rowWindow = buildRowWindow(projection.rows.length, layout)
  return {
    rowWindow,
    taskGrid: buildTaskGrid(document, projection, rowWindow),
    // PROJECT-024: the full state (not just the viewport) flows into the
    // timeline so the dependency-link surface carries the interaction-state
    // reflection (selected/editingField) alongside its geometry.
    // PROJECT-025: the calendar input threads the injected canonical
    // working-time query into the timeline's calendar surfaces (absent →
    // no surfaces, never invented).
    // PROJECT-026: the resources input threads the injected canonical
    // allocation query + the current derived schedule into the
    // resource-visualization surface; the critical-path surface joins
    // automatically from the projection's schedule join.
    timeline: buildTimeline(
      document,
      projection,
      state.viewport,
      rowWindow,
      state,
      calendar,
      resources,
    ),
  }
}

/** What the pointer is over (the geometry→entity inverse). */
export type GanttHitTarget =
  | { readonly kind: 'milestone'; readonly taskId: TaskId }
  | { readonly kind: 'bar'; readonly taskId: TaskId }

/** A pointer position in view space: the absolute visible-row index and
 * the viewport fraction (hosts derive both from pixels). */
export interface GanttHitPoint {
  readonly rowIndex: number
  readonly fraction: number
}

/**
 * Resolves the topmost geometry under a point: milestones before bars
 * (deterministic priority — the smaller target wins when they overlap).
 * `tolerance` widens the hit rectangle for milestones and bars by a
 * fraction-space epsilon (hosts pass `pixelTolerance / timelineWidth`;
 * 0 = exact containment). Pure and deterministic.
 */
export function hitTestGantt(
  timeline: ProjectTimeline,
  point: GanttHitPoint,
  tolerance = 0,
): GanttHitTarget | undefined {
  const epsilon = Number.isFinite(tolerance) ? Math.max(0, tolerance) : 0
  for (const milestone of timeline.milestones) {
    if (
      milestone.rowIndex === point.rowIndex &&
      Math.abs(milestone.fraction - point.fraction) <= epsilon
    ) {
      return { kind: 'milestone', taskId: milestone.taskId }
    }
  }
  for (const bar of timeline.bars) {
    if (
      bar.rowIndex === point.rowIndex &&
      point.fraction >= bar.startFraction - epsilon &&
      point.fraction <= bar.finishFraction + epsilon
    ) {
      return { kind: 'bar', taskId: bar.taskId }
    }
  }
  return undefined
}
