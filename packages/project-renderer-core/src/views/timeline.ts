/**
 * PROJECT-022 — the timeline (right pane) view model.
 *
 * `buildTimeline` composes the time-scaled half of the Gantt view over the
 * SHARED virtualized row window: the accepted `buildTimeAxis` bands for the
 * viewport, the in-window row surface, and the three geometry surfaces —
 * Gantt bars (`./gantt-bars.js`), milestones (`./milestones.js`), and
 * dependency links (`./dependencies.js`) — all addressing rows by their
 * absolute visible-row index, all in fraction space (hosts multiply by
 * pixels; label formatting stays a host/locale concern).
 *
 * Pure composition: the timeline owns no semantics of its own.
 */
import type { ProjectDocument } from '@genoffice/project-contracts'
import type { ProjectTaskRow, ProjectViewProjection } from '../projection.js'
import type { ProjectViewState, TimelineViewport } from '../state.js'
import {
  type TimeAxisBand,
  type TimeAxisLevel,
  axisLevelForSpan,
  buildTimeAxis,
} from '../timeline.js'
import type { ProjectDependencies } from './dependencies.js'
import { buildDependencies } from './dependencies.js'
import type { ProjectGanttBars } from './gantt-bars.js'
import { buildGanttBars, viewportSpanMs } from './gantt-bars.js'
import type { ProjectMilestones } from './milestones.js'
import { buildMilestones } from './milestones.js'
import type { ProjectRowWindow } from './virtualization.js'
import { rowWindowIsEmpty } from './virtualization.js'

/** One in-window timeline row: the absolute index and the projection row
 * joined BY REFERENCE (hosts derive pixel y from
 * `(rowIndex - rowWindow.firstIndex) × rowHeight`). */
export interface ProjectTimelineRow {
  readonly index: number
  readonly row: ProjectTaskRow
}

/** The timeline view model (PROJECT-022). */
export interface ProjectTimeline {
  /** The viewport the geometry is expressed over (echoed by reference). */
  readonly viewport: TimelineViewport
  /** The axis level the bands were built at (chosen deterministically
   * from the viewport span when the caller did not pin one). */
  readonly axisLevel: TimeAxisLevel
  /** The contiguous `[start, finish)` axis bands covering the viewport. */
  readonly bands: readonly TimeAxisBand[]
  readonly rowWindow: ProjectRowWindow
  readonly rows: readonly ProjectTimelineRow[]
  readonly bars: ProjectGanttBars
  readonly milestones: ProjectMilestones
  readonly links: ProjectDependencies
}

/**
 * Builds the timeline view model. Pure and deterministic; never mutates
 * its inputs. An unparseable/degenerate viewport yields an EMPTY model
 * (no bands, no rows, no geometry) rather than invented values.
 *
 * PROJECT-024: the optional `state` parameter threads the interaction
 * state's dependency selection / edit target into the link surface
 * (`selected`/`editingField` reflections) — the PROJECT-023 row-reflection
 * contract's dependency analog. Omitted (or a state with nothing selected
 * / edited), the links carry `selected: false` and no `editingField`; the
 * geometry is IDENTICAL either way (reflection is a pure echo, never a
 * geometry input).
 */
export function buildTimeline(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
  state?: ProjectViewState,
): ProjectTimeline {
  const span = viewportSpanMs(viewport)
  const axisLevel: TimeAxisLevel =
    span !== undefined ? axisLevelForSpan(span.ms) : axisLevelForSpan(0)
  const rows: ProjectTimelineRow[] = []
  if (span !== undefined && !rowWindowIsEmpty(rowWindow)) {
    for (
      let index = rowWindow.firstIndex;
      index <= rowWindow.lastIndex && index < projection.rows.length;
      index += 1
    ) {
      const row = projection.rows[index]
      if (row !== undefined) rows.push({ index, row })
    }
  }
  return {
    viewport,
    axisLevel,
    bands: span !== undefined ? buildTimeAxis(viewport, axisLevel) : [],
    rowWindow,
    rows,
    bars: buildGanttBars(projection, viewport, rowWindow),
    milestones: buildMilestones(projection, viewport, rowWindow),
    links: buildDependencies(document, projection, viewport, rowWindow, state),
  }
}
