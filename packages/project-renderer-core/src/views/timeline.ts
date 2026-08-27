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
 * PROJECT-025 — calendar visualization: when a canonical working-time
 * query is threaded (`CalendarViewInput`), the timeline also carries the
 * calendar surfaces — the background surface (the project default
 * calendar, or the caller's explicit choice) and one surface per in-window
 * row, shaded by that row's task's RESOLVED calendar
 * (`TaskSchedule.resolvedCalendarId` — the scheduling authority's own
 * answer, never re-derived here). Every evaluated instant comes from the
 * injected query (`../calendar.js` — no second calendar engine); rows
 * without a resolved calendar id carry no surface (never invented), and
 * surfaces for identical calendar ids are shared by reference within one
 * build (a per-build memo — nothing is cached across renders,
 * architecture-lock §11).
 *
 * Pure composition: the timeline owns no semantics of its own.
 */
import type { CalendarId, ProjectDocument, TaskId } from '@genoffice/project-contracts'
import type { ProjectTaskRow, ProjectViewProjection } from '../projection.js'
import type { ProjectViewState, TimelineViewport } from '../state.js'
import {
  type CalendarViewInput,
  type CalendarWorkingTimeQuery,
  type ProjectCalendarSurface,
  buildCalendarSurface,
} from '../calendar.js'
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

/** One per-row calendar surface (PROJECT-025): the row's task resolved
 * calendar (`TaskSchedule.resolvedCalendarId` — the scheduling
 * authority's echo) and that calendar's evaluated working-time surface
 * over the timeline viewport. Hosts shade the row's background with the
 * bands; entries exist only for rows whose schedule carries a resolved
 * calendar id. */
export interface ProjectRowCalendar {
  readonly rowIndex: number
  readonly taskId: TaskId
  readonly calendarId: CalendarId
  readonly surface: ProjectCalendarSurface
}

/** The timeline view model (PROJECT-022 + PROJECT-025 calendar surface). */
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
  /** PROJECT-025 — the timeline background's working-time surface (the
   * project default calendar unless `CalendarViewInput.calendarId`
   * overrides it). Present iff a working-time query was threaded and the
   * viewport is a real span; absent otherwise — never invented. */
  readonly calendar?: ProjectCalendarSurface
  /** PROJECT-025 — per-row working-time surfaces (one per in-window row
   * whose task has a resolved calendar id), in ascending row order.
   * Present under the same conditions as `calendar`; EMPTY when the
   * query was threaded but no in-window row carries a resolved calendar
   * (real information — e.g. no scheduler wired). */
  readonly rowCalendars?: readonly ProjectRowCalendar[]
}

/**
 * Builds the timeline view model. Pure and deterministic; never mutates
 * its inputs. An unparseable/degenerate viewport yields an EMPTY model
 * (no bands, no rows, no geometry, no calendar surfaces) rather than
 * invented values.
 *
 * PROJECT-024: the optional `state` parameter threads the interaction
 * state's dependency selection / edit target into the link surface
 * (`selected`/`editingField` reflections) — the PROJECT-023 row-reflection
 * contract's dependency analog. Omitted (or a state with nothing selected
 * / edited), the links carry `selected: false` and no `editingField`; the
 * geometry is IDENTICAL either way (reflection is a pure echo, never a
 * geometry input).
 *
 * PROJECT-025: the optional `calendar` parameter threads the injected
 * canonical working-time query; present query → the background +
 * per-row calendar surfaces. The calendar surfaces are ADDITIVE: with or
 * without them, `bands`/`rows`/`bars`/`milestones`/`links` are
 * byte-identical (reflection/projection never feeds geometry).
 */
export function buildTimeline(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
  state?: ProjectViewState,
  calendar?: CalendarViewInput,
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
  const surfaces =
    span !== undefined && calendar?.workingTime !== undefined
      ? buildTimelineCalendars(
          document,
          projection,
          viewport,
          rowWindow,
          calendar.workingTime,
          calendar.calendarId,
        )
      : undefined
  return {
    viewport,
    axisLevel,
    bands: span !== undefined ? buildTimeAxis(viewport, axisLevel) : [],
    rowWindow,
    rows,
    bars: buildGanttBars(projection, viewport, rowWindow),
    milestones: buildMilestones(projection, viewport, rowWindow),
    links: buildDependencies(document, projection, viewport, rowWindow, state),
    ...(surfaces !== undefined
      ? { calendar: surfaces.calendar, rowCalendars: surfaces.rowCalendars }
      : {}),
  }
}

/**
 * The calendar surfaces of one timeline build: the background surface and
 * the per-row surfaces. Surfaces for identical calendar ids are computed
 * ONCE per build and shared by reference (the query is pure; the memo is
 * build-local — nothing persists across renders, lock §11). The
 * background calendar is `calendarId` when given, else the document's
 * `defaultCalendarId`. Row surfaces read the scheduling authority's
 * `resolvedCalendarId` from the row's verbatim schedule echo — the
 * calendar choice is never re-derived here.
 */
function buildTimelineCalendars(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
  workingTime: CalendarWorkingTimeQuery,
  calendarId: CalendarId | undefined,
): { calendar: ProjectCalendarSurface; rowCalendars: readonly ProjectRowCalendar[] } {
  const memo = new Map<string, ProjectCalendarSurface>()
  const surfaceFor = (id: CalendarId): ProjectCalendarSurface => {
    const key = id as string
    let surface = memo.get(key)
    if (surface === undefined) {
      surface = buildCalendarSurface(document, workingTime, viewport, id)
      memo.set(key, surface)
    }
    return surface
  }
  const calendar = surfaceFor(calendarId ?? document.properties.defaultCalendarId)
  const rowCalendars: ProjectRowCalendar[] = []
  if (!rowWindowIsEmpty(rowWindow)) {
    for (
      let index = rowWindow.firstIndex;
      index <= rowWindow.lastIndex && index < projection.rows.length;
      index += 1
    ) {
      const row = projection.rows[index]
      if (row === undefined) continue
      const resolvedCalendarId = row.schedule?.resolvedCalendarId
      if (resolvedCalendarId === undefined) continue
      rowCalendars.push({
        rowIndex: index,
        taskId: row.taskId,
        calendarId: resolvedCalendarId,
        surface: surfaceFor(resolvedCalendarId),
      })
    }
  }
  return { calendar, rowCalendars }
}
