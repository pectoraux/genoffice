/**
 * PROJECT-025 — the calendar projection layer.
 *
 * The authorized semantic path (verbatim from the work-item directive):
 *
 *     ProjectDocument calendars
 *             +
 *   canonical scheduling semantics
 *             ↓
 *   project-renderer-core calendar projection
 *             ↓
 *   Timeline / Gantt visualization
 *
 * The renderer visualizes canonical calendar semantics ALREADY PRODUCED by
 * the Project domain; it never re-derives them. Concretely there is NO
 * second calendar engine here:
 *
 * - Every "which instants are working time" answer comes from the INJECTED
 *   `CalendarWorkingTimeQuery` — a structural function type satisfied by
 *   binding `resolveCalendar` + `workingIntervals` from
 *   `@genoffice/project-scheduling` (the accepted `ScheduleRunner`
 *   injection precedent; this package stays statically scheduling-free,
 *   architecture-lock §3/§6). The canonical binding hosts write is:
 *
 *     (calendars, calendarId, start, finish) =>
 *       workingIntervals(resolveCalendar({ calendars }, calendarId), start, finish)
 *
 * - The canonical working intervals are joined VERBATIM into the band
 *   surface; `classifyCalendarBands` performs pure interval algebra over
 *   the evaluator's output (clip / sort / merge / complement) — it never
 *   decomposes a date into weekday or date parts, never reads
 *   `workingWeek`, never matches exception dates. Without date
 *   decomposition no calendar evaluation is possible; the architecture
 *   discipline suite scans this module for exactly those markers.
 *
 * - The calendar CATALOG is a pure structural echo of
 *   `ProjectDocument.calendars` (verbatim references) plus display
 *   aggregations of the DECLARATIONS (sums of declared period minutes,
 *   counts) — no evaluation over dates, no inheritance resolution (the
 *   resolved view IS the evaluated band surface).
 *
 * Degradation mirrors the scheduling authority exactly: `schedule()`
 * catches calendar errors and degrades to diagnostics (empty task
 * schedules + a diagnostic echo) — never a crash, never invented values.
 * `buildCalendarSurface` applies the same boundary: an evaluator error
 * that carries a string code (structurally, the canonical `CalendarError`
 * family) degrades to `status: 'unresolvable'` with the diagnostic echoed
 * verbatim and NO bands; an error without a code (a host-binding bug) is
 * re-thrown, exactly like `schedule()` re-throws non-calendar errors. A
 * degenerate window (unparseable or empty span) has no bands and does not
 * consult the evaluator at all.
 *
 * Determinism: every function is pure — the same inputs always produce
 * byte-identical outputs (3×-repeat tested), no wall clock, no randomness,
 * no locale ordering; inputs are never mutated.
 */
import type {
  Calendar,
  CalendarId,
  ImportDiagnostic,
  ISODateTime,
  ProjectDocument,
} from '@genoffice/project-contracts'
import { asISODateTime } from '@genoffice/project-contracts'
import { parseInstant, formatInstant } from './state.js'
import type { TimelineViewport } from './state.js'

/** One working sub-interval of a window, as produced by the canonical
 * working-time evaluator. Structurally identical to the scheduling
 * package's `workingIntervals` element — the canonical binding returns its
 * output verbatim and this layer joins it by value without re-deriving
 * anything. */
export interface CalendarWorkingInterval {
  readonly start: ISODateTime
  readonly finish: ISODateTime
}

/**
 * The injected canonical working-time evaluator. Hosts bind this to the
 * scheduling package's `resolveCalendar` + `workingIntervals` (see the
 * module header); the renderer core NEVER implements it — implementing it
 * here would be a second calendar engine, which the PROJECT-025 directive
 * forbids. The query receives the document's calendars (the canonical
 * calendar book data) so the binding stays a pure, document-independent
 * function hosts create once. Errors: a thrown error carrying a string
 * `code` (the canonical `CalendarError` family) degrades the surface to
 * `unresolvable`; errors without a code propagate (mirroring how
 * `schedule()` degrades calendar errors and re-throws the rest).
 */
export type CalendarWorkingTimeQuery = (
  calendars: readonly Calendar[],
  calendarId: CalendarId,
  start: ISODateTime,
  finish: ISODateTime,
) => readonly CalendarWorkingInterval[]

/** One contiguous `[start, finish)` band of the calendar surface,
 * classified working/non-working. Bands are contiguous and ordered: the
 * union of a surface's bands covers its window exactly, each band's
 * `finish` equals its successor's `start`, and no band is ever empty.
 * Instants are canonical UTC ISO strings (the view-model convention of
 * `TimeAxisBand`; hosts map them to pixels through the viewport math). */
export interface ProjectCalendarBand {
  readonly start: string
  readonly finish: string
  readonly working: boolean
}

/** `ok` — bands present (possibly empty when the window contains no
 * working time). `unresolvable` — the canonical evaluator rejected the
 * calendar (missing id, inheritance cycle, malformed periods); bands are
 * absent, never invented, and the diagnostic is echoed. */
export type CalendarSurfaceStatus = 'ok' | 'unresolvable'

/** The evaluated working-time surface of ONE calendar over ONE window.
 * Working bands are the canonical evaluator's intervals (clipped to the
 * window); non-working bands are their complement inside the window —
 * pure projection, never re-derived. */
export interface ProjectCalendarSurface {
  /** The calendar this surface was evaluated for (explicit id or the
   * document's `defaultCalendarId`). */
  readonly calendarId: CalendarId
  /** The calendar's display name, echoed from the document when the id
   * exists there (pure echo; absent otherwise). */
  readonly name?: string
  readonly status: CalendarSurfaceStatus
  /** Contiguous bands covering `[start, finish)` — present iff
   * `status === 'ok'`. */
  readonly bands?: readonly ProjectCalendarBand[]
  /** The canonical evaluator's error echo, present iff
   * `status === 'unresolvable'` — the same code/message the scheduling
   * authority surfaces for calendar failures (an `ImportDiagnostic`
   * shape, so hosts can feed it into their diagnostics surface). */
  readonly diagnostic?: ImportDiagnostic
  /** The evaluated window, echoed verbatim. */
  readonly start: string
  readonly finish: string
}

/** The fallback diagnostic code for evaluator errors that carry no code
 * of their own. The canonical `CalendarError` family always carries a
 * string code, so this label only ever appears for non-canonical
 * failures of a host-supplied binding — which `buildCalendarSurface`
 * re-throws — making it a defensive guard, not an expected value. */
export const CALENDAR_EVALUATION_FAILED = 'CALENDAR_EVALUATION_FAILED'

/**
 * Classifies a window into contiguous working/non-working bands from the
 * canonical evaluator's working intervals. PURE interval algebra: parse,
 * clip to the window, drop empty spans, sort by start (then finish), merge
 * overlapping/adjacent spans, and complement the gaps — no calendar
 * semantics anywhere (no weekday lookup, no exception matching; the
 * intervals arrive already evaluated). A degenerate window (unparseable
 * or empty span) yields no bands. Deterministic: the output is
 * independent of the input intervals' order.
 */
export function classifyCalendarBands(
  window: TimelineViewport,
  working: readonly CalendarWorkingInterval[],
): readonly ProjectCalendarBand[] {
  const windowStart = parseInstant(window.start)
  const windowFinish = parseInstant(window.finish)
  if (windowStart === undefined || windowFinish === undefined || windowFinish <= windowStart) {
    return []
  }
  const spans: { start: number; finish: number }[] = []
  for (const interval of working) {
    const start = parseInstant(interval.start)
    const finish = parseInstant(interval.finish)
    if (start === undefined || finish === undefined || finish <= start) continue
    const clippedStart = Math.max(start, windowStart)
    const clippedFinish = Math.min(finish, windowFinish)
    if (clippedFinish > clippedStart) spans.push({ start: clippedStart, finish: clippedFinish })
  }
  spans.sort((a, b) => a.start - b.start || a.finish - b.finish)
  const merged: { start: number; finish: number }[] = []
  for (const span of spans) {
    const last = merged[merged.length - 1]
    if (last !== undefined && span.start <= last.finish) {
      if (span.finish > last.finish) last.finish = span.finish
    } else {
      merged.push({ start: span.start, finish: span.finish })
    }
  }
  const bands: ProjectCalendarBand[] = []
  let cursor = windowStart
  for (const span of merged) {
    if (span.start > cursor) {
      bands.push({
        start: formatInstant(cursor),
        finish: formatInstant(span.start),
        working: false,
      })
    }
    bands.push({
      start: formatInstant(span.start),
      finish: formatInstant(span.finish),
      working: true,
    })
    cursor = span.finish
  }
  if (cursor < windowFinish) {
    bands.push({
      start: formatInstant(cursor),
      finish: formatInstant(windowFinish),
      working: false,
    })
  }
  return bands
}

/**
 * Builds the evaluated working-time surface of one calendar over one
 * window. `calendarId` defaults to the document's `defaultCalendarId`
 * (the canonical project calendar). The evaluator is consulted exactly
 * once; its intervals are classified by pure interval algebra. See the
 * module header for the degradation contract (degenerate window → no
 * bands without consulting the evaluator; coded evaluator error →
 * `unresolvable` with the diagnostic echoed; uncoded error → re-thrown).
 * Pure and deterministic; never mutates its inputs.
 */
export function buildCalendarSurface(
  document: ProjectDocument,
  workingTime: CalendarWorkingTimeQuery,
  window: TimelineViewport,
  calendarId?: CalendarId,
): ProjectCalendarSurface {
  const id = calendarId ?? document.properties.defaultCalendarId
  const name = document.calendars.find((calendar) => calendar.id === id)?.name
  const base = {
    calendarId: id,
    ...(name !== undefined ? { name } : {}),
    start: window.start,
    finish: window.finish,
  }
  const startMs = parseInstant(window.start)
  const finishMs = parseInstant(window.finish)
  if (startMs === undefined || finishMs === undefined || finishMs <= startMs) {
    // A degenerate window is not a span — there is nothing to shade and
    // the canonical evaluator is not consulted (never invented values).
    return { ...base, status: 'ok', bands: [] }
  }
  let working: readonly CalendarWorkingInterval[]
  try {
    // The single brand boundary: the window carries view-model strings,
    // the canonical evaluator's contract carries branded instants.
    working = workingTime(
      document.calendars,
      id,
      asISODateTime(window.start),
      asISODateTime(window.finish),
    )
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
  return { ...base, status: 'ok', bands: classifyCalendarBands(window, working) }
}

/** One calendar in the catalog: a VERBATIM echo of the canonical
 * definition (joined by reference) plus pure display aggregations of the
 * DECLARATION. Nothing here evaluates a calendar — the resolved view is
 * the evaluated band surface. */
export interface ProjectCalendarCatalogEntry {
  readonly calendarId: CalendarId
  readonly name: string
  readonly baseCalendarId?: CalendarId
  /** The verbatim canonical `Calendar` (by reference — never copied). */
  readonly calendar: Calendar
  /** Weekday keys (0=Sunday..6=Saturday, the canonical UTC
   * weekday-numbering convention of the scheduling primitives) that carry
   * at least one declared working period — a weekday declared with an
   * empty period list is a non-working day — ascending. */
  readonly workingWeekdays: readonly number[]
  /** Declared working minutes per weekday — the sum of the declared
   * period bounds (`end − start`) for that weekday. A display
   * aggregation of the declaration, never an evaluation over dates. */
  readonly declaredMinutes: Readonly<Record<number, number>>
  /** Total declared working minutes across the week. */
  readonly declaredWeeklyMinutes: number
  /** The number of declared exception dates. */
  readonly exceptionCount: number
  /** How many tasks reference this calendar directly (a pure document
   * count of `task.calendarId` — the RESOLVED calendar per task is the
   * scheduling authority's answer and stays schedule-joined). */
  readonly taskCount: number
}

/** The calendar catalog: every document calendar echoed in document
 * order, with the project's default calendar id echoed alongside. Hosts
 * render this as the calendar view / legend; it requires NO evaluator
 * (pure document projection). */
export interface ProjectCalendarCatalog {
  readonly defaultCalendarId: CalendarId
  readonly calendars: readonly ProjectCalendarCatalogEntry[]
}

/**
 * Projects the document's calendars into the catalog (pure structural
 * echo + display aggregations; see `ProjectCalendarCatalogEntry`).
 * Deterministic: document order, ascending weekday keys, integer sums.
 * Never mutates the document.
 */
export function buildCalendarCatalog(document: ProjectDocument): ProjectCalendarCatalog {
  const entries = document.calendars.map((calendar) => {
    const weekdayKeys = Object.keys(calendar.workingWeek)
      .map((key) => Number(key))
      .filter((key) => Number.isInteger(key) && key >= 0 && key <= 6)
      .filter((key) => (calendar.workingWeek[key] ?? []).length > 0)
      .sort((a, b) => a - b)
    const declaredMinutes: Record<number, number> = {}
    let weekly = 0
    for (const weekday of weekdayKeys) {
      const periods = calendar.workingWeek[weekday] ?? []
      let minutes = 0
      for (const period of periods) {
        const span = period.endMinute - period.startMinute
        if (span > 0) minutes += span
      }
      declaredMinutes[weekday] = minutes
      weekly += minutes
    }
    return {
      calendarId: calendar.id,
      name: calendar.name,
      ...(calendar.baseCalendarId !== undefined ? { baseCalendarId: calendar.baseCalendarId } : {}),
      calendar,
      workingWeekdays: weekdayKeys,
      declaredMinutes,
      declaredWeeklyMinutes: weekly,
      exceptionCount: calendar.exceptions.length,
      taskCount: document.tasks.filter((task) => task.calendarId === calendar.id).length,
    }
  })
  return {
    defaultCalendarId: document.properties.defaultCalendarId,
    calendars: entries,
  }
}

/** The PROJECT-025 calendar visualization inputs for the Gantt/timeline
 * builders: the injected canonical working-time query (required for any
 * evaluated surface) and the calendar the timeline background is shaded
 * by (defaults to the document's `defaultCalendarId`). A per-RENDER
 * input, like the layout inputs — never persisted `ProjectViewState`
 * (the PROJECT-021 boundary between persisted interaction state and
 * ephemeral host render choices). Absent `workingTime` → no calendar
 * surfaces are built (never invented). */
export interface CalendarViewInput {
  readonly workingTime?: CalendarWorkingTimeQuery
  readonly calendarId?: CalendarId
}
