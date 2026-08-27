/**
 * PROJECT-021 — deterministic timeline math.
 *
 * Pure, locale-free, host-independent helpers for the shared timeline
 * surface: viewport scaling/fitting, instant→fraction mapping over the
 * visible window, and time-axis band generation. Hosts (Electron/web) map
 * the fractions and bands to pixels; the math itself is identical in both
 * hosts (R-009 renderer parity) and never consults the wall clock, the
 * locale, or the filesystem. All arithmetic is integer milliseconds over
 * UTC instants — the canonical time model (architecture-lock §5).
 *
 * The timeline layer is pure PROJECTION: it reads no calendar, computes no
 * working time, and derives nothing the scheduling authority owns. Calendar
 * visibility (working/non-working bands) belongs to PROJECT-025 and will be
 * projected from the document's calendars, never re-derived here.
 */
import type { DerivedSchedule, ProjectDocument } from '@genoffice/project-contracts'
import { type TimelineViewport, formatInstant, parseInstant } from './state.js'

export const MIN_VIEWPORT_SPAN_MS = 60_000
export const MAX_VIEWPORT_SPAN_MS = 100 * 366 * 24 * 60 * 60 * 1000

/** The padding ratio applied by `fitViewport`: 2% of the project span on
 * each side, with a minimum fitted total span of one day (so a
 * milestone-only project still yields a usable window). */
export const FIT_PADDING_RATIO = 0.02
export const FIT_MIN_TOTAL_SPAN_MS = 24 * 60 * 60 * 1000

const DAY_MS = 24 * 60 * 60 * 1000
/** Default fitted span when no scheduling information exists: 30 days. */
const DEFAULT_FIT_SPAN_MS = 30 * DAY_MS

const clamp = (value: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, value))

/**
 * Maps an instant to its horizontal position in the viewport as a fraction
 * in `[0, 1]` (clamped): `0` at `viewport.start`, `1` at `viewport.finish`.
 * Hosts multiply the fraction by the pixel width of the timeline surface.
 */
export function viewportFraction(viewport: TimelineViewport, instant: string): number {
  const start = parseInstant(viewport.start) ?? 0
  const finish = parseInstant(viewport.finish) ?? start + DAY_MS
  const at = parseInstant(instant) ?? start
  const span = finish - start
  if (span <= 0) return 0
  return clamp((at - start) / span, 0, 1)
}

/** Maps a viewport fraction back to an instant (inverse of `viewportFraction`). */
export function viewportInstant(viewport: TimelineViewport, fraction: number): string {
  const start = parseInstant(viewport.start) ?? 0
  const finish = parseInstant(viewport.finish) ?? start + DAY_MS
  const at = start + Math.round(clamp(fraction, 0, 1) * (finish - start))
  return formatInstant(at)
}

/**
 * Scales the visible window by `factor` around a focus instant (default: the
 * window midpoint). `factor < 1` zooms in (smaller window); `> 1` zooms out.
 * The focus instant keeps its position; the resulting span is clamped to
 * `[MIN_VIEWPORT_SPAN_MS, MAX_VIEWPORT_SPAN_MS]`. Integer-millisecond
 * arithmetic throughout — the same inputs always produce the same window.
 */
export function scaleViewport(
  viewport: TimelineViewport,
  factor: number,
  focus?: string,
): TimelineViewport {
  if (!Number.isFinite(factor) || factor <= 0) return viewport
  const startMs = parseInstant(viewport.start)
  const finishMs = parseInstant(viewport.finish)
  if (startMs === undefined || finishMs === undefined || finishMs <= startMs) return viewport
  const span = finishMs - startMs
  const focusMs = clamp(
    (focus !== undefined ? parseInstant(focus) : undefined) ?? Math.round((startMs + finishMs) / 2),
    startMs,
    finishMs,
  )
  const nextSpan = Math.round(clamp(span * factor, MIN_VIEWPORT_SPAN_MS, MAX_VIEWPORT_SPAN_MS))
  const leadRatio = span > 0 ? (focusMs - startMs) / span : 0.5
  let nextStart = Math.round(focusMs - nextSpan * leadRatio)
  let nextFinish = nextStart + nextSpan
  if (nextFinish - nextStart < MIN_VIEWPORT_SPAN_MS) {
    nextStart = focusMs - Math.round(MIN_VIEWPORT_SPAN_MS / 2)
    nextFinish = nextStart + MIN_VIEWPORT_SPAN_MS
  }
  return { start: formatInstant(nextStart), finish: formatInstant(nextFinish) }
}

/** The project window to fit: the derived schedule's project span when a
 * schedule is available, else the canonical properties window, else a
 * deterministic default window from the properties start date. A schedule
 * window with `finish === start` (a milestone-only project) is REAL
 * information and is preserved as a zero span — `fitViewport` then pads it
 * to the one-day minimum. Only a missing or inverted finish falls back. */
export function projectWindow(
  document: ProjectDocument,
  schedule?: DerivedSchedule,
): { start: number; finish: number } {
  const startMs =
    (schedule?.projectStart !== undefined
      ? parseInstant(schedule.projectStart)
      : parseInstant(document.properties.startDate)) ?? 0
  let finishMs: number | undefined
  if (schedule?.projectFinish !== undefined) finishMs = parseInstant(schedule.projectFinish)
  else if (document.properties.finishDate !== undefined) {
    finishMs = parseInstant(document.properties.finishDate)
  }
  if (finishMs === undefined || !Number.isFinite(finishMs) || finishMs < startMs) {
    finishMs = startMs + DEFAULT_FIT_SPAN_MS
  }
  return { start: startMs, finish: finishMs }
}

/**
 * Computes the fitted viewport for a project: the project window padded by
 * 2% on each side, with a minimum total span of one day. Deterministic.
 */
export function fitViewport(
  document: ProjectDocument,
  schedule?: DerivedSchedule,
): TimelineViewport {
  const window = projectWindow(document, schedule)
  const span = Math.max(window.finish - window.start, 0)
  const pad = Math.round(span * FIT_PADDING_RATIO)
  let start = window.start - pad
  let finish = window.finish + pad
  if (finish - start < FIT_MIN_TOTAL_SPAN_MS) {
    const midpoint = Math.round((start + finish) / 2)
    start = midpoint - FIT_MIN_TOTAL_SPAN_MS / 2
    finish = midpoint + FIT_MIN_TOTAL_SPAN_MS / 2
  }
  return { start: formatInstant(start), finish: formatInstant(finish) }
}

/** Time-axis band granularity, chosen deterministically from the visible
 * span: `day` below 93 days, `week` below 3 years, `month` beyond. */
export type TimeAxisLevel = 'day' | 'week' | 'month'

/** One axis band: `[start, finish)` in UTC with the band's level. Bands are
 * contiguous, ordered, and carry no formatted label — label formatting is a
 * host/locale concern (the renderer core is locale-free). */
export interface TimeAxisBand {
  readonly start: string
  readonly finish: string
  readonly level: TimeAxisLevel
}

/** Chooses the axis level for a span in milliseconds (deterministic). */
export function axisLevelForSpan(spanMs: number): TimeAxisLevel {
  if (spanMs < 93 * DAY_MS) return 'day'
  if (spanMs < 3 * 366 * DAY_MS) return 'week'
  return 'month'
}

const utcDayStart = (ms: number): number => Math.floor(ms / DAY_MS) * DAY_MS

/** Monday 00:00 UTC of the week containing `ms` (weeks start Monday). */
const utcWeekStart = (ms: number): number => {
  const day = utcDayStart(ms)
  const weekday = ((Math.floor(day / DAY_MS) % 7) + 7) % 7 // 0=Thursday epoch; normalize
  // Epoch day 0 (1970-01-01) was a Thursday (weekday 4 with Monday=0).
  const daysSinceMonday = (weekday + 3) % 7
  return day - daysSinceMonday * DAY_MS
}

/** First day of the UTC month containing `ms`. */
const utcMonthStart = (ms: number): number => {
  const date = new Date(ms)
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1)
}

/** The start of the band following `start` at the given level. */
function nextBandStart(start: number, level: TimeAxisLevel): number {
  if (level === 'day') return utcDayStart(start + DAY_MS)
  if (level === 'week') return utcWeekStart(start + 7 * DAY_MS)
  const date = new Date(start)
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 1)
}

/**
 * Generates the deterministic time-axis bands covering the viewport
 * (contiguous `[start, finish)` bands in ascending order; the first and last
 * bands are clipped to the window). Band boundaries are whole UTC days,
 * Monday-aligned weeks, or UTC calendar months — no locale, no timezone
 * conversion, no wall clock. The band count is bounded: a viewport within
 * the documented span guards yields at most a few hundred bands; the
 * projection contract stays linear in the band count.
 */
export function buildTimeAxis(
  viewport: TimelineViewport,
  level?: TimeAxisLevel,
): readonly TimeAxisBand[] {
  const startMs = parseInstant(viewport.start)
  const finishMs = parseInstant(viewport.finish)
  if (startMs === undefined || finishMs === undefined || finishMs <= startMs) return []
  const resolvedLevel = level ?? axisLevelForSpan(finishMs - startMs)
  const bands: TimeAxisBand[] = []
  let cursor: number
  if (resolvedLevel === 'day') cursor = utcDayStart(startMs)
  else if (resolvedLevel === 'week') cursor = utcWeekStart(startMs)
  else cursor = utcMonthStart(startMs)
  let guard = 0
  while (cursor < finishMs && guard < 2000) {
    guard += 1
    const nextStart = nextBandStart(cursor, resolvedLevel)
    const bandStart = Math.max(cursor, startMs)
    const bandFinish = Math.min(nextStart, finishMs)
    if (bandFinish > bandStart) {
      bands.push({
        start: formatInstant(bandStart),
        finish: formatInstant(bandFinish),
        level: resolvedLevel,
      })
    }
    cursor = nextStart
  }
  return bands
}
