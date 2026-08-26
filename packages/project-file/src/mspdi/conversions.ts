/**
 * PROJECT-015 — Deterministic MSPDI duration / lag / date conversions.
 *
 * MSPDI carries durations as ISO-8601 duration strings (`PT8H0M0S`),
 * lags as integer `<LinkLag>` + `<LinkLagFormat>` pairs, and dates as
 * ISO-8601 date-times that may be naive (no offset), Z-terminated, or carry
 * an explicit `±HH:MM` offset. The canonical model uses integer
 * `WorkingMinutes` for durations/lag and `ISODateTime` (UTC, `.000Z`) for
 * dates. This module is the single deterministic conversion layer.
 *
 * Discipline (per PROJECT-015 brief):
 *   - No host locale, no system timezone as a semantic input. Naive MSPDI
 *     date-times are interpreted as UTC (the only host-independent choice).
 *   - No silent rounding of ambiguous units. Sub-minute remainders, elapsed
 *     durations, and percentage lags are flagged as `invalid` / `unsupported`
 *     so the importer can emit `INVALID_MSPDI_DURATION` /
 *     `UNSUPPORTED_MSPDI_FEATURE` rather than silently approximating.
 *   - Working durations in MSPDI are encoded as ISO-8601 time-parts
 *     (`PT#H#M#S`) where the hour/minute magnitudes are working-time
 *     magnitudes (a 1-working-day task is `PT8H0M0S`, not `P1D`). Date-part
 *     components (`P1D`, `P1W`, `P1M`, `P1Y`) represent elapsed/calendar time
 *     and are `unsupported` (no faithful working-minute conversion without
 *     the project calendar, which the adapter does not consult at import
 *     time — the canonical calendar is applied by the scheduling engine
 *     later).
 *
 * Documented lag convention: MSPDI `<LinkLag>` is stored in **tenths of a
 * minute** (the documented MS Project storage unit). `lagMinutes =
 * LinkLag / 10`. Non-multiple-of-10 lags are sub-minute → `invalid`.
 * Elapsed `LinkLagFormat` values (2/4/6/8/10) and percentage lags are
 * `unsupported` (elapsed/percentage lag has no faithful working-minute
 * representation at import time).
 */
import { asISODateTime, asWorkingMinutes } from '@genoffice/project-contracts'
import type { ISODateTime, WorkingMinutes } from '@genoffice/project-contracts'

export type DurationResult =
  { ok: true; minutes: WorkingMinutes } | { ok: false; reason: 'invalid' | 'unsupported' }

// TimeUnitType enum values (MSPDI `<LinkLagFormat>` / `<DurationFormat>`).
// Working-time units → tenths-of-a-minute storage (odd values below).
// Elapsed variants (even) and percentage (35) are unsupported for lag.
const LAG_FORMAT_MINUTE = 1
const LAG_FORMAT_HOUR = 3
const LAG_FORMAT_DAY = 5
const LAG_FORMAT_WEEK = 7
const LAG_FORMAT_MONTH = 9
const WORKING_LAG_FORMATS = new Set<number>([
  LAG_FORMAT_MINUTE,
  LAG_FORMAT_HOUR,
  LAG_FORMAT_DAY,
  LAG_FORMAT_WEEK,
  LAG_FORMAT_MONTH,
])
const ELAPSED_LAG_FORMATS = new Set<number>([2, 4, 6, 8, 10])
const PERCENT_LAG_FORMAT = 35

const ISO_DURATION_RE =
  /^(-)?P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/

/**
 * Convert an MSPDI ISO-8601 duration string to integer `WorkingMinutes`.
 *
 * - `PT8H0M0S` → 480 (8 working hours).
 * - `PT40H0M0S` → 2400 (5 working days).
 * - `PT0S` / `PT0H0M0S` → 0.
 * - `P1D` / `P1W` / `P1M` / `P1Y` → `unsupported` (date-part components
 *   represent elapsed/calendar time; no faithful working-minute conversion).
 * - Sub-minute seconds (`PT0H0M30S`) → `invalid` (canonical `WorkingMinutes`
 *   is integer; sub-minute remainder is not representable).
 * - Malformed / empty → `invalid`.
 */
export function isoDurationToMinutes(input: string): DurationResult {
  const s = input.trim()
  if (s.length === 0) return { ok: false, reason: 'invalid' }
  const m = ISO_DURATION_RE.exec(s)
  if (m === null) return { ok: false, reason: 'invalid' }
  const negative = m[1] === '-'
  const years = m[2] ? Number(m[2]) : 0
  const months = m[3] ? Number(m[3]) : 0
  const weeks = m[4] ? Number(m[4]) : 0
  const days = m[5] ? Number(m[5]) : 0
  const hours = m[6] ? Number(m[6]) : 0
  const minutes = m[7] ? Number(m[7]) : 0
  const seconds = m[8] ? Number(m[8]) : 0
  if (
    years === 0 &&
    months === 0 &&
    weeks === 0 &&
    days === 0 &&
    hours === 0 &&
    minutes === 0 &&
    seconds === 0
  ) {
    // Empty `P` or `PT` — treat as zero duration.
    return { ok: true, minutes: asWorkingMinutes(0) }
  }
  // Any date-part component means elapsed/calendar time → unsupported.
  if (years !== 0 || months !== 0 || weeks !== 0 || days !== 0) {
    return { ok: false, reason: 'unsupported' }
  }
  if (negative) {
    // Negative working duration is not representable.
    return { ok: false, reason: 'invalid' }
  }
  const totalSeconds = hours * 3600 + minutes * 60 + seconds
  if (totalSeconds % 60 !== 0) {
    return { ok: false, reason: 'invalid' }
  }
  return { ok: true, minutes: asWorkingMinutes(totalSeconds / 60) }
}

/**
 * Convert an MSPDI `<LinkLag>` + `<LinkLagFormat>` pair to integer lag minutes.
 *
 * `LinkLag` is in tenths of a minute. `lagMinutes = LinkLag / 10`.
 *   - Non-multiple-of-10 `LinkLag` → `invalid` (sub-minute lag).
 *   - Elapsed `LinkLagFormat` (2/4/6/8/10) → `unsupported` (elapsed lag has no
 *     faithful working-minute representation).
 *   - Percentage `LinkLagFormat` (35) → `unsupported` (percentage lag requires
 *     the predecessor's working duration, which the adapter does not compute
 *     at import time).
 *   - Missing `LinkLagFormat` is treated as minutes (the most common case and
 *     the storage unit).
 */
export function lagToMinutes(linkLag: number, linkLagFormat: number | undefined): DurationResult {
  if (!Number.isInteger(linkLag)) return { ok: false, reason: 'invalid' }
  const fmt = linkLagFormat ?? LAG_FORMAT_MINUTE
  if (fmt === PERCENT_LAG_FORMAT) return { ok: false, reason: 'unsupported' }
  if (ELAPSED_LAG_FORMATS.has(fmt)) return { ok: false, reason: 'unsupported' }
  if (!WORKING_LAG_FORMATS.has(fmt)) return { ok: false, reason: 'invalid' }
  if (linkLag % 10 !== 0) return { ok: false, reason: 'invalid' }
  return { ok: true, minutes: asWorkingMinutes(linkLag / 10) }
}

const DATE_RE =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.\d+)?(Z|([+-])(\d{2}):(\d{2}))?$/

/** Format a UTC epoch-millisecond timestamp as `YYYY-MM-DDTHH:MM:SS.000Z`
 * (the canonical `ISODateTime` form used across `.gproj` fixtures). */
function formatUtc(ms: number): ISODateTime {
  const d = new Date(ms)
  const pad = (n: number, w = 2) => String(n).padStart(w, '0')
  const y = d.getUTCFullYear()
  const mo = d.getUTCMonth() + 1
  const da = d.getUTCDate()
  const h = d.getUTCHours()
  const mi = d.getUTCMinutes()
  const s = d.getUTCSeconds()
  return asISODateTime(`${pad(y, 4)}-${pad(mo)}-${pad(da)}T${pad(h)}:${pad(mi)}:${pad(s)}.000Z`)
}

/**
 * Normalize an MSPDI date-time to canonical UTC `ISODateTime`.
 *
 * - Naive (no offset) → interpreted as **UTC** (host-independent; never the
 *   system timezone).
 * - `Z` → already UTC.
 * - `±HH:MM` → converted to UTC.
 * - Output is always `YYYY-MM-DDTHH:MM:SS.000Z`.
 * - Malformed → `{ invalid: true }`.
 */
export function normalizeMspdiDate(raw: string): { iso: ISODateTime } | { invalid: true } {
  const s = raw.trim()
  const m = DATE_RE.exec(s)
  if (m === null) return { invalid: true }
  const year = Number(m[1])
  const month = Number(m[2])
  const day = Number(m[3])
  const hour = Number(m[4])
  const minute = Number(m[5])
  const second = Number(m[6])
  // Range-check components (Date.UTC would roll over silently otherwise).
  if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 23 || minute > 59 || second > 59) {
    return { invalid: true }
  }
  let epochMs: number
  if (m[8] === undefined || m[8] === '') {
    // No offset, no Z → naive → treat as UTC.
    epochMs = Date.UTC(year, month - 1, day, hour, minute, second)
  } else if (m[8] === 'Z') {
    epochMs = Date.UTC(year, month - 1, day, hour, minute, second)
  } else {
    const sign = m[9] === '-' ? -1 : 1
    const offH = Number(m[10])
    const offM = Number(m[11])
    if (offH > 23 || offM > 59) return { invalid: true }
    const localMs = Date.UTC(year, month - 1, day, hour, minute, second)
    epochMs = localMs - sign * (offH * 60 + offM) * 60 * 1000
  }
  if (!Number.isFinite(epochMs)) return { invalid: true }
  // Re-validate the rolled-over date (e.g. 2026-02-30 rolled into March).
  const check = new Date(epochMs)
  if (
    check.getUTCFullYear() !== year ||
    check.getUTCMonth() + 1 !== month ||
    check.getUTCDate() !== day
  ) {
    return { invalid: true }
  }
  return { iso: formatUtc(epochMs) }
}

/** Validate an MSPDI calendar-exception date (date-only `YYYY-MM-DD`). */
export function isValidExceptionDate(raw: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return false
  const [y, mo, d] = raw.split('-').map(Number)
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return false
  const ms = Date.UTC(y, mo - 1, d)
  const check = new Date(ms)
  return check.getUTCFullYear() === y && check.getUTCMonth() + 1 === mo && check.getUTCDate() === d
}

/** MSPDI time-of-day string (`HH:MM:SS`) → minute offset from midnight. */
export function mspdiTimeToMinutes(raw: string): number | null {
  const m = /^(\d{2}):(\d{2}):(\d{2})$/.exec(raw.trim())
  if (m === null) return null
  const h = Number(m[1])
  const mi = Number(m[2])
  const s = Number(m[3])
  if (h > 23 || mi > 59 || s > 59) return null
  return h * 60 + mi + s / 60
}
