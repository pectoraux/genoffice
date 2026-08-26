import type {
  Calendar,
  CalendarException,
  CalendarId,
  CalendarPeriod,
  ISODateTime,
  WorkingMinutes,
} from '@genoffice/project-contracts'

const DAY_MS = 86_400_000
const MINUTE_MS = 60_000
const SEARCH_GUARD = 100_000

// Brand-boundary conversions. All calendar arithmetic below runs on plain
// numbers and Date values and is converted back through these helpers, so no
// raw number is ever assigned to a WorkingMinutes/ISODateTime-typed location.
const toISODateTime = (date: Date): ISODateTime => date.toISOString() as ISODateTime
const toWorkingMinutes = (value: number): WorkingMinutes => value as WorkingMinutes
const toPlainMinutes = (value: WorkingMinutes): number => value as number

export class CalendarError extends Error {
  constructor(
    public readonly code:
      | 'MISSING_CALENDAR'
      | 'CALENDAR_CYCLE'
      | 'CALENDAR_PERIOD_MALFORMED'
      | 'CALENDAR_SEARCH_EXHAUSTED',
    message: string,
  ) {
    super(message)
  }
}

export interface CalendarBook {
  calendars: Calendar[]
}

function dayStart(date: Date): Date {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()))
}

function minuteOfDay(date: Date): number {
  return date.getUTCHours() * 60 + date.getUTCMinutes()
}

function dateKey(date: Date): string {
  return date.toISOString().slice(0, 10)
}

function sortedPeriods(periods: CalendarPeriod[]): CalendarPeriod[] {
  return [...periods].sort((a, b) => a.startMinute - b.startMinute || a.endMinute - b.endMinute)
}

function validatePeriods(calendarId: string, periods: CalendarPeriod[]): CalendarPeriod[] {
  const sorted = sortedPeriods(periods)
  let previousEnd = -1
  for (const period of sorted) {
    if (!Number.isInteger(period.startMinute) || !Number.isInteger(period.endMinute)) {
      throw new CalendarError(
        'CALENDAR_PERIOD_MALFORMED',
        `Calendar ${calendarId} has a period with non-integer minute bounds`,
      )
    }
    if (
      period.startMinute < 0 ||
      period.endMinute > 1440 ||
      period.startMinute >= period.endMinute
    ) {
      throw new CalendarError(
        'CALENDAR_PERIOD_MALFORMED',
        `Calendar ${calendarId} has a period outside 00:00-24:00 or with an empty interval`,
      )
    }
    if (period.startMinute < previousEnd) {
      throw new CalendarError(
        'CALENDAR_PERIOD_MALFORMED',
        `Calendar ${calendarId} has overlapping periods`,
      )
    }
    previousEnd = period.endMinute
  }
  return sorted
}

/**
 * Resolves a calendar against its inheritance chain. Weekday working periods
 * defined by a child calendar replace the base calendar's entries for that
 * weekday, and exception dates on the child override base exceptions for the
 * same date. Inheritance cycles and malformed periods are rejected.
 */
export function resolveCalendar(book: CalendarBook, calendarId: CalendarId): Calendar {
  const byId = new Map(book.calendars.map((calendar) => [calendar.id as string, calendar]))
  const chain: Calendar[] = []
  const seen = new Set<string>()
  let currentId: string | null = calendarId as string
  while (currentId !== null) {
    if (seen.has(currentId)) {
      throw new CalendarError('CALENDAR_CYCLE', `Calendar inheritance cycle involves ${currentId}`)
    }
    const calendar = byId.get(currentId)
    if (!calendar) {
      throw new CalendarError('MISSING_CALENDAR', `Unknown calendar ${currentId}`)
    }
    seen.add(currentId)
    chain.push(calendar)
    currentId = calendar.baseCalendarId ? (calendar.baseCalendarId as string) : null
  }

  // chain is ordered child-first; merge from the base upwards so child entries win.
  const workingWeek: Record<number, CalendarPeriod[]> = {}
  const exceptions = new Map<string, CalendarException>()
  for (const calendar of chain.slice().reverse()) {
    for (const key of Object.keys(calendar.workingWeek)) {
      workingWeek[Number(key)] = calendar.workingWeek[Number(key)]
    }
    for (const exception of calendar.exceptions) {
      exceptions.set(exception.date, exception)
    }
  }

  const child = chain[0]
  const resolved: Calendar = {
    id: child.id,
    name: child.name,
    workingWeek: {},
    exceptions: [...exceptions.keys()].sort().map((date) => {
      const exception = exceptions.get(date)!
      return { date, periods: validatePeriods(child.id as string, exception.periods) }
    }),
  }
  for (const key of Object.keys(workingWeek).sort((a, b) => Number(a) - Number(b))) {
    resolved.workingWeek[Number(key)] = validatePeriods(
      child.id as string,
      workingWeek[Number(key)],
    )
  }
  return resolved
}

function periodsFor(calendar: Calendar, date: Date): CalendarPeriod[] {
  const exception = calendar.exceptions.find((item) => item.date === dateKey(date))
  if (exception) return sortedPeriods(exception.periods)
  const periods = calendar.workingWeek[date.getUTCDay()]
  return periods ? sortedPeriods(periods) : []
}

function isWithinPeriods(minute: number, periods: CalendarPeriod[]): boolean {
  return periods.some((period) => minute >= period.startMinute && minute < period.endMinute)
}

/** True when the instant falls inside a working period (period start inclusive, end exclusive). */
export function isWorking(calendar: Calendar, timestamp: ISODateTime): boolean {
  const value = new Date(timestamp)
  return isWithinPeriods(minuteOfDay(value), periodsFor(calendar, value))
}

/**
 * Earliest working instant at or after `timestamp`. When `timestamp` is already
 * working time it is returned unchanged; otherwise the start of the next
 * working period is returned.
 */
export function nextWorkingInstant(calendar: Calendar, timestamp: ISODateTime): ISODateTime {
  let cursor = new Date(timestamp)
  for (let guard = 0; guard < SEARCH_GUARD; guard += 1) {
    const day = dayStart(cursor)
    const minute = minuteOfDay(cursor)
    const periods = periodsFor(calendar, day)
    if (isWithinPeriods(minute, periods)) return toISODateTime(cursor)
    const upcoming = periods.find((period) => period.startMinute >= minute)
    if (upcoming) return toISODateTime(new Date(day.getTime() + upcoming.startMinute * MINUTE_MS))
    cursor = new Date(day.getTime() + DAY_MS)
  }
  throw new CalendarError(
    'CALENDAR_SEARCH_EXHAUSTED',
    'No working instant exists after the given timestamp',
  )
}

/**
 * Latest working instant at or before `timestamp`. When `timestamp` is already
 * working time it is returned unchanged; otherwise the end of the previous
 * working period is returned.
 */
export function prevWorkingInstant(calendar: Calendar, timestamp: ISODateTime): ISODateTime {
  let cursor = new Date(timestamp)
  for (let guard = 0; guard < SEARCH_GUARD; guard += 1) {
    const day = dayStart(cursor)
    const minute = minuteOfDay(cursor)
    const periods = periodsFor(calendar, day)
    if (isWithinPeriods(minute, periods)) return toISODateTime(cursor)
    const earlier = [...periods].reverse().find((period) => period.endMinute <= minute)
    if (earlier) return toISODateTime(new Date(day.getTime() + earlier.endMinute * MINUTE_MS))
    cursor = new Date(day.getTime() - 1)
  }
  throw new CalendarError(
    'CALENDAR_SEARCH_EXHAUSTED',
    'No working instant exists before the given timestamp',
  )
}

/**
 * Advances a working-time duration from `startIso`. The cursor first normalizes
 * forward to a working instant (so a zero duration returns the next working
 * instant, e.g. 17:00 becomes 09:00 the next working day) and then consumes
 * working minutes period by period.
 */
export function addWorkingTime(
  calendar: Calendar,
  startIso: ISODateTime,
  duration: WorkingMinutes,
): ISODateTime {
  const total = toPlainMinutes(duration)
  if (total < 0) return subtractWorkingTime(calendar, startIso, toWorkingMinutes(-total))
  let cursor = new Date(nextWorkingInstant(calendar, startIso))
  let remaining = total
  for (let guard = 0; guard < SEARCH_GUARD && remaining > 0; guard += 1) {
    const day = dayStart(cursor)
    const minute = minuteOfDay(cursor)
    const periods = periodsFor(calendar, day)
    const period = periods.find((candidate) => minute < candidate.endMinute)
    if (!period) {
      cursor = new Date(day.getTime() + DAY_MS)
      continue
    }
    const from = Math.max(minute, period.startMinute)
    const available = period.endMinute - from
    const consumed = Math.min(remaining, available)
    remaining -= consumed
    cursor = new Date(day.getTime() + (from + consumed) * MINUTE_MS)
  }
  if (remaining > 0) {
    throw new CalendarError(
      'CALENDAR_SEARCH_EXHAUSTED',
      'Not enough working time to satisfy the duration',
    )
  }
  return toISODateTime(cursor)
}

/**
 * Moves a working-time duration backwards from `finishIso`. The cursor first
 * normalizes backward to a working instant and then consumes working minutes
 * period by period in reverse.
 */
export function subtractWorkingTime(
  calendar: Calendar,
  finishIso: ISODateTime,
  duration: WorkingMinutes,
): ISODateTime {
  const total = toPlainMinutes(duration)
  if (total < 0) return addWorkingTime(calendar, finishIso, toWorkingMinutes(-total))
  let cursor = new Date(prevWorkingInstant(calendar, finishIso))
  let remaining = total
  for (let guard = 0; guard < SEARCH_GUARD && remaining > 0; guard += 1) {
    const day = dayStart(cursor)
    const minute = minuteOfDay(cursor)
    const periods = periodsFor(calendar, day)
    const period = [...periods].reverse().find((candidate) => minute > candidate.startMinute)
    if (!period) {
      cursor = new Date(day.getTime() - 1)
      continue
    }
    const from = Math.min(minute, period.endMinute)
    const available = from - period.startMinute
    const consumed = Math.min(remaining, available)
    remaining -= consumed
    cursor = new Date(day.getTime() + (from - consumed) * MINUTE_MS)
  }
  if (remaining > 0) {
    throw new CalendarError(
      'CALENDAR_SEARCH_EXHAUSTED',
      'Not enough working time to satisfy the duration',
    )
  }
  return toISODateTime(cursor)
}

/**
 * Working sub-intervals of the half-open span `[startIso, finishIso)`, each
 * clipped to the span. Returns one interval per (calendar period ∩ span) that
 * has positive length, in chronological order. Empty when the span contains no
 * working time. Used by the leveler to intersect assignment demand with the
 * resource's actual working periods so over-allocation is evaluated only where
 * the resource can supply work capacity (PROJECT-013).
 */
export function workingIntervals(
  calendar: Calendar,
  startIso: ISODateTime,
  finishIso: ISODateTime,
): { start: ISODateTime; finish: ISODateTime }[] {
  const from = new Date(startIso)
  const to = new Date(finishIso)
  if (to.getTime() <= from.getTime()) return []
  const firstDay = dayStart(from)
  const lastDay = dayStart(to)
  const firstMinute = minuteOfDay(from)
  const lastMinute = minuteOfDay(to)
  const intervals: { start: ISODateTime; finish: ISODateTime }[] = []
  let day = firstDay
  for (let guard = 0; guard < SEARCH_GUARD && day.getTime() <= lastDay.getTime(); guard += 1) {
    const periods = periodsFor(calendar, day)
    const startMinute = day.getTime() === firstDay.getTime() ? firstMinute : 0
    const endMinute = day.getTime() === lastDay.getTime() ? lastMinute : 1440
    for (const period of periods) {
      const a = Math.max(startMinute, period.startMinute)
      const b = Math.min(period.endMinute, endMinute)
      if (b > a) {
        intervals.push({
          start: toISODateTime(new Date(day.getTime() + a * MINUTE_MS)),
          finish: toISODateTime(new Date(day.getTime() + b * MINUTE_MS)),
        })
      }
    }
    day = new Date(day.getTime() + DAY_MS)
  }
  return intervals
}

/** Working minutes inside the half-open interval [startIso, finishIso). */
export function workingDuration(
  calendar: Calendar,
  startIso: ISODateTime,
  finishIso: ISODateTime,
): WorkingMinutes {
  const from = new Date(startIso)
  const to = new Date(finishIso)
  if (to.getTime() <= from.getTime()) return toWorkingMinutes(0)
  const firstDay = dayStart(from)
  const lastDay = dayStart(to)
  const firstMinute = minuteOfDay(from)
  const lastMinute = minuteOfDay(to)
  let total = 0
  let day = firstDay
  for (let guard = 0; guard < SEARCH_GUARD && day.getTime() <= lastDay.getTime(); guard += 1) {
    const periods = periodsFor(calendar, day)
    const startMinute = day.getTime() === firstDay.getTime() ? firstMinute : 0
    const endMinute = day.getTime() === lastDay.getTime() ? lastMinute : 1440
    for (const period of periods) {
      const a = Math.max(startMinute, period.startMinute)
      const b = Math.min(period.endMinute, endMinute)
      if (b > a) total += b - a
    }
    day = new Date(day.getTime() + DAY_MS)
  }
  return toWorkingMinutes(total)
}

/**
 * Signed working minutes from `startIso` to `finishIso`; negative when finish
 * precedes start. Used for slack computations where constraints can invert the
 * early/late envelope.
 */
export function signedWorkingDuration(
  calendar: Calendar,
  startIso: ISODateTime,
  finishIso: ISODateTime,
): number {
  if (new Date(finishIso).getTime() >= new Date(startIso).getTime()) {
    return toPlainMinutes(workingDuration(calendar, startIso, finishIso))
  }
  return -toPlainMinutes(workingDuration(calendar, finishIso, startIso))
}

/** Exposes the brand conversion so downstream engines share one boundary. */
export const workingMinutesOf = toWorkingMinutes
