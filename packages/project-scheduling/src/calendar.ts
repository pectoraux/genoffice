import type { Calendar, CalendarId, ISODateTime, WorkingMinutes } from '@genoffice/project-contracts'

const DAY_MS = 86_400_000

function dateKey(iso: string): string {
  return iso.slice(0, 10)
}

function dayStart(iso: string): Date {
  const value = new Date(iso)
  return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()))
}

function periodsFor(calendar: Calendar, date: Date): Calendar['workingWeek'][number] {
  const exception = calendar.exceptions.find((item) => item.date === date.toISOString().slice(0, 10))
  return exception ? exception.periods : (calendar.workingWeek[date.getUTCDay()] ?? [])
}

export interface CalendarBook { calendars: Calendar[] }

export function resolveCalendar(book: CalendarBook, calendarId: CalendarId, seen = new Set<string>()): Calendar {
  if (seen.has(calendarId)) throw new Error(`CALENDAR_CYCLE:${calendarId}`)
  const calendar = book.calendars.find((candidate) => candidate.id === calendarId)
  if (!calendar) throw new Error(`MISSING_CALENDAR:${calendarId}`)
  if (!calendar.baseCalendarId) return calendar
  const base = resolveCalendar(book, calendar.baseCalendarId, new Set([...seen, calendarId]))
  const merged: Calendar = {
    ...base,
    ...calendar,
    workingWeek: { ...base.workingWeek, ...calendar.workingWeek },
    exceptions: [...base.exceptions, ...calendar.exceptions],
  }
  return merged
}

export function isWorking(calendar: Calendar, iso: ISODateTime): boolean {
  const value = new Date(iso)
  const periods = periodsFor(calendar, value)
  const minute = value.getUTCHours() * 60 + value.getUTCMinutes() + value.getUTCSeconds() / 60
  return periods.some((period) => minute >= period.startMinute && minute < period.endMinute)
}

export function addWorkingTime(calendar: Calendar, startIso: ISODateTime, duration: WorkingMinutes): ISODateTime {
  if (duration < 0) return subtractWorkingTime(calendar, startIso, (-duration) as WorkingMinutes)
  let cursor = new Date(startIso)
  let remaining = duration
  for (let guard = 0; guard < 100_000 && remaining > 0; guard += 1) {
    const day = dayStart(cursor)
    const periods = periodsFor(calendar, day)
    const nowMinute = cursor.getUTCHours() * 60 + cursor.getUTCMinutes() + cursor.getUTCSeconds() / 60
    const period = periods.find((candidate) => nowMinute < candidate.endMinute)
    if (!period) { cursor = new Date(day.getTime() + DAY_MS); continue }
    const currentMinute = Math.max(nowMinute, period.startMinute)
    if (currentMinute >= period.endMinute) { cursor = new Date(day.getTime() + DAY_MS); continue }
    cursor = new Date(day.getTime() + currentMinute * 60_000)
    const available = period.endMinute - currentMinute
    const consumed = Math.min(remaining, available)
    remaining -= consumed
    cursor = new Date(cursor.getTime() + consumed * 60_000)
    if (remaining > 0) cursor = new Date(day.getTime() + DAY_MS)
  }
  if (remaining > 0) throw new Error('CALENDAR_SEARCH_EXHAUSTED')
  return cursor.toISOString() as ISODateTime
}

export function subtractWorkingTime(calendar: Calendar, finishIso: ISODateTime, duration: WorkingMinutes): ISODateTime {
  if (duration < 0) return addWorkingTime(calendar, finishIso, (-duration) as WorkingMinutes)
  let cursor = new Date(finishIso)
  let remaining = duration
  for (let guard = 0; guard < 100_000 && remaining > 0; guard += 1) {
    const day = dayStart(cursor)
    const periods = periodsFor(calendar, day).slice().sort((a, b) => b.endMinute - a.endMinute)
    const nowMinute = cursor.getUTCHours() * 60 + cursor.getUTCMinutes() + cursor.getUTCSeconds() / 60
    const period = periods.find((candidate) => nowMinute > candidate.startMinute)
    if (!period) { cursor = new Date(day.getTime() - 1); continue }
    const currentMinute = Math.min(nowMinute, period.endMinute)
    if (currentMinute <= period.startMinute) { cursor = new Date(day.getTime() - 1); continue }
    const available = currentMinute - period.startMinute
    const consumed = Math.min(remaining, available)
    remaining -= consumed
    cursor = new Date(day.getTime() + (currentMinute - consumed) * 60_000)
    if (remaining > 0) cursor = new Date(day.getTime() - 1)
  }
  if (remaining > 0) throw new Error('CALENDAR_SEARCH_EXHAUSTED')
  return cursor.toISOString() as ISODateTime
}

export function workingDuration(calendar: Calendar, startIso: ISODateTime, finishIso: ISODateTime): WorkingMinutes {
  if (finishIso <= startIso) return 0 as WorkingMinutes
  let cursor = new Date(startIso)
  const finish = new Date(finishIso)
  let total = 0
  for (let guard = 0; guard < 100_000 && cursor < finish; guard += 1) {
    const day = dayStart(cursor)
    const periods = periodsFor(calendar, day)
    const startMinute = cursor.getUTCDate() === day.getUTCDate() ? cursor.getUTCHours() * 60 + cursor.getUTCMinutes() + cursor.getUTCSeconds() / 60 : 0
    for (const period of periods) {
      const a = Math.max(startMinute, period.startMinute)
      const b = Math.min(period.endMinute, (finish.getTime() - day.getTime()) / 60_000)
      if (b > a) total += b - a
    }
    cursor = new Date(day.getTime() + DAY_MS)
  }
  return total as WorkingMinutes
}

export function calendarDateKey(iso: ISODateTime): string { return dateKey(iso) }
