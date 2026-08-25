import { describe, expect, it } from 'vitest'
import {
  CalendarError,
  addWorkingTime,
  isWorking,
  nextWorkingInstant,
  prevWorkingInstant,
  resolveCalendar,
  subtractWorkingTime,
  workingDuration,
} from '../src/index.js'
import { iso, makeCalendar, holiday, standardWeek, wm } from './fixtures.js'
import type { Calendar, CalendarPeriod } from '@genoffice/project-contracts'
import { asCalendarId } from '@genoffice/project-contracts'

const splitDay = (): CalendarPeriod[] => [
  { startMinute: 540, endMinute: 720 },
  { startMinute: 780, endMinute: 1020 },
]

const splitCalendar = (): Calendar => ({
  ...makeCalendar('split'),
  workingWeek: {
    0: [],
    1: splitDay(),
    2: splitDay(),
    3: splitDay(),
    4: splitDay(),
    5: splitDay(),
    6: [],
  },
})

describe('calendar working-time primitives', () => {
  it('detects working instants with period-start inclusion and period-end exclusion', () => {
    const calendar = makeCalendar('standard')
    expect(isWorking(calendar, iso('2026-08-03T09:00:00.000Z'))).toBe(true)
    expect(isWorking(calendar, iso('2026-08-03T12:30:00.000Z'))).toBe(true)
    expect(isWorking(calendar, iso('2026-08-03T17:00:00.000Z'))).toBe(false)
    expect(isWorking(calendar, iso('2026-08-03T08:59:00.000Z'))).toBe(false)
    expect(isWorking(calendar, iso('2026-08-01T10:00:00.000Z'))).toBe(false) // Saturday
  })

  it('honors exception dates over the weekly pattern', () => {
    const calendar = makeCalendar('standard', { exceptions: [holiday('2026-08-04')] })
    expect(isWorking(calendar, iso('2026-08-04T10:00:00.000Z'))).toBe(false)
    expect(isWorking(calendar, iso('2026-08-05T10:00:00.000Z'))).toBe(true)
    const extra = makeCalendar('standard', {
      exceptions: [{ date: '2026-08-06', periods: [{ startMinute: 540, endMinute: 660 }] }],
    })
    expect(isWorking(extra, iso('2026-08-06T10:00:00.000Z'))).toBe(true)
    expect(isWorking(extra, iso('2026-08-06T11:30:00.000Z'))).toBe(false)
  })

  it('normalizes a zero duration to the next working instant', () => {
    const calendar = makeCalendar('standard')
    expect(addWorkingTime(calendar, iso('2026-08-03T17:00:00.000Z'), wm(0))).toBe(
      '2026-08-04T09:00:00.000Z',
    )
    expect(addWorkingTime(calendar, iso('2026-08-03T09:00:00.000Z'), wm(0))).toBe(
      '2026-08-03T09:00:00.000Z',
    )
    expect(addWorkingTime(calendar, iso('2026-08-07T17:00:00.000Z'), wm(0))).toBe(
      '2026-08-10T09:00:00.000Z',
    )
  })

  it('adds working time across days, weekends, and split periods', () => {
    const standard = makeCalendar('standard')
    expect(addWorkingTime(standard, iso('2026-08-03T09:00:00.000Z'), wm(480))).toBe(
      '2026-08-03T17:00:00.000Z',
    )
    expect(addWorkingTime(standard, iso('2026-08-03T09:00:00.000Z'), wm(960))).toBe(
      '2026-08-04T17:00:00.000Z',
    )
    expect(addWorkingTime(standard, iso('2026-08-03T12:00:00.000Z'), wm(480))).toBe(
      '2026-08-04T12:00:00.000Z',
    )
    expect(addWorkingTime(standard, iso('2026-08-06T09:00:00.000Z'), wm(1440))).toBe(
      '2026-08-10T17:00:00.000Z',
    )

    const split = splitCalendar()
    // 420 working minutes per day: 09:00 Monday plus 480 finishes Tuesday 10:00.
    expect(addWorkingTime(split, iso('2026-08-03T09:00:00.000Z'), wm(480))).toBe(
      '2026-08-04T10:00:00.000Z',
    )
    // Starting at 12:00 (end of the morning period) normalizes to 13:00.
    expect(addWorkingTime(split, iso('2026-08-03T12:00:00.000Z'), wm(60))).toBe(
      '2026-08-03T14:00:00.000Z',
    )
  })

  it('subtracts working time symmetrically', () => {
    const standard = makeCalendar('standard')
    expect(subtractWorkingTime(standard, iso('2026-08-04T09:00:00.000Z'), wm(480))).toBe(
      '2026-08-03T09:00:00.000Z',
    )
    expect(subtractWorkingTime(standard, iso('2026-08-04T17:00:00.000Z'), wm(480))).toBe(
      '2026-08-04T09:00:00.000Z',
    )
    expect(subtractWorkingTime(standard, iso('2026-08-04T09:00:00.000Z'), wm(0))).toBe(
      '2026-08-04T09:00:00.000Z',
    )
    expect(subtractWorkingTime(standard, iso('2026-08-10T09:00:00.000Z'), wm(480))).toBe(
      '2026-08-07T09:00:00.000Z',
    )
    const split = splitCalendar()
    expect(subtractWorkingTime(split, iso('2026-08-04T10:00:00.000Z'), wm(480))).toBe(
      '2026-08-03T09:00:00.000Z',
    )
  })

  it('treats negative durations as the inverse operation', () => {
    const standard = makeCalendar('standard')
    expect(addWorkingTime(standard, iso('2026-08-03T17:00:00.000Z'), wm(-240))).toBe(
      '2026-08-03T13:00:00.000Z',
    )
    expect(subtractWorkingTime(standard, iso('2026-08-04T09:00:00.000Z'), wm(-240))).toBe(
      '2026-08-04T13:00:00.000Z',
    )
  })

  it('measures working duration inside a window', () => {
    const standard = makeCalendar('standard')
    expect(
      workingDuration(standard, iso('2026-08-03T09:00:00.000Z'), iso('2026-08-03T17:00:00.000Z')),
    ).toBe(480)
    expect(
      workingDuration(standard, iso('2026-08-03T09:00:00.000Z'), iso('2026-08-04T17:00:00.000Z')),
    ).toBe(960)
    expect(
      workingDuration(standard, iso('2026-08-03T12:00:00.000Z'), iso('2026-08-04T12:00:00.000Z')),
    ).toBe(480)
    // Friday 09:00 to Monday 09:00 spans a weekend: only Friday counts.
    expect(
      workingDuration(standard, iso('2026-08-07T09:00:00.000Z'), iso('2026-08-10T09:00:00.000Z')),
    ).toBe(480)
    expect(
      workingDuration(standard, iso('2026-08-03T17:00:00.000Z'), iso('2026-08-04T09:00:00.000Z')),
    ).toBe(0)
    expect(
      workingDuration(standard, iso('2026-08-04T09:00:00.000Z'), iso('2026-08-03T09:00:00.000Z')),
    ).toBe(0)
    const holidayCalendar = makeCalendar('standard', { exceptions: [holiday('2026-08-04')] })
    expect(
      workingDuration(
        holidayCalendar,
        iso('2026-08-03T09:00:00.000Z'),
        iso('2026-08-05T09:00:00.000Z'),
      ),
    ).toBe(480)
  })

  it('computes next and previous working instants', () => {
    const standard = makeCalendar('standard')
    expect(nextWorkingInstant(standard, iso('2026-08-03T17:00:00.000Z'))).toBe(
      '2026-08-04T09:00:00.000Z',
    )
    expect(nextWorkingInstant(standard, iso('2026-08-03T10:00:00.000Z'))).toBe(
      '2026-08-03T10:00:00.000Z',
    )
    expect(nextWorkingInstant(standard, iso('2026-08-08T10:00:00.000Z'))).toBe(
      '2026-08-10T09:00:00.000Z',
    )
    expect(prevWorkingInstant(standard, iso('2026-08-03T09:00:00.000Z'))).toBe(
      '2026-08-03T09:00:00.000Z',
    )
    expect(prevWorkingInstant(standard, iso('2026-08-03T07:00:00.000Z'))).toBe(
      '2026-07-31T17:00:00.000Z',
    )
    const split = splitCalendar()
    expect(nextWorkingInstant(split, iso('2026-08-03T12:00:00.000Z'))).toBe(
      '2026-08-03T13:00:00.000Z',
    )
    expect(prevWorkingInstant(split, iso('2026-08-03T12:30:00.000Z'))).toBe(
      '2026-08-03T12:00:00.000Z',
    )
  })
})

describe('calendar inheritance', () => {
  it('lets a child override weekday periods and exception dates', () => {
    const base = makeCalendar('base', {
      exceptions: [holiday('2026-08-05')],
    })
    const child = makeCalendar('child', {
      baseCalendarId: asCalendarId('base'),
      workingWeek: { ...standardWeek(), 3: [] },
      exceptions: [{ date: '2026-08-05', periods: [{ startMinute: 540, endMinute: 1020 }] }],
    })
    const resolved = resolveCalendar({ calendars: [base, child] }, asCalendarId('child'))
    // Wednesday is off in the child override.
    expect(resolved.workingWeek[3]).toEqual([])
    // The child exception overrides the base holiday.
    expect(isWorking(resolved, iso('2026-08-05T10:00:00.000Z'))).toBe(true)
    // Monday still inherits the base week.
    expect(isWorking(resolved, iso('2026-08-03T10:00:00.000Z'))).toBe(true)
  })

  it('keeps base exceptions when the child does not override the date', () => {
    const base = makeCalendar('base', { exceptions: [holiday('2026-08-04')] })
    const child = makeCalendar('child', { baseCalendarId: asCalendarId('base') })
    const resolved = resolveCalendar({ calendars: [base, child] }, asCalendarId('child'))
    expect(resolved.exceptions).toEqual([{ date: '2026-08-04', periods: [] }])
    expect(isWorking(resolved, iso('2026-08-04T10:00:00.000Z'))).toBe(false)
  })

  it('resolves multi-level inheritance chains child-first', () => {
    const base = makeCalendar('base', { workingWeek: { ...standardWeek(), 5: [] } })
    const middle = makeCalendar('middle', {
      baseCalendarId: asCalendarId('base'),
      workingWeek: { ...standardWeek(), 4: [] },
    })
    const child = makeCalendar('child', {
      baseCalendarId: asCalendarId('middle'),
      workingWeek: { 3: [] },
    })
    const resolved = resolveCalendar({ calendars: [base, middle, child] }, asCalendarId('child'))
    expect(resolved.workingWeek[3]).toEqual([])
    expect(resolved.workingWeek[4]).toEqual([])
    expect(resolved.workingWeek[5]).toEqual([{ startMinute: 540, endMinute: 1020 }])
    expect(resolved.id).toBe('child')
  })

  it('rejects inheritance cycles and unknown calendars', () => {
    const a = makeCalendar('a', { baseCalendarId: asCalendarId('b') })
    const b = makeCalendar('b', { baseCalendarId: asCalendarId('a') })
    expect(() => resolveCalendar({ calendars: [a, b] }, asCalendarId('a'))).toThrowError(
      CalendarError,
    )
    expect(() => resolveCalendar({ calendars: [a, b] }, asCalendarId('a'))).toThrowError(/cycle/i)
    expect(() => resolveCalendar({ calendars: [a] }, asCalendarId('missing'))).toThrowError(
      /unknown calendar/i,
    )
  })

  it('rejects malformed working periods', () => {
    const reversed = makeCalendar('reversed', {
      workingWeek: { ...standardWeek(), 1: [{ startMinute: 1020, endMinute: 540 }] },
    })
    expect(() => resolveCalendar({ calendars: [reversed] }, asCalendarId('reversed'))).toThrowError(
      /period/i,
    )
    const outOfRange = makeCalendar('range', {
      workingWeek: { ...standardWeek(), 2: [{ startMinute: 0, endMinute: 1500 }] },
    })
    expect(() => resolveCalendar({ calendars: [outOfRange] }, asCalendarId('range'))).toThrowError(
      /period/i,
    )
    const overlapping = makeCalendar('overlap', {
      workingWeek: {
        ...standardWeek(),
        3: [
          { startMinute: 540, endMinute: 720 },
          { startMinute: 700, endMinute: 900 },
        ],
      },
    })
    expect(() =>
      resolveCalendar({ calendars: [overlapping] }, asCalendarId('overlap')),
    ).toThrowError(/overlap/i)
    const fractional = makeCalendar('fractional', {
      workingWeek: { ...standardWeek(), 4: [{ startMinute: 540.5, endMinute: 1020 }] },
    })
    expect(() =>
      resolveCalendar({ calendars: [fractional] }, asCalendarId('fractional')),
    ).toThrowError(/period/i)
  })

  it('produces identical results on repeated resolution', () => {
    const base = makeCalendar('base')
    const child = makeCalendar('child', {
      baseCalendarId: asCalendarId('base'),
      exceptions: [holiday('2026-08-04')],
    })
    const first = resolveCalendar({ calendars: [base, child] }, asCalendarId('child'))
    const second = resolveCalendar({ calendars: [base, child] }, asCalendarId('child'))
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
  })
})
