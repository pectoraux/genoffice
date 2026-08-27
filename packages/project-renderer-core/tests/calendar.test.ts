/**
 * PROJECT-025 — the calendar projection layer (unit battery).
 *
 * Pure-projection semantics only — NO scheduling package anywhere in this
 * file (the query is a stub; the REAL canonical binding is exercised in
 * `calendar-integration.test.ts`, the accepted test-layer precedent):
 * the catalog echo/aggregation contract, the band classification's pure
 * interval algebra (clip/sort/merge/complement, contiguity, order
 * independence), and the surface builder's degradation contract
 * (degenerate window, coded-error echo, uncoded-error re-throw).
 */
import { describe, expect, it, vi } from 'vitest'
import { asCalendarId, asISODateTime } from '@genoffice/project-contracts'
import type { CalendarWorkingInterval, CalendarWorkingTimeQuery } from '../src/index.js'
import {
  CALENDAR_EVALUATION_FAILED,
  buildCalendarCatalog,
  buildCalendarSurface,
  classifyCalendarBands,
} from '../src/index.js'
import { makeCalendar, makeDocument, makeTask } from './fixtures.js'

const WINDOW = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-07T00:00:00.000Z' }
const interval = (start: string, finish: string): CalendarWorkingInterval => ({
  start: asISODateTime(start),
  finish: asISODateTime(finish),
})

describe('PROJECT-025 buildCalendarCatalog — the pure calendar echo', () => {
  it('echoes the calendars in document order with verbatim references', () => {
    const standard = makeCalendar('standard')
    const night = makeCalendar('night', { name: 'Night shift' })
    const document = makeDocument({ calendars: [standard, night] })
    const catalog = buildCalendarCatalog(document)
    expect(catalog.defaultCalendarId).toBe(asCalendarId('standard'))
    expect(catalog.calendars.map((entry) => entry.calendarId)).toEqual([
      asCalendarId('standard'),
      asCalendarId('night'),
    ])
    // Verbatim echoes — joined by reference, never copied.
    expect(catalog.calendars[0]?.calendar).toBe(standard)
    expect(catalog.calendars[1]?.calendar).toBe(night)
    expect(catalog.calendars[1]?.name).toBe('Night shift')
  })

  it('aggregates declared minutes per weekday (display aggregation of the declaration)', () => {
    const document = makeDocument({
      calendars: [
        makeCalendar('split', {
          workingWeek: {
            0: [],
            1: [
              { startMinute: 540, endMinute: 720 },
              { startMinute: 780, endMinute: 1020 },
            ],
            2: [],
            3: [{ startMinute: 540, endMinute: 720 }],
            4: [],
            5: [],
            6: [],
          },
        }),
      ],
    })
    const entry = buildCalendarCatalog(document).calendars[0]
    expect(entry?.workingWeekdays).toEqual([1, 3])
    expect(entry?.declaredMinutes).toEqual({ 1: 420, 3: 180 })
    expect(entry?.declaredWeeklyMinutes).toBe(600)
  })

  it('echoes base calendars, exception counts, and direct task references', () => {
    const base = makeCalendar('base')
    const child = makeCalendar('child', { baseCalendarId: base.id })
    child.exceptions.push({ date: '2026-08-04', periods: [] })
    child.exceptions.push({ date: '2026-08-05', periods: [{ startMinute: 540, endMinute: 1020 }] })
    const document = makeDocument({
      calendars: [base, child],
      tasks: [
        makeTask({ id: 'a', calendarId: child.id }),
        makeTask({ id: 'b', calendarId: child.id }),
        makeTask({ id: 'c' }), // default calendar — not a direct reference
        makeTask({ id: 'd', calendarId: base.id }),
      ],
    })
    const catalog = buildCalendarCatalog(document)
    expect(catalog.calendars[0]).toMatchObject({
      calendarId: asCalendarId('base'),
      exceptionCount: 0,
      taskCount: 1,
    })
    expect(catalog.calendars[1]).toMatchObject({
      calendarId: asCalendarId('child'),
      baseCalendarId: asCalendarId('base'),
      exceptionCount: 2,
      taskCount: 2,
    })
  })

  it('handles an empty calendar set and drops non-canonical weekday keys', () => {
    const odd = makeCalendar('odd', {
      workingWeek: { 1: [{ startMinute: 0, endMinute: 60 }], 9: [], x: [] } as never,
    })
    const document = makeDocument({ calendars: [odd] })
    const entry = buildCalendarCatalog(document).calendars[0]
    expect(entry?.workingWeekdays).toEqual([1])
    expect(entry?.declaredMinutes).toEqual({ 1: 60 })
    const empty = makeDocument({ calendars: [] })
    const catalog = buildCalendarCatalog(empty)
    expect(catalog.calendars).toEqual([])
    expect(catalog.defaultCalendarId).toBe(asCalendarId('standard'))
  })

  it('is deterministic (3× byte-identical)', () => {
    const document = makeDocument({
      calendars: [makeCalendar('a'), makeCalendar('b', { baseCalendarId: asCalendarId('a') })],
      tasks: [makeTask({ id: 't', calendarId: asCalendarId('b') })],
    })
    const first = JSON.stringify(buildCalendarCatalog(document))
    expect(JSON.stringify(buildCalendarCatalog(document))).toBe(first)
    expect(JSON.stringify(buildCalendarCatalog(document))).toBe(first)
  })
})

describe('PROJECT-025 classifyCalendarBands — pure interval algebra', () => {
  it('returns one non-working band covering the window when nothing works', () => {
    expect(classifyCalendarBands(WINDOW, [])).toEqual([
      { start: WINDOW.start, finish: WINDOW.finish, working: false },
    ])
  })

  it('returns one working band when the evaluator covers the whole window', () => {
    expect(classifyCalendarBands(WINDOW, [interval(WINDOW.start, WINDOW.finish)])).toEqual([
      { start: WINDOW.start, finish: WINDOW.finish, working: true },
    ])
  })

  it('alternates working/non-working and keeps the contiguity invariants', () => {
    const bands = classifyCalendarBands(WINDOW, [
      interval('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z'),
      interval('2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z'),
    ])
    expect(bands.map((band) => [band.working, band.start, band.finish])).toEqual([
      [false, '2026-08-03T00:00:00.000Z', '2026-08-03T09:00:00.000Z'],
      [true, '2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z'],
      [false, '2026-08-03T17:00:00.000Z', '2026-08-04T09:00:00.000Z'],
      [true, '2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z'],
      [false, '2026-08-04T17:00:00.000Z', '2026-08-07T00:00:00.000Z'],
    ])
    // Contiguity: the union covers the window exactly, no empty bands.
    expect(bands[0]?.start).toBe(WINDOW.start)
    expect(bands[bands.length - 1]?.finish).toBe(WINDOW.finish)
    for (let index = 1; index < bands.length; index += 1) {
      expect(bands[index]?.start).toBe(bands[index - 1]?.finish)
    }
    for (const band of bands) {
      expect(band.finish > band.start).toBe(true)
    }
  })

  it('clips evaluator intervals to the window', () => {
    const bands = classifyCalendarBands(WINDOW, [
      interval('2026-08-01T09:00:00.000Z', '2026-08-03T12:00:00.000Z'),
      interval('2026-08-06T09:00:00.000Z', '2026-08-09T12:00:00.000Z'),
    ])
    expect(bands).toEqual([
      { start: WINDOW.start, finish: '2026-08-03T12:00:00.000Z', working: true },
      { start: '2026-08-03T12:00:00.000Z', finish: '2026-08-06T09:00:00.000Z', working: false },
      { start: '2026-08-06T09:00:00.000Z', finish: WINDOW.finish, working: true },
    ])
  })

  it('merges overlapping and adjacent evaluator intervals', () => {
    const bands = classifyCalendarBands(WINDOW, [
      interval('2026-08-03T09:00:00.000Z', '2026-08-03T12:00:00.000Z'),
      interval('2026-08-03T12:00:00.000Z', '2026-08-03T15:00:00.000Z'), // adjacent
      interval('2026-08-03T14:00:00.000Z', '2026-08-03T17:00:00.000Z'), // overlapping
    ])
    expect(bands).toEqual([
      { start: WINDOW.start, finish: '2026-08-03T09:00:00.000Z', working: false },
      { start: '2026-08-03T09:00:00.000Z', finish: '2026-08-03T17:00:00.000Z', working: true },
      { start: '2026-08-03T17:00:00.000Z', finish: WINDOW.finish, working: false },
    ])
  })

  it('is independent of the evaluator intervals order', () => {
    const intervals = [
      interval('2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z'),
      interval('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z'),
    ]
    expect(classifyCalendarBands(WINDOW, [...intervals].reverse())).toEqual(
      classifyCalendarBands(WINDOW, intervals),
    )
  })

  it('drops unparseable, empty, and inverted intervals deterministically', () => {
    expect(
      classifyCalendarBands(WINDOW, [
        interval('not-a-date', '2026-08-03T12:00:00.000Z'),
        interval('2026-08-03T12:00:00.000Z', '2026-08-03T12:00:00.000Z'),
        interval('2026-08-04T12:00:00.000Z', '2026-08-04T09:00:00.000Z'),
      ]),
    ).toEqual([{ start: WINDOW.start, finish: WINDOW.finish, working: false }])
  })

  it('yields no bands for a degenerate window', () => {
    expect(classifyCalendarBands({ start: WINDOW.finish, finish: WINDOW.start }, [])).toEqual([])
    expect(
      classifyCalendarBands({ start: 'garbage', finish: '2026-08-07T00:00:00.000Z' }, []),
    ).toEqual([])
  })

  it('is deterministic (3× byte-identical)', () => {
    const intervals = [
      interval('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z'),
      interval('2026-08-05T09:00:00.000Z', '2026-08-05T17:00:00.000Z'),
    ]
    const first = JSON.stringify(classifyCalendarBands(WINDOW, intervals))
    expect(JSON.stringify(classifyCalendarBands(WINDOW, intervals))).toBe(first)
    expect(JSON.stringify(classifyCalendarBands(WINDOW, intervals))).toBe(first)
  })
})

describe('PROJECT-025 buildCalendarSurface — evaluation + degradation', () => {
  const workingIntervals: CalendarWorkingInterval[] = [
    interval('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z'),
  ]

  it('classifies the canonical evaluator output over the window', () => {
    const query: CalendarWorkingTimeQuery = () => workingIntervals
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const surface = buildCalendarSurface(document, query, WINDOW)
    expect(surface).toEqual({
      calendarId: asCalendarId('standard'),
      name: 'standard',
      status: 'ok',
      start: WINDOW.start,
      finish: WINDOW.finish,
      bands: [
        { start: WINDOW.start, finish: '2026-08-03T09:00:00.000Z', working: false },
        { start: '2026-08-03T09:00:00.000Z', finish: '2026-08-03T17:00:00.000Z', working: true },
        { start: '2026-08-03T17:00:00.000Z', finish: WINDOW.finish, working: false },
      ],
    })
  })

  it('hands the evaluator the calendars by reference, the resolved id, and the window', () => {
    const query = vi.fn<CalendarWorkingTimeQuery>(() => [])
    const calendars = [makeCalendar('standard'), makeCalendar('night')]
    const document = makeDocument({ calendars })
    buildCalendarSurface(document, query, WINDOW)
    expect(query).toHaveBeenCalledWith(
      calendars,
      asCalendarId('standard'),
      WINDOW.start,
      WINDOW.finish,
    )
    expect(query).toHaveBeenCalledTimes(1)
  })

  it('resolves an explicit calendarId over the document default', () => {
    const query = vi.fn<CalendarWorkingTimeQuery>(() => [])
    const document = makeDocument({
      calendars: [makeCalendar('standard'), makeCalendar('night', { name: 'Night' })],
    })
    const surface = buildCalendarSurface(document, query, WINDOW, asCalendarId('night'))
    expect(surface.calendarId).toBe(asCalendarId('night'))
    expect(surface.name).toBe('Night')
    expect(query).toHaveBeenCalledWith(
      document.calendars,
      asCalendarId('night'),
      WINDOW.start,
      WINDOW.finish,
    )
  })

  it('leaves the name absent when the id is not a document calendar', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const surface = buildCalendarSurface(document, () => [], WINDOW, asCalendarId('somewhere-else'))
    expect(surface.name).toBeUndefined()
    expect(surface.calendarId).toBe(asCalendarId('somewhere-else'))
  })

  it('degrades a coded evaluator error to unresolvable with the diagnostic echoed', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const query: CalendarWorkingTimeQuery = () => {
      throw Object.assign(new Error('Unknown calendar missing'), { code: 'MISSING_CALENDAR' })
    }
    const surface = buildCalendarSurface(document, query, WINDOW)
    expect(surface).toEqual({
      calendarId: asCalendarId('standard'),
      name: 'standard',
      status: 'unresolvable',
      diagnostic: {
        code: 'MISSING_CALENDAR',
        severity: 'error',
        message: 'Unknown calendar missing',
      },
      start: WINDOW.start,
      finish: WINDOW.finish,
    })
    expect('bands' in surface).toBe(false)
  })

  it('re-throws evaluator errors that carry no code (host-binding bugs surface)', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const query: CalendarWorkingTimeQuery = () => {
      throw new TypeError('binding bug')
    }
    expect(() => buildCalendarSurface(document, query, WINDOW)).toThrowError(TypeError)
  })

  it('stringifies non-Error coded throws with the fallback-agnostic code echo', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const query: CalendarWorkingTimeQuery = () => {
      throw Object.assign('cycle!', { code: 'CALENDAR_CYCLE' })
    }
    const surface = buildCalendarSurface(document, query, WINDOW)
    expect(surface.status).toBe('unresolvable')
    expect(surface.diagnostic).toEqual({
      code: 'CALENDAR_CYCLE',
      severity: 'error',
      message: 'cycle!',
    })
    expect(CALENDAR_EVALUATION_FAILED).toBe('CALENDAR_EVALUATION_FAILED')
  })

  it('returns an empty band set for a degenerate window without consulting the evaluator', () => {
    const query = vi.fn<CalendarWorkingTimeQuery>(() => [])
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const surface = buildCalendarSurface(document, query, {
      start: WINDOW.finish,
      finish: WINDOW.start,
    })
    expect(surface).toEqual({
      calendarId: asCalendarId('standard'),
      name: 'standard',
      status: 'ok',
      bands: [],
      start: WINDOW.finish,
      finish: WINDOW.start,
    })
    expect(query).not.toHaveBeenCalled()
  })

  it('is deterministic (3× byte-identical, pure query)', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const query: CalendarWorkingTimeQuery = () => workingIntervals
    const first = JSON.stringify(buildCalendarSurface(document, query, WINDOW))
    expect(JSON.stringify(buildCalendarSurface(document, query, WINDOW))).toBe(first)
    expect(JSON.stringify(buildCalendarSurface(document, query, WINDOW))).toBe(first)
  })
})
