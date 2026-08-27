/**
 * PROJECT-025 — calendar visualization integration with the REAL
 * scheduling authority (the no-second-calendar-engine evidence).
 *
 * The canonical working-time query is bound to the REAL scheduling
 * package exactly as hosts bind it (the `ScheduleRunner` injection
 * precedent — the scheduling import lives at the TEST layer only):
 *
 *     (calendars, calendarId, start, finish) =>
 *       workingIntervals(resolveCalendar({ calendars }, calendarId), start, finish)
 *
 * Every evaluated instant in the projection therefore IS the authority's
 * answer: the surface's working bands equal `workingIntervals` output for
 * the same window, non-working bands are its complement, weekend/holiday/
 * inheritance semantics come from `resolveCalendar` — never re-derived.
 * The battery also proves the timeline integration end-to-end (session →
 * projection → buildGanttView: the background surface + per-row surfaces
 * from the schedule's `resolvedCalendarId`), the additive-surface
 * geometry-neutrality contract, the absent-without-query degradation, and
 * the diagnostic-code equivalence with `schedule()`'s own calendar-error
 * degradation.
 */
import { describe, expect, it } from 'vitest'
import {
  asCalendarId,
  asISODateTime,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import type { CalendarWorkingTimeQuery } from '../src/index.js'
import {
  buildCalendarCatalog,
  buildCalendarSurface,
  buildGanttView,
  createRendererSession,
  createViewState,
  projectDocumentView,
  scaleViewport,
} from '../src/index.js'
import { resolveCalendar, workingIntervals } from '@genoffice/project-scheduling'
import { schedule } from '@genoffice/project-scheduling'
import { makeCalendar, makeDocument, makeTask } from './fixtures.js'

/** The canonical binding: the documented host-side adapter. The book
 * spread bridges the renderer's readonly calendars contract into the
 * scheduling package's mutable array type (`resolveCalendar` only reads
 * it). Pure — the same inputs always produce the same intervals (asserted
 * by the determinism scenarios). */
const canonicalQuery: CalendarWorkingTimeQuery = (calendars, calendarId, start, finish) =>
  workingIntervals(resolveCalendar({ calendars: [...calendars] }, calendarId), start, finish)

/** The canonical answer for one calendar/window, straight from the
 * authority — the projection's working bands must equal this. */
const canonicalWorking = (
  document: ProjectDocument,
  calendarId: string,
  start: string,
  finish: string,
) =>
  workingIntervals(
    resolveCalendar({ calendars: [...document.calendars] }, asCalendarId(calendarId)),
    asISODateTime(start),
    asISODateTime(finish),
  )

const WEEK = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-10T00:00:00.000Z' } // Mon..Sun
// 2026-08-03 is a Monday; 08-08/08-09 are the weekend.

describe('PROJECT-025 integration — bands ARE the canonical evaluation', () => {
  it('equals workingIntervals output on the standard week (weekend complement)', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK)
    expect(surface.status).toBe('ok')
    expect(surface.calendarId).toBe(asCalendarId('standard'))
    expect(surface.name).toBe('standard')
    const working = surface.bands?.filter((band) => band.working) ?? []
    const expected = canonicalWorking(document, 'standard', WEEK.start, WEEK.finish)
    // The working bands are the authority's intervals, verbatim.
    expect(working.map((band) => [band.start, band.finish])).toEqual(
      expected.map((interval) => [interval.start, interval.finish]),
    )
    // Five weekday bands 09:00–17:00; weekends are the complement.
    expect(working).toHaveLength(5)
    expect(working[0]).toEqual({
      start: '2026-08-03T09:00:00.000Z',
      finish: '2026-08-03T17:00:00.000Z',
      working: true,
    })
    const nonWorking = surface.bands?.filter((band) => !band.working) ?? []
    expect(nonWorking).toHaveLength(6) // nights ×5 + the weekend span
    // Bands are contiguous and cover the window exactly.
    const all = surface.bands ?? []
    expect(all[0]?.start).toBe(WEEK.start)
    expect(all[all.length - 1]?.finish).toBe(WEEK.finish)
  })

  it('projects a split-period day (morning/afternoon) as alternating bands', () => {
    const split = makeCalendar('split', {
      workingWeek: {
        0: [],
        1: [
          { startMinute: 540, endMinute: 720 },
          { startMinute: 780, endMinute: 1020 },
        ],
        2: [],
        3: [],
        4: [],
        5: [],
        6: [],
      },
    })
    const document = makeDocument({
      calendars: [split],
      startDate: '2026-08-03T00:00:00.000Z',
    })
    const surface = buildCalendarSurface(document, canonicalQuery, {
      start: '2026-08-03T00:00:00.000Z',
      finish: '2026-08-04T00:00:00.000Z',
    })
    expect(surface.bands?.map((band) => [band.working, band.start, band.finish])).toEqual([
      [false, '2026-08-03T00:00:00.000Z', '2026-08-03T09:00:00.000Z'],
      [true, '2026-08-03T09:00:00.000Z', '2026-08-03T12:00:00.000Z'],
      [false, '2026-08-03T12:00:00.000Z', '2026-08-03T13:00:00.000Z'],
      [true, '2026-08-03T13:00:00.000Z', '2026-08-03T17:00:00.000Z'],
      [false, '2026-08-03T17:00:00.000Z', '2026-08-04T00:00:00.000Z'],
    ])
  })

  it('renders a declared non-working exception date as a fully non-working day', () => {
    const holiday = makeCalendar('standard')
    holiday.exceptions.push({ date: '2026-08-04', periods: [] }) // Tuesday off
    const document = makeDocument({ calendars: [holiday] })
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK)
    const working = surface.bands?.filter((band) => band.working) ?? []
    expect(working).toHaveLength(4) // Monday + Wednesday..Friday
    expect(working.some((band) => band.start.startsWith('2026-08-04'))).toBe(false)
  })

  it('renders a declared working exception on a non-working weekday', () => {
    const catchUp = makeCalendar('standard')
    catchUp.exceptions.push({
      date: '2026-08-09', // Sunday
      periods: [{ startMinute: 540, endMinute: 1020 }],
    })
    const document = makeDocument({ calendars: [catchUp] })
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK)
    const working = surface.bands?.filter((band) => band.working) ?? []
    expect(working).toHaveLength(6) // Mon..Fri + the working Sunday
    expect(working[5]).toEqual({
      start: '2026-08-09T09:00:00.000Z',
      finish: '2026-08-09T17:00:00.000Z',
      working: true,
    })
  })

  it('resolves base-calendar inheritance exactly as resolveCalendar does', () => {
    const base = makeCalendar('base')
    base.exceptions.push({
      date: '2026-08-05', // Wednesday short day on the base
      periods: [{ startMinute: 540, endMinute: 780 }],
    })
    const child = makeCalendar('shift', { baseCalendarId: base.id })
    child.workingWeek = { ...child.workingWeek, 6: [{ startMinute: 540, endMinute: 1020 }] } // Saturdays work
    child.exceptions.push({ date: '2026-08-06', periods: [] }) // Thursday off (child exception)
    const document = makeDocument({ calendars: [base, child] })
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK, asCalendarId('shift'))
    const working = surface.bands?.filter((band) => band.working) ?? []
    // Mon, Tue full (week), Wed short (base exception 09:00–13:00 — one
    // band), Fri full, Sat full (child weekday override); Thu off (child
    // exception); Sun off.
    expect(working.map((band) => band.start.slice(0, 10))).toEqual([
      '2026-08-03',
      '2026-08-04',
      '2026-08-05',
      '2026-08-07',
      '2026-08-08',
    ])
    expect(working[2]?.finish).toBe('2026-08-05T13:00:00.000Z') // short day 09:00–13:00
    expect(working[3]?.start).toBe('2026-08-07T09:00:00.000Z')
    // …and the projection equals the authority on the same question.
    expect(working.map((band) => [band.start, band.finish])).toEqual(
      canonicalWorking(document, 'shift', WEEK.start, WEEK.finish).map((interval) => [
        interval.start,
        interval.finish,
      ]),
    )
  })

  it('lets a child exception override the base exception for the same date', () => {
    const base = makeCalendar('base')
    base.exceptions.push({
      date: '2026-08-04',
      periods: [{ startMinute: 540, endMinute: 1020 }], // full day on the base
    })
    const child = makeCalendar('child', { baseCalendarId: base.id })
    child.exceptions.push({ date: '2026-08-04', periods: [] }) // child overrides: off
    const document = makeDocument({ calendars: [base, child] })
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK, asCalendarId('child'))
    const working = surface.bands?.filter((band) => band.working) ?? []
    expect(working.map((band) => band.start.slice(0, 10))).not.toContain('2026-08-04')
  })

  it('re-projects bands when the viewport changes (zoom), clipped to the new window', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const day = {
      start: '2026-08-03T00:00:00.000Z',
      finish: '2026-08-04T00:00:00.000Z',
    }
    const surface = buildCalendarSurface(document, canonicalQuery, day)
    expect(surface.bands).toHaveLength(3) // night, 09–17, night
    const zoomed = scaleViewport(day, 0.5)
    const zoomedSurface = buildCalendarSurface(document, canonicalQuery, zoomed)
    expect(zoomedSurface.start).toBe(zoomed.start)
    expect(zoomedSurface.finish).toBe(zoomed.finish)
    expect(zoomedSurface.bands?.[0]?.start).toBe(zoomed.start)
    expect(zoomedSurface.bands?.[zoomedSurface.bands.length - 1]?.finish).toBe(zoomed.finish)
  })
})

describe('PROJECT-025 integration — the timeline surface through the real pipeline', () => {
  /** A document on two calendars: default `standard` (Mon–Fri 09–17) and
   * `compact` (Mon–Thu 08–18). Task `plain` uses the default; task
   * `dense` overrides to `compact`; the real scheduler resolves each
   * task's calendar and echoes `resolvedCalendarId` on its schedule. */
  const pipeline = () => {
    const compact = makeCalendar('compact', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 480, endMinute: 1080 }],
        2: [{ startMinute: 480, endMinute: 1080 }],
        3: [{ startMinute: 480, endMinute: 1080 }],
        4: [{ startMinute: 480, endMinute: 1080 }],
        5: [],
        6: [],
      },
    })
    const document = makeDocument({
      calendars: [makeCalendar('standard'), compact],
      startDate: '2026-08-03T08:00:00.000Z',
      tasks: [
        makeTask({ id: 'root', uid: 1, outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({
          id: 'plain',
          uid: 2,
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          duration: asWorkingMinutes(480),
          wbs: '1.1',
        }),
        makeTask({
          id: 'dense',
          uid: 3,
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          duration: asWorkingMinutes(600),
          calendarId: compact.id,
          wbs: '1.2',
        }),
      ],
    })
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)
    const projection = projectDocumentView(document, session.schedule, state)
    const layout = { firstRow: 0, visibleRows: 3, overscan: 0 }
    return { document, session, state, projection, layout }
  }

  it('carries the background surface and per-row surfaces resolved by the authority', () => {
    const { document, state, projection, layout } = pipeline()
    const view = buildGanttView(document, projection, state, layout, {
      workingTime: canonicalQuery,
    })
    const timeline = view.timeline
    expect(timeline.calendar?.status).toBe('ok')
    expect(timeline.calendar?.calendarId).toBe(asCalendarId('standard'))
    // Row surfaces: every scheduled row carries the AUTHORITY's resolved
    // calendar id (root/plain → standard, dense → compact).
    expect(
      timeline.rowCalendars?.map((entry) => [entry.rowIndex, entry.taskId, entry.calendarId]),
    ).toEqual([
      [0, asTaskId('root'), asCalendarId('standard')],
      [1, asTaskId('plain'), asCalendarId('standard')],
      [2, asTaskId('dense'), asCalendarId('compact')],
    ])
    // The compact row's Monday band starts 08:00 (its resolved calendar),
    // not 09:00 — per-row shading reflects the task's canonical calendar.
    const dense = timeline.rowCalendars?.[2]
    const denseMonday = dense?.surface.bands?.find(
      (band) => band.working && band.start.startsWith('2026-08-03'),
    )
    expect(denseMonday).toEqual({
      start: '2026-08-03T08:00:00.000Z',
      finish: '2026-08-03T18:00:00.000Z',
      working: true,
    })
    const plainMonday = timeline.rowCalendars?.[1]?.surface.bands?.find(
      (band) => band.working && band.start.startsWith('2026-08-03'),
    )
    expect(plainMonday).toEqual({
      start: '2026-08-03T09:00:00.000Z',
      finish: '2026-08-03T17:00:00.000Z',
      working: true,
    })
    // Surfaces for the same calendar are shared by reference within the
    // build (the per-build memo) — the background and the standard rows.
    expect(timeline.rowCalendars?.[0]?.surface).toBe(timeline.calendar)
    expect(timeline.rowCalendars?.[1]?.surface).toBe(timeline.calendar)
    expect(dense?.surface).not.toBe(timeline.calendar)
    // …and every surface equals the authority on its own question.
    const expected = canonicalWorking(
      document,
      'compact',
      timeline.viewport.start,
      timeline.viewport.finish,
    )
    expect(
      dense?.surface.bands?.filter((band) => band.working).map((band) => [band.start, band.finish]),
    ).toEqual(expected.map((interval) => [interval.start, interval.finish]))
  })

  it('builds no calendar surfaces without the query (never invented)', () => {
    const { document, state, projection, layout } = pipeline()
    const view = buildGanttView(document, projection, state, layout)
    expect(view.timeline.calendar).toBeUndefined()
    expect(view.timeline.rowCalendars).toBeUndefined()
    // The query alone (no calendarId override) still yields surfaces.
    const withQuery = buildGanttView(document, projection, state, layout, {
      workingTime: canonicalQuery,
    })
    expect(withQuery.timeline.calendar).toBeDefined()
  })

  it('shades the background by an explicit calendar when one is given', () => {
    const { document, state, projection, layout } = pipeline()
    const view = buildGanttView(document, projection, state, layout, {
      workingTime: canonicalQuery,
      calendarId: asCalendarId('compact'),
    })
    expect(view.timeline.calendar?.calendarId).toBe(asCalendarId('compact'))
    const monday = view.timeline.calendar?.bands?.find(
      (band) => band.working && band.start.startsWith('2026-08-03'),
    )
    expect(monday?.start).toBe('2026-08-03T08:00:00.000Z')
  })

  it('keeps the geometry surfaces byte-identical with and without the query', () => {
    const { document, state, projection, layout } = pipeline()
    const bare = buildGanttView(document, projection, state, layout)
    const shaded = buildGanttView(document, projection, state, layout, {
      workingTime: canonicalQuery,
    })
    expect(shaded.timeline.bands).toEqual(bare.timeline.bands)
    expect(shaded.timeline.rows).toEqual(bare.timeline.rows)
    expect(shaded.timeline.bars).toEqual(bare.timeline.bars)
    expect(shaded.timeline.milestones).toEqual(bare.timeline.milestones)
    expect(shaded.timeline.links).toEqual(bare.timeline.links)
    expect(shaded.taskGrid).toEqual(bare.taskGrid)
    expect(shaded.rowWindow).toEqual(bare.rowWindow)
  })

  it('carries empty row surfaces when no scheduler is wired (no resolved ids)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'solo', uid: 1, outlineLevel: 1, wbs: '1' })],
    })
    const state = createViewState(document)
    const projection = projectDocumentView(document, undefined, state)
    const view = buildGanttView(
      document,
      projection,
      state,
      { firstRow: 0, visibleRows: 5, overscan: 0 },
      {
        workingTime: canonicalQuery,
      },
    )
    // The background surface exists (the default calendar evaluates over
    // the viewport) but no row has a resolved calendar id — real
    // information, not a missing surface.
    expect(view.timeline.calendar?.status).toBe('ok')
    expect(view.timeline.rowCalendars).toEqual([])
  })

  it('degrades an unresolvable default calendar with the same code schedule() reports', () => {
    const base = makeDocument({
      tasks: [makeTask({ id: 'solo', uid: 1, outlineLevel: 1, wbs: '1' })],
    })
    const document: ProjectDocument = {
      ...base,
      properties: { ...base.properties, defaultCalendarId: asCalendarId('ghost') },
    }
    // The scheduling authority degrades to diagnostics…
    const derived = schedule(document)
    expect(derived.taskSchedules).toEqual({})
    expect(derived.diagnostics.map((diagnostic) => diagnostic.code)).toContain('MISSING_CALENDAR')
    // …and the calendar surface echoes the SAME code, never invented bands.
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK)
    expect(surface.status).toBe('unresolvable')
    expect(surface.diagnostic?.code).toBe('MISSING_CALENDAR')
    expect('bands' in surface).toBe(false)
    // Through the full pipeline: no schedule values AND an unresolvable
    // background surface — coherent degradation everywhere.
    const state = createViewState(document)
    const projection = projectDocumentView(document, derived, state)
    const view = buildGanttView(
      document,
      projection,
      state,
      { firstRow: 0, visibleRows: 5, overscan: 0 },
      {
        workingTime: canonicalQuery,
      },
    )
    expect(projection.hasSchedule).toBe(true) // a schedule object exists…
    expect(projection.rows[0]?.schedule).toBeUndefined() // …but no values for tasks
    expect(view.timeline.calendar?.status).toBe('unresolvable')
    expect(view.timeline.rowCalendars).toEqual([])
  })

  it('degrades an inheritance cycle with the same code schedule() reports', () => {
    const a = makeCalendar('a', { baseCalendarId: asCalendarId('b') })
    const b = makeCalendar('b', { baseCalendarId: asCalendarId('a') })
    const document = makeDocument({ calendars: [a, b] })
    const derived = schedule(document)
    expect(derived.diagnostics.map((diagnostic) => diagnostic.code)).toContain('CALENDAR_CYCLE')
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK, asCalendarId('a'))
    expect(surface.status).toBe('unresolvable')
    expect(surface.diagnostic?.code).toBe('CALENDAR_CYCLE')
  })

  it('projects the catalog over real document calendars', () => {
    const compact = makeCalendar('compact', {
      workingWeek: {
        0: [],
        1: [{ startMinute: 480, endMinute: 1080 }],
        2: [{ startMinute: 480, endMinute: 1080 }],
        3: [{ startMinute: 480, endMinute: 1080 }],
        4: [{ startMinute: 480, endMinute: 1080 }],
        5: [],
        6: [],
      },
    })
    const document = makeDocument({
      calendars: [makeCalendar('standard'), compact],
      tasks: [makeTask({ id: 't', calendarId: compact.id })],
    })
    const catalog = buildCalendarCatalog(document)
    expect(catalog.defaultCalendarId).toBe(asCalendarId('standard'))
    expect(catalog.calendars[0]).toMatchObject({
      calendarId: asCalendarId('standard'),
      workingWeekdays: [1, 2, 3, 4, 5],
      declaredWeeklyMinutes: 2400,
      exceptionCount: 0,
      taskCount: 0,
    })
    expect(catalog.calendars[1]).toMatchObject({
      calendarId: asCalendarId('compact'),
      workingWeekdays: [1, 2, 3, 4],
      declaredWeeklyMinutes: 2400,
      taskCount: 1,
    })
  })

  it('is deterministic end-to-end (3× byte-identical, pure query)', () => {
    const { document, state, projection, layout } = pipeline()
    const build = () =>
      JSON.stringify(
        buildGanttView(document, projection, state, layout, { workingTime: canonicalQuery }),
      )
    const first = build()
    expect(build()).toBe(first)
    expect(build()).toBe(first)
    // The canonical query itself is pure: same inputs, same intervals.
    const once = canonicalQuery(
      document.calendars,
      asCalendarId('standard'),
      asISODateTime(WEEK.start),
      asISODateTime(WEEK.finish),
    )
    expect(
      canonicalQuery(
        document.calendars,
        asCalendarId('standard'),
        asISODateTime(WEEK.start),
        asISODateTime(WEEK.finish),
      ),
    ).toEqual(once)
  })
})
