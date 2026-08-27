/**
 * PROJECT-025 — the golden scenario battery (C01–C12).
 *
 * Each golden drives one acceptance scenario end-to-end through the real
 * machinery — the canonical working-time query bound to the REAL
 * scheduling package (`resolveCalendar` + `workingIntervals`), documents
 * scheduled by the REAL scheduler, surfaces built through
 * `buildCalendarSurface` / the accepted `buildGanttView` pipeline — and
 * asserts the complete observable tuple: the calendar catalog, the
 * background working-time surface (EQUAL to the authority's own answer —
 * never re-derived), the per-row surfaces (each row's task resolved
 * calendar from the real schedule), and — where the scenario touches
 * them — the geometry surfaces (proving the calendar surface is
 * additive). Pure inputs, no wall clock, no randomness; C11 keeps the
 * surfaces stable across session commands; C12 proves a mixed
 * command/viewport sequence replays byte-identical (3×).
 */
import { describe, expect, it } from 'vitest'
import {
  asCalendarId,
  asISODateTime,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type { ProjectCommand, ProjectDocument } from '@genoffice/project-contracts'
import { resolveCalendar, schedule, workingIntervals } from '@genoffice/project-scheduling'
import {
  applyRendererCommand,
  buildCalendarCatalog,
  buildCalendarSurface,
  buildGanttView,
  createRendererSession,
  createViewState,
  projectDocumentView,
  redoRendererCommand,
  scaleViewport,
  undoRendererCommand,
} from '../src/index.js'
import type { CalendarWorkingTimeQuery } from '../src/index.js'
import { makeCalendar, makeDocument, makeTask } from './fixtures.js'

/** The canonical binding — the documented host-side adapter over the REAL
 * scheduling package. The book spread bridges the renderer's readonly
 * calendars contract into the scheduling package's mutable array type
 * (`resolveCalendar` only reads it). */
const canonicalQuery: CalendarWorkingTimeQuery = (calendars, calendarId, start, finish) =>
  workingIntervals(resolveCalendar({ calendars: [...calendars] }, calendarId), start, finish)

const id = (value: string) => asTaskId(value)
const cal = (value: string) => asCalendarId(value)

const WEEK = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-10T00:00:00.000Z' } // Mon..Sun
const DAY = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-04T00:00:00.000Z' } // Monday

/** The canonical answer straight from the authority, for the equivalence
 * half of each golden (the projection must equal it, never re-derive it). */
const canonicalWorking = (document: ProjectDocument, calendarId: string, window = WEEK) =>
  workingIntervals(
    resolveCalendar({ calendars: [...document.calendars] }, cal(calendarId)),
    asISODateTime(window.start),
    asISODateTime(window.finish),
  )

/** The golden projection: the surface's working bands over `window`. */
const goldenWorking = (document: ProjectDocument, calendarId: string, window = WEEK) =>
  (buildCalendarSurface(document, canonicalQuery, window, cal(calendarId)).bands ?? [])
    .filter((band) => band.working)
    .map((band) => [band.start, band.finish])

/** The two-calendar pipeline fixture: `standard` (Mon–Fri 09–17) default,
 * `compact` (Mon–Thu 08–18) — task `dense` overrides to `compact`. */
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
        parentTaskId: id('root'),
        outlineLevel: 2,
        duration: asWorkingMinutes(480),
        wbs: '1.1',
      }),
      makeTask({
        id: 'dense',
        uid: 3,
        parentTaskId: id('root'),
        outlineLevel: 2,
        duration: asWorkingMinutes(600),
        calendarId: compact.id,
        wbs: '1.2',
      }),
    ],
  })
  const session = createRendererSession(document, { schedule })
  const state = createViewState(document, session.schedule)
  return { document, session, state }
}

/** The observable calendar tuple of one timeline build. */
const calendarTuple = (view: ReturnType<typeof buildGanttView>) => ({
  background: view.timeline.calendar,
  rowCalendars: view.timeline.rowCalendars?.map((entry) => ({
    rowIndex: entry.rowIndex,
    taskId: entry.taskId,
    calendarId: entry.calendarId,
    status: entry.surface.status,
    bands: entry.surface.bands,
  })),
})

describe('PROJECT-025 goldens — canonical calendar semantics in the projection', () => {
  it('C01 — standard week: five weekday bands 09:00–17:00, weekends and nights shaded', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const surface = buildCalendarSurface(document, canonicalQuery, WEEK)
    expect(surface).toMatchObject({
      calendarId: cal('standard'),
      name: 'standard',
      status: 'ok',
      start: WEEK.start,
      finish: WEEK.finish,
    })
    const working = surface.bands?.filter((band) => band.working) ?? []
    expect(working).toHaveLength(5)
    expect(working[0]).toEqual({
      start: '2026-08-03T09:00:00.000Z',
      finish: '2026-08-03T17:00:00.000Z',
      working: true,
    })
    expect(working[4]).toEqual({
      start: '2026-08-07T09:00:00.000Z',
      finish: '2026-08-07T17:00:00.000Z',
      working: true,
    })
    // The complement: nights + the full weekend, contiguous cover.
    const all = surface.bands ?? []
    expect(all).toHaveLength(11) // 5 working + 4 nights + weekend + final night
    expect(all[0]?.start).toBe(WEEK.start)
    expect(all[all.length - 1]?.finish).toBe(WEEK.finish)
    // Equivalence with the authority on the same question.
    expect(working.map((band) => [band.start, band.finish])).toEqual(
      canonicalWorking(document, 'standard').map((interval) => [interval.start, interval.finish]),
    )
  })

  it('C02 — split periods: one working day renders as morning/afternoon bands', () => {
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
    const document = makeDocument({ calendars: [split] })
    const surface = buildCalendarSurface(document, canonicalQuery, DAY)
    expect(surface.bands?.map((band) => [band.working, band.start, band.finish])).toEqual([
      [false, '2026-08-03T00:00:00.000Z', '2026-08-03T09:00:00.000Z'],
      [true, '2026-08-03T09:00:00.000Z', '2026-08-03T12:00:00.000Z'],
      [false, '2026-08-03T12:00:00.000Z', '2026-08-03T13:00:00.000Z'],
      [true, '2026-08-03T13:00:00.000Z', '2026-08-03T17:00:00.000Z'],
      [false, '2026-08-03T17:00:00.000Z', '2026-08-04T00:00:00.000Z'],
    ])
    expect(goldenWorking(document, 'split', DAY)).toEqual([
      ['2026-08-03T09:00:00.000Z', '2026-08-03T12:00:00.000Z'],
      ['2026-08-03T13:00:00.000Z', '2026-08-03T17:00:00.000Z'],
    ])
  })

  it('C03 — holiday exception: the declared date carries no working band', () => {
    const holiday = makeCalendar('standard')
    holiday.exceptions.push({ date: '2026-08-04', periods: [] }) // Tuesday off
    const document = makeDocument({ calendars: [holiday] })
    const working = goldenWorking(document, 'standard')
    expect(working).toHaveLength(4)
    expect(working.map(([start]) => start.slice(0, 10))).not.toContain('2026-08-04')
    // …and the catalog echoes the exception declaration.
    expect(buildCalendarCatalog(document).calendars[0]?.exceptionCount).toBe(1)
  })

  it('C04 — working-Sunday exception: the weekend weekday gains a band', () => {
    const catchUp = makeCalendar('standard')
    catchUp.exceptions.push({
      date: '2026-08-09',
      periods: [{ startMinute: 540, endMinute: 1020 }],
    })
    const document = makeDocument({ calendars: [catchUp] })
    const working = goldenWorking(document, 'standard')
    expect(working).toHaveLength(6)
    expect(working[5]).toEqual(['2026-08-09T09:00:00.000Z', '2026-08-09T17:00:00.000Z'])
  })

  it('C05 — inheritance: child weekday override + base exception both resolve', () => {
    const base = makeCalendar('base')
    base.exceptions.push({
      date: '2026-08-05',
      periods: [{ startMinute: 540, endMinute: 780 }],
    })
    const child = makeCalendar('shift', { baseCalendarId: base.id })
    child.workingWeek = { ...child.workingWeek, 6: [{ startMinute: 540, endMinute: 1020 }] }
    child.exceptions.push({ date: '2026-08-06', periods: [] })
    const document = makeDocument({ calendars: [base, child] })
    expect(goldenWorking(document, 'shift')).toEqual([
      ['2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z'],
      ['2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z'],
      ['2026-08-05T09:00:00.000Z', '2026-08-05T13:00:00.000Z'], // base short day inherited
      ['2026-08-07T09:00:00.000Z', '2026-08-07T17:00:00.000Z'],
      ['2026-08-08T09:00:00.000Z', '2026-08-08T17:00:00.000Z'], // child Saturday override
    ])
    // The catalog echoes the DECLARATION (the child declares weekday 6 and
    // one exception; the base's declaration stays the base's own entry).
    const catalog = buildCalendarCatalog(document)
    expect(catalog.calendars[1]).toMatchObject({
      calendarId: cal('shift'),
      baseCalendarId: cal('base'),
      workingWeekdays: [1, 2, 3, 4, 5, 6],
      exceptionCount: 1,
    })
    expect(catalog.calendars[0]).toMatchObject({
      calendarId: cal('base'),
      workingWeekdays: [1, 2, 3, 4, 5],
      exceptionCount: 1,
    })
  })

  it('C06 — exception override: the child declaration replaces the base for the same date', () => {
    const base = makeCalendar('base')
    base.exceptions.push({
      date: '2026-08-04',
      periods: [{ startMinute: 540, endMinute: 1020 }],
    })
    const child = makeCalendar('child', { baseCalendarId: base.id })
    child.exceptions.push({ date: '2026-08-04', periods: [] })
    const document = makeDocument({ calendars: [base, child] })
    expect(goldenWorking(document, 'child').map(([start]) => start.slice(0, 10))).not.toContain(
      '2026-08-04',
    )
    // The base itself keeps its own working Tuesday.
    expect(goldenWorking(document, 'base').map(([start]) => start.slice(0, 10))).toContain(
      '2026-08-04',
    )
  })

  it('C07 — per-task calendars: row surfaces carry the authority-resolved calendar', () => {
    const { session, state, document } = pipeline()
    const view = buildGanttView(
      document,
      projectDocumentView(document, session.schedule, state),
      state,
      { firstRow: 0, visibleRows: 3, overscan: 0 },
      { workingTime: canonicalQuery },
    )
    const tuple = calendarTuple(view)
    expect(tuple.background?.calendarId).toBe(cal('standard'))
    expect(tuple.rowCalendars).toHaveLength(3)
    expect(tuple.rowCalendars?.[0]).toMatchObject({
      rowIndex: 0,
      taskId: id('root'),
      calendarId: cal('standard'),
    })
    expect(tuple.rowCalendars?.[1]).toMatchObject({
      rowIndex: 1,
      taskId: id('plain'),
      calendarId: cal('standard'),
    })
    expect(tuple.rowCalendars?.[2]).toMatchObject({
      rowIndex: 2,
      taskId: id('dense'),
      calendarId: cal('compact'),
    })
    // The viewport is the project window padded to the one-day minimum
    // (Mon 08:00 → Tue 08:00): the compact row shades Mon 08:00–18:00
    // working then the night; the standard background shades 09:00–17:00
    // working inside it — per-row calendars visibly differ.
    expect(tuple.rowCalendars?.[2]?.bands).toEqual([
      { start: '2026-08-03T08:00:00.000Z', finish: '2026-08-03T18:00:00.000Z', working: true },
      { start: '2026-08-03T18:00:00.000Z', finish: '2026-08-04T08:00:00.000Z', working: false },
    ])
    expect(tuple.background?.bands).toEqual([
      { start: '2026-08-03T08:00:00.000Z', finish: '2026-08-03T09:00:00.000Z', working: false },
      { start: '2026-08-03T09:00:00.000Z', finish: '2026-08-03T17:00:00.000Z', working: true },
      { start: '2026-08-03T17:00:00.000Z', finish: '2026-08-04T08:00:00.000Z', working: false },
    ])
    // The authority scheduled `dense` inside its resolved calendar.
    expect(session.schedule?.taskSchedules[id('dense')]?.resolvedCalendarId).toBe(cal('compact'))
    expect(session.schedule?.taskSchedules[id('plain')]?.resolvedCalendarId).toBe(cal('standard'))
    expect(session.schedule?.taskSchedules[id('dense')]?.scheduledFinish).toBe(
      '2026-08-03T18:00:00.000Z',
    )
    expect(state.viewport.start).toBe('2026-08-03T08:00:00.000Z')
  })

  it('C08 — no query: surfaces absent, geometry byte-identical to the shaded build', () => {
    const { session, state, document } = pipeline()
    const projection = projectDocumentView(document, session.schedule, state)
    const layout = { firstRow: 0, visibleRows: 3, overscan: 0 }
    const bare = buildGanttView(document, projection, state, layout)
    const shaded = buildGanttView(document, projection, state, layout, {
      workingTime: canonicalQuery,
    })
    expect(bare.timeline.calendar).toBeUndefined()
    expect(bare.timeline.rowCalendars).toBeUndefined()
    expect(JSON.stringify(shaded.timeline.bands)).toBe(JSON.stringify(bare.timeline.bands))
    expect(JSON.stringify(shaded.timeline.bars)).toBe(JSON.stringify(bare.timeline.bars))
    expect(JSON.stringify(shaded.timeline.milestones)).toBe(
      JSON.stringify(bare.timeline.milestones),
    )
    expect(JSON.stringify(shaded.timeline.links)).toBe(JSON.stringify(bare.timeline.links))
  })

  it('C09 — unresolvable default calendar: the surface echoes the authority diagnostic', () => {
    const base = makeDocument({
      tasks: [makeTask({ id: 'solo', uid: 1, outlineLevel: 1, wbs: '1' })],
    })
    const document: ProjectDocument = {
      ...base,
      properties: { ...base.properties, defaultCalendarId: cal('ghost') },
    }
    const derived = schedule(document)
    expect(derived.diagnostics.map((diagnostic) => diagnostic.code)).toContain('MISSING_CALENDAR')
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)
    const view = buildGanttView(
      document,
      projectDocumentView(document, session.schedule, state),
      state,
      { firstRow: 0, visibleRows: 1, overscan: 0 },
      { workingTime: canonicalQuery },
    )
    expect(view.timeline.calendar?.status).toBe('unresolvable')
    expect(view.timeline.calendar?.diagnostic?.code).toBe('MISSING_CALENDAR')
    expect('bands' in (view.timeline.calendar ?? {})).toBe(false)
    expect(view.timeline.rowCalendars).toEqual([])
  })

  it('C10 — viewport zoom: bands re-project clipped to the new window', () => {
    const document = makeDocument({ calendars: [makeCalendar('standard')] })
    const day = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-04T00:00:00.000Z' }
    const daySurface = buildCalendarSurface(document, canonicalQuery, day)
    expect(daySurface.bands).toHaveLength(3) // night, 09–17, night
    const zoomed = scaleViewport(day, 0.5) // Mon 06:00 → Mon 18:00
    const zoomedSurface = buildCalendarSurface(document, canonicalQuery, zoomed)
    expect(zoomedSurface.start).toBe(zoomed.start)
    expect(zoomedSurface.finish).toBe(zoomed.finish)
    expect(zoomedSurface.bands).toEqual([
      { start: '2026-08-03T06:00:00.000Z', finish: '2026-08-03T09:00:00.000Z', working: false },
      { start: '2026-08-03T09:00:00.000Z', finish: '2026-08-03T17:00:00.000Z', working: true },
      { start: '2026-08-03T17:00:00.000Z', finish: '2026-08-03T18:00:00.000Z', working: false },
    ])
    // The canonical evaluator answers the same question identically.
    expect(canonicalWorking(document, 'standard', zoomed)).toEqual([
      { start: '2026-08-03T09:00:00.000Z', finish: '2026-08-03T17:00:00.000Z' },
    ])
  })

  it('C11 — session commands move the schedule; the calendar surfaces stay stable', () => {
    const { session, state, document } = pipeline()
    const layout = { firstRow: 0, visibleRows: 3, overscan: 0 }
    const buildFor = (doc: ProjectDocument, derived: typeof session.schedule) =>
      calendarTuple(
        buildGanttView(doc, projectDocumentView(doc, derived, state), state, layout, {
          workingTime: canonicalQuery,
        }),
      )
    const before = buildFor(document, session.schedule)
    const command: ProjectCommand = {
      type: 'SetTaskDuration',
      taskId: id('dense'),
      duration: asWorkingMinutes(1080),
    }
    const applied = applyRendererCommand(session, command)
    expect(applied.result.accepted).toBe(true)
    // The schedule moved (dense now spills into Tuesday: 600 min Monday
    // + 480 min Tuesday → Tue 16:00)…
    expect(applied.session.schedule?.taskSchedules[id('dense')]?.scheduledFinish).toBe(
      '2026-08-04T16:00:00.000Z',
    )
    // …but the state's viewport (and the calendars) are untouched, so the
    // calendar surfaces are byte-identical across the revision.
    expect(JSON.stringify(buildFor(applied.session.document, applied.session.schedule))).toBe(
      JSON.stringify(before),
    )
    // Undo restores the exact prior schedule; surfaces stay identical.
    const undone = undoRendererCommand(applied.session)
    expect(undone.session.schedule?.taskSchedules[id('dense')]?.scheduledFinish).toBe(
      '2026-08-03T18:00:00.000Z',
    )
    expect(JSON.stringify(buildFor(undone.session.document, undone.session.schedule))).toBe(
      JSON.stringify(before),
    )
    // Redo replays forward to the same tuple as the applied step.
    const redone = redoRendererCommand(undone.session)
    expect(JSON.stringify(buildFor(redone.session.document, redone.session.schedule))).toBe(
      JSON.stringify(buildFor(applied.session.document, applied.session.schedule)),
    )
  })

  it('C12 — mixed command/viewport sequence replays byte-identical (3× determinism)', () => {
    const run = () => {
      const { document, session, state } = pipeline()
      let currentSession = session
      let currentState = state
      const steps: string[] = []
      const build = () =>
        JSON.stringify(
          calendarTuple(
            buildGanttView(
              currentSession.document,
              projectDocumentView(currentSession.document, currentSession.schedule, currentState),
              currentState,
              { firstRow: 0, visibleRows: 3, overscan: 0 },
              { workingTime: canonicalQuery },
            ),
          ),
        )
      steps.push(build())
      const command: ProjectCommand = {
        type: 'SetTaskDuration',
        taskId: id('plain'),
        duration: asWorkingMinutes(960),
      }
      const applied = applyRendererCommand(currentSession, command)
      expect(applied.result.accepted).toBe(true)
      currentSession = applied.session
      steps.push(build())
      currentState = { ...currentState, viewport: scaleViewport(currentState.viewport, 2) }
      steps.push(build())
      const undone = undoRendererCommand(currentSession)
      currentSession = undone.session
      steps.push(build())
      return JSON.stringify(steps)
    }
    const first = run()
    expect(run()).toBe(first)
    expect(run()).toBe(first)
  })
})
