/**
 * PROJECT-026 — the critical-path projection layer (unit battery).
 *
 * Pure-projection semantics only — NO scheduling package anywhere in this
 * file (the schedules are hand-authored fixtures; the REAL scheduler's
 * criticality/float equality is exercised in
 * `critical-resource-integration.test.ts` and the golden battery, the
 * accepted test-layer precedent): the verbatim echo contract (every float
 * value equals the authority's own TaskSchedule values — the "visual layers
 * match derived schedule" acceptance evidence), the slack-bar geometry
 * (clamping, edge flags, absent for zero/inverted spans), the
 * both-endpoints critical-link classification, the windowing and
 * degradation contracts, and determinism/purity.
 */
import { describe, expect, it } from 'vitest'
import type { DerivedSchedule, ProjectDocument } from '@genoffice/project-contracts'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import {
  buildCriticalPath,
  buildTimeline,
  createViewState,
  projectDocumentView,
} from '../src/index.js'
import { makeDependency, makeDocument, makeScheduleEntry, makeTask } from './fixtures.js'

const VIEWPORT = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-04T00:00:00.000Z' } // Monday

/** The critical-path fixture document: a critical task, a floating task, a
 * negative-slack task, an unscheduled task, and a task without a late
 * finish — plus the dependency mix for the link classification. */
const criticalDocument = (): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({ id: 'crit', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'float', duration: asWorkingMinutes(180) }),
      makeTask({ id: 'neg', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'unsched', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'nolate', duration: asWorkingMinutes(480) }),
      makeTask({ id: 'snapped', duration: asWorkingMinutes(480) }),
    ],
    dependencies: [
      makeDependency('d1', 'crit', 'float'), // one critical endpoint — NOT critical
      makeDependency('d2', 'crit', 'neg'), // both critical — classified
      makeDependency('d3', 'unsched', 'crit'), // unscheduled endpoint — NOT classified
      makeDependency('d4', 'nolate', 'float'), // nolate is not critical — NOT classified
      makeDependency('d5', 'crit', 'snapped'), // both critical — classified
    ],
  })

const criticalSchedule = (): DerivedSchedule => ({
  taskSchedules: {
    [asTaskId('crit')]: makeScheduleEntry(
      'crit',
      '2026-08-03T09:00:00.000Z',
      '2026-08-03T17:00:00.000Z',
      {
        critical: true,
        totalSlack: 0,
        freeSlack: 0,
        lateFinish: asISODateTime('2026-08-03T17:00:00.000Z'), // zero slack span — no geometry
      },
    ),
    [asTaskId('float')]: makeScheduleEntry(
      'float',
      '2026-08-03T09:00:00.000Z',
      '2026-08-03T12:00:00.000Z',
      {
        critical: false,
        totalSlack: 360,
        freeSlack: 120,
        lateFinish: asISODateTime('2026-08-03T18:00:00.000Z'), // slack 12:00 → 18:00
      },
    ),
    [asTaskId('neg')]: makeScheduleEntry(
      'neg',
      '2026-08-03T09:00:00.000Z',
      '2026-08-03T17:00:00.000Z',
      {
        critical: true,
        totalSlack: -240,
        freeSlack: 0,
        lateFinish: asISODateTime('2026-08-03T13:00:00.000Z'), // INVERTED (negative slack) — no geometry
      },
    ),
    // 'unsched' deliberately has no schedule entry.
    [asTaskId('nolate')]: makeScheduleEntry(
      'nolate',
      '2026-08-03T09:00:00.000Z',
      '2026-08-03T17:00:00.000Z',
      {
        critical: false,
        totalSlack: 480,
        freeSlack: 480,
        // no lateFinish — no slack geometry
      },
    ),
    // The snapped-late-date case: ZERO total slack while the authority's
    // lateFinish sits at the NEXT working instant (a wall-clock-later
    // window with zero WORKING time inside it). No float → no geometry.
    [asTaskId('snapped')]: makeScheduleEntry(
      'snapped',
      '2026-08-03T09:00:00.000Z',
      '2026-08-03T17:00:00.000Z',
      {
        critical: true,
        totalSlack: 0,
        freeSlack: 0,
        lateFinish: asISODateTime('2026-08-04T09:00:00.000Z'),
      },
    ),
  },
  diagnostics: [],
})

describe('PROJECT-026 buildCriticalPath — the verbatim canonical echo', () => {
  it("echoes the authority's critical/float values verbatim (never recomputed, never clamped)", () => {
    const document = criticalDocument()
    const schedule = criticalSchedule()
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const surface = buildCriticalPath(document, projection, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 5,
    })
    // One float per SCHEDULED row (unsched carries none — never invented).
    expect(surface.floats.map((float) => float.taskId as string)).toEqual([
      'crit',
      'float',
      'neg',
      'nolate',
      'snapped',
    ])
    expect(surface.floats.map((float) => float.rowIndex)).toEqual([0, 1, 2, 4, 5])
    // The echoes equal the authority's own values, byte for byte.
    for (const float of surface.floats) {
      const authoritative = schedule.taskSchedules[float.taskId]
      expect(float.critical).toBe(authoritative?.critical)
      expect(float.totalSlack).toBe(authoritative?.totalSlack)
      expect(float.freeSlack).toBe(authoritative?.freeSlack)
    }
    expect(surface.floats[0]).toMatchObject({ critical: true, totalSlack: 0, freeSlack: 0 })
    expect(surface.floats[1]).toMatchObject({ critical: false, totalSlack: 360, freeSlack: 120 })
    // Negative slack is echoed NEGATIVE — never clamped, never interpreted.
    expect(surface.floats[2]).toMatchObject({ critical: true, totalSlack: -240 })
  })

  it('projects the slack geometry over the viewport (clamped, edge flags)', () => {
    const document = criticalDocument()
    const projection = projectDocumentView(document, criticalSchedule(), createViewState(document))
    const surface = buildCriticalPath(document, projection, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 5,
    })
    const floats = new Map(surface.floats.map((float) => [float.taskId as string, float]))
    // float: slack window 12:00 → 18:00 over the Monday viewport.
    expect(floats.get('float')?.slack).toEqual({
      startFraction: 0.5,
      finishFraction: 0.75,
      startsBefore: false,
      finishesAfter: false,
    })
    // Zero-slack critical (including the SNAPPED late date — a
    // wall-clock-later window with zero working time inside), inverted
    // (negative), and late-finish-less rows carry NO slack geometry —
    // never invented.
    expect(floats.get('crit')?.slack).toBeUndefined()
    expect(floats.get('snapped')?.slack).toBeUndefined()
    expect(floats.get('neg')?.slack).toBeUndefined()
    expect(floats.get('nolate')?.slack).toBeUndefined()
  })

  it('flags slack windows extending beyond the viewport on either side', () => {
    // A task finishing BEFORE the viewport with its late finish inside, and
    // a task whose late finish lands AFTER the viewport.
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'early', duration: asWorkingMinutes(480) }),
        makeTask({ id: 'late', duration: asWorkingMinutes(480) }),
      ],
    })
    const schedule: DerivedSchedule = {
      taskSchedules: {
        [asTaskId('early')]: makeScheduleEntry(
          'early',
          '2026-07-31T09:00:00.000Z',
          '2026-07-31T17:00:00.000Z',
          {
            critical: false,
            totalSlack: 3000,
            freeSlack: 3000,
            lateFinish: asISODateTime('2026-08-03T12:00:00.000Z'),
          },
        ),
        [asTaskId('late')]: makeScheduleEntry(
          'late',
          '2026-08-03T09:00:00.000Z',
          '2026-08-03T17:00:00.000Z',
          {
            critical: false,
            totalSlack: 3000,
            freeSlack: 3000,
            lateFinish: asISODateTime('2026-08-04T18:00:00.000Z'),
          },
        ),
      },
      diagnostics: [],
    }
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const surface = buildCriticalPath(document, projection, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 1,
    })
    const floats = new Map(surface.floats.map((float) => [float.taskId as string, float]))
    // early: the window [Fri 17:00, Mon 12:00) starts before the viewport.
    expect(floats.get('early')?.slack).toEqual({
      startFraction: 0,
      finishFraction: 0.5,
      startsBefore: true,
      finishesAfter: false,
    })
    // late: the window [Mon 17:00, Tue 18:00) finishes after the viewport.
    expect(floats.get('late')?.slack).toEqual({
      startFraction: 17 / 24,
      finishFraction: 1,
      startsBefore: false,
      finishesAfter: true,
    })
  })

  it('classifies a link critical iff BOTH canonical endpoint tasks are critical (the projection convention)', () => {
    const document = criticalDocument()
    const projection = projectDocumentView(document, criticalSchedule(), createViewState(document))
    const surface = buildCriticalPath(document, projection, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 5,
    })
    // d2 (crit → neg) and d5 (crit → snapped) are the both-endpoints-critical
    // links; d1 has a non-critical endpoint, d3 an unscheduled one, d4 a
    // non-critical one. Document order preserved.
    expect(surface.criticalDependencyIds.map((id) => id as string)).toEqual(['d2', 'd5'])
  })

  it('scopes floats to the in-window rows', () => {
    const document = criticalDocument()
    const projection = projectDocumentView(document, criticalSchedule(), createViewState(document))
    // Window covers only rows 0–2 (crit, float, neg).
    const surface = buildCriticalPath(document, projection, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 2,
    })
    expect(surface.floats.map((float) => float.taskId as string)).toEqual(['crit', 'float', 'neg'])
    // The link classification is not windowed — it is the document-order
    // classification over the projection's schedule join.
    expect(surface.criticalDependencyIds.map((id) => id as string)).toEqual(['d2', 'd5'])
  })

  it('yields an EMPTY surface for a degenerate viewport (never invented values)', () => {
    const document = criticalDocument()
    const projection = projectDocumentView(document, criticalSchedule(), createViewState(document))
    const surface = buildCriticalPath(
      document,
      projection,
      { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-03T00:00:00.000Z' },
      { firstIndex: 0, lastIndex: 4 },
    )
    expect(surface).toEqual({ floats: [], criticalDependencyIds: [] })
    const unparseable = buildCriticalPath(
      document,
      projection,
      { start: 'not-a-date', finish: '2026-08-04T00:00:00.000Z' },
      { firstIndex: 0, lastIndex: 4 },
    )
    expect(unparseable).toEqual({ floats: [], criticalDependencyIds: [] })
  })

  it('is deterministic (3× byte-identical) and never mutates its inputs', () => {
    const document = criticalDocument()
    const schedule = criticalSchedule()
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const documentBefore = JSON.stringify(document)
    const scheduleBefore = JSON.stringify(schedule)
    const first = JSON.stringify(
      buildCriticalPath(document, projection, VIEWPORT, { firstIndex: 0, lastIndex: 4 }),
    )
    for (let i = 0; i < 2; i += 1) {
      expect(
        JSON.stringify(
          buildCriticalPath(document, projection, VIEWPORT, { firstIndex: 0, lastIndex: 4 }),
        ),
      ).toBe(first)
    }
    expect(JSON.stringify(document)).toBe(documentBefore)
    expect(JSON.stringify(schedule)).toBe(scheduleBefore)
  })
})

describe('PROJECT-026 timeline integration — the additive critical-path surface', () => {
  it('joins the critical-path surface iff the projection carries a schedule', () => {
    const document = criticalDocument()
    const schedule = criticalSchedule()
    const state = createViewState(document)
    const withSchedule = projectDocumentView(document, schedule, state)
    const timeline = buildTimeline(document, withSchedule, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 5,
    })
    expect(timeline.criticalPath).toBeDefined()
    // The joined surface equals the standalone projection's value.
    expect(timeline.criticalPath).toEqual(
      buildCriticalPath(document, withSchedule, VIEWPORT, { firstIndex: 0, lastIndex: 5 }),
    )
    // Without a schedule there is no criticality to project — never invented.
    const withoutSchedule = projectDocumentView(document, undefined, state)
    const bare = buildTimeline(document, withoutSchedule, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 5,
    })
    expect(bare.criticalPath).toBeUndefined()
    // No bars without a schedule (dates are never invented), while the
    // scheduled timeline carries real geometry.
    expect(bare.bars).toEqual([])
    expect(timeline.bars.length).toBeGreaterThan(0)
  })

  it('carries an EMPTY critical-path surface when the schedule joined but no window row is scheduled', () => {
    const document = makeDocument({ tasks: [makeTask({ id: 'only' })] })
    const schedule: DerivedSchedule = { taskSchedules: {}, diagnostics: [] }
    const projection = projectDocumentView(document, schedule, createViewState(document))
    const timeline = buildTimeline(document, projection, VIEWPORT, {
      firstIndex: 0,
      lastIndex: 0,
    })
    expect(timeline.criticalPath).toEqual({ floats: [], criticalDependencyIds: [] })
  })
})
