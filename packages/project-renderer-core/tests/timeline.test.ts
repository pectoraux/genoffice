import { describe, expect, it } from 'vitest'
import { asISODateTime } from '@genoffice/project-contracts'
import {
  MIN_VIEWPORT_SPAN_MS,
  MAX_VIEWPORT_SPAN_MS,
  axisLevelForSpan,
  buildTimeAxis,
  fitViewport,
  projectWindow,
  scaleViewport,
  viewportFraction,
  viewportInstant,
} from '../src/index.js'
import { makeDocument } from './fixtures.js'

const DAY_MS = 24 * 60 * 60 * 1000

describe('PROJECT-021 timeline — fraction mapping', () => {
  const viewport = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-11T00:00:00.000Z' }

  it('maps window edges to 0 and 1 and inner instants linearly', () => {
    expect(viewportFraction(viewport, '2026-08-01T00:00:00.000Z')).toBe(0)
    expect(viewportFraction(viewport, '2026-08-11T00:00:00.000Z')).toBe(1)
    expect(viewportFraction(viewport, '2026-08-06T00:00:00.000Z')).toBe(0.5)
  })

  it('clamps out-of-window instants to [0, 1]', () => {
    expect(viewportFraction(viewport, '2026-07-01T00:00:00.000Z')).toBe(0)
    expect(viewportFraction(viewport, '2026-12-01T00:00:00.000Z')).toBe(1)
  })

  it('viewportInstant is the inverse of viewportFraction', () => {
    expect(viewportInstant(viewport, 0)).toBe('2026-08-01T00:00:00.000Z')
    expect(viewportInstant(viewport, 0.5)).toBe('2026-08-06T00:00:00.000Z')
    expect(viewportInstant(viewport, 1)).toBe('2026-08-11T00:00:00.000Z')
    expect(viewportInstant(viewport, 2)).toBe('2026-08-11T00:00:00.000Z')
  })
})

describe('PROJECT-021 timeline — viewport scaling', () => {
  it('halving and doubling round-trips the window exactly', () => {
    const viewport = { start: '2026-08-01T12:00:00.000Z', finish: '2026-08-31T12:00:00.000Z' }
    const halved = scaleViewport(viewport, 0.5)
    const restored = scaleViewport(halved, 2)
    expect(restored).toEqual(viewport)
  })

  it('clamps the span to the documented minimum of one minute', () => {
    const viewport = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-01T00:10:00.000Z' }
    const zoomed = scaleViewport(viewport, 0.000001)
    expect(Date.parse(zoomed.finish) - Date.parse(zoomed.start)).toBe(MIN_VIEWPORT_SPAN_MS)
  })

  it('clamps the span to the documented maximum of 100 years', () => {
    const viewport = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-02T00:00:00.000Z' }
    const zoomedOut = scaleViewport(viewport, 1e9)
    expect(Date.parse(zoomedOut.finish) - Date.parse(zoomedOut.start)).toBe(MAX_VIEWPORT_SPAN_MS)
  })

  it('is a no-op for invalid factors and unparseable windows', () => {
    const viewport = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-02T00:00:00.000Z' }
    expect(scaleViewport(viewport, 0)).toBe(viewport)
    expect(scaleViewport(viewport, -1)).toBe(viewport)
    expect(scaleViewport({ start: 'x', finish: 'y' }, 2)).toEqual({ start: 'x', finish: 'y' })
  })
})

describe('PROJECT-021 timeline — project window and fit', () => {
  it('prefers the derived schedule window, then properties, then a 30-day default', () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      finishDate: '2026-08-20T09:00:00.000Z',
    })
    const fromProperties = projectWindow(document)
    expect(fromProperties.start).toBe(Date.parse('2026-08-03T09:00:00.000Z'))
    expect(fromProperties.finish).toBe(Date.parse('2026-08-20T09:00:00.000Z'))
    const schedule = {
      taskSchedules: {},
      projectStart: asISODateTime('2026-08-05T09:00:00.000Z'),
      projectFinish: asISODateTime('2026-08-10T09:00:00.000Z'),
      diagnostics: [],
    }
    const fromSchedule = projectWindow(document, schedule)
    expect(fromSchedule.start).toBe(Date.parse('2026-08-05T09:00:00.000Z'))
    expect(fromSchedule.finish).toBe(Date.parse('2026-08-10T09:00:00.000Z'))
    const bare = projectWindow(makeDocument({ startDate: '2026-08-03T09:00:00.000Z' }))
    expect(bare.finish - bare.start).toBe(30 * DAY_MS)
  })

  it('fitViewport pads 2% each side with a one-day minimum total span', () => {
    const document = makeDocument({ startDate: '2026-08-10T09:00:00.000Z' })
    const schedule = {
      taskSchedules: {},
      projectStart: asISODateTime('2026-08-10T09:00:00.000Z'),
      projectFinish: asISODateTime('2026-08-10T09:00:00.000Z'), // milestone-only: zero span
      diagnostics: [],
    }
    const fitted = fitViewport(document, schedule)
    expect(Date.parse(fitted.finish) - Date.parse(fitted.start)).toBe(DAY_MS)
    const normal = fitViewport(makeDocument({ startDate: '2026-08-03T09:00:00.000Z' }), {
      taskSchedules: {},
      projectStart: asISODateTime('2026-08-03T09:00:00.000Z'),
      projectFinish: asISODateTime('2026-08-13T09:00:00.000Z'),
      diagnostics: [],
    })
    const span = Date.parse(normal.finish) - Date.parse(normal.start)
    expect(span).toBe(10 * DAY_MS + Math.round(10 * DAY_MS * 0.02) * 2)
  })
})

describe('PROJECT-021 timeline — time axis', () => {
  it('chooses the band level deterministically from the span', () => {
    expect(axisLevelForSpan(30 * DAY_MS)).toBe('day')
    expect(axisLevelForSpan(92 * DAY_MS)).toBe('day')
    expect(axisLevelForSpan(93 * DAY_MS)).toBe('week')
    expect(axisLevelForSpan(2 * 366 * DAY_MS)).toBe('week')
    expect(axisLevelForSpan(3 * 366 * DAY_MS)).toBe('month')
  })

  it('generates contiguous UTC-day bands clipped to the window', () => {
    const viewport = { start: '2026-08-01T06:00:00.000Z', finish: '2026-08-04T18:00:00.000Z' }
    const bands = buildTimeAxis(viewport, 'day')
    expect(bands.map((band) => band.level)).toEqual(['day', 'day', 'day', 'day'])
    expect(bands[0]!.start).toBe('2026-08-01T06:00:00.000Z')
    expect(bands[0]!.finish).toBe('2026-08-02T00:00:00.000Z')
    expect(bands[1]!.start).toBe('2026-08-02T00:00:00.000Z')
    expect(bands[3]!.start).toBe('2026-08-04T00:00:00.000Z')
    expect(bands[3]!.finish).toBe('2026-08-04T18:00:00.000Z')
    // Contiguity: every band starts exactly where the previous finished.
    for (let index = 1; index < bands.length; index += 1) {
      expect(bands[index]!.start).toBe(bands[index - 1]!.finish)
    }
  })

  it('generates Monday-aligned week bands', () => {
    // 2026-08-03 is a Monday; the window starts the preceding Saturday.
    const viewport = { start: '2026-08-01T00:00:00.000Z', finish: '2026-08-21T00:00:00.000Z' }
    const bands = buildTimeAxis(viewport, 'week')
    expect(bands.map((band) => band.start)).toEqual([
      '2026-08-01T00:00:00.000Z',
      '2026-08-03T00:00:00.000Z',
      '2026-08-10T00:00:00.000Z',
      '2026-08-17T00:00:00.000Z',
    ])
    expect(bands[1]!.finish).toBe('2026-08-10T00:00:00.000Z')
  })

  it('generates calendar-month bands', () => {
    const viewport = { start: '2026-08-15T00:00:00.000Z', finish: '2026-10-05T00:00:00.000Z' }
    const bands = buildTimeAxis(viewport, 'month')
    expect(bands.map((band) => band.start)).toEqual([
      '2026-08-15T00:00:00.000Z',
      '2026-09-01T00:00:00.000Z',
      '2026-10-01T00:00:00.000Z',
    ])
    expect(bands[1]!.finish).toBe('2026-10-01T00:00:00.000Z')
  })

  it('returns no bands for empty/degenerate windows and is deterministic across runs', () => {
    expect(buildTimeAxis({ start: 'x', finish: 'y' })).toEqual([])
    expect(
      buildTimeAxis({ start: '2026-08-01T00:00:00.000Z', finish: '2026-08-01T00:00:00.000Z' }),
    ).toEqual([])
    const viewport = { start: '2026-01-01T00:00:00.000Z', finish: '2026-04-01T00:00:00.000Z' }
    expect(JSON.stringify(buildTimeAxis(viewport))).toBe(JSON.stringify(buildTimeAxis(viewport)))
    expect(JSON.stringify(buildTimeAxis(viewport))).toBe(JSON.stringify(buildTimeAxis(viewport)))
  })

  it('uses the span-derived level when none is given (a year yields week bands; four years yield month bands)', () => {
    const yearViewport = { start: '2026-01-01T00:00:00.000Z', finish: '2027-01-01T00:00:00.000Z' }
    const yearBands = buildTimeAxis(yearViewport)
    expect(yearBands.every((band) => band.level === 'week')).toBe(true)
    expect(yearBands).toHaveLength(53) // 52 whole weeks + the clipped first band
    const multiYearViewport = {
      start: '2026-01-01T00:00:00.000Z',
      finish: '2030-01-01T00:00:00.000Z',
    }
    const monthBands = buildTimeAxis(multiYearViewport)
    expect(monthBands.every((band) => band.level === 'month')).toBe(true)
    expect(monthBands).toHaveLength(48)
  })
})
