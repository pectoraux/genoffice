/**
 * Unit tests — /office/workbooks/save request validation for the
 * chartEdits family and chart visualAdditions (EXCEL-023 — Charts).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical chart payloads
 * (chart additions with a supported chart type, 1-24 bounded series, and
 * bounded style options; chart edits keyed by canonical xl/charts paths
 * carrying at least one bounded property change) and rejects malformed
 * chart types, invalid locators, bad ranges, unknown fields, out-of-bounds
 * percentages, empty series sets, and excessive counts with 400s — nothing
 * unvalidated reaches the engine.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(
  savePlanExtras: Record<string, unknown>,
): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: { edits: [], ...savePlanExtras },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

const ANCHOR = {
  fromRow: 2,
  fromColumn: 1,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRow: 12,
  toColumn: 6,
  toRowOffset: 0,
  toColumnOffset: 0,
}

function canonicalChartAddition(): Record<string, unknown> & {
  chart: Record<string, unknown>
} {
  return {
    sheetName: 'Data',
    anchor: { ...ANCHOR },
    chart: {
      chartType: 'column',
      title: 'Sales',
      series: [
        {
          name: 'Revenue',
          categories: ['Q1', 'Q2'],
          values: [10, 20],
          valuesRef: 'Data!$B$2:$B$3',
          categoriesRef: 'Data!$A$2:$A$3',
          color: '#4472C4',
        },
      ],
      legend: 'bottom',
      dataLabels: 'value',
      gapWidthPct: 150,
    },
  }
}

function canonicalChartEdit(): Record<string, unknown> {
  return {
    chartPath: 'xl/charts/chart1.xml',
    title: 'Renamed',
  }
}

describe('workbooks/save chartAdditions (visualAdditions.chart) validation', () => {
  it('accepts a canonical chart addition', async () => {
    const res = await save({ visualAdditions: [canonicalChartAddition()] })
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts every addable chart type', async () => {
    for (const chartType of [
      'column',
      'bar',
      'line',
      'area',
      'pie',
      'scatter',
      'radar',
      'doughnut',
      'combo',
    ]) {
      const res = await save({
        visualAdditions: [
          { ...canonicalChartAddition(), chart: { ...canonicalChartAddition().chart, chartType } },
        ],
      })
      expect(res.body.error).toBe('malformed')
    }
  })

  it('rejects an unsupported chart type', async () => {
    const res = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: { ...canonicalChartAddition().chart, chartType: 'bubble' },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('chartType')
  })

  it('rejects an empty or oversized series set', async () => {
    const empty = await save({
      visualAdditions: [
        { ...canonicalChartAddition(), chart: { ...canonicalChartAddition().chart, series: [] } },
      ],
    })
    expect(empty.status).toBe(400)
    expect(empty.body.message).toContain('series must not be empty')

    const many = Array.from({ length: 25 }, () => ({
      name: 'S',
      categories: ['a'],
      values: [1],
    }))
    const oversized = await save({
      visualAdditions: [
        { ...canonicalChartAddition(), chart: { ...canonicalChartAddition().chart, series: many } },
      ],
    })
    expect(oversized.status).toBe(400)
    expect(oversized.body.message).toContain('series exceeds')
  })

  it('rejects non-finite values and out-of-bounds series fields', async () => {
    const badValues = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: {
            ...canonicalChartAddition().chart,
            series: [{ name: 'S', categories: ['a'], values: [Number.NaN] }],
          },
        },
      ],
    })
    expect(badValues.status).toBe(400)
    expect(badValues.body.message).toContain('finite number')

    const badColor = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: {
            ...canonicalChartAddition().chart,
            series: [{ name: 'S', categories: ['a'], values: [1], color: '4472C4' }],
          },
        },
      ],
    })
    expect(badColor.status).toBe(400)
    expect(badColor.body.message).toContain('#RRGGBB')

    const badPointKey = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: {
            ...canonicalChartAddition().chart,
            series: [
              { name: 'S', categories: ['a'], values: [1], pointColors: { 1000: '#FF0000' } },
            ],
          },
        },
      ],
    })
    expect(badPointKey.status).toBe(400)
    expect(badPointKey.body.message).toContain('index')
  })

  it('rejects out-of-bounds style options', async () => {
    const gap = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: { ...canonicalChartAddition().chart, gapWidthPct: 501 },
        },
      ],
    })
    expect(gap.status).toBe(400)
    expect(gap.body.message).toContain('gapWidthPct')

    const hole = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: { ...canonicalChartAddition().chart, holeSizePct: 5 },
        },
      ],
    })
    expect(hole.status).toBe(400)
    expect(hole.body.message).toContain('holeSizePct')

    const legend = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: { ...canonicalChartAddition().chart, legend: 'diagonal' },
        },
      ],
    })
    expect(legend.status).toBe(400)
    expect(legend.body.message).toContain('legend')
  })

  it('rejects unknown fields on the chart and its nested objects', async () => {
    const unknown = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: { ...canonicalChartAddition().chart, trendlines: true },
        },
      ],
    })
    expect(unknown.status).toBe(400)
    expect(unknown.body.message).toContain('unknown field')

    const unknownSeriesField = await save({
      visualAdditions: [
        {
          ...canonicalChartAddition(),
          chart: {
            ...canonicalChartAddition().chart,
            series: [{ name: 'S', categories: ['a'], values: [1], marker: 'circle' }],
          },
        },
      ],
    })
    expect(unknownSeriesField.status).toBe(400)
    expect(unknownSeriesField.body.message).toContain('unknown field')
  })
})

describe('workbooks/save chartEdits validation', () => {
  it('accepts a canonical chart edit', async () => {
    const res = await save({ chartEdits: [canonicalChartEdit()] })
    expect(res.body.error).toBe('malformed')
  })

  it('accepts every convertible chart type', async () => {
    for (const chartType of ['column', 'bar', 'line', 'area', 'pie', 'doughnut']) {
      const res = await save({
        chartEdits: [{ ...canonicalChartEdit(), chartType }],
      })
      expect(res.body.error).toBe('malformed')
    }
  })

  it('rejects scatter/combo type edits (not convertible)', async () => {
    for (const chartType of ['scatter', 'combo', 'bubble']) {
      const res = await save({ chartEdits: [{ ...canonicalChartEdit(), chartType }] })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('chartType')
    }
  })

  it('rejects chart paths outside xl/charts', async () => {
    for (const chartPath of [
      'xl/drawings/drawing1.xml',
      'xl/charts/chart1.xml.exe',
      '../charts/chart1.xml',
      'xl/charts/../charts/chart1.xml',
    ]) {
      const res = await save({ chartEdits: [{ ...canonicalChartEdit(), chartPath }] })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('chartPath')
    }
  })

  it('rejects an edit with no property change', async () => {
    const res = await save({ chartEdits: [{ chartPath: 'xl/charts/chart1.xml' }] })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('at least one property')
  })

  it('rejects malformed series sets and series edits', async () => {
    const emptySet = await save({
      chartEdits: [{ ...canonicalChartEdit(), seriesSet: [] }],
    })
    expect(emptySet.status).toBe(400)
    expect(emptySet.body.message).toContain('seriesSet must not be empty')

    const badSeries = await save({
      chartEdits: [{ ...canonicalChartEdit(), series: [{ index: 0 }] }],
    })
    expect(badSeries.status).toBe(400)
    expect(badSeries.body.message).toContain('needs a name or data')

    const missingIndex = await save({
      chartEdits: [{ ...canonicalChartEdit(), series: [{ name: 'S' }] }],
    })
    expect(missingIndex.status).toBe(400)
    expect(missingIndex.body.message).toContain('index')
  })

  it('rejects malformed source ranges', async () => {
    const longRef = await save({
      chartEdits: [
        {
          ...canonicalChartEdit(),
          seriesSet: [
            {
              name: 'S',
              values: [1],
              valuesRef: 'x'.repeat(513),
            },
          ],
        },
      ],
    })
    expect(longRef.status).toBe(400)
    expect(longRef.body.message).toContain('512')
  })

  it('rejects out-of-bounds numeric options', async () => {
    const badGap = await save({
      chartEdits: [{ ...canonicalChartEdit(), gapWidthPct: -1 }],
    })
    expect(badGap.status).toBe(400)
    expect(badGap.body.message).toContain('gapWidthPct')

    const badHole = await save({
      chartEdits: [{ ...canonicalChartEdit(), holeSizePct: 91 }],
    })
    expect(badHole.status).toBe(400)
    expect(badHole.body.message).toContain('holeSizePct')

    const badExplosion = await save({
      chartEdits: [{ ...canonicalChartEdit(), explosionPct: 401 }],
    })
    expect(badExplosion.status).toBe(400)
    expect(badExplosion.body.message).toContain('explosionPct')

    const emptyAxis = await save({
      chartEdits: [{ ...canonicalChartEdit(), valueAxis: {} }],
    })
    expect(emptyAxis.status).toBe(400)
    expect(emptyAxis.body.message).toContain('valueAxis needs min or max')
  })

  it('rejects oversized payloads', async () => {
    const many = Array.from({ length: 201 }, () => canonicalChartEdit())
    const res = await save({ chartEdits: many })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('chartEdits exceeds')
  })

  it('rejects unknown fields on the edit', async () => {
    const res = await save({
      chartEdits: [{ ...canonicalChartEdit(), sheetName: 'Data' }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field')
  })

  it('rejects invalid point colors keyed by non-index strings', async () => {
    const res = await save({
      chartEdits: [{ ...canonicalChartEdit(), pointColors: { abc: { 0: '#FF0000' } } }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('index')
  })
})
