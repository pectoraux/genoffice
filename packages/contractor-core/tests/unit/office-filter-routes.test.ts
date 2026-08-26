/**
 * Unit tests — /office/workbooks/save request validation for the filterStates
 * family (Phase 4 Increment 4 — Data → Filter).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical SheetFilterState
 * payloads and rejects arbitrary/unknown filter objects, unsupported
 * operators, malformed ranges, and out-of-bounds coordinates with 400s —
 * nothing unvalidated reaches the engine.
 *
 * The XLSX application itself is covered by @genoffice/xlsx-gateway tests;
 * these tests pin the WIRE contract.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Minimal valid XLSX bytes (a real zip is not required for validation-only
 * tests — decodeFileBytes accepts any base64 payload; invalid archives only
 * fail later at the engine stage). */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

interface SavePlan {
  edits?: unknown[]
  filterStates?: unknown[]
}

async function save(plan: SavePlan): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: { edits: [], ...plan },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

describe('workbooks/save filterStates validation', () => {
  it('accepts a canonical value-filter state (range + values + hidden rows)', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'Produce',
          filter: {
            range: { startRow: 0, endRow: 7, startColumn: 0, endColumn: 3 },
            columns: [{ colId: 0, values: ['Fruit', 'Veg'] }],
          },
          hiddenRows: [2, 4, 6],
          visibilityRange: { startRow: 0, endRow: 7, startColumn: 0, endColumn: 3 },
        },
      ],
    })
    // Validation passes; the engine stage then fails on the placeholder
    // bytes with a MALFORMED (not validation) error — proving the filter
    // state itself cleared validation.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts the cleared-filter state (filter: null)', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'Produce',
          filter: null,
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 7, startColumn: 0, endColumn: 3 },
        },
      ],
    })
    expect(res.body.error).toBe('malformed')
  })

  it('accepts blank + custom criteria with supported operators', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 2 },
            columns: [
              { colId: 0, blank: true },
              {
                colId: 1,
                customs: {
                  and: true,
                  filters: [
                    { val: 5, operator: 'greaterThan' },
                    { val: 10, operator: 'lessThanOrEqual' },
                  ],
                },
              },
            ],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 2 },
        },
      ],
    })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects a non-object filter state entry', async () => {
    const res = await save({ filterStates: ['not-a-filter'] as unknown[] })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
    expect(res.body.message).toContain('filterStates[0]')
  })

  it('rejects an unknown filter "kind" object (arbitrary filter shapes)', async () => {
    const res = await save({
      filterStates: [{ sheetName: 'S', kind: 'colorFilter', colors: ['#fff'] } as unknown],
    })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
    // filter is required — the arbitrary keys did not slip through.
    expect(res.body.message).toContain('filter')
  })

  it('rejects a missing filter field', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('.filter is required')
  })

  it('rejects unsupported custom operators', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
            columns: [{ colId: 0, customs: { filters: [{ val: 'x', operator: 'beginsWith' }] } }],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('not a supported filter condition')
  })

  it('rejects a malformed range (end before start)', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 5, endRow: 2, startColumn: 0, endColumn: 1 },
            columns: [{ colId: 0, values: ['a'] }],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 5, endRow: 2, startColumn: 0, endColumn: 1 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('end must be >= start')
  })

  it('rejects a column whose colId lies outside the filter range', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
            columns: [{ colId: 5, values: ['a'] }],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('outside the filter range')
  })

  it('rejects a criteria-less filter column', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
            columns: [{ colId: 0 }],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('no criteria')
  })

  it('rejects non-string filter values', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
            columns: [{ colId: 0, values: [42 as unknown as string] }],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('values[0] must be a string')
  })

  it('rejects more than two custom conditions', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: {
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
            columns: [
              {
                colId: 0,
                customs: {
                  filters: [
                    { val: 1, operator: 'equal' },
                    { val: 2, operator: 'equal' },
                    { val: 3, operator: 'equal' },
                  ],
                },
              },
            ],
          },
          hiddenRows: [],
          visibilityRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('1..2 conditions')
  })

  it('rejects negative hidden row indexes', async () => {
    const res = await save({
      filterStates: [
        {
          sheetName: 'S',
          filter: null,
          hiddenRows: [-1],
          visibilityRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('hiddenRows[0]')
  })

  it('rejects an absurd number of filter states', async () => {
    const states = Array.from({ length: 101 }, () => ({
      sheetName: 'S',
      filter: null,
      hiddenRows: [],
      visibilityRange: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    }))
    const res = await save({ filterStates: states })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 100')
  })
})
