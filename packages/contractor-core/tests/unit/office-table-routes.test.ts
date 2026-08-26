/**
 * Unit tests — /office/workbooks/save request validation for the
 * tableAdditions family (EXCEL-021 — Insert → Table).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical table addition
 * payloads (the desktop preload's workbookTableAddSchema shape) and rejects
 * missing sheet names, over-long names, empty or oversized column lists,
 * over-long column names, non-builtin style names, non-boolean bandedRows,
 * unordered or fractional areas, unknown fields, and excessive counts with
 * 400s — nothing unvalidated reaches the engine.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(
  tableAdditions: unknown,
): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: {
        edits: [],
        tableAdditions,
      },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

function canonical(): Record<string, unknown> {
  return {
    sheetName: 'Data',
    area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
    name: 'Table1',
    columnNames: ['Item', 'Qty'],
    style: 'TableStyleMedium2',
    bandedRows: true,
  }
}

describe('workbooks/save tableAdditions validation', () => {
  it('accepts a canonical table addition', async () => {
    const res = await save([canonical()])
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts an addition without the optional style (engine default)', async () => {
    const { style: _style, ...withoutStyle } = canonical()
    const res = await save([withoutStyle])
    expect(res.body.error).toBe('malformed')
  })

  it('accepts an empty tableAdditions array (no table change)', async () => {
    const res = await save([])
    expect(res.body.error).toBe('malformed')
  })

  it('rejects a missing or empty sheet name', async () => {
    const missing = await save([{ ...canonical(), sheetName: undefined }])
    expect(missing.status).toBe(400)
    expect(missing.body.message).toContain('sheetName')

    const empty = await save([{ ...canonical(), sheetName: '' }])
    expect(empty.status).toBe(400)
    expect(empty.body.message).toContain('sheetName')
  })

  it('rejects a missing or over-long table name', async () => {
    const missing = await save([{ ...canonical(), name: undefined }])
    expect(missing.status).toBe(400)
    expect(missing.body.message).toContain('name')

    const long = await save([{ ...canonical(), name: 'x'.repeat(256) }])
    expect(long.status).toBe(400)
    expect(long.body.message).toContain('255')
  })

  it('rejects an empty or oversized columnNames list', async () => {
    const empty = await save([{ ...canonical(), columnNames: [] }])
    expect(empty.status).toBe(400)
    expect(empty.body.message).toContain('columnNames')

    const many = await save([
      { ...canonical(), columnNames: Array.from({ length: 1_001 }, () => 'C') },
    ])
    expect(many.status).toBe(400)
    expect(many.body.message).toContain('1000')
  })

  it('rejects over-long or non-string column names', async () => {
    const long = await save([{ ...canonical(), columnNames: ['x'.repeat(256)] }])
    expect(long.status).toBe(400)
    expect(long.body.message).toContain('255')

    const notString = await save([{ ...canonical(), columnNames: [42] }])
    expect(notString.status).toBe(400)
    expect(notString.body.message).toContain('columnNames')
  })

  it('rejects a non-builtin style name', async () => {
    for (const bad of [
      'MyHouseStyle',
      'tablestylelight1',
      'TableStyleMedium',
      'PivotStyleLight16',
    ]) {
      const res = await save([{ ...canonical(), style: bad }])
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('style')
    }
  })

  it('accepts every built-in style family shape', async () => {
    for (const good of [
      'TableStyleLight1',
      'TableStyleLight21',
      'TableStyleMedium2',
      'TableStyleMedium28',
      'TableStyleDark11',
    ]) {
      const res = await save([{ ...canonical(), style: good }])
      expect(res.body.error).toBe('malformed')
    }
  })

  it('rejects a non-boolean bandedRows flag', async () => {
    for (const bad of ['1', 1, null, undefined]) {
      const res = await save([{ ...canonical(), bandedRows: bad }])
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('bandedRows')
    }
  })

  it('rejects unordered, fractional, or negative areas', async () => {
    const unordered = await save([
      { ...canonical(), area: { startRow: 3, startColumn: 0, endRow: 1, endColumn: 1 } },
    ])
    expect(unordered.status).toBe(400)
    expect(unordered.body.message).toContain('ordered')

    const fractional = await save([
      { ...canonical(), area: { startRow: 0.5, startColumn: 0, endRow: 2, endColumn: 1 } },
    ])
    expect(fractional.status).toBe(400)
    expect(fractional.body.message).toContain('startRow')

    const negative = await save([
      { ...canonical(), area: { startRow: -1, startColumn: 0, endRow: 2, endColumn: 1 } },
    ])
    expect(negative.status).toBe(400)
    expect(negative.body.message).toContain('startRow')
  })

  it('rejects unknown fields', async () => {
    const res = await save([{ ...canonical(), totalsRow: true }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field')
    expect(res.body.message).toContain('totalsRow')
  })

  it('rejects a non-array tableAdditions payload', async () => {
    const res = await save({ sheetName: 'S' })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('tableAdditions')
  })

  it('rejects an oversized tableAdditions payload', async () => {
    const entries = Array.from({ length: 51 }, (_, i) => ({
      ...canonical(),
      name: `Table${i + 1}`,
    }))
    const res = await save(entries)
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds')
  })
})
