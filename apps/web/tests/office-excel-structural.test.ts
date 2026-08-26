/**
 * Unit tests — Excel structural operations wire contract.
 *
 * Tests insert/remove row/column round-trips through the canonical engine
 * (applyStructuralOps replay), the combined structural + cell-edit flow
 * (post-op coordinates), byte preservation of untouched content, and the
 * runtime validation error shape for malformed structural ops.
 *
 * These exercise the pure routeOffice function directly (node environment);
 * the real browser → HTTP → engine → bytes path is covered by the Playwright
 * suite (tests/e2e/excel-structural.spec.ts).
 */
import { describe, expect, it } from 'vitest'
import { routeOffice } from '@contractor/core/api'
import { buildExcelFixture } from './e2e/fixtures'

interface WireSheet {
  id: string
  name: string
  cells: Record<string, { value: unknown; formula?: string }>
  merges?: string[]
  rowHeights?: Record<string, number>
  colWidths?: Record<string, number>
}

const b64 = (b: Buffer) => b.toString('base64')

async function openBook(bytes: Buffer): Promise<WireSheet[]> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/workbooks/open',
    body: { fileName: 'fixture.xlsx', fileBytes: b64(bytes) },
  })
  expect(res?.status).toBe(200)
  return (res?.body as { snapshot: { sheets: WireSheet[] } }).snapshot.sheets
}

async function saveBook(
  bytes: Buffer,
  structuralOps: Array<{
    sheetName: string
    ops: Array<{ kind: string; index: number; count: number }>
  }>,
  edits: Array<Record<string, unknown>> = [],
): Promise<Buffer> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/workbooks/save',
    body: {
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: { edits, structuralOps },
    },
  })
  if (res?.status !== 200) {
    throw new Error(`save failed: ${res?.status} ${JSON.stringify(res?.body).slice(0, 300)}`)
  }
  return Buffer.from((res?.body as { fileBytes: string }).fileBytes, 'base64')
}

/** Expect a validation error (400) and return its body. */
async function expectValidation(body: unknown): Promise<{ error: string; message: string }> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/workbooks/save',
    body,
  })
  expect(res?.status).toBe(400)
  const err = res?.body as { error: string; message: string }
  expect(err.error).toBe('validation')
  return err
}

describe('Excel structural ops — insert rows', () => {
  it('inserts 2 rows and shifts cells, formulas, merges, heights', async () => {
    const bytes = await buildExcelFixture()
    const saved = await saveBook(bytes, [
      { sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] },
    ])
    const sheets = await openBook(saved)
    const data = sheets[0]
    // Original A1 "Original Text" is now A3.
    expect(data.cells.A3?.value).toBe('Original Text')
    expect(data.cells.B3?.value).toBe(10)
    expect(data.cells.C3?.formula).toBe('=SUM(B3:B3)')
    // Merge shifted: A3:B3 → A5:B5.
    expect(data.merges).toEqual(['A5:B5'])
    // Row height shifted: row 5 → row 7.
    expect(data.rowHeights).toEqual({ '7': 30 })
    // Column width unchanged.
    expect(data.colWidths).toEqual({ A: 173 })
    // Hidden sheet untouched.
    expect(sheets[1].name).toBe('HiddenSheet')
    expect(sheets[1].cells.A1?.value).toBe('Hidden Value')
  })

  it('combined structural + cell edit in post-op coordinates', async () => {
    const bytes = await buildExcelFixture()
    const saved = await saveBook(
      bytes,
      [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] }],
      // Post-shift: the original A1 is now A3.
      [{ sheetName: 'Data', row: 2, column: 0, writeValue: true, cell: { value: 'EDITED' } }],
    )
    const sheets = await openBook(saved)
    const data = sheets[0]
    // The cell edit landed on the shifted A3.
    expect(data.cells.A3?.value).toBe('EDITED')
    // Untouched cells survived.
    expect(data.cells.B3?.value).toBe(10)
    expect(data.cells.C3?.formula).toBe('=SUM(B3:B3)')
  })
})

describe('Excel structural ops — remove rows', () => {
  it('removes the first row and shifts content up', async () => {
    const bytes = await buildExcelFixture()
    const saved = await saveBook(bytes, [
      { sheetName: 'Data', ops: [{ kind: 'remove-rows', index: 0, count: 1 }] },
    ])
    const sheets = await openBook(saved)
    const data = sheets[0]
    // Original A2 (empty) is now A1; original A3 "Merged Header" is now A2.
    expect(data.cells.A2?.value).toBe('Merged Header')
    // Original B1=10 is gone (removed with row 1). B2 (empty) is now B1.
    // The merge shifted from A3:B3 to A2:B2.
    expect(data.merges).toEqual(['A2:B2'])
    // Row height shifted from row 5 to row 4.
    expect(data.rowHeights).toEqual({ '4': 30 })
  })
})

describe('Excel structural ops — insert/remove columns', () => {
  it('inserts a column and shifts cells', async () => {
    const bytes = await buildExcelFixture()
    const saved = await saveBook(bytes, [
      { sheetName: 'Data', ops: [{ kind: 'insert-cols', index: 0, count: 1 }] },
    ])
    const sheets = await openBook(saved)
    const data = sheets[0]
    // Original A1 is now B1.
    expect(data.cells.B1?.value).toBe('Original Text')
    expect(data.cells.C1?.value).toBe(10)
    expect(data.cells.D1?.formula).toBe('=SUM(C1:C1)')
    // Column width shifted: A → B.
    expect(data.colWidths).toEqual({ B: 173 })
    // Merge shifted from A3:B3 to B3:C3.
    expect(data.merges).toEqual(['B3:C3'])
  })

  it('removes a column and shifts cells left', async () => {
    const bytes = await buildExcelFixture()
    const saved = await saveBook(bytes, [
      { sheetName: 'Data', ops: [{ kind: 'remove-cols', index: 0, count: 1 }] },
    ])
    const sheets = await openBook(saved)
    const data = sheets[0]
    // Original B1=10 is now A1.
    expect(data.cells.A1?.value).toBe(10)
    // Original C1 formula is now B1.
    expect(data.cells.B1?.formula).toBe('=SUM(A1:A1)')
    // The custom column A width is gone (column A removed).
    expect(data.colWidths).toBeUndefined()
  })
})

describe('Excel structural ops — validation (malformed payloads → 400)', () => {
  it('rejects an invalid kind', async () => {
    const bytes = await buildExcelFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [],
        structuralOps: [
          { sheetName: 'Data', ops: [{ kind: 'teleport-rows', index: 0, count: 1 }] },
        ],
      },
    })
    expect(err.message).toContain('kind')
  })

  it('rejects a negative index', async () => {
    const bytes = await buildExcelFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [],
        structuralOps: [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: -1, count: 1 }] }],
      },
    })
    expect(err.message).toContain('index')
  })

  it('rejects a zero count', async () => {
    const bytes = await buildExcelFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [],
        structuralOps: [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 0 }] }],
      },
    })
    expect(err.message).toContain('count')
  })

  it('rejects a non-object op', async () => {
    const bytes = await buildExcelFixture()
    await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [],
        structuralOps: [{ sheetName: 'Data', ops: ['not-an-object'] }],
      },
    })
  })

  it('rejects a missing sheet name', async () => {
    const bytes = await buildExcelFixture()
    await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [],
        structuralOps: [{ ops: [{ kind: 'insert-rows', index: 0, count: 1 }] }],
      },
    })
  })
})
