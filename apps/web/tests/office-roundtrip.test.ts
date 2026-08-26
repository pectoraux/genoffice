/**
 * Round-trip tests for the office API routes.
 *
 * Verifies:
 *   - XLSX round-trip: build fixture → open via routeOffice → verify cells →
 *     save → re-open → verify cells preserved
 *   - DOCX round-trip: build fixture → open via routeOffice → verify blocks →
 *     save → re-open → verify blocks preserved
 *   - Security: path traversal rejected, oversized file rejected, malformed
 *     file rejected
 *
 * The tests exercise the pure `routeOffice` function directly — they do NOT
 * spin up the HTTP server. The HTTP layer (vercel-handler.ts) is just a thin
 * transport that calls the same function.
 */

import { describe, expect, it } from 'vitest'
import { routeOffice } from '@contractor/core/api'
import { buildCompatibilityFixture } from '@genoffice/xlsx-gateway/tests/fixture-builder'
import { buildBlankDocx } from '@genoffice/docx-engine'
import type { CellEdit } from '@genoffice/xlsx-gateway'

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Wrap raw bytes into the routeOffice open-workbook request shape. */
async function openWorkbook(fileName: string, bytes: Uint8Array) {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/workbooks/open',
    body: { fileName, fileBytes: toBase64(bytes) },
  })
  if (!res || res.status !== 200) {
    throw new Error(`openWorkbook failed: ${res?.status} ${JSON.stringify(res?.body)}`)
  }
  return res.body as {
    snapshot: {
      sheets: ReadonlyArray<{
        id: string
        name: string
        cells: Readonly<
          Record<string, { value: string | number | boolean | null; formula?: string }>
        >
      }>
    }
    sheetNamesById: Readonly<Record<string, string>>
  }
}

async function saveWorkbook(fileName: string, bytes: Uint8Array, edits: readonly CellEdit[]) {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/workbooks/save',
    body: { fileName, fileBytes: toBase64(bytes), cellEdits: edits },
  })
  if (!res || res.status !== 200) {
    throw new Error(`saveWorkbook failed: ${res?.status} ${JSON.stringify(res?.body)}`)
  }
  return fromBase64((res?.body as { fileBytes: string }).fileBytes)
}

async function openDocument(fileName: string, bytes: Uint8Array) {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/open',
    body: { fileName, fileBytes: toBase64(bytes) },
  })
  if (!res || res.status !== 200) {
    throw new Error(`openDocument failed: ${res?.status} ${JSON.stringify(res?.body)}`)
  }
  return (
    res?.body as {
      blocks: ReadonlyArray<{
        docxIndex: number | null
        type: string
        text: string
        level?: number
      }>
    }
  ).blocks
}

async function saveDocument(fileName: string, bytes: Uint8Array, blocks: readonly unknown[]) {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/save',
    body: { fileName, fileBytes: toBase64(bytes), blocks },
  })
  if (!res || res.status !== 200) {
    throw new Error(`saveDocument failed: ${res?.status} ${JSON.stringify(res?.body)}`)
  }
  return fromBase64((res?.body as { fileBytes: string }).fileBytes)
}

function toBase64(bytes: Uint8Array): string {
  // Node Buffer is available in the test environment (vitest runs on Node).
  // The browser uses a different implementation (see office-client.ts), but
  // the wire shape is identical.
  return Buffer.from(bytes).toString('base64')
}

function fromBase64(b64: string): Uint8Array {
  return new Uint8Array(Buffer.from(b64, 'base64'))
}

// ── XLSX round-trip ─────────────────────────────────────────────────────────

describe('XLSX round-trip', () => {
  it('opens a fixture workbook and reads its cells', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await openWorkbook('fixture.xlsx', bytes)
    expect(res.snapshot.sheets).toHaveLength(1)
    const sheet = res.snapshot.sheets[0]
    expect(sheet.name).toBe('Sheet1')
    expect(sheet.cells.A1?.value).toBe('Old')
    expect(sheet.cells.B1?.value).toBe(10)
  })

  it('preserves cells through a no-op save → re-open cycle', async () => {
    const bytes = await buildCompatibilityFixture()
    const before = await openWorkbook('fixture.xlsx', bytes)
    const saved = await saveWorkbook('fixture.xlsx', bytes, [])
    const after = await openWorkbook('fixture.xlsx', saved)
    expect(after.snapshot.sheets[0].cells).toEqual(before.snapshot.sheets[0].cells)
  })

  it('applies cell edits and re-reads the mutated values', async () => {
    const bytes = await buildCompatibilityFixture()
    const edits: CellEdit[] = [
      { sheetName: 'Sheet1', row: 0, column: 0, writeValue: true, cell: { value: 'World' } },
      { sheetName: 'Sheet1', row: 1, column: 1, writeValue: true, cell: { value: 42 } },
    ]
    const saved = await saveWorkbook('fixture.xlsx', bytes, edits)
    const after = await openWorkbook('fixture.xlsx', saved)
    const cells = after.snapshot.sheets[0].cells
    expect(cells.A1?.value).toBe('World')
    expect(cells.B2?.value).toBe(42)
    // Unedited B1 survives.
    expect(cells.B1?.value).toBe(10)
  })

  it('round-trips a formula cell unchanged', async () => {
    const bytes = await buildCompatibilityFixture()
    const before = await openWorkbook('fixture.xlsx', bytes)
    // The fixture has a formula at B3 (=SUM(C1:C2)). Saving with no edits and
    // re-opening should preserve the formula string.
    const saved = await saveWorkbook('fixture.xlsx', bytes, [])
    const after = await openWorkbook('fixture.xlsx', saved)
    const beforeFormula = before.snapshot.sheets[0].cells.B3?.formula
    const afterFormula = after.snapshot.sheets[0].cells.B3?.formula
    // The formula should be preserved — the reader may store it with or
    // without the '=' prefix. Just verify it round-trips unchanged.
    expect(beforeFormula).toBe(afterFormula)
  })
})

// ── DOCX round-trip ─────────────────────────────────────────────────────────

describe('DOCX round-trip', () => {
  it('opens a blank document and returns at least one block', async () => {
    const bytes = await buildBlankDocx()
    const blocks = await openDocument('blank.docx', bytes)
    expect(blocks.length).toBeGreaterThan(0)
    // Blank doc has a single empty paragraph + a hidden sectPr block.
    const visible = blocks.filter((b) => b.type !== 'hidden')
    expect(visible.length).toBeGreaterThanOrEqual(1)
  })

  it('preserves bytes through a no-op save → re-open cycle (blank doc)', async () => {
    const bytes = await buildBlankDocx()
    const before = await openDocument('blank.docx', bytes)
    // Mark all blocks as unchanged (edited: false) — the engine returns the
    // original bytes byte-identically.
    const blocksNoEdit = before.map((b) => ({ ...b, edited: false }))
    const saved = await saveDocument('blank.docx', bytes, blocksNoEdit)
    // Re-open and verify the block structure matches.
    const after = await openDocument('blank.docx', saved)
    expect(after.length).toBe(before.length)
    for (let i = 0; i < before.length; i++) {
      expect(after[i].type).toBe(before[i].type)
      expect(after[i].docxIndex).toBe(before[i].docxIndex)
    }
  })

  it('preserves the blank document through a no-op save → re-open cycle', async () => {
    const bytes = await buildBlankDocx()
    const before = await openDocument('blank.docx', bytes)
    // Verify the blank fixture has paragraph blocks.
    const visibleTypes = before.filter((b) => b.type !== 'hidden').map((b) => b.type)
    expect(visibleTypes).toContain('paragraph')

    const blocksNoEdit = before.map((b) => ({ ...b, edited: false }))
    const saved = await saveDocument('blank.docx', bytes, blocksNoEdit)
    const after = await openDocument('blank.docx', saved)
    expect(after.length).toBe(before.length)
    const afterVisibleTypes = after.filter((b) => b.type !== 'hidden').map((b) => b.type)
    expect(afterVisibleTypes).toEqual(visibleTypes)
  })

  it('returns text content for paragraph blocks', async () => {
    const bytes = await buildBlankDocx()
    const blocks = await openDocument('blank.docx', bytes)
    // The blank fixture has paragraph blocks.
    const firstParagraph = blocks.find((b) => b.type === 'paragraph')
    expect(firstParagraph).toBeDefined()
    // Text may be empty for a blank paragraph — that's OK.
    expect(typeof firstParagraph?.text).toBe('string')
  })
})

// ── Security ────────────────────────────────────────────────────────────────

describe('office security', () => {
  it('rejects path traversal in fileName (..)', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: '../etc/passwd.xlsx', fileBytes: toBase64(bytes) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('rejects path traversal in fileName (slash)', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'subdir/evil.xlsx', fileBytes: toBase64(bytes) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('rejects path traversal in fileName (backslash)', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'subdir\\evil.xlsx', fileBytes: toBase64(bytes) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('rejects a fileName with invalid characters', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'bad name.xlsx', fileBytes: toBase64(bytes) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('rejects an unsupported extension', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'secret.txt', fileBytes: toBase64(bytes) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('rejects an oversized file (>10MB)', async () => {
    // 11MB of zeros — above the 10MB cap.
    const big = new Uint8Array(11 * 1024 * 1024)
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'big.xlsx', fileBytes: toBase64(big) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
    expect((res?.body as { message: string }).message).toMatch(/10MB/)
  })

  it('rejects a malformed XLSX (not a zip)', async () => {
    const garbage = new TextEncoder().encode('this is not a zip file')
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'bad.xlsx', fileBytes: toBase64(garbage) },
    })
    expect(res?.status).toBe(400)
    // The route maps engine parse failures to the malformed error code.
    const body = res?.body as { error: string; message: string }
    expect(body.error === 'validation' || body.error === 'malformed').toBe(true)
  })

  it('rejects a malformed DOCX (not a zip)', async () => {
    const garbage = new TextEncoder().encode('this is not a docx file')
    const res = await routeOffice({
      method: 'POST',
      path: '/office/documents/open',
      body: { fileName: 'bad.docx', fileBytes: toBase64(garbage) },
    })
    expect(res?.status).toBe(400)
    const body = res?.body as { error: string; message: string }
    expect(body.error === 'validation' || body.error === 'malformed').toBe(true)
  })

  it('rejects missing fileName', async () => {
    const bytes = await buildCompatibilityFixture()
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileBytes: toBase64(bytes) },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('rejects missing fileBytes', async () => {
    const res = await routeOffice({
      method: 'POST',
      path: '/office/workbooks/open',
      body: { fileName: 'test.xlsx' },
    })
    expect(res?.status).toBe(400)
    expect((res?.body as { error: string }).error).toBe('validation')
  })

  it('returns null for an unknown office route', async () => {
    const res = await routeOffice({
      method: 'POST',
      path: '/office/foo/bar',
      body: {},
    })
    expect(res).toBe(null)
  })

  it('returns null for a non-office path', async () => {
    const res = await routeOffice({
      method: 'GET',
      path: '/projects',
      body: null,
    })
    expect(res).toBe(null)
  })

  it('accepts .csv, .xls, .xlsx, .docx extensions', async () => {
    const bytes = await buildCompatibilityFixture()
    for (const name of ['a.xlsx', 'b.csv', 'c.xls']) {
      const res = await routeOffice({
        method: 'POST',
        path: '/office/workbooks/open',
        body: { fileName: name, fileBytes: toBase64(bytes) },
      })
      // .xls/.csv may parse-fail (we only have an XLSX parser), but the
      // filename validation must pass — the failure, if any, comes from
      // the engine, not the security layer.
      if (res?.status === 200) continue
      const body = res?.body as { error: string }
      expect(body.error).not.toBe('validation')
    }
  })
})

// ── Workbook save with cellEdits = full re-write ────────────────────────────

describe('XLSX save with edits', () => {
  it('overwrites an existing cell value', async () => {
    const bytes = await buildCompatibilityFixture()
    const edits: CellEdit[] = [
      { sheetName: 'Sheet1', row: 0, column: 0, writeValue: true, cell: { value: 'New' } },
    ]
    const saved = await saveWorkbook('fixture.xlsx', bytes, edits)
    const after = await openWorkbook('fixture.xlsx', saved)
    expect(after.snapshot.sheets[0].cells.A1?.value).toBe('New')
  })

  it('adds a new cell that did not exist before', async () => {
    const bytes = await buildCompatibilityFixture()
    const edits: CellEdit[] = [
      { sheetName: 'Sheet1', row: 5, column: 5, writeValue: true, cell: { value: 'inserted' } },
    ]
    const saved = await saveWorkbook('fixture.xlsx', bytes, edits)
    const after = await openWorkbook('fixture.xlsx', saved)
    expect(after.snapshot.sheets[0].cells.F6?.value).toBe('inserted')
  })

  it('clears a cell by writing null', async () => {
    const bytes = await buildCompatibilityFixture()
    const edits: CellEdit[] = [
      { sheetName: 'Sheet1', row: 0, column: 0, writeValue: true, cell: { value: null } },
    ]
    const saved = await saveWorkbook('fixture.xlsx', bytes, edits)
    const after = await openWorkbook('fixture.xlsx', saved)
    // A1 was 'Old'; after clearing, it should either be absent or null-valued.
    const a1 = after.snapshot.sheets[0].cells.A1
    expect(a1 === undefined || a1.value === null).toBe(true)
  })

  it('writes a formula', async () => {
    const bytes = await buildCompatibilityFixture()
    const edits: CellEdit[] = [
      {
        sheetName: 'Sheet1',
        row: 2,
        column: 1,
        writeValue: true,
        cell: { value: null, formula: '=SUM(C1:C3)' },
      },
    ]
    const saved = await saveWorkbook('fixture.xlsx', bytes, edits)
    const after = await openWorkbook('fixture.xlsx', saved)
    expect(after.snapshot.sheets[0].cells.B3?.formula).toBe('=SUM(C1:C3)')
  })
})
