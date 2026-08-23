/**
 * Unit tests — editable table wire contract (Phase 3 Increment 7).
 *
 * Tests the office-routes table payload: TypedTable serialization on open,
 * canonical regeneration on save (generateTableModelXml path), byte
 * preservation of unchanged tables, and the runtime validation error shape
 * for malformed table payloads.
 *
 * These exercise the pure routeOffice function directly (node environment);
 * the real browser → HTTP → engine → bytes path is covered by the Playwright
 * suite (tests/e2e/word-table.spec.ts).
 */
import { describe, expect, it } from 'vitest'
import { routeOffice } from '@contractor/core/api'
import { buildWordTableFixture } from './e2e/fixtures'
import { readZipEntry } from './e2e/fixtures'

interface WireRun {
  text: string
  bold?: boolean
  italic?: boolean
}
interface WireCell {
  paras: string[]
  richParas?: Array<{ runs: WireRun[]; align?: string }>
  colSpan?: number
  vMerge?: 'restart' | 'continue'
  fill?: string
  rawTcPr?: string
}
interface WireTable {
  rows: WireCell[][]
  colWidthsPct?: number[]
  headerRows?: boolean[]
}
interface WireBlock {
  docxIndex: number | null
  type: string
  text: string
  table?: WireTable
  edited?: boolean
  hidden?: boolean
}

const b64 = (b: Buffer) => b.toString('base64')

async function openDoc(bytes: Buffer): Promise<WireBlock[]> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/open',
    body: { fileName: 'fixture.docx', fileBytes: b64(bytes) },
  })
  expect(res?.status).toBe(200)
  return (res?.body as { blocks: WireBlock[] }).blocks
}

async function saveDoc(bytes: Buffer, blocks: WireBlock[]): Promise<Buffer> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/save',
    body: { fileName: 'fixture.docx', fileBytes: b64(bytes), blocks },
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
    path: '/office/documents/save',
    body,
  })
  expect(res?.status).toBe(400)
  const err = res?.body as { error: string; message: string }
  expect(err.error).toBe('validation')
  expect(typeof err.message).toBe('string')
  return err
}

describe('Word table wire contract', () => {
  it('open serializes the typed table payload (structure, marks, fill, vMerge)', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const tableBlock = blocks.find((b) => b.type === 'table')
    expect(tableBlock).toBeDefined()
    expect(tableBlock!.docxIndex).toBe(2)
    const table = tableBlock!.table!
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0]).toHaveLength(3)
    expect(table.colWidthsPct).toHaveLength(3)
    expect(table.rows[0][0].vMerge).toBe('restart')
    expect(table.rows[1][0].vMerge).toBe('continue')
    expect(table.rows[0][0].richParas?.[0].runs).toEqual([
      { text: 'Merged ' },
      { text: 'bold', bold: true },
      { text: ' and ' },
      { text: 'italic', italic: true },
      { text: ' cell' },
    ])
    expect(table.rows[0][1].fill).toBe('FFF2CC')
    expect(table.rows[0][1].richParas?.[0].align).toBe('center')
    // rawTcPr echoes for byte preservation of unmodeled properties.
    expect(table.rows[0][1].rawTcPr).toContain('<w:shd')
  })

  it('unchanged table (edited=false) is byte-preserved through the original path', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    // The original table bytes are copied verbatim (kind: 'original'):
    // the fixture's exact indentation survives regeneration-free.
    expect(xml).toContain(
      '<w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>',
    )
    expect(xml).toContain('Bottom middle')
    expect(xml).toContain('w:fill="FFF2CC"')
  })

  it('edited table regenerates through the canonical engine generator', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    tbl.table!.rows[1][1].richParas![0].runs[0] = { text: 'UNIT EDITED cell' }
    tbl.table!.rows[1][1].paras[0] = 'UNIT EDITED cell'
    // Structural edit: add a header row flag.
    tbl.table!.headerRows = [false, false]
    const saved = await saveDoc(bytes, edited)

    const xml = await readZipEntry(saved, 'word/document.xml')
    const tblXml = xml.match(/<w:tbl>[\s\S]*<\/w:tbl>/)?.[0] ?? ''
    expect(tblXml).toContain('UNIT EDITED cell')
    expect(tblXml).toContain('Bottom right')
    expect(tblXml).toContain('<w:vMerge w:val="restart"/>')
    expect(/<w:vMerge\/>/.test(tblXml)).toBe(true)
    expect(tblXml).toContain('w:fill="FFF2CC"')
    expect((tblXml.match(/<w:gridCol/g) ?? []).length).toBe(3)
    // tblPr is preserved from the original bytes (borders + fixed layout).
    expect(tblXml).toContain('<w:tblLayout w:type="fixed"/>')

    // Reopen: the edit round-trips.
    const reopened = await openDoc(saved)
    const rt = reopened.find((b) => b.type === 'table')?.table
    expect(rt?.rows[1][1].paras[0]).toBe('UNIT EDITED cell')
    expect(rt?.rows[0][0].vMerge).toBe('restart')
    expect(rt?.rows[1][0].vMerge).toBe('continue')
    expect(rt?.rows[0][1].fill).toBe('FFF2CC')
  })

  it('headerRows flag is patched into trPr tblHeader on save', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    tbl.table!.headerRows = [true, false]
    const saved = await saveDoc(bytes, edited)
    const xml = await readZipEntry(saved, 'word/document.xml')
    const tblXml = xml.match(/<w:tbl>[\s\S]*<\/w:tbl>/)?.[0] ?? ''
    expect(tblXml).toContain('<w:tblHeader/>')
    // Reopen: the header row flag round-trips.
    const reopened = await openDoc(saved)
    const rt = reopened.find((b) => b.type === 'table')?.table
    expect(rt?.headerRows?.[0]).toBe(true)
    expect(rt?.headerRows?.[1]).toBe(false)
  })
})

describe('Word table wire validation (malformed payloads → 400 validation error)', () => {
  it('rejects an invalid vMerge value', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    ;(tbl.table!.rows[0][0] as unknown as { vMerge: string }).vMerge = 'sideways'
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('vMerge')
  })

  it('rejects an out-of-range colSpan', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    tbl.table!.rows[0][0].colSpan = 999
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('colSpan')
  })

  it('rejects a table payload on a non-table block', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    edited[0].table = edited[2].table
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('table')
  })

  it('rejects an edited table without a payload', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    delete tbl.table
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('table payload')
  })

  it('rejects a malformed rawTcPr fragment', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    tbl.table!.rows[0][0].rawTcPr = '<w:tcPr><script>alert(1)</script></w:tcPr>'
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('rawTcPr')
  })

  it('rejects malformed table rows (not arrays of arrays)', async () => {
    const bytes = await buildWordTableFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const tbl = edited.find((b) => b.type === 'table')!
    tbl.edited = true
    ;(tbl.table as unknown as { rows: unknown }).rows = 'not-an-array'
    await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
  })
})
