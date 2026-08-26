/**
 * Unit tests — Excel cell formatting wire contract (Phase 3 Increment 9).
 *
 * Tests the presentation snapshot (styles/merges/rowHeights/colWidths) on
 * open, canonical style-delta round-trips on save, byte preservation, and
 * the runtime validation error shape for malformed WorkbookStyleEdit
 * payloads.
 *
 * These exercise the pure routeOffice function directly (node environment);
 * the real browser → HTTP → engine → bytes path is covered by the Playwright
 * suite (tests/e2e/excel-format.spec.ts).
 */
import { describe, expect, it } from 'vitest'
import { routeOffice } from '@contractor/core/api'
import { buildExcelFormatFixture, readZipEntry } from './e2e/fixtures'

interface WireFormat {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  fontFamily?: string
  fontSize?: number
  fontColor?: string
  fillColor?: string
  horizontalAlign?: string
  verticalAlign?: string
  wrapText?: boolean
}
interface WireSheet {
  id: string
  name: string
  cells: Record<string, { value: unknown; formula?: string }>
  styles?: Record<string, WireFormat>
  merges?: string[]
  rowHeights?: Record<string, number>
  colWidths?: Record<string, number>
}
interface WireEdit {
  sheetName: string
  row: number
  column: number
  writeValue: boolean
  cell: { value: unknown; formula?: string }
  style?: Record<string, unknown>
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

async function saveBook(bytes: Buffer, edits: WireEdit[]): Promise<Buffer> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/workbooks/save',
    body: { fileName: 'fixture.xlsx', fileBytes: b64(bytes), savePlan: { edits } },
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
  expect(typeof err.message).toBe('string')
  return err
}

describe('Excel presentation snapshot (open)', () => {
  it('exposes cell styles, merges, row heights and column widths', async () => {
    const bytes = await buildExcelFormatFixture()
    const sheets = await openBook(bytes)
    const formats = sheets[0]
    // Every style variant resolved.
    expect(formats.styles?.A1?.bold).toBe(true)
    expect(formats.styles?.B1?.italic).toBe(true)
    expect(formats.styles?.C1?.underline).toBe(true)
    expect(formats.styles?.C1?.strikethrough).toBe(true)
    expect(formats.styles?.C1?.fontColor).toBe('C00000')
    expect(formats.styles?.D1?.fontSize).toBe(14)
    expect(formats.styles?.E1?.fillColor).toBe('FFD966')
    expect(formats.styles?.A2?.horizontalAlign).toBe('center')
    expect(formats.styles?.A2?.verticalAlign).toBe('center')
    expect(formats.styles?.A2?.wrapText).toBe(true)
    expect(formats.styles?.B2?.horizontalAlign).toBe('right')
    // Plain cell carries no style entry.
    expect(formats.styles?.B4).toBeUndefined()
    // Merges, row heights, column widths.
    expect(formats.merges).toEqual(['A3:B3'])
    expect(formats.rowHeights).toEqual({ '5': 30 })
    expect(formats.colWidths).toEqual({ A: 173 })
    // Second sheet has no presentation data (absent, not empty objects).
    expect(sheets[1].styles).toBeUndefined()
    expect(sheets[1].merges).toBeUndefined()
  })
})

describe('Excel style edits (canonical WorkbookStyleEdit round-trips)', () => {
  it('bold delta on a plain cell round-trips', async () => {
    const bytes = await buildExcelFormatFixture()
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 3,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: { bold: true },
      },
    ])
    const reopened = await openBook(saved)
    expect(reopened[0].styles?.B4?.bold).toBe(true)
  })

  it('fill + fontColor + alignment + wrap compose on one cell', async () => {
    const bytes = await buildExcelFormatFixture()
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 3,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: {
          fillColor: '#D9EAF7',
          fontColor: '#C00000',
          horizontalAlignment: 'center',
          wrapText: true,
        },
      },
    ])
    const reopened = await openBook(saved)
    const style = reopened[0].styles?.B4
    expect(style?.fillColor).toBe('D9EAF7')
    expect(style?.fontColor).toBe('C00000')
    expect(style?.horizontalAlign).toBe('center')
    expect(style?.wrapText).toBe(true)
  })

  it('a style delta applies on top of the existing format (merge semantics)', async () => {
    const bytes = await buildExcelFormatFixture()
    // E1 already has fill FFD966; add bold on top → both survive.
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 0,
        column: 4,
        writeValue: false,
        cell: { value: null },
        style: { bold: true },
      },
    ])
    const reopened = await openBook(saved)
    const style = reopened[0].styles?.E1
    expect(style?.bold).toBe(true)
    expect(style?.fillColor).toBe('FFD966')
  })

  it('bold:false clears bold while keeping other marks', async () => {
    const bytes = await buildExcelFormatFixture()
    // C1 has underline + strike + red font; remove nothing but bold (it has
    // none — use A1 which is bold; clearing bold keeps the font family).
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { bold: false },
      },
    ])
    const reopened = await openBook(saved)
    const style = reopened[0].styles?.A1
    expect(style?.bold).toBeUndefined()
  })

  it('combined value + style edit applies both', async () => {
    const bytes = await buildExcelFormatFixture()
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 3,
        column: 1,
        writeValue: true,
        cell: { value: 'Styled + valued' },
        style: { bold: true, fillColor: '#D9EAF7' },
      },
    ])
    const reopened = await openBook(saved)
    expect(reopened[0].cells.B4?.value).toBe('Styled + valued')
    expect(reopened[0].styles?.B4?.bold).toBe(true)
    expect(reopened[0].styles?.B4?.fillColor).toBe('D9EAF7')
  })

  it('untouched styles/merges/heights/widths are byte-preserved', async () => {
    const bytes = await buildExcelFormatFixture()
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 3,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: { bold: true },
      },
    ])
    const reopened = await openBook(saved)
    // Original formatting untouched.
    expect(reopened[0].styles?.A1?.bold).toBe(true)
    expect(reopened[0].styles?.E1?.fillColor).toBe('FFD966')
    expect(reopened[0].merges).toEqual(['A3:B3'])
    expect(reopened[0].rowHeights).toEqual({ '5': 30 })
    expect(reopened[0].colWidths).toEqual({ A: 173 })
    // The other sheet's cell untouched.
    expect(reopened[1].cells.A1?.value).toBe('Untouched')
  })

  it('saved XML adds a deduped xf without touching existing entries', async () => {
    const bytes = await buildExcelFormatFixture()
    const saved = await saveBook(bytes, [
      {
        sheetName: 'Formats',
        row: 3,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: { bold: true },
      },
    ])
    const stylesXml = await readZipEntry(saved, 'xl/styles.xml')
    // Existing entries preserved verbatim (copy-on-write).
    expect(stylesXml).toContain('<font><b/><sz val="11"/><name val="Calibri"/></font>')
    expect(stylesXml).toContain('<fgColor rgb="FFFFD966"/>')
    // Full dedup: B4's delta resolved to the EXISTING bold xf 1 — the
    // stylesheet is byte-identical (no new fonts, no new xfs; the regex
    // counts 1 cellStyleXfs entry + 8 cellXfs entries).
    expect((stylesXml.match(/<font>/g) ?? []).length).toBe(5)
    expect((stylesXml.match(/<xf /g) ?? []).length).toBe(9)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<c r="B4" s="1"')
  })
})

describe('Excel style validation (malformed deltas → 400)', () => {
  it('rejects a color without the # prefix', async () => {
    const bytes = await buildExcelFormatFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: { fillColor: 'FFD966' },
          },
        ],
      },
    })
    expect(err.message).toContain('fillColor')
  })

  it('rejects an invalid horizontal alignment', async () => {
    const bytes = await buildExcelFormatFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: { horizontalAlignment: 'diagonal' },
          },
        ],
      },
    })
    expect(err.message).toContain('horizontalAlignment')
  })

  it('rejects an out-of-range fontSize', async () => {
    const bytes = await buildExcelFormatFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: { fontSize: 1000 },
          },
        ],
      },
    })
    expect(err.message).toContain('fontSize')
  })

  it('rejects an invalid underlineStyle', async () => {
    const bytes = await buildExcelFormatFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: { underlineStyle: 'wavy' },
          },
        ],
      },
    })
    expect(err.message).toContain('underlineStyle')
  })

  it('rejects a non-object style payload', async () => {
    const bytes = await buildExcelFormatFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: '<script>alert(1)</script>',
          },
        ],
      },
    })
    expect(err.message).toContain('style')
  })

  it('rejects a non-boolean bold', async () => {
    const bytes = await buildExcelFormatFixture()
    await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: { bold: 'yes' },
          },
        ],
      },
    })
  })

  it('rejects an invalid border style', async () => {
    const bytes = await buildExcelFormatFixture()
    const err = await expectValidation({
      fileName: 'fixture.xlsx',
      fileBytes: b64(bytes),
      savePlan: {
        edits: [
          {
            sheetName: 'Formats',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: { borderTop: { style: 'zigzag', color: '#000000' } },
          },
        ],
      },
    })
    expect(err.message).toContain('borderTop')
  })
})
