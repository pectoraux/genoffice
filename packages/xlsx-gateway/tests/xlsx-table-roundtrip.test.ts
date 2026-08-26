/**
 * EXCEL-021 — Tables round-trip tests.
 *
 * Proves the canonical table family end-to-end at the gateway:
 *
 *   read: <tableParts> (per worksheet, resolved through the rels) →
 *         WorksheetState.tables — metadata (area, header/totals rows,
 *         name, columns, style flags) plus PRE-RESOLVED banding colors
 *         (theme accents + Excel's HSL tint transform for the built-in
 *         Light/Medium/Dark families; custom <tableStyle> dxfs from
 *         styles.xml). Unreadable table wiring fails closed PER SHEET;
 *         parts without a readable ref are skipped per table.
 *
 *   write: applyCellEditsToXlsx(tableAdditions) — the new trailing
 *         parameter flows to planCellEditsToXlsx's slot 15 →
 *         applyTableAdditions writes xl/tables/tableN.xml, the
 *         worksheet's <tableParts>, the relationship, and the
 *         [Content_Types] override.
 *
 *   round-trip: create → save → readBasicWorkbook → table metadata with
 *         the Medium2 palette; a no-op save (empty families) preserves
 *         the source bytes exactly, tables included.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'
import type { SheetTableAddition } from '../src/gateway/xlsx-gateway'

/// Minimal single-sheet workbook with one table wired through the rels.
/// The theme is optional; the table part is optional (omission builds a
/// worksheet that REFERENCES a table whose part is missing).
async function buildTableFixture(options?: {
  readonly tableXml?: string
  readonly themeXml?: string
  readonly stylesXml?: string
  readonly dropTablePart?: boolean
  readonly dropRels?: boolean
  readonly sheetName?: string
}): Promise<Buffer> {
  const zip = new JSZip()
  const hasTable = options?.tableXml !== undefined
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  ${hasTable ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ''}
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${options?.sheetName ?? 'Data'}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  if (options?.themeXml !== undefined) {
    zip.file('xl/theme/theme1.xml', options.themeXml)
  }
  if (options?.stylesXml !== undefined) {
    zip.file('xl/styles.xml', options.stylesXml)
  }
  if (options?.dropRels !== true) {
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${hasTable ? '<Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' : ''}
</Relationships>`,
    )
  }
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Widget</t></is></c><c r="B2"><v>2</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Gadget</t></is></c><c r="B3"><v>5</v></c></row>
  </sheetData>
  ${hasTable ? '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>' : ''}
</worksheet>`,
  )
  if (hasTable && options?.dropTablePart !== true) {
    zip.file('xl/tables/table1.xml', options!.tableXml!)
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function tableXml(options?: {
  readonly ref?: string
  readonly name?: string
  readonly displayName?: string
  readonly headerRowCount?: string
  readonly totalsRowCount?: string
  readonly columns?: readonly string[]
  readonly styleInfo?: string
}): string {
  const ref = options?.ref ?? 'A1:B3'
  const name = options?.name ?? 'ReportTable'
  const displayName = options?.displayName ?? name
  const columns = (options?.columns ?? ['Item', 'Qty'])
    .map((column, index) => `<tableColumn id="${index + 1}" name="${column}"/>`)
    .join('')
  const headerRowCount =
    options?.headerRowCount === undefined ? '' : ` headerRowCount="${options.headerRowCount}"`
  const totalsRowCount =
    options?.totalsRowCount === undefined ? '' : ` totalsRowCount="${options.totalsRowCount}"`
  const styleInfo =
    options?.styleInfo === undefined
      ? '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
      : options.styleInfo
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${name}" displayName="${displayName}" ref="${ref}"${headerRowCount}${totalsRowCount}>
  <autoFilter ref="${ref}"/>
  <tableColumns count="${(options?.columns ?? ['Item', 'Qty']).length}">${columns}</tableColumns>
  ${styleInfo}
</table>`
}

/// Theme with distinctive accents so palette math is hand-verifiable:
/// dk1 = 000000, accent1 = FF0000 (pure red).
const RED_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
  <a:themeElements>
    <a:clrScheme name="Test">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
      <a:accent3><a:srgbClr val="0000FF"/></a:accent3>
      <a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
      <a:accent5><a:srgbClr val="FF00FF"/></a:accent5>
      <a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`

function addition(overrides: Partial<SheetTableAddition> = {}): SheetTableAddition {
  return {
    sheetName: 'Data',
    area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
    name: 'NewTable',
    columnNames: ['Item', 'Qty'],
    bandedRows: true,
    ...overrides,
  }
}

describe('table read', () => {
  it('parses table metadata through the worksheet rels', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        name: 'SalesTable',
        displayName: 'SalesDisplay',
        ref: 'A1:B3',
        totalsRowCount: '1',
        styleInfo:
          '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="1"/>',
      }),
      themeXml: RED_THEME,
    })
    const imported = await readBasicWorkbook(buffer)
    const tables = imported.snapshot.sheets[0]!.tables
    expect(tables).toBeDefined()
    expect(tables).toHaveLength(1)
    const table = tables![0]!
    expect(table.area).toEqual({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 })
    expect(table.headerRowCount).toBe(1)
    expect(table.totalsRowCount).toBe(1)
    // displayName is the token structured references use; name is a fallback.
    expect(table.name).toBe('SalesDisplay')
    expect(table.columns).toEqual(['Item', 'Qty'])
    expect(table.styleName).toBe('TableStyleMedium2')
    expect(table.showRowStripes).toBe(true)
    expect(table.showColumnStripes).toBe(true)
    // Medium2 with accent1 = FF0000: solid red header, white font,
    // stripe = red tinted 0.8 (#FFCCCC).
    expect(table.headerFill).toBe('#FF0000')
    expect(table.headerFontColor).toBe('#FFFFFF')
    expect(table.stripeFill).toBe('#FFCCCC')
  })

  it('resolves Medium2 against the DEFAULT Office accents without a theme', async () => {
    const buffer = await buildTableFixture({ tableXml: tableXml() })
    const imported = await readBasicWorkbook(buffer)
    const table = imported.snapshot.sheets[0]!.tables![0]!
    expect(table.headerFill).toBe('#4472C4')
    expect(table.headerFontColor).toBe('#FFFFFF')
    // Excel's tint transform applied to accent1 (HSL luminance 0.8).
    expect(table.stripeFill).toBe('#DAE3F3')
  })

  it('applies the Light family rules — unfilled bold accent header + frame', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo: '<tableStyleInfo name="TableStyleLight2" showRowStripes="1"/>',
      }),
      themeXml: RED_THEME,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    // Light 1-7: unfilled bold header in the accent color (accent1 = red).
    expect(table.headerFill).toBeUndefined()
    expect(table.headerFontColor).toBe('#FF0000')
    // Stripe = accent tinted 0.85.
    expect(table.stripeFill).toBe('#FFD9D9')
    // Light 1-7 draw their frame in the base color.
    expect(table.borderColor).toBe('#FF0000')
  })

  it('applies the Dark family rules — darkened header, 0.4 stripe', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo: '<tableStyleInfo name="TableStyleDark2" showRowStripes="1"/>',
      }),
      themeXml: RED_THEME,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerFill).toBe('#BF0000')
    expect(table.headerFontColor).toBe('#FFFFFF')
    expect(table.stripeFill).toBe('#FF6666')
    expect(table.borderColor).toBeUndefined()
  })

  it('applies the Medium full-color block (8-14) — darker 0.6 stripes', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo: '<tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/>',
      }),
      themeXml: RED_THEME,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerFill).toBe('#FF0000')
    expect(table.stripeFill).toBe('#FF9999')
  })

  it('treats unknown style names as Medium2 (desktop parity)', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo: '<tableStyleInfo name="Corporate Blue" showRowStripes="1"/>',
      }),
      themeXml: RED_THEME,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerFill).toBe('#FF0000')
    expect(table.headerFontColor).toBe('#FFFFFF')
    expect(table.stripeFill).toBe('#FFCCCC')
  })

  it('paints nothing for a nameless tableStyleInfo (style "None")', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo:
          '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>',
      }),
      themeXml: RED_THEME,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerFill).toBeUndefined()
    expect(table.headerFontColor).toBeUndefined()
    expect(table.stripeFill).toBeUndefined()
    expect(table.borderColor).toBeUndefined()
    expect(table.styleName).toBeUndefined()
  })

  it('resolves custom tableStyle bands through the styles.xml dxfs', async () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dxfs count="3">
    <dxf><font><color theme="1" tint="0.5"/></font><fill><patternFill><bgColor rgb="FF112233"/></patternFill></fill></dxf>
    <dxf><fill><patternFill><bgColor rgb="FF445566"/></patternFill></fill></dxf>
    <dxf><fill><patternFill><bgColor rgb="FF778899"/></patternFill></fill></dxf>
  </dxfs>
  <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
    <tableStyle name="MyHouseStyle" pivot="0" count="4">
      <tableStyleElement type="wholeTable" dxfId="2"/>
      <tableStyleElement type="headerRow" dxfId="0"/>
      <tableStyleElement type="firstRowStripe" dxfId="1"/>
      <tableStyleElement type="firstColumn" dxfId="2"/>
    </tableStyle>
  </tableStyles>
</styleSheet>`
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo: '<tableStyleInfo name="MyHouseStyle" showFirstColumn="1" showRowStripes="1"/>',
      }),
      themeXml: RED_THEME,
      stylesXml,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerFill).toBe('#112233')
    // theme=1 (dk1 = 000000) tinted 0.5 → #808080.
    expect(table.headerFontColor).toBe('#808080')
    expect(table.stripeFill).toBe('#445566')
    expect(table.wholeTableFill).toBe('#778899')
    // First-column emphasis is gated by showFirstColumn.
    expect(table.firstColumnFill).toBe('#778899')
    // Custom styles carry no frame color.
    expect(table.borderColor).toBeUndefined()
  })

  it('hides first-column emphasis when showFirstColumn is off', async () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dxfs count="1">
    <dxf><fill><patternFill><bgColor rgb="FF778899"/></patternFill></fill></dxf>
  </dxfs>
  <tableStyles count="1">
    <tableStyle name="MyHouseStyle">
      <tableStyleElement type="firstColumn" dxfId="0"/>
      <tableStyleElement type="firstHeaderCell" dxfId="0"/>
    </tableStyle>
  </tableStyles>
</styleSheet>`
    const buffer = await buildTableFixture({
      tableXml: tableXml({
        styleInfo: '<tableStyleInfo name="MyHouseStyle" showFirstColumn="0"/>',
      }),
      stylesXml,
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.firstColumnFill).toBeUndefined()
    expect(table.firstHeaderCellFontColor).toBeUndefined()
  })

  it('parses headerless tables (headerRowCount=0)', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml({ headerRowCount: '0', ref: 'A1:B2' }),
    })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerRowCount).toBe(0)
  })

  it('defaults headerRowCount to 1 and omits totalsRowCount when absent', async () => {
    const buffer = await buildTableFixture({ tableXml: tableXml() })
    const table = (await readBasicWorkbook(buffer)).snapshot.sheets[0]!.tables![0]!
    expect(table.headerRowCount).toBe(1)
    expect(table.totalsRowCount).toBeUndefined()
  })

  it('skips table parts without a readable ref', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml().replace(' ref="A1:B3"', ''),
    })
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.tables).toBeUndefined()
  })

  it('fails closed per sheet when the table relationship is missing', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml(),
      dropRels: true,
    })
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.tables).toBeUndefined()
  })

  it('surfaces no tables for worksheets without tableParts', async () => {
    const buffer = await buildTableFixture()
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.tables).toBeUndefined()
  })
})

describe('table write (wrapper slot)', () => {
  it('appliesCellEditsToXlsx persists tableAdditions end-to-end', async () => {
    const buffer = await buildTableFixture()
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      [],
      null,
      [addition()],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const tableXmlOut = await zip.file('xl/tables/table1.xml')?.async('text')
    expect(tableXmlOut).toBeDefined()
    expect(tableXmlOut).toContain('name="NewTable"')
    expect(tableXmlOut).toContain('ref="A1:B3"')
    expect(tableXmlOut).toContain('showRowStripes="1"')
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(worksheet).toContain('<tableParts count="1">')
    const rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('text')
    expect(rels).toContain('/table')
    const contentTypes = await zip.file('[Content_Types].xml')?.async('text')
    expect(contentTypes).toContain('/xl/tables/table1.xml')
  })
})

describe('table round-trip', () => {
  it('create → save → reopen reads back metadata and the Medium2 palette', async () => {
    const buffer = await buildTableFixture()
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      [],
      null,
      [
        addition({
          name: 'RoundTrip',
          columnNames: ['Item', 'Qty'],
          area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
          style: 'TableStyleMedium2',
          bandedRows: true,
        }),
      ],
    )
    const reopened = await readBasicWorkbook(mutation.buffer)
    const tables = reopened.snapshot.sheets[0]!.tables
    expect(tables).toHaveLength(1)
    const table = tables![0]!
    expect(table.name).toBe('RoundTrip')
    expect(table.area).toEqual({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 })
    expect(table.columns).toEqual(['Item', 'Qty'])
    expect(table.styleName).toBe('TableStyleMedium2')
    expect(table.headerFill).toBe('#4472C4')
    expect(table.headerFontColor).toBe('#FFFFFF')
    expect(table.stripeFill).toBe('#DAE3F3')
    expect(table.showRowStripes).toBe(true)
  })

  it('a no-op save preserves a workbook with tables entry-for-entry', async () => {
    const buffer = await buildTableFixture({
      tableXml: tableXml(),
      themeXml: RED_THEME,
    })
    const mutation = await applyCellEditsToXlsx(buffer, [])
    // The engine's no-op save touches only workbook.xml (calc metadata —
    // pre-existing engine behavior). Every table-related entry — the
    // table part, the worksheet's <tableParts>, the rels, the content-type
    // override, and the theme — keeps its exact bytes (per-entry SHA-256,
    // the canonical preservation proof).
    expect(mutation.touchedEntries).toEqual(['xl/workbook.xml'])
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
    for (const path of [
      'xl/tables/table1.xml',
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
      '[Content_Types].xml',
      'xl/theme/theme1.xml',
    ]) {
      expect(before.has(path), `${path} present before`).toBe(true)
      expect(after.get(path), `${path} preserved`).toBe(before.get(path))
    }
  })
})
