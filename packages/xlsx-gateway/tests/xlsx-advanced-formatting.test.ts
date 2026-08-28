import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'

/**
 * EXCEL-027 — Advanced cell formatting: per-edge borders, line styles,
 * border colors, text rotation, and indentation.
 *
 * These tests prove the full read → write → reopen contract through the REAL
 * gateway entry points (readBasicWorkbook / applyCellEditsToXlsx):
 *
 * READ side (StylesheetReader): per-edge border state (multiple styles and
 * colors on different sides of one cell), textRotation (45 / 135 clockwise /
 * 255 stacked), indent, absent defaults, and the fail-closed policy for
 * unmodelable values (unknown ST_BorderStyle names, rotations outside the
 * OOXML domain, malformed indents) — such values stay in the file's own XML
 * and are simply not modeled.
 *
 * WRITE side (StylesheetEditor via the canonical WorkbookStyleEdit deltas):
 * every one of the 13 OOXML line styles, border colors, side isolation (an
 * edit of one edge preserves the other edges, the diagonal, and every
 * non-border property), border clears (null edges), rotation write/clear/
 * vertical, indent write/clear, the combined
 * font+fill+numberFormat+alignment+border+rotation fixture, and byte-level
 * preservation of an untouched sibling sheet.
 */

/// All 13 OOXML ST_BorderStyle values the canonical model round-trips.
const ALL_LINE_STYLES = [
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
] as const

/**
 * Two-sheet fixture. Sheet "Data" exercises the advanced-formatting matrix:
 *   A1 — border 1 (top thin red + bottom medium blue, no color on top)
 *   B1 — border 2 (left dashed + DIAGONAL hair + diagonalUp/Down attrs)
 *   C1 — rotation 45
 *   D1 — rotation 135 (clockwise down)
 *   E1 — rotation 255 (vertical stacked)
 *   F1 — indent 2
 *   G1 — the COMBINED cell: bold + fill + custom number format + alignment
 *        (center/bottom/wrap) + border 3 (double green all four sides) +
 *        rotation 60 — editing any single property must preserve the rest.
 *   H1 — border 4: an UNKNOWN style ("wave" — not ST_BorderStyle) → the
 *        reader must skip the edge (fail closed for modeling, XML kept).
 *   I1 — rotation 200 (outside the OOXML domain) → not modeled.
 * Sheet "Other" carries one plain cell — the byte-preservation control.
 */
async function buildAdvancedFormattingFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>' +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts>' +
      '<fonts count="2"><font/><font><b/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="5">' +
      '<border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left/><right/><top style="thin"><color rgb="FFC00000"/></top><bottom style="medium"><color rgb="FF0000FF"/></bottom><diagonal/></border>' +
      '<border diagonalDown="1" diagonalUp="1"><left style="dashed"/><right/><top/><bottom/><diagonal style="hair"/></border>' +
      '<border><left style="double"><color rgb="FF00B050"/></left><right style="double"><color rgb="FF00B050"/></right><top style="double"><color rgb="FF00B050"/></top><bottom style="double"><color rgb="FF00B050"/></bottom><diagonal/></border>' +
      '<border><left style="wave"/><right/><top/><bottom/><diagonal/></border>' +
      '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="10">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="45"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="135"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="255"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment indent="2"/></xf>' +
      '<xf numFmtId="164" fontId="1" fillId="1" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="bottom" wrapText="1" textRotation="60"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="200"/></xf>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1">' +
      '<c r="A1" s="1"><v>1</v></c>' +
      '<c r="B1" s="2"><v>2</v></c>' +
      '<c r="C1" s="3"><v>3</v></c>' +
      '<c r="D1" s="4"><v>4</v></c>' +
      '<c r="E1" s="5"><v>5</v></c>' +
      '<c r="F1" s="6"><v>6</v></c>' +
      '<c r="G1" s="7"><v>7</v></c>' +
      '<c r="H1" s="8"><v>8</v></c>' +
      '<c r="I1" s="9"><v>9</v></c>' +
      '</row></sheetData>' +
      '</worksheet>',
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="A1"><v>100</v></c></row></sheetData>' +
      '</worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/** Runs the save path and re-reads the result's snapshot styles. */
async function editedStyles(
  edits: readonly unknown[],
): Promise<Record<string, Record<string, unknown>> | undefined> {
  const mutation = await applyCellEditsToXlsx(
    await buildAdvancedFormattingFixture(),
    edits as never,
  )
  return stylesOf(mutation.buffer)
}

async function stylesOf(
  buffer: Buffer,
): Promise<Record<string, Record<string, unknown>> | undefined> {
  const imported = await readBasicWorkbook(buffer)
  return imported.snapshot.sheets[0]!.styles
}

async function savedStylesXml(edits: readonly unknown[]): Promise<string> {
  const mutation = await applyCellEditsToXlsx(
    await buildAdvancedFormattingFixture(),
    edits as never,
  )
  const zip = await JSZip.loadAsync(mutation.buffer)
  return zip.file('xl/styles.xml')!.async('string')
}

describe('EXCEL-027 reader — per-edge borders', () => {
  it('resolves mixed border styles and colors on different sides of one cell', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    expect(styles!.A1).toEqual({
      border: {
        top: { style: 'thin', color: 'C00000' },
        bottom: { style: 'medium', color: '0000FF' },
      },
    })
  })

  it('resolves single-edge borders and skips the diagonal', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    // B1 carries border 2: left dashed (no color) + diagonal hair +
    // diagonalUp/Down attributes. Only the modeled left edge is exposed.
    expect(styles!.B1).toEqual({ border: { left: { style: 'dashed' } } })
  })

  it('resolves all four sides of a fully bordered cell', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    expect(styles!.G1!.border).toEqual({
      top: { style: 'double', color: '00B050' },
      bottom: { style: 'double', color: '00B050' },
      left: { style: 'double', color: '00B050' },
      right: { style: 'double', color: '00B050' },
    })
  })

  it('fails closed for an unknown ST_BorderStyle name — the edge is not modeled', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    // H1 carries border 4 with style="wave" — not a canonical style name.
    expect(styles!.H1).toBeUndefined()
  })

  it('exposes no border for the default (empty) border', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    // C1's xf has borderId 0 (all edges empty).
    expect(styles!.C1!.border).toBeUndefined()
  })
})

describe('EXCEL-027 reader — text rotation and indent', () => {
  it('resolves counterclockwise, clockwise, and vertical (stacked) rotation', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    expect(styles!.C1).toEqual({ textRotation: 45 })
    expect(styles!.D1).toEqual({ textRotation: 135 })
    expect(styles!.E1).toEqual({ textRotation: 'vertical' })
  })

  it('resolves the indent step count', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    expect(styles!.F1).toEqual({ indent: 2 })
  })

  it('ignores a rotation outside the OOXML domain (200)', async () => {
    const styles = await stylesOf(await buildAdvancedFormattingFixture())
    // I1 carries textRotation="200" — 181..254 is not a legal value; the
    // reader models nothing and the file's own attribute is untouched.
    expect(styles!.I1).toBeUndefined()
  })
})

describe('EXCEL-027 writer — border deltas through the canonical save path', () => {
  it('round-trips every one of the 13 OOXML line styles', async () => {
    for (const style of ALL_LINE_STYLES) {
      const styles = await editedStyles([
        {
          sheetName: 'Data',
          row: 0,
          column: 8, // I1 — a plain, unbordered cell
          writeValue: false,
          cell: { value: null },
          style: { borderTop: { style } },
        },
      ])
      expect(styles!.I1, `style ${style} must round-trip`).toEqual({
        border: { top: { style } },
      })
    }
  })

  it('round-trips the border color', async () => {
    const styles = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 8,
        writeValue: false,
        cell: { value: null },
        style: { borderLeft: { style: 'thick', color: '#FF0055' } },
      },
    ])
    expect(styles!.I1).toEqual({
      border: { left: { style: 'thick', color: 'FF0055' } },
    })
  })

  it('edits ONE side while preserving the other sides, the diagonal, and the border attributes', async () => {
    // B1 carries border 2: left dashed + diagonal hair + diagonalUp/Down.
    // Edit only the RIGHT side; left, diagonal, and the attributes survive.
    const mutation = await applyCellEditsToXlsx(await buildAdvancedFormattingFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: { borderRight: { style: 'medium', color: '#123456' } },
      },
    ])
    const styles = await stylesOf(mutation.buffer)
    expect(styles!.B1!.border).toEqual({
      left: { style: 'dashed' },
      right: { style: 'medium', color: '123456' },
    })
    // The diagonal edge + its diagonalUp/Down attributes ride along verbatim
    // on the newly interned border entry.
    const zip = await JSZip.loadAsync(mutation.buffer)
    const stylesXml = await zip.file('xl/styles.xml')!.async('string')
    expect(stylesXml).toContain('diagonalUp="1"')
    expect(stylesXml).toContain('<diagonal style="hair"/>')
  })

  it('clears one border side with a null edge while preserving the others', async () => {
    // A1 carries top thin red + bottom medium blue. Clear ONLY the top.
    const styles = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { borderTop: null },
      },
    ])
    expect(styles!.A1).toEqual({
      border: { bottom: { style: 'medium', color: '0000FF' } },
    })
  })

  it('clears every side (the No Border preset)', async () => {
    const styles = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: {
          borderTop: null,
          borderBottom: null,
          borderLeft: null,
          borderRight: null,
        },
      },
    ])
    // A1's cellXfs entry derives from a base whose only property was the
    // border; with every edge cleared the resolved format carries nothing.
    expect(styles!.A1).toBeUndefined()
  })

  it('does not disturb unrelated cells or the untouched sibling sheet', async () => {
    const mutation = await applyCellEditsToXlsx(await buildAdvancedFormattingFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 8,
        writeValue: false,
        cell: { value: null },
        style: { borderTop: { style: 'thin' } },
      },
    ])
    const styles = await stylesOf(mutation.buffer)
    // Every pre-existing advanced format survives an unrelated cell's edit.
    expect(styles!.A1).toEqual({
      border: {
        top: { style: 'thin', color: 'C00000' },
        bottom: { style: 'medium', color: '0000FF' },
      },
    })
    expect(styles!.C1).toEqual({ textRotation: 45 })
    expect(styles!.F1).toEqual({ indent: 2 })
    // The untouched sheet's XML is byte-identical.
    const before = new Map(mutation.beforeEntries.map((e) => [e.path, e.sha256]))
    const after = new Map(mutation.afterEntries.map((e) => [e.path, e.sha256]))
    expect(after.get('xl/worksheets/sheet2.xml')).toBe(before.get('xl/worksheets/sheet2.xml'))
  })
})

describe('EXCEL-027 writer — the combined-format fixture (accidental-replacement guard)', () => {
  it('a border edit preserves font, fill, number format, alignment, rotation, and the other edges', async () => {
    // G1 = bold + fill FFF2CC + numFmt 164 (#,##0.000) + center/bottom/wrap
    // + rotation 60 + double green border on all four sides. Edit ONE edge.
    const styles = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 6,
        writeValue: false,
        cell: { value: null },
        style: { borderTop: { style: 'thick', color: '#000000' } },
      },
    ])
    expect(styles!.G1).toEqual({
      bold: true,
      fillColor: 'FFF2CC',
      numberFormat: '#,##0.000',
      horizontalAlign: 'center',
      verticalAlign: 'bottom',
      wrapText: true,
      textRotation: 60,
      border: {
        top: { style: 'thick', color: '000000' },
        bottom: { style: 'double', color: '00B050' },
        left: { style: 'double', color: '00B050' },
        right: { style: 'double', color: '00B050' },
      },
    })
  })

  it('a rotation edit preserves the border and every other property', async () => {
    const styles = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 6,
        writeValue: false,
        cell: { value: null },
        style: { textRotation: 90 },
      },
    ])
    expect(styles!.G1!.textRotation).toBe(90)
    expect(styles!.G1!.border).toEqual({
      top: { style: 'double', color: '00B050' },
      bottom: { style: 'double', color: '00B050' },
      left: { style: 'double', color: '00B050' },
      right: { style: 'double', color: '00B050' },
    })
    expect(styles!.G1!.bold).toBe(true)
    expect(styles!.G1!.numberFormat).toBe('#,##0.000')
  })

  it('a bold edit (the basic family) preserves the border and rotation', async () => {
    const styles = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 6,
        writeValue: false,
        cell: { value: null },
        style: { italic: true },
      },
    ])
    expect(styles!.G1!.italic).toBe(true)
    expect(styles!.G1!.textRotation).toBe(60)
    expect(styles!.G1!.border!.left).toEqual({ style: 'double', color: '00B050' })
  })
})

describe('EXCEL-027 writer — rotation and indent round-trips', () => {
  it('writes, reopens, and re-reads every rotation form (45 / 135 / 255)', async () => {
    for (const [delta, expected] of [
      [{ textRotation: 45 }, 45],
      [{ textRotation: 135 }, 135],
      [{ textRotation: 255 }, 'vertical'],
    ] as const) {
      const styles = await editedStyles([
        {
          sheetName: 'Data',
          row: 0,
          column: 8,
          writeValue: false,
          cell: { value: null },
          style: delta,
        },
      ])
      expect(styles!.I1, `rotation ${JSON.stringify(delta)} must round-trip`).toEqual({
        textRotation: expected,
      })
    }
  })

  it('clears a rotation with the 0 sentinel and writes no textRotation attribute', async () => {
    const mutation = await applyCellEditsToXlsx(await buildAdvancedFormattingFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 2, // C1 carries rotation 45
        writeValue: false,
        cell: { value: null },
        style: { textRotation: 0 },
      },
    ])
    const styles = await stylesOf(mutation.buffer)
    expect(styles!.C1).toBeUndefined()
    // The copy-on-write editor APPENDS a new xf for C1 (the old entry stays
    // for any other cell referencing it) — prove C1's OWN cell now points at
    // an xf whose alignment carries no textRotation.
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    const stylesXml = await zip.file('xl/styles.xml')!.async('string')
    const xfIndex = Number(/<c r="C1" s="(\d+)"/.exec(sheet1)?.[1] ?? '-1')
    expect(xfIndex).toBeGreaterThanOrEqual(0)
    const xfs = [
      ...stylesXml
        .slice(stylesXml.indexOf('<cellXfs'), stylesXml.indexOf('</cellXfs>'))
        .matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g),
    ].map((m) => m[0])
    expect(xfs[xfIndex]).not.toContain('textRotation')
  })

  it('writes and clears the indent', async () => {
    const set = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 8,
        writeValue: false,
        cell: { value: null },
        style: { indent: 3 },
      },
    ])
    expect(set!.I1).toEqual({ indent: 3 })
    const cleared = await editedStyles([
      {
        sheetName: 'Data',
        row: 0,
        column: 5, // F1 carries indent 2
        writeValue: false,
        cell: { value: null },
        style: { indent: 0 },
      },
    ])
    expect(cleared!.F1).toBeUndefined()
  })

  it('keeps rotation and border edits style-only — cell values and formulas untouched', async () => {
    const mutation = await applyCellEditsToXlsx(await buildAdvancedFormattingFixture(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 2,
        writeValue: false,
        cell: { value: null },
        style: { textRotation: 90, borderTop: { style: 'thin' } },
      },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    // C1's stored value 3 survives verbatim next to the new style index.
    expect(sheet1).toMatch(/<c r="C1" s="\d+"><v>3<\/v><\/c>/)
  })
})

describe('EXCEL-027 writer — serialization details', () => {
  it('writes border edges in schema order with the ARGB color', async () => {
    const stylesXml = await savedStylesXml([
      {
        sheetName: 'Data',
        row: 0,
        column: 8,
        writeValue: false,
        cell: { value: null },
        style: {
          borderTop: { style: 'dashDot', color: '#AB12CD' },
          borderBottom: { style: 'slantDashDot' },
          borderLeft: { style: 'hair' },
          borderRight: { style: 'mediumDashed', color: '#001122' },
        },
      },
    ])
    expect(stylesXml).toContain(
      '<left style="hair"/><right style="mediumDashed"><color rgb="FF001122"/></right>' +
        '<top style="dashDot"><color rgb="FFAB12CD"/></top>' +
        '<bottom style="slantDashDot"/>',
    )
  })

  it('interns identical borders — one new entry, not four', async () => {
    const stylesXml = await savedStylesXml([
      {
        sheetName: 'Data',
        row: 0,
        column: 8,
        writeValue: false,
        cell: { value: null },
        style: { borderTop: { style: 'thin', color: '#010203' } },
      },
      {
        sheetName: 'Data',
        row: 1,
        column: 8,
        writeValue: false,
        cell: { value: null },
        style: { borderTop: { style: 'thin', color: '#010203' } },
      },
    ])
    // The fixture's 5 borders + exactly one new interned border entry.
    const bordersSection = /<borders count="(\d+)">([\s\S]*?)<\/borders>/.exec(stylesXml)
    expect(bordersSection).not.toBeNull()
    expect(Number(bordersSection![1])).toBe(6)
    const entries = bordersSection![2].match(/<border\b/g)?.length ?? 0
    expect(entries).toBe(6)
  })
})
