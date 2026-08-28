/**
 * Deterministic office file fixtures for the Playwright browser E2E tests.
 *
 * Built from raw OOXML with JSZip and FIXED zip entry dates, so every run
 * produces byte-identical fixtures (content determinism).
 *
 * These run in the Playwright Node process (NOT in the browser) — they are
 * test data builders, not app code.
 */
import JSZip from 'jszip'

/** Fixed zip entry date — keeps fixture bytes deterministic across runs. */
const FIXED_DATE = new Date(Date.UTC(2024, 0, 2, 3, 4, 5))

function addFile(zip: JSZip, path: string, content: string): void {
  zip.file(path, content, { date: FIXED_DATE, createFolders: false })
}

async function toBytes(zip: JSZip): Promise<Buffer> {
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

// ── XLSX fixture ────────────────────────────────────────────────────────────

/**
 * Deterministic XLSX exercising every fidelity surface the browser editor
 * must preserve:
 *
 *   Sheet "Data" (visible):
 *     A1 — text "Original Text", styled (bold + light-yellow fill, xf 1)
 *     B1 — number 10
 *     C1 — formula =SUM(B1:B1) with cached value (independent of A1 so the
 *          browser's A1 edit never triggers a recalc write-back for C1)
 *     A3 — "Merged Header", merged across A3:B3
 *     row 5 — custom height 30pt (ht + customHeight)
 *     col A — custom width 24
 *   Sheet "HiddenSheet" (state="hidden" in workbook.xml):
 *     A1 — "Hidden Value"
 */
// ── Excel formatting fixture (cell-formatting E2E) ────────────────────────────

/**
 * Deterministic XLSX exercising the cell-formatting surfaces:
 *
 *   Sheet "Formats" (visible):
 *     A1 "Bold"          — bold (font 1)
 *     B1 "Italic"        — italic (font 2)
 *     C1 "Decorated"     — underline + strike + fontColor C00000 (font 3)
 *     D1 "Big red"       — fontSize 14 + fontColor C00000 (font 4)
 *     E1 "Filled"        — solid fill FFD966 (fill 2)
 *     A2 "Centered"      — alignment horizontal=center vertical=center
 *                          wrapText=1 (xf 6)
 *     B2 "Right"         — alignment horizontal=right (xf 7)
 *     A3 "Merged"        — merge A3:B3
 *     B4 "Plain"         — no formatting (control)
 *     row 5              — custom height 30pt
 *     col A              — custom width 24 chars
 *   Sheet "Other" (visible):
 *     A1 "Untouched"     — no formatting (dirty-isolation control)
 */
export async function buildExcelFormatFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Formats" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="8">
  <si><t>Bold</t></si>
  <si><t>Italic</t></si>
  <si><t>Decorated</t></si>
  <si><t>Big red</t></si>
  <si><t>Filled</t></si>
  <si><t>Centered</t></si>
  <si><t>Right</t></si>
  <si><t>Merged</t></si>
  <si><t>Plain</t></si>
  <si><t>Untouched</t></si>
</sst>`,
  )

  // fonts: 0 default, 1 bold, 2 italic, 3 underline+strike+color, 4 size14+color
  // fills: 0 none, 1 gray125, 2 solid FFD966
  // xf:   0 default, 1 font1, 2 font2, 3 font3, 4 font4, 5 fill2, 6 align center/center/wrap, 7 align right
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="5">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><name val="Calibri"/></font>
    <font><u val="single"/><strike/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>
    <font><sz val="14"/><color rgb="FFC00000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFD966"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="2"><v>1</v></c>
      <c r="C1" t="s" s="3"><v>2</v></c>
      <c r="D1" t="s" s="4"><v>3</v></c>
      <c r="E1" t="s" s="5"><v>4</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s" s="6"><v>5</v></c>
      <c r="B2" t="s" s="7"><v>6</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>7</v></c>
    </row>
    <row r="4">
      <c r="B4" t="s"><v>8</v></c>
    </row>
    <row r="5" ht="30" customHeight="1"/>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A3:B3"/>
  </mergeCells>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>9</v></c></row>
  </sheetData>
</worksheet>`,
  )

  return toBytes(zip)
}

// ── Excel formula fixture (formula-bar fidelity E2E) ──────────────────────────

/**
 * Deterministic XLSX exercising the formula-editing surfaces:
 *
 *   Sheet1 (visible):
 *     A1 = 10, A2 = 20, A3 = SUM(A1:A2)          (styled: bold, fill FFF2CC)
 *     B1 = 5, B2 = 7, B3 = B1*B2
 *     C1 = "static", C2 = "Hello " & C1
 *     merge A5:B5, row 5 height 30pt, col A width 24
 *   Sheet2 (visible):
 *     A1 = 100, A2 = Sheet1!A3 + 1               (cross-sheet reference)
 */
export async function buildExcelFormulaFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>static</t></si>
  <si><t>Merged Header</t></si>
</sst>`,
  )

  // xf 1 = bold + fill FFF2CC (the styled formula cell A3)
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><v>5</v></c>
      <c r="C1" t="s"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>20</v></c>
      <c r="B2"><v>7</v></c>
      <c r="C2"><f>"Hello " &amp; C1</f><v>Hello static</v></c>
    </row>
    <row r="3">
      <c r="A3" s="1"><f>SUM(A1:A2)</f><v>30</v></c>
      <c r="B3"><f>B1*B2</f><v>35</v></c>
    </row>
    <row r="5" ht="30" customHeight="1">
      <c r="A5" t="s"><v>1</v></c>
      <c r="B5" s="1"/>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A5:B5"/>
  </mergeCells>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>100</v></c></row>
    <row r="2"><c r="A2"><f>Sheet1!A3 + 1</f><v>31</v></c></row>
  </sheetData>
</worksheet>`,
  )

  return toBytes(zip)
}

export async function buildExcelFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  // Two sheets: "Data" (visible) and "HiddenSheet" (state="hidden").
  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="HiddenSheet" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Original Text</t></si>
  <si><t>Merged Header</t></si>
  <si><t>Hidden Value</t></si>
</sst>`,
  )

  // xf 1 = bold font (fontId 1) + solid light-yellow fill (fillId 2).
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C3B00"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1"><v>10</v></c>
      <c r="C1"><f>SUM(B1:B1)</f><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>1</v></c>
      <c r="B3" s="1"/>
    </row>
    <row r="5" ht="30" customHeight="1"/>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A3:B3"/>
  </mergeCells>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
  )

  return toBytes(zip)
}

// ── Excel sort fidelity fixture (ribbon-data E2E) ───────────────────────────

/**
 * Deterministic XLSX exercising the canonical sort save path's fidelity
 * surfaces — the regression fixture the architect mandated:
 *
 *   Sheet "Sort" (visible):
 *     Row 1 — header (bold + light-yellow fill via xf 1, NOT in sort range)
 *       A1="Name" (bold + fill, xf 1), B1="Qty", C1="Price", D1="Subtotal"
 *     Row 2 — Banana (italic via font 2 → xf 2)
 *       A2="Banana" (italic, xf 2), B2=30, C2=1.50 (currency numfmt via
 *         xf 3 / numFmtId 164), D2=`=B2*C2` (formula with RELATIVE refs,
 *         cached value 45), hyperlink on A2 → rId1
 *     Row 3 — Cherry (regular, xf 0)
 *       A3="Cherry", B3=10, C3=3.00 (currency numfmt, xf 3),
 *         D3=`=B3*C3` (cached value 30)
 *     Row 4 — Apple (bold via font 1 → xf 4)
 *       A4="Apple" (bold, xf 4), B4=20, C4=2.00 (currency numfmt, xf 3),
 *         D4=`=B4*C4` (cached value 40)
 *   Worksheet-level <hyperlinks>:
 *     <hyperlink ref="A2" r:id="rId1"/> (banana hyperlink — Univer's
 *       ReorderRangeMutation does NOT move worksheet-level hyperlink
 *       definitions; the gateway mirrors that, so the hyperlink ref
 *       stays at A2 while the cell content at A2 changes)
 *
 * When the user selects A2:D4 and clicks Sort Asc (alphabetical by
 * column A), the rows reorder to [Apple, Banana, Cherry]. The gateway's
 * `reorder-rows` structural op permutes the <row> blocks atomically —
 * styles (italic on Banana, bold on Apple), numfmt (currency on C
 * column), formula text (verbatim `=B2*C2` etc. — Univer's deepClone
 * does NOT rewrite relative refs, matching Univer's live state), and
 * worksheet-level hyperlink definitions all travel verbatim.
 */
export async function buildExcelSortFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sort" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  // Shared strings: header labels + fruit names (A column values).
  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="7">
  <si><t>Name</t></si>
  <si><t>Qty</t></si>
  <si><t>Price</t></si>
  <si><t>Subtotal</t></si>
  <si><t>Banana</t></si>
  <si><t>Cherry</t></si>
  <si><t>Apple</t></si>
</sst>`,
  )

  // Styles: 5 cellXfs.
  //   xf 0 — default (regular)
  //   xf 1 — header (bold font 1 + light-yellow fill 2)
  //   xf 2 — italic (font 2 italic)
  //   xf 3 — currency numfmt (numFmtId 164, applied via applyNumberFormat)
  //   xf 4 — bold (font 1 bold)
  // Custom numFmt 164 = "$"#,##0.00 (currency).
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C3B00"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  // Worksheet: header row 1, data rows 2-4. The <hyperlinks> element
  // (worksheet-level) carries the banana hyperlink on A2 — Univer's
  // ReorderRangeMutation does NOT move it, and the gateway mirrors that.
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
      <c r="C1" t="s" s="1"><v>2</v></c>
      <c r="D1" t="s" s="1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s" s="2"><v>4</v></c>
      <c r="B2"><v>30</v></c>
      <c r="C2" s="3"><v>1.5</v></c>
      <c r="D2"><f>B2*C2</f><v>45</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>5</v></c>
      <c r="B3"><v>10</v></c>
      <c r="C3" s="3"><v>3</v></c>
      <c r="D3"><f>B3*C3</f><v>30</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s" s="4"><v>6</v></c>
      <c r="B4"><v>20</v></c>
      <c r="C4" s="3"><v>2</v></c>
      <c r="D4"><f>B4*C4</f><v>40</v></c>
    </row>
  </sheetData>
  <hyperlinks count="1">
    <hyperlink ref="A2" r:id="rId1"/>
  </hyperlinks>
</worksheet>`,
  )

  // Worksheet rels — carry the banana hyperlink relationship target.
  addFile(
    zip,
    'xl/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/banana" TargetMode="External"/>
</Relationships>`,
  )

  return toBytes(zip)
}

// ── Excel sort formula-semantics fixture (ribbon-data E2E) ──────────────────

/**
 * Deterministic XLSX for the sort formula-semantics regression gate — the
 * architect-mandated proof that Univer's in-browser formula rewrite and the
 * gateway's verbatim row permutation stay synchronized:
 *
 *   Sheet "SortFx" (visible):
 *     Row 1 — header: A1="Qty", B1="Calc", C1="Plus" (bold, xf 1)
 *             D1 = 5 (constant, OUTSIDE the sort range in both axes)
 *     Row 2 — A2=10,  B2==A2*10   (cached v=100),  C2==A2+$D$1 (cached v=15)
 *     Row 3 — A3=30,  B3==A3*10   (cached v=300),  C3==A3+$D$1 (cached v=35)
 *
 *   Sort A2:C3 DESCENDING by column A (30 before 10) swaps the data rows:
 *     NEW row 2 = OLD row 3: A2=30, B2 must be =A2*10   (rewritten from =A3*10)
 *                                  C2 must be =A2+$D$1 (rewritten from =A3+$D$1 —
 *                                  relative A3→A2 shifted; $D$1 ABSOLUTE untouched)
 *     NEW row 3 = OLD row 2: A3=10, B3 must be =A3*10,  C3 must be =A3+$D$1
 *
 * The two reference classes prove the split responsibility:
 *   - RELATIVE refs to cells that MOVE WITH the sorted row: Univer's
 *     FormulaReorderController rewrites them by the row delta in the browser
 *     (journaled as formula CellEdits), so the same-row relationship survives
 *     — B2 recalculates to 30*10=300, NOT 10*10=100 (the verbatim-text trap).
 *   - ABSOLUTE refs ($D$1) pointing OUTSIDE the sort range: moveFormulaRefOffset
 *     skips AbsoluteRefType.ALL references, and the gateway's reorder-rows
 *     never rewrites formulas at all — the reference stays $D$1 verbatim.
 *
 * If either side regressed (browser rewrite dropped, or gateway started
 * rewriting), B2/C2 would carry the verbatim old text (=A3*10 / =A3+$D$1)
 * and recalculate to the WRONG values (100 / 15) — the assertions catch it.
 */
export async function buildExcelSortFormulaFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="SortFx" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  // Shared strings: header labels only (the data rows are numbers + formulas).
  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Qty</t></si>
  <si><t>Calc</t></si>
  <si><t>Plus</t></si>
</sst>`,
  )

  // xf 1 = bold header font.
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>
</styleSheet>`,
  )

  // D1=5 sits OUTSIDE the sort range (A2:C3) in both axes: row 1 is above
  // the sorted rows AND column D is right of the sorted columns. $D$1 is the
  // absolute-reference probe; the D1 cell itself must never move.
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
      <c r="C1" t="s" s="1"><v>2</v></c>
      <c r="D1"><v>5</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>10</v></c>
      <c r="B2"><f>A2*10</f><v>100</v></c>
      <c r="C2"><f>A2+$D$1</f><v>15</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>30</v></c>
      <c r="B3"><f>A3*10</f><v>300</v></c>
      <c r="C3"><f>A3+$D$1</f><v>35</v></c>
    </row>
  </sheetData>
</worksheet>`,
  )

  return toBytes(zip)
}

// ── Excel AutoFilter fixtures (ribbon-filter E2E) ───────────────────────────

/**
 * Deterministic XLSX for the Data → Filter E2E — a produce-price table with
 * the architect-mandated surfaces:
 *
 *   Sheet "Produce" (visible):
 *     Row 1 (header, bold + fill via xf 1):
 *       A1="Category" B1="Item" C1="Qty" D1="Total"
 *     Rows 2-8 (7 data rows, mixed values + ONE blank):
 *       A2=Fruit  B2=Apple   C2=10  D2==C2*2 (v=20)   — formula
 *       A3=Veg    B3=Carrot  C3=5   D3==C3*2 (v=10)
 *       A4=Fruit  B4=Banana  C4=20  D4==C4*2 (v=40)
 *       A5=Veg    B5=Pea     C5=(blank) D5=(blank)   — the blank row
 *       A6=Fruit  B6=Cherry  C6=15  D6==C6*2 (v=30)
 *       A7=Veg    B7=Kale    C7=8   D7==C7*2 (v=16)
 *       A8=Fruit  B8=Melon   C8=30  D8==C8*2 (v=60)
 *   C column carries the currency numfmt (xf 3, numFmtId 164) on data rows
 *   A2 is italic (xf 2); other category cells are regular — mixed styles.
 *   Worksheet-level hyperlink on B2 → rId1 (https://example.com/produce).
 *
 * buildExcelFilterFixture(): NO autoFilter — the E2E applies one.
 * buildExcelFilteredFixture(): carries
 *   <autoFilter ref="A1:D8"><filterColumn colId="0"><filters>
 *     <filter val="Fruit"/></filters></filterColumn></autoFilter>
 *   with the Veg rows (3, 5, 7 — 1-based) hidden="1" — the state Excel
 *   would store for Category="Fruit".
 */
export async function buildExcelFilterFixture(): Promise<Buffer> {
  return buildProduceFilterFixture(false)
}

export async function buildExcelFilteredFixture(): Promise<Buffer> {
  return buildProduceFilterFixture(true)
}

async function buildProduceFilterFixture(withExistingFilter: boolean): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Produce" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  // Shared strings: headers + categories + items.
  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="13" uniqueCount="13">
  <si><t>Category</t></si>
  <si><t>Item</t></si>
  <si><t>Qty</t></si>
  <si><t>Total</t></si>
  <si><t>Fruit</t></si>
  <si><t>Veg</t></si>
  <si><t>Apple</t></si>
  <si><t>Carrot</t></si>
  <si><t>Banana</t></si>
  <si><t>Pea</t></si>
  <si><t>Cherry</t></si>
  <si><t>Kale</t></si>
  <si><t>Melon</t></si>
</sst>`,
  )

  // xf 1 = header (bold font 1 + fill 2), xf 2 = italic, xf 3 = currency.
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  // Data rows 2-8. The "with filter" variant hides the Veg rows (1-based
  // 3, 5, 7) and carries the Category="Fruit" autoFilter.
  const hiddenVeg = (row: number) =>
    withExistingFilter && (row === 3 || row === 5 || row === 7) ? ' hidden="1"' : ''
  const autoFilter = withExistingFilter
    ? '<autoFilter ref="A1:D8"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>'
    : ''
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
      <c r="C1" t="s" s="1"><v>2</v></c>
      <c r="D1" t="s" s="1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s" s="2"><v>4</v></c>
      <c r="B2" t="s"><v>6</v></c>
      <c r="C2" s="3"><v>10</v></c>
      <c r="D2"><f>C2*2</f><v>20</v></c>
    </row>
    <row r="3"${hiddenVeg(3)}>
      <c r="A3" t="s"><v>5</v></c>
      <c r="B3" t="s"><v>7</v></c>
      <c r="C3" s="3"><v>5</v></c>
      <c r="D3"><f>C3*2</f><v>10</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>4</v></c>
      <c r="B4" t="s"><v>8</v></c>
      <c r="C4" s="3"><v>20</v></c>
      <c r="D4"><f>C4*2</f><v>40</v></c>
    </row>
    <row r="5"${hiddenVeg(5)}>
      <c r="A5" t="s"><v>5</v></c>
      <c r="B5" t="s"><v>9</v></c>
    </row>
    <row r="6">
      <c r="A6" t="s"><v>4</v></c>
      <c r="B6" t="s"><v>10</v></c>
      <c r="C6" s="3"><v>15</v></c>
      <c r="D6"><f>C6*2</f><v>30</v></c>
    </row>
    <row r="7"${hiddenVeg(7)}>
      <c r="A7" t="s"><v>5</v></c>
      <c r="B7" t="s"><v>11</v></c>
      <c r="C7" s="3"><v>8</v></c>
      <c r="D7"><f>C7*2</f><v>16</v></c>
    </row>
    <row r="8">
      <c r="A8" t="s"><v>4</v></c>
      <c r="B8" t="s"><v>12</v></c>
      <c r="C8" s="3"><v>30</v></c>
      <c r="D8"><f>C8*2</f><v>60</v></c>
    </row>
  </sheetData>
  ${autoFilter}
  <hyperlinks count="1">
    <hyperlink ref="B2" r:id="rId1"/>
  </hyperlinks>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/produce" TargetMode="External"/>
</Relationships>`,
  )

  return toBytes(zip)
}

// ── Excel data-validation fixtures (ribbon-data-validation E2E) ─────────────

/**
 * Deterministic XLSX for the Data → Data Validation E2E — a survey sheet:
 *
 *   Sheet "Survey" (visible):
 *     Row 1 (header, bold + fill via xf 1): A1="Count" B1="Choice" C1="Check"
 *     Rows 2-6 (5 data rows):
 *       A2=5   B2=Apple   C2=1
 *       A3=120 B3=Carrot  C3=0
 *       A4=50  B4=Banana  C4=1
 *       A5=(blank — allowBlank probe) B5=Pea  C5=0
 *       A6=7   B6=Kale    C6=1
 *   Currency numfmt (xf 3) on the C data cells; hyperlink on B2 → rId1.
 *
 * buildExcelDvFixture(): NO validations — the E2E creates them.
 * buildExcelDvExistingFixture(): carries THREE validations:
 *   A2:A6 whole between 1..100 (with error message) — A3=120 violates it
 *   B2:B6 list "Fruit,Vegetable,Grain" (dropdown shown)
 *   C2:C6 custom =ISNUMBER(C2)
 */
export async function buildExcelDvFixture(): Promise<Buffer> {
  return buildSurveyDvFixture(false)
}

export async function buildExcelDvExistingFixture(): Promise<Buffer> {
  return buildSurveyDvFixture(true)
}

async function buildSurveyDvFixture(withValidations: boolean): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Survey" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="9">
  <si><t>Count</t></si>
  <si><t>Choice</t></si>
  <si><t>Check</t></si>
  <si><t>Apple</t></si>
  <si><t>Carrot</t></si>
  <si><t>Banana</t></si>
  <si><t>Pea</t></si>
  <si><t>Kale</t></si>
</sst>`,
  )

  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  const validations = withValidations
    ? `  <dataValidations count="3">
    <dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" errorTitle="Bad count" error="Enter 1-100" sqref="A2:A6"><formula1>1</formula1><formula2>100</formula2></dataValidation>
    <dataValidation type="list" sqref="B2:B6"><formula1>"Fruit,Vegetable,Grain"</formula1></dataValidation>
    <dataValidation type="custom" sqref="C2:C6"><formula1>ISNUMBER(C2)</formula1></dataValidation>
  </dataValidations>
`
    : ''
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
      <c r="C1" t="s" s="1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>5</v></c>
      <c r="B2" t="s"><v>3</v></c>
      <c r="C2" s="2"><v>1</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>120</v></c>
      <c r="B3" t="s"><v>4</v></c>
      <c r="C3" s="2"><v>0</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>50</v></c>
      <c r="B4" t="s"><v>5</v></c>
      <c r="C4" s="2"><v>1</v></c>
    </row>
    <row r="5">
      <c r="B5" t="s"><v>6</v></c>
      <c r="C5" s="2"><v>0</v></c>
    </row>
    <row r="6">
      <c r="A6"><v>7</v></c>
      <c r="B6" t="s"><v>7</v></c>
      <c r="C6" s="2"><v>1</v></c>
    </row>
  </sheetData>
${validations}  <hyperlinks count="1">
    <hyperlink ref="B2" r:id="rId1"/>
  </hyperlinks>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/survey" TargetMode="External"/>
</Relationships>`,
  )

  return toBytes(zip)
}

// ── Excel notes fixture (ribbon-review-notes E2E) ────────────────────────────

/**
 * Deterministic XLSX for the Review → Notes/Comments E2E — a small ledger:
 *
 *   Sheet "Ledger" (visible):
 *     Row 1 (header, bold + fill via xf 1): A1="Item" B1="Amount"
 *     Rows 2-5: A2=Fee B2=10 / A3=Tax B3=5 / A4=Tip B4=2 / A5=Total B5=17
 *
 * buildExcelNotesFixture(): NO comments part — the E2E creates notes.
 * buildExcelNotedFixture(): carries a comments part (xl/comments1.xml) with
 *   TWO notes wired through the worksheet rels + a VML drawing with Note
 *   shapes (the full Excel presentation chain the writer also emits):
 *     B2 → author "Reviewer", text "Verify the fee <amount> & tax"
 *     A4 → author "", text "second note"
 */
export async function buildExcelNotesFixture(): Promise<Buffer> {
  return buildLedgerNotesFixture(false)
}

export async function buildExcelNotedFixture(): Promise<Buffer> {
  return buildLedgerNotesFixture(true)
}

async function buildLedgerNotesFixture(withNotes: boolean): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  ${withNotes ? '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' : ''}
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Ledger" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">
  <si><t>Item</t></si>
  <si><t>Amount</t></si>
  <si><t>Fee</t></si>
  <si><t>Tax</t></si>
  <si><t>Tip</t></si>
  <si><t>Total</t></si>
</sst>`,
  )

  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/></cellXfs>
</styleSheet>`,
  )

  const noteRels = withNotes
    ? `  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
`
    : ''
  addFile(
    zip,
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${noteRels}</Relationships>`,
  )

  const legacyDrawing = withNotes ? '<legacyDrawing r:id="rId6"/>' : ''
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><v>5</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>4</v></c>
      <c r="B4"><v>2</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s"><v>5</v></c>
      <c r="B5"><v>17</v></c>
    </row>
  </sheetData>
  ${legacyDrawing}
</worksheet>`,
  )

  if (withNotes) {
    addFile(
      zip,
      'xl/comments1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Reviewer</author><author></author></authors>
  <commentList>
    <comment ref="B2" authorId="0"><text><t>Verify the fee &lt;amount&gt; &amp; tax</t></text></comment>
    <comment ref="A4" authorId="1"><text><t>second note</t></text></comment>
  </commentList>
</comments>`,
    )
    addFile(
      zip,
      'xl/drawings/vmlDrawing1.vml',
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2,15,1,2,5,15,5,2</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>
<v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;z-index:2;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1,15,3,2,4,15,7,2</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>3</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
</xml>`,
    )
  }

  return toBytes(zip)
}

// ── DOCX fixture ────────────────────────────────────────────────────────────

/**
 * Deterministic DOCX exercising the browser Word editor's block surfaces:
 *
 *   idx 0 — plain paragraph "This is a plain paragraph."
 *   idx 1 — Heading1 "Fixture Document Heading"
 *   idx 2 — paragraph with bold + italic runs
 *   idx 3 — paragraph with NESTED marks (bold + bold-italic)
 *   idx 4 — bullet list item "First bullet item"  (numId 1)
 *   idx 5 — bullet list item "Second bullet item" (numId 1)
 *   idx 6 — ordered list item "First ordered item"  (numId 2)
 *   idx 7 — ordered list item "Second ordered item" (numId 2)
 *   trailing w:sectPr (hidden block — engine re-appends automatically)
 */
export async function buildWordFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
</w:styles>`,
  )

  // numId 1 = bullet, numId 2 = decimal (matches the blank template ids the
  // docx-engine's generated blocks reference).
  addFile(
    zip,
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`,
  )

  const body = `
    <w:p><w:r><w:t xml:space="preserve">This is a plain paragraph.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Fixture Document Heading</w:t></w:r></w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">This paragraph has </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold text</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">italic text</w:t></w:r>
      <w:r><w:t xml:space="preserve">.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">Nested marks: </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">outer bold </w:t></w:r>
      <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">bold italic</w:t></w:r>
    </w:p>
    <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t xml:space="preserve">First bullet item</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t xml:space="preserve">Second bullet item</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t xml:space="preserve">First ordered item</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t xml:space="preserve">Second ordered item</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`

  addFile(
    zip,
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${body}
  </w:body>
</w:document>`,
  )

  return toBytes(zip)
}

// ── Saved-file verification helpers (unzip + inspect real bytes) ────────────

/** Read a text entry from an in-memory zip (Buffer). */
export async function readZipEntry(buffer: Buffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  if (!entry) throw new Error(`zip entry not found: ${path}`)
  return entry.async('string')
}

/** Read a zip entry as base64 (for binary parts like word/media/*.png). */
export async function readZipEntryBase64(buffer: Buffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  if (!entry) throw new Error(`zip entry not found: ${path}`)
  return entry.async('base64')
}

/** List all zip entry paths (for structural assertions). */
export async function listZipEntries(buffer: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer)
  return Object.keys(zip.files).filter((p) => !zip.files[p].dir)
}

// ── Word table fixture (editable-table E2E) ─────────────────────────────────

/**
 * Deterministic DOCX exercising the editable-table path:
 *
 *   idx 0 — plain paragraph "This is a plain paragraph."
 *   idx 1 — Heading1 "Table Fixture Heading"
 *   idx 2 — 2×3 table:
 *             row 1: [vMerge restart: bold/italic runs] [fill FFF2CC, center] [plain]
 *             row 2: [vMerge continue]                 [plain]              [plain]
 *   idx 3 — trailing plain paragraph "Paragraph after the table."
 *   trailing w:sectPr (hidden)
 *
 * The vertical merge spans both rows of column 1; the merged cell carries
 * rich run content; cell (1,2) has a fill and centered paragraph.
 */
// ── Deterministic PNG builder (for the image fixtures) ───────────────────────

/** CRC32 (IEEE) for PNG chunks — deterministic, no zlib dependency. */
const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/** Adler-32 for the zlib stream (PNG IDAT uses stored/uncompressed deflate). */
function adler32(bytes: Uint8Array): number {
  let a = 1
  let b = 0
  for (const byte of bytes) {
    a = (a + byte) % 65521
    b = (b + a) % 65521
  }
  return ((b << 16) | a) >>> 0
}

function u32(v: number): Uint8Array {
  return new Uint8Array([(v >>> 24) & 0xff, (v >>> 16) & 0xff, (v >>> 8) & 0xff, v & 0xff])
}

/** One PNG chunk: length + type + data + CRC. */
function pngChunk(type: string, data: Uint8Array): Buffer {
  const typeBytes = Buffer.from(type, 'latin1')
  const crcBytes = u32(crc32(Buffer.concat([typeBytes, Buffer.from(data)])))
  return Buffer.concat([u32(data.length), typeBytes, Buffer.from(data), crcBytes])
}

/**
 * Build a deterministic solid-color PNG (no zlib compression — stored
 * deflate blocks). w/h ≤ 32 keeps it tiny; color is [r,g,b].
 */
export function buildSolidPng(w: number, h: number, rgb: [number, number, number]): Buffer {
  const ihdr = Buffer.concat([
    u32(w),
    u32(h),
    new Uint8Array([8, 2, 0, 0, 0]), // 8-bit truecolor RGB
  ])
  // Raw scanlines: filter byte 0 + w RGB pixels
  const raw: number[] = []
  for (let y = 0; y < h; y++) {
    raw.push(0)
    for (let x = 0; x < w; x++) raw.push(rgb[0], rgb[1], rgb[2])
  }
  const rawBytes = new Uint8Array(raw)
  // zlib stream with stored (uncompressed) deflate blocks
  const maxStored = 65535
  const blocks: number[] = [0x78, 0x01]
  for (let off = 0; off < rawBytes.length; off += maxStored) {
    const end = Math.min(off + maxStored, rawBytes.length)
    const isLast = end === rawBytes.length ? 1 : 0
    const len = end - off
    blocks.push(isLast, 0, len & 0xff, (len >> 8) & 0xff, ~len & 0xff, (~len >> 8) & 0xff)
    for (let i = off; i < end; i++) blocks.push(rawBytes[i])
  }
  blocks.push(...u32(adler32(rawBytes)))
  const idat = new Uint8Array(blocks)
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', idat),
    pngChunk('IEND', new Uint8Array(0)),
  ])
}

// ── Word image fixture (editable-image E2E) ──────────────────────────────────

const EMU = 9525 // per CSS px

function imageDrawingXml(
  rId: string,
  opts: {
    cx: number
    cy: number
    floating?: boolean
    wrap?: 'square' | 'topBottom' | 'none' | 'behind'
    posOffsetX?: number
    posOffsetY?: number
    rotDeg?: number
    flipH?: boolean
    flipV?: boolean
    srcRect?: { l: number; t: number; r: number; b: number }
    align?: 'center' | 'right'
    docPrId: number
    /** accessibility alt text (wp:docPr descr); absent when undefined */
    descr?: string
  },
): string {
  const { cx, cy } = opts
  let xfrmAttrs = ''
  if (opts.rotDeg) xfrmAttrs += ` rot="${opts.rotDeg * 60000}"`
  if (opts.flipH) xfrmAttrs += ' flipH="1"'
  if (opts.flipV) xfrmAttrs += ' flipV="1"'
  const srcRect = opts.srcRect
    ? `<a:srcRect l="${Math.round(opts.srcRect.l * 100000)}" t="${Math.round(opts.srcRect.t * 100000)}" r="${Math.round(opts.srcRect.r * 100000)}" b="${Math.round(opts.srcRect.b * 100000)}"/>`
    : ''
  const pic = `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${opts.docPrId}" name="Picture ${opts.docPrId}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
  const graphic = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">${pic}</a:graphicData></a:graphic>`
  const jc = opts.align ? `<w:pPr><w:jc w:val="${opts.align}"/></w:pPr>` : ''
  const descrAttr =
    opts.descr !== undefined
      ? ` descr="${opts.descr.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')}"`
      : ''
  const docPr = `<wp:docPr id="${opts.docPrId}" name="Picture ${opts.docPrId}"${descrAttr}/>`
  if (!opts.floating) {
    return `<w:p>${jc}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>${docPr}${graphic}</wp:inline></w:drawing></w:r></w:p>`
  }
  const behind = opts.wrap === 'behind' ? '1' : '0'
  let wrapEl = '<wp:wrapNone/>'
  if (opts.wrap === 'square') wrapEl = '<wp:wrapSquare wrapText="bothSides"/>'
  else if (opts.wrap === 'topBottom') wrapEl = '<wp:wrapTopAndBottom/>'
  const posH = `<wp:positionH relativeFrom="column"><wp:posOffset>${opts.posOffsetX ?? 0}</wp:posOffset></wp:positionH>`
  const posV = `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${opts.posOffsetY ?? 0}</wp:posOffset></wp:positionV>`
  return `<w:p>${jc}<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${posH}${posV}${wrapEl}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${docPr}${graphic}</wp:anchor></w:drawing></w:r></w:p>`
}

/**
 * Deterministic DOCX exercising the editable-image surfaces:
 *
 *   idx 0 — plain paragraph "First paragraph."
 *   idx 1 — INLINE image (red 16×16, 64×64 px)
 *   idx 2 — FLOATING image (green 16×16, square wrap left, posOffset)
 *   idx 3 — NON-DEFAULT SIZE image (blue 16×16, 240×120 px)
 *   idx 4 — WRAPPING image (cyan 16×16, topBottom wrap)
 *   idx 5 — ROTATED image (magenta 16×16, 90°)
 *   idx 6 — FLIPPED image (yellow 16×16, flipH + flipV)
 *   idx 7 — CROPPED image (gray 16×16, srcRect l=10% t=20% r=10% b=20%)
 *   idx 8 — trailing plain paragraph "Last paragraph."
 *   trailing w:sectPr (hidden)
 *
 * Media: three PNG parts (image1/2/3.png) shared across drawings — the
 * canonical parser resolves each blip to a data URL.
 */
export async function buildWordImageFixture(): Promise<Buffer> {
  const zip = new JSZip()
  const png1 = buildSolidPng(16, 16, [200, 30, 30]) // red
  const png2 = buildSolidPng(16, 16, [30, 160, 60]) // green
  const png3 = buildSolidPng(16, 16, [40, 60, 200]) // blue
  const png4 = buildSolidPng(16, 16, [30, 170, 170]) // cyan
  const png5 = buildSolidPng(16, 16, [190, 40, 170]) // magenta
  const png6 = buildSolidPng(16, 16, [210, 200, 40]) // yellow
  const png7 = buildSolidPng(16, 16, [128, 128, 128]) // gray

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rIdImg2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
  <Relationship Id="rIdImg3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/>
  <Relationship Id="rIdImg4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image4.png"/>
  <Relationship Id="rIdImg5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image5.png"/>
  <Relationship Id="rIdImg6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image6.png"/>
  <Relationship Id="rIdImg7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image7.png"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`,
  )

  // Seven image paragraphs: inline / floating square / big size / topBottom
  // wrap / rotated / flipped / cropped. Images 1, 3, 5 carry accessibility
  // alt text (wp:docPr descr) so the alt-render/edit/clear E2E + unit tests
  // have real descr to assert; images 2, 4, 6, 7 have none (verifies the
  // absent-alt path round-trips without picking up a spurious descr).
  const imgs = [
    imageDrawingXml('rIdImg1', { cx: 64 * EMU, cy: 64 * EMU, docPrId: 101, descr: 'A red square' }),
    imageDrawingXml('rIdImg2', {
      cx: 80 * EMU,
      cy: 80 * EMU,
      floating: true,
      wrap: 'square',
      posOffsetX: 200000,
      posOffsetY: 100000,
      docPrId: 102,
    }),
    imageDrawingXml('rIdImg3', {
      cx: 240 * EMU,
      cy: 120 * EMU,
      docPrId: 103,
      descr: 'Wide blue banner',
    }),
    imageDrawingXml('rIdImg4', {
      cx: 100 * EMU,
      cy: 60 * EMU,
      floating: true,
      wrap: 'topBottom',
      docPrId: 104,
    }),
    imageDrawingXml('rIdImg5', {
      cx: 70 * EMU,
      cy: 70 * EMU,
      rotDeg: 90,
      docPrId: 105,
      descr: 'Rotated magenta',
    }),
    imageDrawingXml('rIdImg6', {
      cx: 70 * EMU,
      cy: 70 * EMU,
      flipH: true,
      flipV: true,
      docPrId: 106,
    }),
    imageDrawingXml('rIdImg7', {
      cx: 90 * EMU,
      cy: 90 * EMU,
      srcRect: { l: 0.1, t: 0.2, r: 0.1, b: 0.2 },
      docPrId: 107,
    }),
  ]

  const body = `
    <w:p><w:r><w:t xml:space="preserve">First paragraph.</w:t></w:r></w:p>
    ${imgs.join('\n')}
    <w:p><w:r><w:t xml:space="preserve">Last paragraph.</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`

  addFile(
    zip,
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>${body}
  </w:body>
</w:document>`,
  )

  zip.file('word/media/image1.png', png1, { date: FIXED_DATE, createFolders: false })
  zip.file('word/media/image2.png', png2, { date: FIXED_DATE, createFolders: false })
  zip.file('word/media/image3.png', png3, { date: FIXED_DATE, createFolders: false })
  zip.file('word/media/image4.png', png4, { date: FIXED_DATE, createFolders: false })
  zip.file('word/media/image5.png', png5, { date: FIXED_DATE, createFolders: false })
  zip.file('word/media/image6.png', png6, { date: FIXED_DATE, createFolders: false })
  zip.file('word/media/image7.png', png7, { date: FIXED_DATE, createFolders: false })

  return toBytes(zip)
}

export async function buildWordTableFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
</w:styles>`,
  )

  addFile(
    zip,
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`,
  )

  const border = (name: string) => `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const tableXml = `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
      <w:p>
        <w:r><w:t xml:space="preserve">Merged </w:t></w:r>
        <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold</w:t></w:r>
        <w:r><w:t xml:space="preserve"> and </w:t></w:r>
        <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">italic</w:t></w:r>
        <w:r><w:t xml:space="preserve"> cell</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Filled center</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t xml:space="preserve">Top right</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:vMerge/></w:tcPr>
      <w:p/>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t xml:space="preserve">Bottom middle</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t xml:space="preserve">Bottom right</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>`

  const body = `
    <w:p><w:r><w:t xml:space="preserve">This is a plain paragraph.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Table Fixture Heading</w:t></w:r></w:p>
    ${tableXml}
    <w:p><w:r><w:t xml:space="preserve">Paragraph after the table.</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`

  addFile(
    zip,
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${body}
  </w:body>
</w:document>`,
  )

  return toBytes(zip)
}

/**
 * Deterministic DOCX exercising the remaining inline marks:
 *
 *   idx 0 — paragraph: underlined run, struck run, plain run
 *   idx 1 — paragraph with an external hyperlink (rId10) around one run
 *   idx 2 — Heading1 (control block)
 *   trailing w:sectPr (hidden)
 *
 * The hyperlink relationship lives in word/_rels/document.xml.rels
 * (TargetMode="External") — the parser resolves href from it.
 */
export async function buildWordMarksFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  )

  // rId10 = external hyperlink relationship (resolved to run.link.href on parse).
  addFile(
    zip,
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/fixture" TargetMode="External"/>
</Relationships>`,
  )

  addFile(
    zip,
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>
    <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`,
  )

  addFile(
    zip,
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`,
  )

  const body = `
    <w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Underlined words </w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve">struck words </w:t></w:r><w:r><w:t xml:space="preserve">plain words.</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">Before link </w:t></w:r><w:hyperlink r:id="rId10"><w:r><w:t xml:space="preserve">linked text</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> after link.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Marks Fixture Heading</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`

  addFile(
    zip,
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${body}
  </w:body>
</w:document>`,
  )

  return toBytes(zip)
}

// ── Excel Remove Duplicates fixture (EXCEL-018 E2E) ─────────────────────────

/**
 * Deterministic XLSX for the EXCEL-018 Remove Duplicates save/reopen
 * verification matrix. Proves the canonical cell-edit round-trip:
 *   - basic duplicate rows (header preserved when hasHeader=true)
 *   - multiple selected columns act as the comparison key
 *   - duplicate rows with styles (header is bold+filled; data rows have
 *     per-column styles that must survive the rewrite)
 *   - duplicate rows with formulas (a formula row that gets overwritten
 *     by a moved row's computed value — the formula is replaced with
 *     the literal; an unchanged row's formula survives)
 *   - save/reopen yields the exact expected rows
 *
 * Sheet "Dedupe" (visible):
 *   Row 1 — header: A1="Name" (s=1 bold+fill), B1="Qty" (s=1)
 *   Row 2 — A2="Apple"  (s=2 italic),  B2=10
 *   Row 3 — A3="Apple"  (s=2 italic),  B3=10   ← full duplicate of row 2
 *   Row 4 — A4="Banana" (s=4 bold),    B4=20
 *   Row 5 — A5="Apple"  (s=2 italic),  B5=10   ← full duplicate of row 2
 *   Row 6 — A6="Cherry" (s=0 default), B6=30
 *   Row 7 — A7="Apple"  (s=2 italic),  B7==B6  (formula, cached v=30)
 *                                          ← NOT a duplicate (B7 result is 30, B2 result is 10)
 *
 * Dedupe with hasHeader=true on A1:B7 must keep rows 1, 2, 4, 6, 7
 * and blank rows 3, 5. Total of 2 rows removed. The header at row 1
 * is preserved verbatim. The style on A1/B1 (xf 1) survives because
 * the rewrite does not touch styles.
 */
export async function buildExcelDedupeFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Dedupe" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  // Shared strings: header labels + fruit names.
  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="5">
  <si><t>Name</t></si>
  <si><t>Qty</t></si>
  <si><t>Apple</t></si>
  <si><t>Banana</t></si>
  <si><t>Cherry</t></si>
</sst>`,
  )

  // Styles:
  //   xf 0 — default
  //   xf 1 — header (bold font 1 + light-yellow fill 2)
  //   xf 2 — italic (font 2 italic)
  //   xf 4 — bold (font 1 bold)  [matches the sort fixture's xf 4]
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C3B00"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  // Worksheet: 7 rows. The dedupe A1:B7 with hasHeader=true must keep
  // rows 1,2,4,6,7 and blank rows 3,5. Row 7 carries a formula
  // =B6 with cached value 30 — its B-column RESULT (30) differs from
  // row 2's B-column value (10), so row 7 is NOT a duplicate of row 2.
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s" s="2"><v>2</v></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s" s="2"><v>2</v></c>
      <c r="B3"><v>10</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s" s="4"><v>3</v></c>
      <c r="B4"><v>20</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s" s="2"><v>2</v></c>
      <c r="B5"><v>10</v></c>
    </row>
    <row r="6">
      <c r="A6" t="s"><v>4</v></c>
      <c r="B6"><v>30</v></c>
    </row>
    <row r="7">
      <c r="A7" t="s" s="2"><v>2</v></c>
      <c r="B7"><f>B6</f><v>30</v></c>
    </row>
  </sheetData>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Deterministic XLSX for the EXCEL-018 architect's mandatory second
 * regression case — formula preservation across reference types
 * (absolute / relative / mixed / external-to-selection).
 *
 * The selection is A1:B7 (the dedupe range). The dedupe MUST preserve
 * the formulas on the survivor row that gets compacted (Banana at
 * row 4 → row 3), and the gateway MUST rewrite every reference type
 * to track the moved referenced cell (Cherry at row 6 → row 4 after
 * the two dedupes; the Anchor cell at D6 → D4).
 *
 * Sheet "DedupeMixed" (visible):
 *   Row 1 — header: A1="Name" (s=1), B1="Qty" (s=1)
 *   Row 2 — A2="Apple"  (s=2 italic), B2=10
 *   Row 3 — A3="Apple"  (s=2 italic), B3=10   ← full duplicate of row 2 (DELETE)
 *   Row 4 — A4="Banana" (s=4 bold),    B4=20
 *                                     C4="=$D$6"  ← absolute ref to D6 (outside selection)
 *                                     D4="=A6"    ← relative ref to A6 (inside selection)
 *                                     E4="=$A6"   ← mixed: col $, row relative — to A6
 *                                     F4="=A$6"   ← mixed: col relative, row $ — to A6
 *   Row 5 — A5="Apple"  (s=2 italic), B5=10   ← full duplicate of row 2 (DELETE)
 *   Row 6 — A6="Cherry" (s=0),        B6=30
 *                                     D6="Anchor" ← target of C4's $D$6 (outside selection)
 *   Row 7 — A7="Apple"  (s=2 italic), B7=10   ← full duplicate of row 2 (DELETE)
 *
 * Dedupe with hasHeader=true on A1:B7 must DELETE rows 3, 5, 7 (3
 * duplicates). The runtime issues ws.deleteRows(startRow+6, 1),
 * ws.deleteRows(startRow+4, 1), ws.deleteRows(startRow+2, 1) in
 * DESCENDING order. The survivor at row 4 (Banana, with the four
 * formulas in Cols C-F) compacts to row 3; the survivor at row 6
 * (Cherry, the target of the formulas) compacts to row 4. The
 * gateway's transformFormulas rewrites every reference:
 *   - $D$6 → $D$4   (absolute: $ preserved, row 6→4 below deletions)
 *   - A6 → A4       (relative: row 6→4)
 *   - $A6 → $A4     (mixed col $, row relative: $ preserved, row 6→4)
 *   - A$6 → A$4     (mixed col relative, row $: $ preserved, row 6→4)
 *
 * After dedupe:
 *   Row 1: Name, Qty (header preserved)
 *   Row 2: Apple, 10 (first occurrence preserved)
 *   Row 3: Banana, 20 with formulas C3="=$D$4" D3="=A4" E3="=$A4" F3="=A$4"
 *   Row 4: Cherry, 30 with D4="Anchor" (was D6)
 *
 * The formulas' COMPUTED VALUES all become "Cherry" (from A4) and
 * "Anchor" (from D4) — but the test must distinguish the formula
 * TEXT surviving from the computed result happening to be the same.
 */
export async function buildExcelDedupeMixedReferencesFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="DedupeMixed" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  // Shared strings: header labels + fruit names + the Anchor cell text.
  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="6">
  <si><t>Name</t></si>
  <si><t>Qty</t></si>
  <si><t>Apple</t></si>
  <si><t>Banana</t></si>
  <si><t>Cherry</t></si>
  <si><t>Anchor</t></si>
</sst>`,
  )

  // Styles — same 5 cellXfs as the basic dedupe fixture (parity).
  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C3B00"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  )

  // Worksheet: 7 rows in cols A,B + an Anchor cell at D6 + four
  // formulas at C4,D4,E4,F4 (on the survivor Banana row). Row 4's
  // B-column value (20) differs from row 2's B-column value (10), so
  // row 4 is NOT a duplicate of row 2 — kept.
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s" s="2"><v>2</v></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s" s="2"><v>2</v></c>
      <c r="B3"><v>10</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s" s="4"><v>3</v></c>
      <c r="B4"><v>20</v></c>
      <c r="C4"><f>$D$6</f><v>Anchor</v></c>
      <c r="D4"><f>A6</f><v>Cherry</v></c>
      <c r="E4"><f>$A6</f><v>Cherry</v></c>
      <c r="F4"><f>A$6</f><v>Cherry</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s" s="2"><v>2</v></c>
      <c r="B5"><v>10</v></c>
    </row>
    <row r="6">
      <c r="A6" t="s"><v>4</v></c>
      <c r="B6"><v>30</v></c>
      <c r="D6" t="s"><v>5</v></c>
    </row>
    <row r="7">
      <c r="A7" t="s" s="2"><v>2</v></c>
      <c r="B7"><v>10</v></c>
    </row>
  </sheetData>
</worksheet>`,
  )

  return toBytes(zip)
}

// ── EXCEL-020 Protection fixtures ────────────────────────────────────────────

/**
 * Deterministic XLSX for the Review → Protection E2E (EXCEL-020) — the
 * notes fixture's ledger shape (header row + 4 data rows, shared strings,
 * styles with cellXfs) with three variants:
 *
 *   buildExcelProtectionFixture():     NO protection elements (protect
 *                                      from scratch, editable-vs-locked).
 *   buildExcelProtectedFixture():      worksheet <sheetProtection sheet="1"
 *                                      objects="1" scenarios="1"/> + workbook
 *                                      <workbookProtection lockStructure="1"/>
 *                                      (read + unprotect + reopen).
 *   buildExcelPasswordFixture():      password-bearing elements (legacy hash
 *                                      form) on BOTH levels — the negative
 *                                      authorization cases (unprotect must be
 *                                      refused up front; the gateway fails
 *                                      closed on the write too).
 */
export async function buildExcelProtectionFixture(): Promise<Buffer> {
  return buildProtectionLedgerFixture({})
}

export async function buildExcelProtectedFixture(): Promise<Buffer> {
  return buildProtectionLedgerFixture({
    sheetProtection: '<sheetProtection sheet="1" objects="1" scenarios="1"/>',
    workbookProtection: '<workbookProtection lockStructure="1"/>',
  })
}

export async function buildExcelPasswordFixture(): Promise<Buffer> {
  return buildProtectionLedgerFixture({
    sheetProtection: '<sheetProtection sheet="1" password="83AF"/>',
    workbookProtection: '<workbookProtection lockStructure="1" workbookPassword="83AF"/>',
  })
}

async function buildProtectionLedgerFixture(options: {
  readonly sheetProtection?: string
  readonly workbookProtection?: string
}): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${options.workbookProtection ?? ''}
  <sheets>
    <sheet name="Ledger" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">
  <si><t>Item</t></si>
  <si><t>Amount</t></si>
  <si><t>Fee</t></si>
  <si><t>Tax</t></si>
  <si><t>Tip</t></si>
  <si><t>Total</t></si>
</sst>`,
  )

  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>
</styleSheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><v>5</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>4</v></c>
      <c r="B4"><v>2</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s"><v>5</v></c>
      <c r="B5"><v>17</v></c>
    </row>
  </sheetData>
  ${options.sheetProtection ?? ''}
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * EXCEL-021 — table fixtures.
 *
 *   buildExcelTableFixture():        one existing table (SalesTable,
 *                                    TableStyleMedium2, header + 3 data
 *                                    rows, autoFilter INSIDE the table part
 *                                    so the worksheet carries none — the
 *                                    sheet's filter origin IS the table).
 *   buildExcelTableCreateFixture():  the same ledger WITHOUT any table —
 *                                    the create-from-scratch cases.
 */
export async function buildExcelTableFixture(): Promise<Buffer> {
  return buildTableLedgerFixture({ withTable: true })
}

export async function buildExcelTableCreateFixture(): Promise<Buffer> {
  return buildTableLedgerFixture({ withTable: false })
}

async function buildTableLedgerFixture(options: { readonly withTable: boolean }): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  ${options.withTable ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ''}
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Ledger" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">
  <si><t>Item</t></si>
  <si><t>Amount</t></si>
  <si><t>Fee</t></si>
  <si><t>Tax</t></si>
  <si><t>Tip</t></si>
  <si><t>Total</t></si>
</sst>`,
  )

  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>
</styleSheet>`,
  )

  if (options.withTable) {
    addFile(
      zip,
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
    )
    addFile(
      zip,
      'xl/tables/table1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:B4" totalsRowCount="0" totalsRowShown="0">
  <autoFilter ref="A1:B4"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Amount"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`,
    )
  }

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s" s="1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><v>5</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>4</v></c>
      <c r="B4"><v>2</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s"><v>5</v></c>
      <c r="B5"><v>17</v></c>
    </row>
  </sheetData>
  ${options.withTable ? '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>' : ''}
</worksheet>`,
  )

  return toBytes(zip)
}

// ── Worksheet image fixtures (EXCEL-022) ─────────────────────────────────────

/** Minimal valid 1×1 JPEG (baseline, no EXIF) — deterministic bytes. */
const MINIMAL_JPEG_BASE64 =
  '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwcJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD3+iiigD//2Q=='

/** One spreadsheetDrawing anchor carrying a picture. */
function sheetDrawingAnchorXml(options: {
  kind: 'twoCellAnchor' | 'oneCellAnchor' | 'absoluteAnchor'
  embedId: string
  name?: string
  from?: { col: number; colOff: number; row: number; rowOff: number }
  to?: { col: number; colOff: number; row: number; rowOff: number }
  ext?: { cx: number; cy: number }
}): string {
  const marker = (m: { col: number; colOff: number; row: number; rowOff: number }) =>
    `<xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff>` +
    `<xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff>`
  const from = options.from ?? { col: 1, colOff: 0, row: 2, rowOff: 0 }
  const to = options.to ?? { col: 6, colOff: 0, row: 12, rowOff: 0 }
  const ext = options.ext ?? { cx: 190500, cy: 95250 }
  const name = options.name ?? 'Picture'
  const head =
    options.kind === 'absoluteAnchor'
      ? `<xdr:pos x="47625" y="9525"/><xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>`
      : `<xdr:from>${marker(from)}</xdr:from>` +
        (options.kind === 'twoCellAnchor' ? `<xdr:to>${marker(to)}</xdr:to>` : '') +
        (options.kind === 'oneCellAnchor' ? `<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>` : '')
  const pic =
    '<xdr:pic><xdr:nvPicPr>' +
    `<xdr:cNvPr id="2" name="${name}"/>` +
    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>' +
    '</xdr:nvPicPr>' +
    `<xdr:blipFill><a:blip r:embed="${options.embedId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
  return `<xdr:${options.kind}>${head}${pic}<xdr:clientData/></xdr:${options.kind}>`
}

interface SheetImageSpec {
  anchors: Array<{
    kind: 'twoCellAnchor' | 'oneCellAnchor' | 'absoluteAnchor'
    embedId: string
    name?: string
    from?: { col: number; colOff: number; row: number; rowOff: number }
    to?: { col: number; colOff: number; row: number; rowOff: number }
  }>
  rels: string[]
}

/** Builds an image-carrying workbook: sharedStrings-free single- or dual-sheet. */
async function buildSheetImageFixture(options: {
  sheets: Array<{ name: string; images: SheetImageSpec }>
  media: Record<string, Buffer | string>
  contentTypes: string
}): Promise<Buffer> {
  const zip = new JSZip()
  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${options.contentTypes}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${options.sheets
    .map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('')}
  ${options.sheets
    .map(
      (_, index) =>
        `<Override PartName="/xl/drawings/drawing${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
    )
    .join('')}
</Types>`,
  )
  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${options.sheets
      .map(
        (sheet, index) =>
          `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
      )
      .join('\n    ')}
  </sheets>
</workbook>`,
  )
  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${options.sheets
    .map(
      (_, index) =>
        `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
    )
    .join('\n  ')}
</Relationships>`,
  )
  options.sheets.forEach((sheet, index) => {
    const number = index + 1
    addFile(
      zip,
      `xl/worksheets/sheet${number}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1"><v>${number * 10}</v></c></row>
  </sheetData>
  <drawing r:id="rIdDrw1"/>
</worksheet>`,
    )
    addFile(
      zip,
      `xl/worksheets/_rels/sheet${number}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${number}.xml"/>
</Relationships>`,
    )
    addFile(
      zip,
      `xl/drawings/drawing${number}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${sheet.images.anchors.map((anchor) => sheetDrawingAnchorXml(anchor)).join('\n')}
</xdr:wsDr>`,
    )
    addFile(
      zip,
      `xl/drawings/_rels/drawing${number}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheet.images.rels.join('\n')}
</Relationships>`,
    )
  })
  for (const [path, content] of Object.entries(options.media)) {
    zip.file(path, content, { createFolders: false })
  }
  return toBytes(zip)
}

const IMAGE_REL = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${target}"/>`

const PNG_DEFAULT = '<Default Extension="png" ContentType="image/png"/>'

/** One PNG picture (red 16×16) with a two-cell anchor B3 → G13. */
export async function buildExcelImageFixture(): Promise<Buffer> {
  return buildSheetImageFixture({
    sheets: [
      {
        name: 'Data',
        images: {
          anchors: [{ kind: 'twoCellAnchor', embedId: 'rId1', name: 'Red dot' }],
          rels: [IMAGE_REL('rId1', 'image1.png')],
        },
      },
    ],
    media: { 'xl/media/image1.png': buildSolidPng(16, 16, [200, 40, 40]) },
    contentTypes: PNG_DEFAULT,
  })
}

/** Two PNG pictures on one sheet (isolation) — red + blue, distinct media. */
export async function buildExcelImagesFixture(): Promise<Buffer> {
  return buildSheetImageFixture({
    sheets: [
      {
        name: 'Data',
        images: {
          anchors: [
            { kind: 'twoCellAnchor', embedId: 'rId1', name: 'Red dot' },
            {
              kind: 'twoCellAnchor',
              embedId: 'rId2',
              name: 'Blue dot',
              from: { col: 8, colOff: 0, row: 4, rowOff: 0 },
              to: { col: 13, colOff: 0, row: 14, rowOff: 0 },
            },
          ],
          rels: [IMAGE_REL('rId1', 'image1.png'), IMAGE_REL('rId2', 'image2.png')],
        },
      },
    ],
    media: {
      'xl/media/image1.png': buildSolidPng(16, 16, [200, 40, 40]),
      'xl/media/image2.png': buildSolidPng(16, 16, [40, 40, 200]),
    },
    contentTypes: PNG_DEFAULT,
  })
}

/** One PNG per sheet across two sheets. */
export async function buildExcelMultiSheetImageFixture(): Promise<Buffer> {
  const images = (embedId: string): SheetImageSpec => ({
    anchors: [{ kind: 'twoCellAnchor', embedId, name: 'Dot' }],
    rels: [IMAGE_REL(embedId, 'image1.png')],
  })
  return buildSheetImageFixture({
    sheets: [
      { name: 'First', images: images('rId1') },
      { name: 'Second', images: images('rId1') },
    ],
    // BOTH pictures share one media part — cross-sheet reference parity.
    media: { 'xl/media/image1.png': buildSolidPng(16, 16, [40, 160, 40]) },
    contentTypes: PNG_DEFAULT,
  })
}

/** One JPEG picture (writer-supported type parity). */
export async function buildExcelJpegImageFixture(): Promise<Buffer> {
  return buildSheetImageFixture({
    sheets: [
      {
        name: 'Data',
        images: {
          anchors: [{ kind: 'twoCellAnchor', embedId: 'rId1', name: 'Photo' }],
          rels: [IMAGE_REL('rId1', 'image1.jpeg')],
        },
      },
    ],
    media: { 'xl/media/image1.jpeg': Buffer.from(MINIMAL_JPEG_BASE64, 'base64') },
    contentTypes: '<Default Extension="jpeg" ContentType="image/jpeg"/>',
  })
}

/** One oneCellAnchor picture — read-only, edits fail closed. */
export async function buildExcelOneCellImageFixture(): Promise<Buffer> {
  return buildSheetImageFixture({
    sheets: [
      {
        name: 'Data',
        images: {
          anchors: [{ kind: 'oneCellAnchor', embedId: 'rId1', name: 'One cell' }],
          rels: [IMAGE_REL('rId1', 'image1.png')],
        },
      },
    ],
    media: { 'xl/media/image1.png': buildSolidPng(16, 16, [120, 40, 160]) },
    contentTypes: PNG_DEFAULT,
  })
}

/**
 * One absoluteAnchor picture — omitted from the browser model (architect
 * review, PR #20 blocker 2): fixed-sheet geometry has no two-cell
 * representation, so the reader must fail closed and never relocate it.
 */
export async function buildExcelAbsoluteImageFixture(): Promise<Buffer> {
  return buildSheetImageFixture({
    sheets: [
      {
        name: 'Data',
        images: {
          anchors: [{ kind: 'absoluteAnchor', embedId: 'rId1', name: 'Absolute' }],
          rels: [IMAGE_REL('rId1', 'image1.png')],
        },
      },
    ],
    media: { 'xl/media/image1.png': buildSolidPng(16, 16, [40, 120, 160]) },
    contentTypes: PNG_DEFAULT,
  })
}

/** Read a ZIP entry as raw bytes (binary-safe). */
export async function readZipEntryBytes(buffer: Buffer, path: string): Promise<Buffer | null> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  return entry === null ? null : ((await entry.async('nodebuffer')) as Buffer)
}

// ── EXCEL-023: chart-carrying fixtures ─────────────────────────────────────

const CHART_REL = (id: string, target: string): string =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/${target}"/>`

interface ChartAnchorSpec {
  kind: 'twoCellAnchor' | 'oneCellAnchor' | 'absoluteAnchor'
  chartRelId?: string
  name?: string
  from?: { col: number; colOff: number; row: number; rowOff: number }
  to?: { col: number; colOff: number; row: number; rowOff: number }
  ext?: { cx: number; cy: number }
  picture?: boolean
}

function chartAnchorXml(spec: ChartAnchorSpec): string {
  const marker = (m: { col: number; colOff: number; row: number; rowOff: number }) =>
    `<xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff>` +
    `<xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff>`
  const from = spec.from ?? { col: 1, colOff: 0, row: 2, rowOff: 0 }
  const to = spec.to ?? { col: 8, colOff: 0, row: 18, rowOff: 0 }
  const ext = spec.ext ?? { cx: 914400, cy: 609600 }
  const relId = spec.chartRelId ?? 'rIdChart1'
  const name = spec.name ?? 'Chart 1'
  const body = spec.picture
    ? '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Pic"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip r:embed="rIdPic"/></xdr:blipFill>' +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></a:xfrm></xdr:spPr></xdr:pic>`
    : '<xdr:graphicFrame macro="">' +
      `<xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="${name}"/></xdr:nvGraphicFramePr></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></xdr:xfrm>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/>` +
      '</a:graphicData></a:graphic></xdr:graphicFrame>'
  const head =
    spec.kind === 'absoluteAnchor'
      ? `<xdr:pos x="47625" y="9525"/><xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>`
      : `<xdr:from>${marker(from)}</xdr:from>` +
        (spec.kind === 'twoCellAnchor' ? `<xdr:to>${marker(to)}</xdr:to>` : '') +
        (spec.kind === 'oneCellAnchor' ? `<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>` : '')
  return `<xdr:${spec.kind}>${head}${body}<xdr:clientData/></xdr:${spec.kind}>`
}

function chartSeriesXml(
  name: string,
  values: readonly number[],
  options: {
    color?: string
    valuesRef?: string
  } = {},
): string {
  const nameRef =
    '<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/>' +
    `<c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
  const spPr =
    options.color !== undefined
      ? `<c:spPr><a:solidFill><a:srgbClr val="${options.color}"/></a:solidFill></c:spPr>`
      : ''
  const valuesRef = options.valuesRef ?? 'Data!$B$2:$B$4'
  const points = values
    .map((value, idx) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`)
    .join('')
  return (
    `<c:ser><c:idx val="0"/><c:order val="0"/>${nameRef}${spPr}` +
    '<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>' +
    '<c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>' +
    '<c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
    `<c:val><c:numRef><c:f>${valuesRef}</c:f><c:numCache><c:formatCode>General</c:formatCode>` +
    `<c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef></c:val></c:ser>`
  )
}

function barChartXml(
  series: readonly string[],
  options: { title?: string; legendPos?: string; gapWidth?: number } = {},
): string {
  const title =
    options.title !== undefined
      ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
        `<a:t>${options.title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
      : ''
  const legend =
    options.legendPos !== undefined
      ? `<c:legend><c:legendPos val="${options.legendPos}"/></c:legend>`
      : ''
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:chart>' +
    title +
    '<c:autoTitleDeleted val="0"/>' +
    '<c:plotArea><c:layout/>' +
    '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
    '<c:varyColors val="0"/>' +
    series.join('') +
    `<c:gapWidth val="${options.gapWidth ?? 150}"/>` +
    '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
    '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>' +
    '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>' +
    '</c:plotArea>' +
    legend +
    '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
  )
}

function pieChartXml(series: readonly string[], title?: string): string {
  const titleElement =
    title !== undefined
      ? '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r>' +
        `<a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
      : ''
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:chart>' +
    titleElement +
    '<c:plotArea><c:layout/>' +
    '<c:pieChart><c:varyColors val="1"/>' +
    series.join('') +
    '<c:firstSliceAng val="0"/></c:pieChart></c:plotArea>' +
    '<c:legend><c:legendPos val="r"/></c:legend>' +
    '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
  )
}

function threeDeeChartXml(series: readonly string[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:chart><c:plotArea><c:layout/>' +
    '<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
    series.join('') +
    '<c:axId val="111"/><c:axId val="222"/></c:bar3DChart>' +
    '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
  )
}

/**
 * Builds a chart-carrying workbook: one 'Data' sheet with typed cell data
 * (headers B1/C1, categories A2:A4, values B2:C4 as inline strings +
 * numbers) plus a drawing part with configurable chart anchors and chart
 * parts.
 */
async function buildSheetChartFixture(options: {
  anchors: ChartAnchorSpec[]
  chartParts: Record<string, string>
  chartRels: string[]
  dataRows?: boolean
}): Promise<Buffer> {
  const zip = new JSZip()
  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  ${Object.keys(options.chartParts)
    .map(
      (path) =>
        `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    )
    .join('\n  ')}
</Types>`,
  )
  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  const data =
    options.dataRows === false
      ? ''
      : `
    <row r="1"><c r="A1" t="inlineStr"><is><t>Quarter</t></is></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c><c r="C1" t="inlineStr"><is><t>Cost</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>5</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>8</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Q3</t></is></c><c r="B4"><v>30</v></c><c r="C4"><v>11</v></c></row>`
  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>${data}</sheetData>
  <drawing r:id="rIdDrw1"/>
</worksheet>`,
  )
  addFile(
    zip,
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
  )
  addFile(
    zip,
    'xl/drawings/drawing1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${options.anchors
      .map((anchor) => chartAnchorXml(anchor))
      .join('')}</xdr:wsDr>`,
  )
  addFile(
    zip,
    'xl/drawings/_rels/drawing1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${options.chartRels.join('')}</Relationships>`,
  )
  for (const [path, xml] of Object.entries(options.chartParts)) {
    addFile(zip, path, xml)
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/** One two-cell bar chart (2 series with refs + caches). */
export async function buildExcelChartFixture(): Promise<Buffer> {
  return buildSheetChartFixture({
    anchors: [{ kind: 'twoCellAnchor', chartRelId: 'rIdChart1', name: 'Sales chart' }],
    chartParts: {
      'xl/charts/chart1.xml': barChartXml(
        [
          chartSeriesXml('Revenue', [10, 20, 30], { color: '4472C4' }),
          chartSeriesXml('Cost', [5, 8, 11], { color: 'ED7D31', valuesRef: 'Data!$C$2:$C$4' }),
        ],
        { title: 'Sales &amp; Cost', legendPos: 'r' },
      ),
    },
    chartRels: [CHART_REL('rIdChart1', 'chart1.xml')],
  })
}

/** Two charts on one sheet (bar + pie) — isolation evidence. */
export async function buildExcelChartsFixture(): Promise<Buffer> {
  return buildSheetChartFixture({
    anchors: [
      { kind: 'twoCellAnchor', chartRelId: 'rIdChart1', name: 'Sales chart' },
      {
        kind: 'twoCellAnchor',
        chartRelId: 'rIdChart2',
        name: 'Share chart',
        from: { col: 10, colOff: 0, row: 2, rowOff: 0 },
        to: { col: 16, colOff: 0, row: 18, rowOff: 0 },
      },
    ],
    chartParts: {
      'xl/charts/chart1.xml': barChartXml(
        [chartSeriesXml('Revenue', [10, 20, 30], { color: '4472C4' })],
        { title: 'Revenue', legendPos: 'b' },
      ),
      'xl/charts/chart2.xml': pieChartXml([chartSeriesXml('Share', [40, 35, 25])], 'Share'),
    },
    chartRels: [CHART_REL('rIdChart1', 'chart1.xml'), CHART_REL('rIdChart2', 'chart2.xml')],
  })
}

/** One supported bar chart + one 3-D chart (unsupported — omitted). The
 *  3-D chart is the FIRST anchor: the survivor keeps drawingIndex 1,
 *  proving the omitted anchor still counts toward locator parity. */
export async function buildExcel3DChartFixture(): Promise<Buffer> {
  return buildSheetChartFixture({
    anchors: [
      { kind: 'twoCellAnchor', chartRelId: 'rIdChart2', name: 'Unsupported 3D' },
      {
        kind: 'twoCellAnchor',
        chartRelId: 'rIdChart1',
        name: 'Supported',
        from: { col: 10, colOff: 0, row: 2, rowOff: 0 },
        to: { col: 16, colOff: 0, row: 18, rowOff: 0 },
      },
    ],
    chartParts: {
      'xl/charts/chart1.xml': barChartXml(
        [chartSeriesXml('Revenue', [10, 20, 30], { color: '4472C4' })],
        { title: 'Revenue' },
      ),
      'xl/charts/chart2.xml': threeDeeChartXml([chartSeriesXml('Cost', [5, 8, 11])]),
    },
    chartRels: [CHART_REL('rIdChart1', 'chart1.xml'), CHART_REL('rIdChart2', 'chart2.xml')],
  })
}

/** One oneCellAnchor chart — move-only (resize fails closed). */
export async function buildExcelOneCellChartFixture(): Promise<Buffer> {
  return buildSheetChartFixture({
    anchors: [{ kind: 'oneCellAnchor', chartRelId: 'rIdChart1', ext: { cx: 476250, cy: 285750 } }],
    chartParts: {
      'xl/charts/chart1.xml': barChartXml([chartSeriesXml('Revenue', [10, 20, 30])], {
        title: 'One cell',
      }),
    },
    chartRels: [CHART_REL('rIdChart1', 'chart1.xml')],
  })
}

/** One absoluteAnchor chart (omitted) + one supported two-cell chart. */
export async function buildExcelAbsoluteChartFixture(): Promise<Buffer> {
  return buildSheetChartFixture({
    anchors: [
      { kind: 'absoluteAnchor', chartRelId: 'rIdChart1', name: 'Absolute' },
      {
        kind: 'twoCellAnchor',
        chartRelId: 'rIdChart2',
        name: 'Supported',
        from: { col: 10, colOff: 0, row: 2, rowOff: 0 },
        to: { col: 16, colOff: 0, row: 18, rowOff: 0 },
      },
    ],
    chartParts: {
      'xl/charts/chart1.xml': barChartXml([chartSeriesXml('A', [1, 2, 3])], { title: 'Absolute' }),
      'xl/charts/chart2.xml': barChartXml([chartSeriesXml('Revenue', [10, 20, 30])], {
        title: 'Supported',
      }),
    },
    chartRels: [CHART_REL('rIdChart1', 'chart1.xml'), CHART_REL('rIdChart2', 'chart2.xml')],
  })
}

/** Data-only workbook (A1:C4 typed cells, no charts) — insert-chart input. */
export async function buildExcelChartDataOnlyFixture(): Promise<Buffer> {
  return buildSheetChartFixture({
    anchors: [],
    chartParts: {},
    chartRels: [],
  })
}

// ── Excel conditional-formatting fixtures (EXCEL-024 E2E) ─────────────────────

/**
 * Deterministic XLSX exercising the conditional-formatting matrix.
 *
 *   Sheet "Data" (visible):
 *     A1 "Name"    A2:A6 text (alpha / urgent / gamma / delta / epsilon)
 *     B1 "Score"   B2:B6 numbers (10 / 55 / 80 / 3 / 95)
 *     C1 "Scale"   C2:C6 numbers 1..5 (color scale range C2:C6)
 *     D1 "Bar"     D2:D6 numbers (data bar)
 *     E1 "Icons"   E2:E6 numbers (icon set)
 *     F1 "Dup"     F2:F6 with a duplicate value
 *     G1 "Top"     G2:G6 numbers (top-10 rule)
 *     H1 "Avg"     H2:H6 numbers (above-average rule)
 *     I1 "Expr"    I2:I6 numbers (expression rule =B2>60)
 *     J1 "Blank"   J2:J6 with blanks (containsBlanks rule)
 *     K1 "Multi"   K2:K6 + M2:M6 numbers (multi-range, 2 rules with
 *                  stopIfTrue on the higher-priority one)
 *
 *   CF rules (dxfs: 0 = bold red font on pink fill, 1 = italic green font,
 *   2 = yellow fill):
 *     B2:B6  cellIs greaterThan 50        dxf 0  priority 1
 *     A2:A6  containsText "urgent"        dxf 1  priority 2
 *     C2:C6  3-color scale (min/50th/max)         priority 3
 *     D2:D6  data bar (min/max, FF638EC6)          priority 4
 *     E2:E6  iconSet 3TrafficLights1 (0/33/67)    priority 5
 *     F2:F6  duplicateValues             dxf 2  priority 6
 *     G2:G6  top10 rank 2                 dxf 0  priority 7
 *     H2:H6  aboveAverage                 dxf 2  priority 8
 *     I2:I6  expression =B2>60            dxf 1  priority 9
 *     J2:J6  containsBlanks               dxf 2  priority 10
 *     K2:K6 M2:M6  cellIs lessThan 5      dxf 2  priority 11 stopIfTrue
 *     K2:K6 M2:M6  cellIs greaterThan 0   dxf 1  priority 12
 *
 *   Sheet "Other" (visible): A1 "Untouched"; B2:B4 cellIs greaterThan 10
 *   (dxf 0) — the unrelated-sheet preservation control.
 */
export async function buildExcelCfFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="18" uniqueCount="18">
  <si><t>Name</t></si><si><t>alpha</t></si><si><t>urgent</t></si><si><t>gamma</t></si><si><t>delta</t></si><si><t>epsilon</t></si>
  <si><t>Score</t></si><si><t>Scale</t></si><si><t>Bar</t></si><si><t>Icons</t></si><si><t>Dup</t></si>
  <si><t>Top</t></si><si><t>Avg</t></si><si><t>Expr</t></si><si><t>Blank</t></si><si><t>Multi</t></si>
  <si><t>x</t></si><si><t>Untouched</t></si>
</sst>`,
  )

  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf/></cellXfs>
  <dxfs count="3">
    <dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
    <dxf><font><i/><color rgb="FF006100"/></font></dxf>
    <dxf><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>
  </dxfs>
</styleSheet>`,
  )

  const cfSections = `
  <conditionalFormatting sqref="B2:B6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="A2:A6"><cfRule type="containsText" dxfId="1" priority="2" text="urgent" operator="containsText"><formula>NOT(ISERROR(SEARCH("urgent",A2)))</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="C2:C6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="D2:D6"><cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="E2:E6"><cfRule type="iconSet" priority="5"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="num" val="33"/><cfvo type="num" val="67"/></iconSet></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="F2:F6"><cfRule type="duplicateValues" dxfId="2" priority="6"/></conditionalFormatting>
  <conditionalFormatting sqref="G2:G6"><cfRule type="top10" dxfId="0" priority="7" rank="2"/></conditionalFormatting>
  <conditionalFormatting sqref="H2:H6"><cfRule type="aboveAverage" dxfId="2" priority="8"/></conditionalFormatting>
  <conditionalFormatting sqref="I2:I6"><cfRule type="expression" dxfId="1" priority="9"><formula>B2&gt;60</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="J2:J6"><cfRule type="containsBlanks" dxfId="2" priority="10"><formula>LEN(TRIM(J2))=0</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="K2:K6 M2:M6"><cfRule type="cellIs" dxfId="2" priority="11" stopIfTrue="1" operator="lessThan"><formula>5</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="12" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>6</v></c><c r="C1" t="s"><v>7</v></c><c r="D1" t="s"><v>8</v></c><c r="E1" t="s"><v>9</v></c><c r="F1" t="s"><v>10</v></c><c r="G1" t="s"><v>11</v></c><c r="H1" t="s"><v>12</v></c><c r="I1" t="s"><v>13</v></c><c r="J1" t="s"><v>14</v></c><c r="K1" t="s"><v>15</v></c><c r="M1" t="s"><v>15</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>10</v></c><c r="C2"><v>1</v></c><c r="D2"><v>5</v></c><c r="E2"><v>3</v></c><c r="F2" t="s"><v>16</v></c><c r="G2"><v>1</v></c><c r="H2"><v>10</v></c><c r="I2"><v>10</v></c><c r="K2"><v>4</v></c><c r="M2"><v>2</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3"><v>55</v></c><c r="C3"><v>3</v></c><c r="D3"><v>3</v></c><c r="E3"><v>2</v></c><c r="F3" t="s"><v>16</v></c><c r="G3"><v>2</v></c><c r="H3"><v>20</v></c><c r="I3"><v>20</v></c><c r="K3"><v>2</v></c><c r="M3"><v>6</v></c></row>
    <row r="4"><c r="A4" t="s"><v>3</v></c><c r="B4"><v>80</v></c><c r="C4"><v>5</v></c><c r="D4"><v>8</v></c><c r="E4"><v>1</v></c><c r="F4" t="s"><v>17</v></c><c r="G4"><v>3</v></c><c r="H4"><v>30</v></c><c r="I4"><v>30</v></c><c r="J4" t="s"><v>17</v></c><c r="K4"><v>7</v></c><c r="M4"><v>1</v></c></row>
    <row r="5"><c r="A5" t="s"><v>4</v></c><c r="B5"><v>3</v></c><c r="C5"><v>2</v></c><c r="D5"><v>1</v></c><c r="E5"><v>2</v></c><c r="F5" t="s"><v>17</v></c><c r="G5"><v>4</v></c><c r="H5"><v>40</v></c><c r="I5"><v>40</v></c><c r="K5"><v>9</v></c><c r="M5"><v>3</v></c></row>
    <row r="6"><c r="A6" t="s"><v>5</v></c><c r="B6"><v>95</v></c><c r="C6"><v>4</v></c><c r="D6"><v>9</v></c><c r="E6"><v>3</v></c><c r="F6" t="s"><v>16</v></c><c r="G6"><v>5</v></c><c r="H6"><v>50</v></c><c r="I6"><v>50</v></c><c r="K6"><v>6</v></c><c r="M6"><v>8</v></c></row>
  </sheetData>${cfSections}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>17</v></c></row>
    <row r="2"><c r="B2"><v>5</v></c></row>
    <row r="3"><c r="B3"><v>15</v></c></row>
    <row r="4"><c r="B4"><v>25</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Fail-closed fixture: Sheet "Data" carries one representable rule PLUS an
 * x14-linked data bar (cfRule extLst) and a timePeriod rule — the sheet
 * must open with cfLocked (no rules installed, edits refused) while
 * Sheet "Other" still parses its own rule.
 */
export async function buildExcelCfLockedFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Untouched</t></si></sst>`,
  )

  addFile(
    zip,
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf/></cellXfs>
  <dxfs count="1">
    <dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
  </dxfs>
</styleSheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>5</v></c><c r="C2"><v>45000</v></c></row>
    <row r="3"><c r="A3"><v>60</v></c><c r="B3"><v>8</v></c><c r="C3"><v>44000</v></c></row>
    <row r="4"><c r="A4"><v>30</v></c><c r="B4"><v>2</v></c><c r="C4"><v>46000</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A2:A4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>25</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="B2:B4"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-BA67-4E7D1D9ED3A3}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{F7C31AE2-4BEF-4D70-8C5B-56A5B3D9E1A2}</x14:id></ext></extLst></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="C2:C4"><cfRule type="timePeriod" dxfId="0" priority="3" timePeriod="yesterday"><formula>FLOOR(C2,1)=TODAY()-1</formula></cfRule></conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="B2"><v>5</v></c></row>
    <row r="3"><c r="B3"><v>15</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

// ── EXCEL-025: defined-names fixtures ───────────────────────────────────────

/**
 * Workbook carrying the full name classification matrix:
 *
 * - GlobalTotal (workbook scope, Data!$A$1:$A$5) — modeled, resolvable,
 *   consumed by the B7 formula =SUM(GlobalTotal).
 * - LocalTotal (sheet scope on "Other", Other!$B$2:$B$4) — modeled.
 * - Print_Titles (_xlnm built-in, sheet-scoped) — auto-preserved.
 * - SecretRate (hidden) — auto-preserved.
 * - A1 (cell-ref lookalike) — reader preserve list.
 * - Excel_Version ×2 (workbook + sheet scope duplicates) — ranked winner
 *   modeled, loser preserved.
 *
 * Column A carries five numbers the names point at; B7 carries the
 * name-consuming formula.
 */
export async function buildExcelNamesFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>
    <definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>
    <definedName name="LocalTotal" localSheetId="1">Other!$B$2:$B$4</definedName>
    <definedName name="SecretRate" hidden="1">Data!$C$1</definedName>
    <definedName name="A1">Data!$A$1</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="C1"><v>99</v></c><c r="D1"><v>2026</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c></row>
    <row r="3"><c r="A3"><v>30</v></c></row>
    <row r="4"><c r="A4"><v>40</v></c></row>
    <row r="5"><c r="A5"><v>50</v></c></row>
    <row r="7"><c r="B7"><f>SUM(GlobalTotal)</f></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c r="B2"><v>5</v></c></row>
    <row r="3"><c r="B3"><v>15</v></c></row>
    <row r="4"><c r="B4"><v>25</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Fail-closed fixture: the workbook's <definedNames> section carries an
 * element with no name attribute — structurally unparseable. The workbook
 * must open with namesLocked (no names installed, Name Manager edits
 * refused) and a no-op save must preserve the section byte-for-byte.
 */
export async function buildExcelNamesLockedFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="GlobalTotal">Data!$A$1</definedName>
    <definedName localSheetId="0">Data!$A$2</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Collision fixture: a GENUINE same-scope duplicate — the name
 * "Excel_Version" is defined TWICE at workbook scope (a live definition
 * plus a #REF! residue, which Excel itself never writes; hand-crafted
 * files can). The reader models the winner and preserves the loser, so
 * ANY name edit's declarative save fails closed with the writer's
 * collision guard (the architect-endorsed fail-closed semantics for true
 * same-scope duplicates — as opposed to a same-name pair at DIFFERENT
 * scopes, which is a legal Excel construct and must save cleanly; see
 * buildExcelNamesSameNameFixture).
 */
export async function buildExcelNamesCollisionFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>
    <definedName name="Excel_Version">#REF!</definedName>
    <definedName name="Excel_Version">Data!$D$1</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="D1"><v>2026</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Same-name cross-scope pair fixture (the architect's EXCEL-025 blocker
 * scenario): "Total" is defined BOTH at workbook scope (Data!$B$2:$B$4) AND
 * at Data-sheet scope (localSheetId=0, Data!$C$7:$C$9) — two LEGITIMATE
 * Excel definitions that resolve differently by context. The reader must
 * model both, the browser must hold/list/resolve/edit them independently,
 * and a save rewriting one must preserve the other byte-verbatim. The
 * built-in print titles ride along for the preservation assertions.
 */
export async function buildExcelNamesSameNameFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>
    <definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>
    <definedName name="Total">Data!$B$2:$B$4</definedName>
    <definedName name="Total" localSheetId="0">Data!$C$7:$C$9</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c></row>
    <row r="3"><c r="A3"><v>30</v></c></row>
    <row r="4"><c r="A4"><v>40</v></c></row>
    <row r="5"><c r="A5"><v>50</v></c></row>
    <row r="7"><c r="C7"><v>7</v></c></row>
    <row r="8"><c r="C8"><v>8</v></c></row>
    <row r="9"><c r="C9"><v>9</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c r="B2"><v>5</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Case-variant collision fixture (the architect's second EXCEL-025 blocker
 * scenario): the name "Foo" and its case variant "foo" are BOTH defined at
 * workbook scope — a GENUINE same-scope duplicate under Excel's
 * case-insensitive name resolution (which Excel itself never writes;
 * hand-crafted files can). The reader models the winner and preserves the
 * loser, and the canonical writer's CASE-INSENSITIVE collision guard must
 * fail the names-dirty save closed — a case-sensitive guard would let both
 * variants serialize, writing a duplicate Excel cannot distinguish.
 * "GlobalTotal" rides along so the test can dirty the names model by
 * editing an unrelated name.
 */
export async function buildExcelNamesCaseCollisionFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>
    <definedName name="Foo">Data!$D$1</definedName>
    <definedName name="foo">#REF!</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="D1"><v>99</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * Case-variant cross-scope pair fixture (the architect's positive EXCEL-025
 * regression): "Total" at workbook scope and its case variant "total" at
 * Data-sheet scope (localSheetId=0). Scope dominates case in the canonical
 * uniqueness key (case-insensitive name, scope), so this is a LEGAL Excel
 * pair: both are modeled, both install, and a names-dirty save must
 * round-trip BOTH — proving the case-insensitive collision guards do not
 * over-block the cross-scope construct. The built-in print titles ride
 * along for the preservation assertions.
 */
export async function buildExcelNamesCasePairFixture(): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>
    <definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>
    <definedName name="Total">Data!$B$2:$B$4</definedName>
    <definedName name="total" localSheetId="0">Data!$C$7:$C$9</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c></row>
    <row r="3"><c r="A3"><v>30</v></c></row>
    <row r="4"><c r="A4"><v>40</v></c></row>
    <row r="5"><c r="A5"><v>50</v></c></row>
    <row r="7"><c r="C7"><v>7</v></c></row>
    <row r="8"><c r="C8"><v>8</v></c></row>
    <row r="9"><c r="C9"><v>9</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c r="B2"><v>5</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}

/**
 * EXCEL-026 — View / Page Layout persistence fixture. Two visible sheets:
 * "Data" (formula cells + a configurable <sheetView>) and "Other" (default
 * view). `dataView` is the sheetView element XML for sheet 1 (e.g.
 * '<sheetView showGridLines="0" showFormulas="1" workbookViewId="0"/>');
 * `dataTail` appends extra worksheet elements (pageSetup, etc.).
 */
export async function buildExcelViewFixture(
  dataView = '<sheetView workbookViewId="0"/>',
  dataTail = '',
): Promise<Buffer> {
  const zip = new JSZip()

  addFile(
    zip,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )

  addFile(
    zip,
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="191029"/>
</workbook>`,
  )

  addFile(
    zip,
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>${dataView}</sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><v>32</v></c>
      <c r="C1"><f>SUM(A1:B1)</f><v>42</v></c>
      <c r="D1"><f>A1*2</f><v>20</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>5</v></c>
      <c r="B2"><v>7</v></c>
    </row>
  </sheetData>
  ${dataTail}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  addFile(
    zip,
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="2"><c r="B2"><v>5</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`,
  )

  return toBytes(zip)
}
