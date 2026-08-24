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
