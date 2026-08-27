import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'

/**
 * EXCEL-018 independent byte proof (architect verification debt closure).
 *
 * These tests inspect the resulting XLSX PACKAGE directly — no browser,
 * no HTTP, no Univer. They replicate the exact structural-op sequences
 * the browser runtime issues for Remove Duplicates (descending
 * `{kind:'remove-rows', index, count:1}` per duplicate, journaled from
 * `sheet.mutation.remove-rows`) and prove the three package-level
 * invariants the architect demanded:
 *
 *   1. survivor `<c>` records remain FORMULAS (never rewritten to
 *      computed literals through a value-level compaction);
 *   2. style references (`s=` indices) travel with survivor rows;
 *   3. unrelated rows / cells / package parts remain untouched
 *      (byte-identical where the surgical save promises preservation).
 *
 * Scenario A is the architect's mandatory retained-formula regression
 * (B7 `=B6` must compact to B5 `=B4`, NOT the literal 30).
 * Scenario B is the mixed-reference regression ($D$6 / A6 / $A6 / A$6
 * all rewritten to row 4 with `$` markers preserved).
 */

async function entryText(buffer: Buffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  if (!entry) throw new Error(`Missing ${path}`)
  return entry.async('text')
}

/** Minimal styles with FOUR distinct xf records so `s=` travel is provable. */
const dedupeStyles = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font/><font><b/></font></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="4"><xf/><xf fontId="1" applyFont="1"/><xf fillId="0" applyAlignment="1"/><xf fontId="0" applyFont="0"/></cellXfs>
</styleSheet>`

const sharedStrings = (count: number, texts: readonly string[]): string =>
  `<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${texts.length}">${texts
    .map((t) => `<si><t>${t}</t></si>`)
    .join('')}</sst>`

interface FixtureSpec {
  readonly worksheet: string
  readonly strings: readonly string[]
}

async function buildDedupeFixture(spec: FixtureSpec): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
  )
  zip.file('xl/worksheets/sheet1.xml', spec.worksheet)
  // An unrelated second sheet the dedupe never touches — byte-identity proof.
  zip.file(
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>111</v></c></row></sheetData></worksheet>`,
  )
  zip.file('xl/styles.xml', dedupeStyles)
  zip.file('xl/sharedStrings.xml', sharedStrings(spec.strings.length, spec.strings))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/**
 * Scenario A fixture (the architect's mandatory retained-formula case):
 *
 *   row 1  header      (xf 1 — bold)
 *   row 2  Apple / 10  (xf 2 — first occurrence, kept)
 *   row 3  Apple / 10  (xf 2 — DUPLICATE, index 2 — deleted)
 *   row 4  Banana / 20 (xf 3 — survivor)
 *   row 5  Apple / 10  (xf 2 — DUPLICATE, index 4 — deleted)
 *   row 6  Cherry / 30 (xf 1 — the referenced row)
 *   row 7  Apple / =B6 (xf 2 — formula survivor, computed 30)
 *   row 9  Far / 99    (xf 1 — far-away row OUTSIDE the A1:B7 selection)
 */
const scenarioAWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B9"/>
  <sheetData>
    <row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2" s="2" t="s"><v>2</v></c><c r="B2" s="2"><v>10</v></c></row>
    <row r="3"><c r="A3" s="2" t="s"><v>2</v></c><c r="B3" s="2"><v>10</v></c></row>
    <row r="4"><c r="A4" s="3" t="s"><v>3</v></c><c r="B4" s="3"><v>20</v></c></row>
    <row r="5"><c r="A5" s="2" t="s"><v>2</v></c><c r="B5" s="2"><v>10</v></c></row>
    <row r="6"><c r="A6" s="1" t="s"><v>4</v></c><c r="B6" s="1"><v>30</v></c></row>
    <row r="7"><c r="A7" s="2" t="s"><v>2</v></c><c r="B7" s="2"><f>B6</f><v>30</v></c></row>
    <row r="9"><c r="A9" s="1" t="s"><v>5</v></c><c r="B9" s="1"><v>99</v></c></row>
  </sheetData>
</worksheet>`

const scenarioAStrings = ['Name', 'Value', 'Apple', 'Banana', 'Cherry', 'Far']

/**
 * Scenario B fixture (the architect's mixed-reference case):
 *
 *   row 1  header
 *   row 2  Apple / 10            — first occurrence
 *   row 3  Apple / 10            — DUPLICATE, index 2
 *   row 4  Banana + C4 =$D$6, D4 =A6, E4 =$A6, F4 =A$6   — survivor
 *   row 5  Apple / 10            — DUPLICATE, index 4
 *   row 6  Cherry / 30 / D6=Anchor  — the referenced target row
 *   row 7  Apple / 10            — DUPLICATE, index 6
 */
const scenarioBWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:F7"/>
  <sheetData>
    <row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3"><v>10</v></c></row>
    <row r="4"><c r="A4" s="3" t="s"><v>3</v></c><c r="B4" s="3"><v>20</v></c><c r="C4" s="3"><f>$D$6</f></c><c r="D4" s="3"><f>A6</f></c><c r="E4" s="3"><f>$A6</f></c><c r="F4" s="3"><f>A$6</f></c></row>
    <row r="5"><c r="A5" t="s"><v>2</v></c><c r="B5"><v>10</v></c></row>
    <row r="6"><c r="A6" t="s"><v>4</v></c><c r="B6"><v>30</v></c><c r="D6" t="s"><v>5</v></c></row>
    <row r="7"><c r="A7" t="s"><v>2</v></c><c r="B7"><v>10</v></c></row>
  </sheetData>
</worksheet>`

const scenarioBStrings = ['Name', 'Value', 'Apple', 'Banana', 'Cherry', 'Anchor']

describe('EXCEL-018 independent byte proof — structural remove-rows dedupe', () => {
  it('scenario A: the retained formula survives as a <f> record, styles travel, unrelated content is untouched', async () => {
    const source = await buildDedupeFixture({
      worksheet: scenarioAWorksheet,
      strings: scenarioAStrings,
    })
    // The EXACT op sequence the browser runtime issues (descending order):
    // delete index 4 (row 5) first, then index 2 (row 3).
    const mutation = await applyCellEditsToXlsx(
      source,
      [],
      [
        {
          sheetName: 'Data',
          ops: [
            { kind: 'remove-rows', index: 4, count: 1 },
            { kind: 'remove-rows', index: 2, count: 1 },
          ],
        },
      ],
    )
    const sheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')

    // --- (1) The duplicate rows are GONE (structural deletion). ---
    expect(sheet).not.toContain('r="A3" t="s"')
    expect(sheet).not.toMatch(/<row r="3"[^>]*>\s*<c r="A3"[^>]*t="s"[^>]*>\s*<v>2<\/v>/)
    // Rows 3 and 5 (the two Apple duplicates) no longer exist as data rows:
    // after compaction there are exactly the survivors 1,2,3,4,5,7.
    const rowNumbers = [...sheet.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]))
    expect(rowNumbers).toEqual([1, 2, 3, 4, 5, 7])

    // --- (2) Survivor formulas remain FORMULAS (the architect's case). ---
    // B7 (=B6, computed 30) compacted to B5 and its reference rewrote B6→B4
    // (Cherry moved from row 6 to row 4). The cell MUST carry <f>B4</f>.
    expect(sheet).toMatch(/<c r="B5"[^>]*><f>B4<\/f>/)
    // The value-level failure mode the architect explicitly rejected: B5 as
    // the literal 30 with NO <f> element (a setValues compaction).
    expect(sheet).not.toMatch(/<c r="B5"[^>]*>(?!<f>)(?:<v>30<\/v>)?<\/c>/)

    // --- (3) Style references travel with survivor rows. ---
    // Banana (was row 4, xf 3) lands at row 3 still on xf 3.
    expect(sheet).toMatch(/<c r="A3" s="3" t="s"><v>3<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="B3" s="3"><v>20<\/v><\/c>/)
    // Cherry (was row 6, xf 1) lands at row 4 still on xf 1.
    expect(sheet).toMatch(/<c r="A4" s="1" t="s"><v>4<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="B4" s="1"><v>30<\/v><\/c>/)
    // The formula survivor (was row 7, xf 2) lands at row 5 still on xf 2.
    expect(sheet).toMatch(/<c r="B5" s="2"><f>B4<\/f>/)
    // The first occurrence (row 2, xf 2) is untouched verbatim.
    expect(sheet).toMatch(/<c r="A2" s="2" t="s"><v>2<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="B2" s="2"><v>10<\/v><\/c>/)

    // --- (4) Cells OUTSIDE the selection are untouched (content + style
    // travel verbatim; only their row coordinate shifts per Excel
    // semantics — Remove Duplicates deletes WHOLE rows). ---
    expect(sheet).toMatch(/<c r="A7" s="1" t="s"><v>5<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="B7" s="1"><v>99<\/v><\/c>/)

    // --- (5) Unrelated package parts stay byte-identical. ---
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const before = new Map(mutation.beforeEntries.map((e) => [e.path, e.sha256]))
    const after = new Map(mutation.afterEntries.map((e) => [e.path, e.sha256]))
    expect(after.get('xl/styles.xml')).toBe(before.get('xl/styles.xml'))
    expect(after.get('xl/sharedStrings.xml')).toBe(before.get('xl/sharedStrings.xml'))
    expect(after.get('xl/worksheets/sheet2.xml')).toBe(before.get('xl/worksheets/sheet2.xml'))
    expect(after.get('_rels/.rels')).toBe(before.get('_rels/.rels'))
    expect(after.get('xl/_rels/workbook.xml.rels')).toBe(before.get('xl/_rels/workbook.xml.rels'))

    // --- (6) Reopen: the canonical snapshot carries the formula. ---
    const imported = await readBasicWorkbook(mutation.buffer)
    const data = imported.snapshot.sheets.find((s) => s.name === 'Data')
    expect(data).toBeDefined()
    expect(data?.cells.B5?.formula).toBe('=B4')
    expect(data?.cells.B5?.value).not.toBe(30)
    // The reopened sheet carries exactly the survivor geometry.
    expect(data?.cells.A3?.value).toBe('Banana')
    expect(data?.cells.A4?.value).toBe('Cherry')
    expect(data?.cells.A7?.value).toBe('Far')
  })

  it('scenario B: absolute / relative / mixed references all track the moved row with $ markers preserved', async () => {
    const source = await buildDedupeFixture({
      worksheet: scenarioBWorksheet,
      strings: scenarioBStrings,
    })
    // Duplicates at indices 2, 4, 6 — deleted in DESCENDING order exactly
    // as the runtime issues them.
    const mutation = await applyCellEditsToXlsx(
      source,
      [],
      [
        {
          sheetName: 'Data',
          ops: [
            { kind: 'remove-rows', index: 6, count: 1 },
            { kind: 'remove-rows', index: 4, count: 1 },
            { kind: 'remove-rows', index: 2, count: 1 },
          ],
        },
      ],
    )
    const sheet = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')

    // Banana (was row 4) compacts to row 3; Cherry/Anchor (was row 6)
    // compacts to row 4. All four references rewrite row 6 → row 4 while
    // preserving every $ marker:
    expect(sheet).toMatch(/<c r="C3"[^>]*><f>\$D\$4<\/f><\/c>/)
    expect(sheet).toMatch(/<c r="D3"[^>]*><f>A4<\/f><\/c>/)
    expect(sheet).toMatch(/<c r="E3"[^>]*><f>\$A4<\/f><\/c>/)
    expect(sheet).toMatch(/<c r="F3"[^>]*><f>A\$4<\/f><\/c>/)
    // The referenced content actually moved there (D6 Anchor → D4).
    expect(sheet).toMatch(/<c r="D4" t="s"><v>5<\/v><\/c>/)
    // All four survivors are still FORMULAS — none collapsed to literals.
    const row3 = /<row r="3"[^>]*>([\s\S]*?)<\/row>/.exec(sheet)?.[1] ?? ''
    expect(row3.match(/<f>/g)?.length).toBe(4)

    // Styles travel: Banana's xf 3 stays on every survivor cell.
    expect(sheet).toMatch(/<c r="A3" s="3" t="s"><v>3<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="C3" s="3"><f>\$D\$4<\/f><\/c>/)

    // Unrelated parts stay byte-identical.
    const before = new Map(mutation.beforeEntries.map((e) => [e.path, e.sha256]))
    const after = new Map(mutation.afterEntries.map((e) => [e.path, e.sha256]))
    expect(after.get('xl/styles.xml')).toBe(before.get('xl/styles.xml'))
    expect(after.get('xl/sharedStrings.xml')).toBe(before.get('xl/sharedStrings.xml'))
    expect(after.get('xl/worksheets/sheet2.xml')).toBe(before.get('xl/worksheets/sheet2.xml'))

    // Reopen: the snapshot carries all four rewritten formulas.
    const imported = await readBasicWorkbook(mutation.buffer)
    const data = imported.snapshot.sheets.find((s) => s.name === 'Data')
    expect(data?.cells.C3?.formula).toBe('=$D$4')
    expect(data?.cells.D3?.formula).toBe('=A4')
    expect(data?.cells.E3?.formula).toBe('=$A4')
    expect(data?.cells.F3?.formula).toBe('=A$4')
  })

  it('zero-removal no-op: applying NO structural ops changes no worksheet byte', async () => {
    const source = await buildDedupeFixture({
      worksheet: scenarioAWorksheet,
      strings: scenarioAStrings,
    })
    // The runtime's `removed === 0` branch fires NO mutation at all; the
    // equivalent package-level contract is an empty ops list leaving the
    // worksheet untouched (no rewrite, no calcPr flip on the sheet part).
    const before = await entryText(source, 'xl/worksheets/sheet1.xml')
    expect(before).toContain('<f>B6</f>')
    // A no-op save (empty edits + empty ops) preserves the sheet bytes.
    const mutation = await applyCellEditsToXlsx(source, [], [])
    expect(await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')).toBe(before)
  })
})
