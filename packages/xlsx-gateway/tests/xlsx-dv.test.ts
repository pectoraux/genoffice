import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyDvRules,
  DvEditError,
  DvReadError,
  parseDataValidations,
} from '../src/gateway/xlsx-dv'
import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<dataValidations count="1"><dataValidation type="whole" operator="equal" sqref="B1">' +
  '<formula1>7</formula1></dataValidation></dataValidations>' +
  '<hyperlinks><hyperlink ref="G1" r:id="rId1"/></hyperlinks>' +
  '<pageMargins left="0.7"/></worksheet>'

const range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }

describe('applyDvRules', () => {
  it('replaces the existing section with the snapshot, before hyperlinks', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: {
          type: 'decimal',
          operator: 'greaterThan',
          formula1: '3',
          allowBlank: true,
        },
      },
    ])
    expect(xml).not.toContain('type="whole"')
    expect(xml).toContain(
      '<dataValidations count="1"><dataValidation type="decimal" operator="greaterThan" ' +
        'allowBlank="1" sqref="A1:A5"><formula1>3</formula1></dataValidation>' +
        '</dataValidations><hyperlinks',
    )
  })

  it('removes the section when the snapshot is empty', () => {
    const xml = applyDvRules(SHEET, [])
    expect(xml).not.toContain('dataValidations')
    expect(xml).toContain('<hyperlinks>')
  })

  it('quotes list literals and strips = from list references', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range], rule: { type: 'list', formula1: 'Yes,No', showDropDown: true } },
      {
        ranges: [{ ...range, startColumn: 1, endColumn: 1 }],
        rule: { type: 'list', formula1: "='My Sheet'!$D$1:$D$9", showDropDown: false },
      },
    ])
    expect(xml).toContain('<formula1>"Yes,No"</formula1>')
    expect(xml).toContain("<formula1>'My Sheet'!$D$1:$D$9</formula1>")
    // OOXML showDropDown="1" means "suppress the dropdown" — only the
    // dropdown-off rule carries it.
    expect(xml).toContain('<dataValidation type="list" showDropDown="1" sqref="B1:B5">')
    expect(xml).toContain('<dataValidation type="list" sqref="A1:A5">')
  })

  it('writes prompts, errors, and the inverse error-style mapping', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: {
          type: 'textLength',
          operator: 'lessThanOrEqual',
          formula1: '10',
          showInputMessage: true,
          promptTitle: 'Limit',
          prompt: 'Max 10 & "short"',
          showErrorMessage: true,
          errorStyle: 2,
          errorTitle: 'Too long',
          error: 'Shorten it',
        },
      },
    ])
    expect(xml).toContain(
      'type="textLength" operator="lessThanOrEqual" showInputMessage="1" ' +
        'showErrorMessage="1" errorStyle="warning" errorTitle="Too long" error="Shorten it" ' +
        'promptTitle="Limit" prompt="Max 10 &amp; &quot;short&quot;" sqref="A1:A5"',
    )
  })

  it('maps any back to the attribute-less default type and skips between', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: { type: 'any', showInputMessage: true, prompt: 'Read me' },
      },
      {
        ranges: [{ ...range, startColumn: 2, endColumn: 2 }],
        rule: { type: 'whole', operator: 'between', formula1: '1', formula2: '9' },
      },
    ])
    expect(xml).toContain('<dataValidation showInputMessage="1" prompt="Read me" sqref="A1:A5"/>')
    expect(xml).toContain(
      '<dataValidation type="whole" sqref="C1:C5"><formula1>1</formula1>' +
        '<formula2>9</formula2></dataValidation>',
    )
  })

  it('converts panel date/time strings to Excel serials, keeps numbers', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: {
          type: 'date',
          operator: 'notBetween',
          formula1: '2024-01-01',
          formula2: '2024-12-31 12:00:00',
        },
      },
      {
        ranges: [{ ...range, startColumn: 3, endColumn: 3 }],
        rule: { type: 'time', operator: 'lessThan', formula1: '06:00' },
      },
      {
        ranges: [{ ...range, startColumn: 4, endColumn: 4 }],
        rule: { type: 'date', operator: 'greaterThan', formula1: '45123' },
      },
    ])
    expect(xml).toContain('<formula1>45292</formula1>')
    expect(xml).toContain('<formula2>45657.5</formula2>')
    expect(xml).toContain('<formula1>0.25</formula1>')
    expect(xml).toContain('<formula1>45123</formula1>')
  })

  it('strips = from custom formulas and escapes their XML', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: { type: 'custom', formula1: '=AND($A1>0,$A1<5)' },
      },
    ])
    expect(xml).toContain('<formula1>AND($A1&gt;0,$A1&lt;5)</formula1>')
  })

  it('degrades checkbox rules to a two-value list', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range], rule: { type: 'checkbox', allowBlank: true } },
      {
        ranges: [{ ...range, startColumn: 1, endColumn: 1 }],
        rule: { type: 'checkbox', formula1: 'Yes', formula2: 'No' },
      },
    ])
    expect(xml).toContain(
      '<dataValidation type="list" allowBlank="1" sqref="A1:A5">' +
        '<formula1>"1,0"</formula1></dataValidation>',
    )
    expect(xml).toContain(
      '<dataValidation type="list" sqref="B1:B5"><formula1>"Yes,No"</formula1></dataValidation>',
    )
  })

  it('fails closed on Univer-only rule types', () => {
    expect(() =>
      applyDvRules(SHEET, [{ ranges: [range], rule: { type: 'listMultiple', formula1: 'a,b' } }]),
    ).toThrow(DvEditError)
    expect(() =>
      applyDvRules(SHEET, [{ ranges: [range], rule: { type: 'listMultiple', formula1: 'a,b' } }]),
    ).toThrow(/Multi-select/)
  })

  it('fails closed on x14 extended validations', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      '<extLst><ext><x14:dataValidations count="1"><x14:dataValidation type="list"/>' +
        '</x14:dataValidations></ext></extLst>',
    )
    expect(() =>
      applyDvRules(x14Sheet, [
        { ranges: [range], rule: { type: 'whole', operator: 'equal', formula1: '1' } },
      ]),
    ).toThrow(/x14/)
  })

  it('appends before </worksheet> when no later section exists', () => {
    const bare = '<worksheet><sheetData/></worksheet>'
    const xml = applyDvRules(bare, [
      { ranges: [range], rule: { type: 'whole', operator: 'equal', formula1: '1' } },
    ])
    expect(xml).toBe(
      '<worksheet><sheetData/><dataValidations count="1">' +
        '<dataValidation type="whole" operator="equal" sqref="A1:A5">' +
        '<formula1>1</formula1></dataValidation></dataValidations></worksheet>',
    )
  })
})

// ── Read path (Phase 4 Increment 5 — Data → Data Validation) ─────────────────

describe('parseDataValidations (read path)', () => {
  const dv = (attrs: string, body = '') =>
    `<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
    `<dataValidations count="1"><dataValidation ${attrs}>${body}</dataValidation></dataValidations></worksheet>`

  it('1. parses a whole-number between validation', () => {
    const rules = parseDataValidations(
      dv('type="whole" sqref="A2:A6"', '<formula1>1</formula1><formula2>100</formula2>'),
    )
    expect(rules).toEqual([
      {
        ranges: [{ startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }],
        rule: { type: 'whole', operator: 'between', formula1: '1', formula2: '100' },
      },
    ])
  })

  it('2. parses a decimal validation with an explicit operator', () => {
    const rules = parseDataValidations(
      dv('type="decimal" operator="greaterThan" sqref="B2"', '<formula1>3.5</formula1>'),
    )
    expect(rules[0]!.rule).toEqual({ type: 'decimal', operator: 'greaterThan', formula1: '3.5' })
  })

  it('3. parses a list validation (literal and formula forms)', () => {
    const literal = parseDataValidations(
      dv('type="list" sqref="B2:B6"', '<formula1>"Fruit,Vegetable,Grain"</formula1>'),
    )
    expect(literal[0]!.rule).toEqual({ type: 'list', formula1: '"Fruit,Vegetable,Grain"' })
    const reference = parseDataValidations(
      dv('type="list" sqref="B2:B6"', '<formula1>$D$1:$D$3</formula1>'),
    )
    expect(reference[0]!.rule).toEqual({ type: 'list', formula1: '$D$1:$D$3' })
  })

  it('4. parses a date validation (serial formula preserved verbatim)', () => {
    const rules = parseDataValidations(
      dv(
        'type="date" operator="between" sqref="C2"',
        '<formula1>45292</formula1><formula2>45657</formula2>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'date',
      operator: 'between',
      formula1: '45292',
      formula2: '45657',
    })
  })

  it('5. parses a time validation', () => {
    const rules = parseDataValidations(
      dv(
        'type="time" operator="notBetween" sqref="D2"',
        '<formula1>0.25</formula1><formula2>0.75</formula2>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'time',
      operator: 'notBetween',
      formula1: '0.25',
      formula2: '0.75',
    })
  })

  it('6. parses a textLength validation', () => {
    const rules = parseDataValidations(
      dv('type="textLength" operator="lessThanOrEqual" sqref="E2"', '<formula1>10</formula1>'),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'textLength',
      operator: 'lessThanOrEqual',
      formula1: '10',
    })
  })

  it('7. parses a custom (formula-backed) validation', () => {
    const rules = parseDataValidations(
      dv('type="custom" sqref="F2:F6"', '<formula1>ISNUMBER(F2)</formula1>'),
    )
    expect(rules[0]!.rule).toEqual({ type: 'custom', formula1: 'ISNUMBER(F2)' })
  })

  it('8. parses multiple ranges in one sqref (and single-cell refs)', () => {
    const rules = parseDataValidations(dv('type="whole" sqref="A2 B3:B5 C1"'))
    expect(rules[0]!.ranges).toEqual([
      { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      { startRow: 2, endRow: 4, startColumn: 1, endColumn: 1 },
      { startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 },
    ])
  })

  it('9. parses allowBlank', () => {
    const rules = parseDataValidations(
      dv('type="whole" allowBlank="1" sqref="A2"', '<formula1>5</formula1>'),
    )
    expect(rules[0]!.rule.allowBlank).toBe(true)
    const without = parseDataValidations(dv('type="whole" sqref="A2"', '<formula1>5</formula1>'))
    expect(without[0]!.rule.allowBlank).toBeUndefined()
  })

  it('10. parses prompt/error metadata', () => {
    const rules = parseDataValidations(
      dv(
        'type="whole" showInputMessage="1" showErrorMessage="1" errorTitle="Bad value" ' +
          'error="Enter 1-100" promptTitle="Hint" prompt="Whole numbers only" sqref="A2"',
        '<formula1>1</formula1><formula2>100</formula2>',
      ),
    )
    expect(rules[0]!.rule).toMatchObject({
      showInputMessage: true,
      showErrorMessage: true,
      errorTitle: 'Bad value',
      error: 'Enter 1-100',
      promptTitle: 'Hint',
      prompt: 'Whole numbers only',
    })
  })

  it('11. parses dropdown semantics (OOXML inverted attribute)', () => {
    // showDropDown="1" SUPPRESSES the dropdown.
    const suppressed = parseDataValidations(
      dv('type="list" showDropDown="1" sqref="B2"', '<formula1>"a,b"</formula1>'),
    )
    expect(suppressed[0]!.rule.showDropDown).toBe(false)
    const shown = parseDataValidations(dv('type="list" sqref="B2"', '<formula1>"a,b"</formula1>'))
    expect(shown[0]!.rule.showDropDown).toBeUndefined()
    // Non-list types keep the attribute verbatim so the write side re-emits it.
    const whole = parseDataValidations(
      dv('type="whole" showDropDown="0" sqref="C2"', '<formula1>1</formula1>'),
    )
    expect(whole[0]!.rule.showDropDown).toBe(true)
  })

  it('12. round-trips: parse(apply(rules)) recovers the rules', () => {
    const rules = [
      {
        ranges: [{ startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }],
        rule: {
          type: 'whole',
          operator: 'between',
          formula1: '1',
          formula2: '100',
          allowBlank: true,
          showErrorMessage: true,
          errorTitle: 'Bad',
          error: 'Nope',
          errorStyle: 2,
        },
      },
      {
        ranges: [{ startRow: 1, endRow: 5, startColumn: 1, endColumn: 1 }],
        rule: { type: 'list', formula1: 'Fruit,Vegetable,Grain', showDropDown: true },
      },
      {
        ranges: [{ startRow: 1, endRow: 5, startColumn: 2, endColumn: 2 }],
        rule: { type: 'custom', formula1: '=ISNUMBER(C2)' },
      },
    ] as const
    const written = applyDvRules(SHEET, rules)
    const reparsed = parseDataValidations(written)
    // between is omitted as the OOXML default on write and re-defaulted on
    // read; list literals are quoted on write and kept quoted on read (the
    // browser install unquotes them); custom strips the '='. All expected.
    expect(reparsed).toEqual([
      {
        ranges: rules[0]!.ranges,
        rule: {
          type: 'whole',
          operator: 'between',
          formula1: '1',
          formula2: '100',
          allowBlank: true,
          showErrorMessage: true,
          errorTitle: 'Bad',
          error: 'Nope',
          errorStyle: 2,
        },
      },
      {
        ranges: rules[1]!.ranges,
        // showDropDown:true (dropdown shown) is the OOXML default — the
        // writer omits the attribute, so the read leaves it undefined.
        rule: { type: 'list', formula1: '"Fruit,Vegetable,Grain"' },
      },
      {
        ranges: rules[2]!.ranges,
        rule: { type: 'custom', formula1: 'ISNUMBER(C2)' },
      },
    ])
  })

  it('13. clear: an empty rule list serializes away and re-parses as []', () => {
    const cleared = applyDvRules(SHEET, [])
    expect(parseDataValidations(cleared)).toEqual([])
    expect(cleared).not.toContain('dataValidations')
  })

  it('14. fails closed on unsupported operators', () => {
    expect(() => parseDataValidations(dv('type="whole" operator="beginsWith" sqref="A2"'))).toThrow(
      DvReadError,
    )
  })

  it('15. fails closed on unsupported types', () => {
    expect(() => parseDataValidations(dv('type="iconSet" sqref="A2"'))).toThrow(DvReadError)
  })

  it('16. fails closed on x14 extensions', () => {
    const withX14 =
      '<worksheet><sheetData/><dataValidations count="1"><dataValidation type="whole" sqref="A2"/></dataValidations>' +
      '<extLst><ext xmlns:x14="x"><x14:dataValidation sqref="B2"><x14:formula1><xm:f>ISNUMBER(B2)</xm:f></x14:formula1></x14:dataValidation></ext></extLst></worksheet>'
    expect(() => parseDataValidations(withX14)).toThrow(DvReadError)
  })

  it('17. fails closed on malformed ranges', () => {
    expect(() => parseDataValidations(dv('type="whole" sqref="not-a-ref"'))).toThrow(DvReadError)
    expect(() => parseDataValidations(dv('type="whole" sqref=""'))).toThrow(DvReadError)
    expect(() => parseDataValidations(dv('type="whole"'))).toThrow(DvReadError)
  })

  it('18. no-op: a worksheet without DV parses to [] (saves byte-preserved per test 19)', () => {
    expect(parseDataValidations('<worksheet><sheetData/></worksheet>')).toEqual([])
    // The no-op SAVE proof (no dvStates → untouched worksheet entry) is
    // covered by test 19 with a real archive.
  })

  it('19. unrelated validations are preserved verbatim by a no-op save', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c></row>
  </sheetData>
  <dataValidations count="2">
    <dataValidation type="whole" operator="between" sqref="A1:A2"><formula1>1</formula1><formula2>10</formula2></dataValidation>
    <dataValidation type="list" sqref="B1"><formula1>"x,y"</formula1></dataValidation>
  </dataValidations>
</worksheet>`,
    )
    const fixture = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const mutation = await applyCellEditsToXlsx(fixture, [])
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
    const out = await JSZip.loadAsync(mutation.buffer)
    const sheet = await out.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<dataValidations count="2">')
    expect(sheet).toContain('sqref="A1:A2"')
    expect(sheet).toContain('<formula1>"x,y"</formula1>')
  })

  it('20. parses multiple validations on one sheet (readBasicWorkbook integration)', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <dataValidations count="3">
    <dataValidation type="whole" operator="between" sqref="A2:A6"><formula1>1</formula1><formula2>100</formula2></dataValidation>
    <dataValidation type="list" sqref="B2:B6"><formula1>"Fruit,Vegetable,Grain"</formula1></dataValidation>
    <dataValidation type="custom" sqref="C2:C6"><formula1>ISNUMBER(C2)</formula1></dataValidation>
  </dataValidations>
</worksheet>`,
    )
    const imported = await readBasicWorkbook(
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
    const sheet = imported.snapshot.sheets[0]!
    expect(sheet.dvRules).toHaveLength(3)
    expect(sheet.dvRules![0]).toEqual({
      ranges: [{ startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }],
      rule: { type: 'whole', operator: 'between', formula1: '1', formula2: '100' },
    })
    expect(sheet.dvRules![1]!.rule).toEqual({ type: 'list', formula1: '"Fruit,Vegetable,Grain"' })
    expect(sheet.dvRules![2]!.rule).toEqual({ type: 'custom', formula1: 'ISNUMBER(C2)' })
  })

  it('fail-closed per sheet: an unrepresentable section surfaces no dvRules but the workbook opens', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row>
  </sheetData>
  <dataValidations count="1">
    <dataValidation type="whole" operator="beginsWith" sqref="A2"><formula1>x</formula1></dataValidation>
  </dataValidations>
</worksheet>`,
    )
    const imported = await readBasicWorkbook(
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
    expect(imported.snapshot.sheets[0]!.cells.A1?.value).toBe('ok')
    expect(imported.snapshot.sheets[0]!.dvRules).toBeUndefined()
  })
})
