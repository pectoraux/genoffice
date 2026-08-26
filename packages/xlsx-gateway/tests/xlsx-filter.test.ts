import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'
import {
  applyFilterState,
  FilterEditError,
  FilterReadError,
  parseAutoFilter,
} from '../src/gateway/xlsx-filter'
import { buildStructureFixture } from './fixture-builder'

const AREA = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 }

const worksheet = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c></row>
    <row r="3" hidden="1"><c r="A3"><v>2</v></c></row>
    <row r="5" ht="20" customHeight="1"><c r="A5"><v>4</v></c></row>
  </sheetData>
</worksheet>`

describe('applyFilterState', () => {
  it('writes a values filter with hidden rows and creates missing row elements', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: {
        range: AREA,
        columns: [{ colId: 0, values: ['1', '4'], blank: true }],
      },
      hiddenRows: [2, 3],
      visibilityRange: AREA,
    })
    expect(xml).toContain(
      '<autoFilter ref="A1:D10"><filterColumn colId="0">' +
        '<filters blank="1"><filter val="1"/><filter val="4"/></filters>' +
        '</filterColumn></autoFilter>',
    )
    // Row 3 (index 2) stays hidden, row 4 (index 3) has no element yet and is
    // created hidden, row 5 keeps its height and gets unhidden semantics.
    expect(xml).toContain('<row r="3" hidden="1">')
    expect(xml).toContain('<row r="4" hidden="1"/>')
    expect(xml).toContain('<row r="5" ht="20" customHeight="1">')
    // Header row is never touched.
    expect(xml).toContain('<row r="1">')
  })

  it('unhides previously hidden rows that the filter no longer excludes', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: { range: AREA, columns: [] },
      hiddenRows: [],
      visibilityRange: AREA,
    })
    expect(xml).toContain('<row r="3"><c r="A3">')
    expect(xml).toContain('<autoFilter ref="A1:D10"/>')
  })

  it('replaces an existing autoFilter in place and can remove it', () => {
    const withFilter = `<worksheet><sheetData><row r="2" hidden="1"><c r="A2"><v>1</v></c></row></sheetData><autoFilter ref="A1:B5"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells></worksheet>`
    const replaced = applyFilterState(withFilter, {
      sheetName: 'Data',
      filter: { range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }, columns: [] },
      hiddenRows: [],
      visibilityRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
    })
    expect(replaced).toContain('<autoFilter ref="A1:B5"/><mergeCells')
    expect(replaced).not.toContain('filterColumn')
    expect(replaced).toContain('<row r="2"><c r="A2">')

    const removed = applyFilterState(withFilter, {
      sheetName: 'Data',
      filter: null,
      hiddenRows: [],
      visibilityRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
    })
    expect(removed).not.toContain('autoFilter')
    expect(removed).toContain('<row r="2"><c r="A2">')
  })

  it('serializes custom criteria and rejects unknown operators', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: {
        range: AREA,
        columns: [
          {
            colId: 2,
            customs: {
              and: true,
              filters: [
                { val: 5, operator: 'greaterThan' },
                { val: '*end', operator: 'equal' },
              ],
            },
          },
        ],
      },
      hiddenRows: [],
      visibilityRange: AREA,
    })
    expect(xml).toContain(
      '<filterColumn colId="2"><customFilters and="1">' +
        '<customFilter operator="greaterThan" val="5"/><customFilter val="*end"/>' +
        '</customFilters></filterColumn>',
    )
    expect(() =>
      applyFilterState(worksheet, {
        sheetName: 'Data',
        filter: {
          range: AREA,
          columns: [{ colId: 0, customs: { filters: [{ val: 1, operator: 'aboveAverage' }] } }],
        },
        hiddenRows: [],
        visibilityRange: AREA,
      }),
    ).toThrow(FilterEditError)
  })

  it('escapes filter values', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: {
        range: AREA,
        columns: [{ colId: 0, values: ['a<b>&"\''] }],
      },
      hiddenRows: [],
      visibilityRange: AREA,
    })
    expect(xml).toContain('<filter val="a&lt;b&gt;&amp;&quot;&apos;"/>')
  })
})

describe('filter save integration', () => {
  it('writes the autoFilter through the preservation pipeline', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      [],
      [],
      [],
      undefined,
      [
        {
          sheetName: 'Data',
          filter: {
            range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 },
            columns: [{ colId: 0, values: ['1', '10'] }],
          },
          hiddenRows: [1, 3],
          visibilityRange: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 },
        },
      ],
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<autoFilter ref="A1:D10"><filterColumn colId="0">')
    expect(sheet).toContain('<row r="2" hidden="1">')
    expect(sheet).toContain('<row r="4" hidden="1">')
    expect(sheet).toContain('<row r="10"><c r="A10">')
  })
})

// ── Read path (Phase 4 Increment 4 — Data → Filter) ─────────────────────────

const FILTERED_WORKSHEET = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Fruit</t></is></c></row>
    <row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>Veg</t></is></c></row>
    <row r="4" hidden="1"><c r="A4"/></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>Fruit</t></is></c></row>
    <row r="7" hidden="1"><c r="A7" t="inlineStr"><is><t>ManualHide</t></is></c></row>
  </sheetData>
  <autoFilter ref="A1:A5"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>
</worksheet>`

describe('parseAutoFilter (read path)', () => {
  it('reads an existing autoFilter: range, value criteria, and hidden rows', () => {
    const state = parseAutoFilter(FILTERED_WORKSHEET, 'Data')
    expect(state).not.toBeNull()
    expect(state!.filter).toEqual({
      range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
      columns: [{ colId: 0, values: ['Fruit'] }],
    })
    // Rows 3 and 4 (1-based) are hidden inside the filter span (0-based 2
    // and 3); row 7's manual hide is OUTSIDE the span and must not be part
    // of the filter state.
    expect(state!.hiddenRows).toEqual([2, 3])
    expect(state!.visibilityRange).toEqual({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 })
  })

  it('returns null when no autoFilter exists', () => {
    expect(parseAutoFilter(worksheet, 'Data')).toBeNull()
  })

  it('reads blank filtering (filters blank="1")', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
      `<autoFilter ref="A1:B3"><filterColumn colId="0"><filters blank="1"/></filterColumn></autoFilter></worksheet>`
    const state = parseAutoFilter(xml, 'Data')
    expect(state!.filter!.columns).toEqual([{ colId: 0, blank: true }])
  })

  it('reads blank + values together', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
      `<autoFilter ref="A1:B3"><filterColumn colId="0"><filters blank="1"><filter val="a"/></filters></filterColumn></autoFilter></worksheet>`
    const state = parseAutoFilter(xml, 'Data')
    expect(state!.filter!.columns).toEqual([{ colId: 0, values: ['a'], blank: true }])
  })

  it('reads supported custom filters with operators, and-joining, and numeric vals', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
      `<autoFilter ref="A1:C4"><filterColumn colId="1"><customFilters and="1">` +
      `<customFilter operator="greaterThan" val="5"/><customFilter operator="lessThanOrEqual" val="10"/>` +
      `</customFilters></filterColumn></autoFilter></worksheet>`
    const state = parseAutoFilter(xml, 'Data')
    expect(state!.filter!.columns).toEqual([
      {
        colId: 1,
        customs: {
          and: true,
          filters: [
            { val: 5, operator: 'greaterThan' },
            { val: 10, operator: 'lessThanOrEqual' },
          ],
        },
      },
    ])
  })

  it('keeps string vals as strings (no numeric coercion of non-numbers)', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
      `<autoFilter ref="A1:C4"><filterColumn colId="0"><customFilters>` +
      `<customFilter operator="notEqual" val="abc&amp;def"/>` +
      `</customFilters></filterColumn></autoFilter></worksheet>`
    const state = parseAutoFilter(xml, 'Data')
    expect(state!.filter!.columns).toEqual([
      {
        colId: 0,
        customs: { filters: [{ val: 'abc&def', operator: 'notEqual' }] },
      },
    ])
  })

  it('reads an empty-range-only autoFilter (dropdowns, no criteria)', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row><row r="2" hidden="1"><c r="A2"/></row></sheetData>` +
      `<autoFilter ref="A1:B2"/></worksheet>`
    const state = parseAutoFilter(xml, 'Data')
    expect(state!.filter).toEqual({
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      columns: [],
    })
    expect(state!.hiddenRows).toEqual([1])
  })

  it('fails closed on unsupported criteria (top10, dynamicFilter, iconFilter, dateGroup)', () => {
    for (const unsupported of [
      '<top10 val="10"/>',
      '<dynamicFilter type="aboveAverage" val="5"/>',
      '<iconFilter iconId="1"/>',
      '<dateGroupItem val="2024-01-01T00:00:00"/>',
    ]) {
      const xml =
        `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
        `<autoFilter ref="A1:B3"><filterColumn colId="0">${unsupported}</filterColumn></autoFilter></worksheet>`
      expect(() => parseAutoFilter(xml, 'Data')).toThrow(FilterReadError)
    }
  })

  it('fails closed on unsupported custom operators in the file', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
      `<autoFilter ref="A1:B3"><filterColumn colId="0"><customFilters>` +
      `<customFilter operator="beginsWith" val="x"/>` +
      `</customFilters></filterColumn></autoFilter></worksheet>`
    expect(() => parseAutoFilter(xml, 'Data')).toThrow(FilterReadError)
  })

  it('fails closed on a criteria-less filterColumn', () => {
    const xml =
      `<worksheet><sheetData><row r="1"><c r="A1"/></row></sheetData>` +
      `<autoFilter ref="A1:B3"><filterColumn colId="0"/></filterColumn></autoFilter></worksheet>`
    expect(() => parseAutoFilter(xml, 'Data')).toThrow(FilterReadError)
  })

  it('round-trips: parse(apply(state)) === state for values, blank, and customs', () => {
    const state = {
      sheetName: 'Data',
      filter: {
        range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 },
        columns: [
          { colId: 0, values: ['Fruit', 'Veg'], blank: true },
          { colId: 2, customs: { and: true, filters: [{ val: 5, operator: 'greaterThan' }] } },
        ],
      },
      hiddenRows: [2, 4],
      visibilityRange: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 },
    } as const
    const written = applyFilterState(FILTERED_WORKSHEET, state)
    const reparsed = parseAutoFilter(written, 'Data')
    // hiddenRows from the round-trip: the visibility span is A1:C6
    // (0-based rows 0..5) so rows 2, 4 are hidden inside it; the manual
    // hide on row 7 is outside the span.
    expect(reparsed!.filter).toEqual(state.filter)
    expect(reparsed!.hiddenRows).toEqual([2, 4])
    expect(reparsed!.visibilityRange).toEqual(state.visibilityRange)
  })
})

describe('filter read integration (readBasicWorkbook)', () => {
  it('surfaces filterState on the snapshot for a filtered worksheet', async () => {
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
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Fruit</t></is></c></row>
    <row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>Veg</t></is></c></row>
  </sheetData>
  <autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>
</worksheet>`,
    )
    const imported = await readBasicWorkbook(
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
    const sheet = imported.snapshot.sheets[0]!
    expect(sheet.filterState).toEqual({
      sheetName: 'Data',
      filter: {
        range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
        columns: [{ colId: 0, values: ['Fruit'] }],
      },
      hiddenRows: [2],
      visibilityRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    })
  })

  it('omits filterState (fail closed per filter) on unsupported criteria, but opens the workbook', async () => {
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
    <row r="1"><c r="A1" t="inlineStr"><is><t>N</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c></row>
  </sheetData>
  <autoFilter ref="A1:A2"><filterColumn colId="0"><top10 val="10"/></filterColumn></autoFilter>
</worksheet>`,
    )
    const imported = await readBasicWorkbook(
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
    // The workbook still opens with its cells; only the filter is skipped.
    expect(imported.snapshot.sheets[0]!.cells.A1?.value).toBe('N')
    expect(imported.snapshot.sheets[0]!.filterState).toBeUndefined()
  })

  it('no-op save (empty plan) preserves the filter XML byte-for-byte', async () => {
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
    const originalSheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Fruit</t></is></c></row>
    <row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>Veg</t></is></c></row>
  </sheetData>
  <autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>
</worksheet>`
    zip.file('xl/worksheets/sheet1.xml', originalSheet)
    const fixture = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    // Empty save plan: no edits, no structural ops, NO filter states — the
    // sheet is not filter-dirty, so the file's own XML must survive.
    const mutation = await applyCellEditsToXlsx(fixture, [])
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
    const outZip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await outZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain(
      '<autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>',
    )
    expect(sheet).toContain('<row r="3" hidden="1">')
  })

  it('clearing a filter removes the autoFilter and restores hidden rows', async () => {
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
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Fruit</t></is></c></row>
    <row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>Veg</t></is></c></row>
  </sheetData>
  <autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>
</worksheet>`,
    )
    const fixture = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const mutation = await applyCellEditsToXlsx(fixture, [], [], [], undefined, [
      {
        sheetName: 'Data',
        filter: null,
        hiddenRows: [],
        visibilityRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
      },
    ])
    const outZip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await outZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).not.toContain('autoFilter')
    expect(sheet).toContain('<row r="3"><c r="A3"')
  })

  it('unrelated rows outside the visibility span stay untouched by a filter save', () => {
    const xml = applyFilterState(FILTERED_WORKSHEET, {
      sheetName: 'Data',
      filter: {
        range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
        columns: [{ colId: 0, values: ['Fruit'] }],
      },
      hiddenRows: [2],
      visibilityRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
    })
    // Row 7's manual hide (outside the filter span) survives untouched.
    expect(xml).toContain('<row r="7" hidden="1">')
  })
})
