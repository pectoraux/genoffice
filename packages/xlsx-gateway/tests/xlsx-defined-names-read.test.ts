import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import {
  DefinedNamesReadError,
  parseDefinedNamesState,
  definedNameIsSaveable,
} from '../src/gateway/xlsx-defined-names'
import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'

/// ── Pure reader tests (parseDefinedNamesState) ──────────────────────────────

const wb = (inner: string): string =>
  `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>${inner}<calcPr/></workbook>`

describe('parseDefinedNamesState', () => {
  it('reads workbook- and sheet-scoped names with their formulas', () => {
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="GlobalTotal">Data!$A$1</definedName>' +
          '<definedName name="LocalTotal" localSheetId="1">Other!$B$2</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([
      { name: 'GlobalTotal', formula: 'Data!$A$1' },
      { name: 'LocalTotal', formula: 'Other!$B$2', sheetIndex: 1 },
    ])
    expect(state.preserveNames).toEqual([])
  })

  it('returns empty state for a workbook without a section', () => {
    expect(parseDefinedNamesState(wb(''), 2)).toEqual({ names: [], preserveNames: [] })
    expect(parseDefinedNamesState(wb('<definedNames/>'), 2)).toEqual({
      names: [],
      preserveNames: [],
    })
    expect(parseDefinedNamesState(wb('<definedNames></definedNames>'), 2)).toEqual({
      names: [],
      preserveNames: [],
    })
  })

  it('skips _xlnm built-ins and hidden names entirely (auto-preserved by the writer)', () => {
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>' +
          '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$C$4</definedName>' +
          '<definedName name="Secret" hidden="1">Data!$Z$9</definedName>' +
          '<definedName name="Visible">Data!$A$2</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([{ name: 'Visible', formula: 'Data!$A$2' }])
    expect(state.preserveNames).toEqual([])
  })

  it('preserves names that fail the writer-s validation rules', () => {
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="A1">Data!$A$1</definedName>' +
          '<definedName name="R2C3">Data!$A$1</definedName>' +
          '<definedName name="true">Data!$A$1</definedName>' +
          '<definedName name="1Starts">Data!$A$1</definedName>' +
          '<definedName name="Has Space">Data!$A$1</definedName>' +
          '<definedName name="Good">Data!$A$1</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([{ name: 'Good', formula: 'Data!$A$1' }])
    expect(state.preserveNames).toEqual(['A1', 'R2C3', 'true', '1Starts', 'Has Space'])
  })

  it('preserves out-of-range scopes and empty bodies instead of modeling them', () => {
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="Dangling" localSheetId="7">Data!$A$1</definedName>' +
          '<definedName name="EmptyBody">   </definedName>' +
          '<definedName name="Good">Data!$A$1</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([{ name: 'Good', formula: 'Data!$A$1' }])
    expect(state.preserveNames).toEqual(['Dangling', 'EmptyBody'])
  })

  it('ranks duplicate names the way the desktop installs them and preserves the losers', () => {
    // Excel_Version exists as a #REF! sheet-scoped residue, a live
    // sheet-scoped definition, and a live workbook-level definition — the
    // desktop's applyDefinedNames ordering test exactly. The workbook-level
    // definition wins; the losers are preserved (one deduped preserve
    // entry covers every element with that name — the writer checks
    // membership by name).
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="Excel_Version" localSheetId="0">#REF!</definedName>' +
          '<definedName name="Excel_Version" localSheetId="1">Other!$H$9</definedName>' +
          '<definedName name="Excel_Version">Data!$H$9</definedName>' +
          '<definedName name="Unrelated">Data!$A$1</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([
      { name: 'Excel_Version', formula: 'Data!$H$9' },
      { name: 'Unrelated', formula: 'Data!$A$1' },
    ])
    expect(state.preserveNames).toEqual(['Excel_Version'])
  })

  it('groups case-variant duplicates case-insensitively (engine resolution parity)', () => {
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="Foo">Data!$A$1</definedName>' +
          '<definedName name="foo">Data!$A$2</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([{ name: 'Foo', formula: 'Data!$A$1' }])
    expect(state.preserveNames).toEqual(['foo'])
  })

  it('keeps #REF! residues modelable (they round-trip verbatim)', () => {
    const state = parseDefinedNamesState(
      wb('<definedNames><definedName name="Version">#REF!</definedName></definedNames>'),
      2,
    )
    expect(state.names).toEqual([{ name: 'Version', formula: '#REF!' }])
  })

  it('unescapes XML entities in names and formulas', () => {
    const state = parseDefinedNamesState(
      wb(
        '<definedNames>' +
          '<definedName name="Cmp">IF(A1&lt;5,"a &amp; b","c")</definedName>' +
          '</definedNames>',
      ),
      2,
    )
    expect(state.names).toEqual([{ name: 'Cmp', formula: 'IF(A1<5,"a & b","c")' }])
  })

  it('fails closed on a non-numeric localSheetId', () => {
    expect(() =>
      parseDefinedNamesState(
        wb(
          '<definedNames><definedName name="X" localSheetId="abc">Data!$A$1</definedName></definedNames>',
        ),
        2,
      ),
    ).toThrow(DefinedNamesReadError)
  })

  it('fails closed on an element with no name attribute', () => {
    expect(() =>
      parseDefinedNamesState(
        wb('<definedNames><definedName localSheetId="0">Data!$A$1</definedName></definedNames>'),
        2,
      ),
    ).toThrow(DefinedNamesReadError)
  })

  it('fails closed when the section carries elements the scanner cannot account for', () => {
    // A nested construct the element scanner never matches leaves
    // unaccounted <definedName> openings — the whole family locks.
    expect(() =>
      parseDefinedNamesState(
        wb(
          '<definedNames><definedName name="A">Data!$A$1</definedName>' +
            '<definedName name="B">Data!$A$1</definedName></definedNames>',
        ).replace('</definedName><definedName name="B">', '</definedName><definedName name="B">'),
        2,
      ),
    ).not.toThrow()
    expect(() =>
      parseDefinedNamesState(
        wb(
          '<definedNames><definedName name="A">Data!$A$1<definedName name="B">Data!$A$1</definedName></definedName></definedNames>',
        ),
        2,
      ),
    ).toThrow(DefinedNamesReadError)
  })
})

describe('definedNameIsSaveable (the single canonical name predicate)', () => {
  it('accepts Excel-legal names including Unicode', () => {
    for (const name of ['Revenue', 'Tax_Rate', 'sales.total', '\\Backslash', '数据区域']) {
      expect(definedNameIsSaveable(name)).toBe(true)
    }
  })
  it('rejects the writer-forbidden forms', () => {
    for (const name of [
      '',
      'A1',
      'XFD1048576',
      'R2C3',
      'TRUE',
      'false',
      '_xlnm.Print_Area',
      '1Starts',
      'Has Space',
      'α[name]',
      'x'.repeat(256),
    ]) {
      expect(definedNameIsSaveable(name)).toBe(false)
    }
  })
})

/// ── Snapshot-level tests (readBasicWorkbook) ────────────────────────────────

async function buildNamesFixture(workbookDefinedNames: string): Promise<Buffer> {
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
      workbookDefinedNames +
      '<calcPr calcId="191029"/>' +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '</Relationships>',
  )
  const worksheet =
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="A2"><v>2</v></c></row></sheetData>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>'
  zip.file('xl/worksheets/sheet1.xml', worksheet)
  zip.file('xl/worksheets/sheet2.xml', worksheet)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function unzipWorkbookXml(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  return (await zip.file('xl/workbook.xml')!.async('string')) as string
}

describe('readBasicWorkbook — definedNames on the snapshot', () => {
  it('surfaces modeled names and the preserve list', async () => {
    const buffer = await buildNamesFixture(
      '<definedNames>' +
        '<definedName name="GlobalTotal">Data!$A$1</definedName>' +
        '<definedName name="LocalTotal" localSheetId="1">Other!$A$2</definedName>' +
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>' +
        '<definedName name="Hidden" hidden="1">Data!$A$2</definedName>' +
        '<definedName name="A1">Data!$A$1</definedName>' +
        '</definedNames>',
    )
    const { snapshot } = await readBasicWorkbook(buffer)
    expect(snapshot.definedNames).toEqual({
      names: [
        { name: 'GlobalTotal', formula: 'Data!$A$1' },
        { name: 'LocalTotal', formula: 'Other!$A$2', sheetIndex: 1 },
      ],
      preserveNames: ['A1'],
    })
    expect(snapshot.namesLocked).toBeUndefined()
  })

  it('omits the field entirely when the workbook carries no names', async () => {
    const buffer = await buildNamesFixture('')
    const { snapshot } = await readBasicWorkbook(buffer)
    expect(snapshot.definedNames).toBeUndefined()
    expect(snapshot.namesLocked).toBeUndefined()
  })

  it('marks an unparseable section namesLocked (fail closed, workbook still opens)', async () => {
    const buffer = await buildNamesFixture(
      '<definedNames><definedName localSheetId="0">Data!$A$1</definedName></definedNames>',
    )
    const { snapshot } = await readBasicWorkbook(buffer)
    expect(snapshot.definedNames).toBeUndefined()
    expect(snapshot.namesLocked).toBe(true)
    expect(snapshot.sheets.length).toBe(2)
  })
})

/// ── Save-family + byte-preservation tests (the critical invariant) ─────────

describe('definedNamesState save — preservation invariant', () => {
  const NAMES_XML =
    '<definedNames>' +
    '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>' +
    '<definedName name="GlobalTotal">Data!$A$1</definedName>' +
    '<definedName name="SecondName">Data!$A$2</definedName>' +
    '<definedName name="Hidden" hidden="1">Data!$A$2</definedName>' +
    '<definedName name="Unparseable">EXOTIC(1)</definedName>' +
    '</definedNames>'

  it('a save WITHOUT the family preserves the section content (no-op)', async () => {
    const buffer = await buildNamesFixture(NAMES_XML)
    const before = await unzipWorkbookXml(buffer)
    const mutation = await applyCellEditsToXlsx(buffer, [])
    const after = await unzipWorkbookXml(mutation.buffer)
    // The no-op save rewrites workbook.xml only for the documented global
    // fullCalcOnLoad marker; the <definedNames> section itself survives
    // byte-for-byte.
    expect(after).toContain(
      '<definedNames>' +
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>' +
        '<definedName name="GlobalTotal">Data!$A$1</definedName>' +
        '<definedName name="SecondName">Data!$A$2</definedName>' +
        '<definedName name="Hidden" hidden="1">Data!$A$2</definedName>' +
        '<definedName name="Unparseable">EXOTIC(1)</definedName>' +
        '</definedNames>',
    )
    expect(after.replace(' fullCalcOnLoad="1"', '')).toBe(before)
  })

  it('editing ONE name preserves every sibling, the hidden name, and the print titles', async () => {
    const buffer = await buildNamesFixture(NAMES_XML)
    const mutation = await applyCellEditsToXlsx(buffer, [], [], [], undefined, [], [], [], [], [], {
      names: [
        { name: 'GlobalTotal', formula: 'Data!$A$1:$A$9' },
        { name: 'SecondName', formula: 'Data!$A$2' },
      ],
      preserveNames: [],
    })
    const xml = await unzipWorkbookXml(mutation.buffer)
    expect(xml).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$9</definedName>')
    // The untouched sibling survives (never silently dropped).
    expect(xml).toContain('<definedName name="SecondName">Data!$A$2</definedName>')
    // The print titles and hidden names survive byte-verbatim.
    expect(xml).toContain(
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>',
    )
    expect(xml).toContain('<definedName name="Hidden" hidden="1">Data!$A$2</definedName>')
  })

  it('deleting one name preserves the others (declarative full-model snapshot)', async () => {
    const buffer = await buildNamesFixture(NAMES_XML)
    const mutation = await applyCellEditsToXlsx(buffer, [], [], [], undefined, [], [], [], [], [], {
      names: [{ name: 'SecondName', formula: 'Data!$A$2' }],
      preserveNames: [],
    })
    const xml = await unzipWorkbookXml(mutation.buffer)
    expect(xml).not.toContain('name="GlobalTotal"')
    expect(xml).toContain('<definedName name="SecondName">Data!$A$2</definedName>')
    expect(xml).toContain('name="_xlnm.Print_Titles"')
    expect(xml).toContain('name="Hidden"')
  })

  it('preserve-listed names survive while modeled ones rewrite', async () => {
    const buffer = await buildNamesFixture(NAMES_XML)
    const mutation = await applyCellEditsToXlsx(buffer, [], [], [], undefined, [], [], [], [], [], {
      names: [{ name: 'SecondName', formula: 'Data!$A$2' }],
      preserveNames: ['Unparseable'],
    })
    const xml = await unzipWorkbookXml(mutation.buffer)
    expect(xml).toContain('name="Unparseable"')
    expect(xml).toContain('name="_xlnm.Print_Titles"')
  })

  it('write → reopen round-trip: the reader reads back what the writer wrote', async () => {
    const buffer = await buildNamesFixture('')
    const state = {
      names: [
        { name: 'Revenue', formula: 'Data!$A$1:$A$9' },
        { name: 'LocalTotal', formula: 'Other!$A$2', sheetIndex: 1 },
      ],
      preserveNames: [],
    }
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
      state,
    )
    const { snapshot } = await readBasicWorkbook(mutation.buffer)
    expect(snapshot.definedNames).toEqual(state)
  })

  it('rejects names + structural ops in the SAME save (the split-save contract)', async () => {
    const buffer = await buildNamesFixture(NAMES_XML)
    await expect(
      applyCellEditsToXlsx(
        buffer,
        [],
        [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] }],
        [],
        undefined,
        [],
        [],
        [],
        [],
        [],
        { names: [{ name: 'Revenue', formula: 'Data!$A$1' }], preserveNames: [] },
      ),
    ).rejects.toThrow(/cannot be saved together/)
  })

  it('the browser split-save: structural phase then names phase shifts references consistently', async () => {
    const buffer = await buildNamesFixture(
      '<definedNames><definedName name="Revenue">Data!$A$1:$A$5</definedName></definedNames>',
    )
    // Phase 1: insert 2 rows at the top (names NOT dirty — the section
    // shifts through the structural replay).
    const phase1 = await applyCellEditsToXlsx(
      buffer,
      [],
      [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] }],
    )
    expect(await unzipWorkbookXml(phase1.buffer)).toContain(
      '<definedName name="Revenue">Data!$A$3:$A$7</definedName>',
    )
    // Phase 2: the names snapshot (post-shift coordinates, exactly what the
    // live model reports after Univer's UpdateDefinedNameController) rewrites
    // the section against the phase-1 bytes.
    const phase2 = await applyCellEditsToXlsx(
      phase1.buffer,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      { names: [{ name: 'Revenue', formula: 'Data!$A$3:$A$7' }], preserveNames: [] },
    )
    expect(await unzipWorkbookXml(phase2.buffer)).toContain(
      '<definedName name="Revenue">Data!$A$3:$A$7</definedName>',
    )
    const { snapshot } = await readBasicWorkbook(phase2.buffer)
    expect(snapshot.definedNames).toEqual({
      names: [{ name: 'Revenue', formula: 'Data!$A$3:$A$7' }],
      preserveNames: [],
    })
  })

  it('only workbook.xml changes on a names-only save (package-level byte evidence)', async () => {
    const buffer = await buildNamesFixture(NAMES_XML)
    const mutation = await applyCellEditsToXlsx(buffer, [], [], [], undefined, [], [], [], [], [], {
      names: [
        { name: 'GlobalTotal', formula: 'Data!$A$1:$A$9' },
        { name: 'SecondName', formula: 'Data!$A$2' },
      ],
      preserveNames: [],
    })
    const before = await JSZip.loadAsync(buffer)
    const after = await JSZip.loadAsync(mutation.buffer)
    expect(await before.file('xl/worksheets/sheet1.xml')!.async('string')).toBe(
      await after.file('xl/worksheets/sheet1.xml')!.async('string'),
    )
    expect(await before.file('xl/worksheets/sheet2.xml')!.async('string')).toBe(
      await after.file('xl/worksheets/sheet2.xml')!.async('string'),
    )
    expect(await before.file('[Content_Types].xml')!.async('string')).toBe(
      await after.file('[Content_Types].xml')!.async('string'),
    )
    expect(await before.file('xl/_rels/workbook.xml.rels')!.async('string')).toBe(
      await after.file('xl/_rels/workbook.xml.rels')!.async('string'),
    )
  })
})
