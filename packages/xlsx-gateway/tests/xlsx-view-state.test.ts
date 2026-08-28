import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'

/**
 * EXCEL-026 — View / Page Setup persistence.
 *
 * Reader: parseSheetViewState exposes ONLY non-default <sheetView> display
 * flags (showGridLines/showRowColHeaders off, showFormulas on) plus the
 * frozen pane (parseFrozenPane). Writer: the canonical applyPageSetupState
 * (via the applyCellEditsToXlsx pageSetupStates slot) writes sheetView
 * attributes, the <pane> element, and the print page-setup family.
 *
 * These tests prove the full read → write → reopen contract through the
 * REAL gateway entry points (readBasicWorkbook / applyCellEditsToXlsx):
 * parse existing state, absent state, explicit defaults, malformed values,
 * write → reopen, edit → save → reopen, clear/reset, no-op preservation,
 * and unrelated worksheet/package preservation.
 */

/// Two-sheet fixture with a configurable first <sheetView> per sheet.
async function buildViewFixture(
  sheet1View: string,
  sheet2View: string,
  extraSheet1Tail = '',
): Promise<Buffer> {
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
  const worksheet = (view: string, tail: string): string =>
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<sheetViews>${view}</sheetViews>` +
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>' +
    tail +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>'
  zip.file('xl/worksheets/sheet1.xml', worksheet(sheet1View, extraSheet1Tail))
  zip.file('xl/worksheets/sheet2.xml', worksheet(sheet2View, ''))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const readEntry = async (buffer: Buffer, entry: string): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer)
  return zip.file(entry)!.async('string')
}

/// applyCellEditsToXlsx with pageSetupStates in its positional slot.
async function saveWithPageSetup(
  buffer: Buffer,
  states: readonly Record<string, unknown>[],
): Promise<Buffer> {
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
    states as never,
  )
  return mutation.buffer
}

describe('EXCEL-026 reader — parseSheetViewState', () => {
  it('parses existing non-default view state (gridlines hidden, formulas shown)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="0" showFormulas="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toEqual({
      showGridlines: false,
      showFormulas: true,
    })
    expect(imported.snapshot.sheets[1]!.view).toBeUndefined()
  })

  it('parses hidden headings (showRowColHeaders off)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showRowColHeaders="0" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toEqual({ showHeadings: false })
  })

  it('absent state — no sheetView attributes means no view field', async () => {
    const buffer = await buildViewFixture(
      '<sheetView tabSelected="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toBeUndefined()
    expect(imported.snapshot.sheets[1]!.view).toBeUndefined()
  })

  it('explicit defaults are NOT exposed (showGridLines="1", showFormulas="0")', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="1" showFormulas="0" showRowColHeaders="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toBeUndefined()
  })

  it('accepts xsd:boolean literals "true"/"false"', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="false" showFormulas="true" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toEqual({
      showGridlines: false,
      showFormulas: true,
    })
  })

  it('malformed values are ignored for modeling (treated as default)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="2" showFormulas="yes" showRowColHeaders="maybe" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toBeUndefined()
  })

  it('malformed value does not hide a valid sibling flag', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="2" showFormulas="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toEqual({ showFormulas: true })
  })

  it('no sheetViews section at all — no view field, workbook still opens', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
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
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    )
    const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.view).toBeUndefined()
  })
})

describe('EXCEL-026 reader — parseFrozenPane (unit gap fill)', () => {
  it('parses a frozen pane (rows + columns)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/></sheetView>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.freeze).toEqual({ frozenRows: 3, frozenColumns: 2 })
  })

  it('a split (non-frozen) pane is not a freeze', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" state="split"/></sheetView>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.freeze).toBeUndefined()
  })

  it('frozenSplit counts (Excel writes it after dragging split bars)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" state="frozenSplit"/></sheetView>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.freeze).toEqual({ frozenRows: 2, frozenColumns: 0 })
  })

  it('zero splits mean no freeze', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"><pane state="frozen"/></sheetView>',
      '<sheetView workbookViewId="0"/>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.freeze).toBeUndefined()
  })
})

describe('EXCEL-026 write → reopen round-trips', () => {
  it('existing view state round-trips when re-emitted verbatim (write → reopen)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="0" showFormulas="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [
      { sheetName: 'Data', showGridlines: false, showFormulas: true },
    ])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('showGridLines="0"')
    expect(xml).toContain('showFormulas="1"')
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.view).toEqual({
      showGridlines: false,
      showFormulas: true,
    })
    // Unrelated sheet untouched.
    expect(reopened.snapshot.sheets[1]!.view).toBeUndefined()
  })

  it('edit → save → reopen: hiding gridlines persists', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [{ sheetName: 'Data', showGridlines: false }])
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.view).toEqual({ showGridlines: false })
  })

  it('edit → save → reopen: enabling formula view persists', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [{ sheetName: 'Data', showFormulas: true }])
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.view).toEqual({ showFormulas: true })
  })

  it('clear/reset: restoring gridlines drops the attribute and reopens as default', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="0" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [{ sheetName: 'Data', showGridlines: true }])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('showGridLines')
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.view).toBeUndefined()
  })

  it('clear/reset: turning formula view off drops the attribute and reopens as default', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showFormulas="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [{ sheetName: 'Data', showFormulas: false }])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('showFormulas')
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.view).toBeUndefined()
  })

  it('clear/reset: freeze 0/0 removes the pane and reopens unfrozen', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [
      { sheetName: 'Data', frozenRows: 0, frozenColumns: 0 },
    ])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('<pane')
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.freeze).toBeUndefined()
  })

  it('a user toggle replaces a malformed attribute with a definite canonical value', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="2" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const saved = await saveWithPageSetup(buffer, [{ sheetName: 'Data', showGridlines: false }])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('showGridLines="0"')
    expect(xml).not.toContain('showGridLines="2"')
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.view).toEqual({ showGridlines: false })
  })

  it('freeze edit round-trips through the slot (set → reopen → clear → reopen)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const frozen = await saveWithPageSetup(buffer, [
      { sheetName: 'Data', frozenRows: 3, frozenColumns: 2 },
    ])
    const xml = await readEntry(frozen, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('<pane xSplit="2" ySplit="3"')
    expect(xml).toContain('state="frozen"')
    const reopened = await readBasicWorkbook(frozen)
    expect(reopened.snapshot.sheets[0]!.freeze).toEqual({ frozenRows: 3, frozenColumns: 2 })
    const cleared = await saveWithPageSetup(frozen, [
      { sheetName: 'Data', frozenRows: 0, frozenColumns: 0 },
    ])
    const reopened2 = await readBasicWorkbook(cleared)
    expect(reopened2.snapshot.sheets[0]!.freeze).toBeUndefined()
  })
})

describe('EXCEL-026 print page-setup family through the slot', () => {
  it('orientation / margins / paperSize / scale / fit reach the worksheet XML', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
      '<pageSetup paperSize="9" orientation="portrait" scale="100"/>',
    )
    const saved = await saveWithPageSetup(buffer, [
      {
        sheetName: 'Data',
        orientation: 'landscape',
        margins: 'wide',
        paperSize: 1,
        scale: 75,
        fitToWidth: 2,
        fitToHeight: 3,
        fitToPage: true,
      },
    ])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('orientation="landscape"')
    expect(xml).toContain('paperSize="1"')
    expect(xml).toContain('scale="75"')
    expect(xml).toContain('fitToWidth="2"')
    expect(xml).toContain('fitToHeight="3"')
    expect(xml).toContain('<pageSetUpPr fitToPage="1"')
    const margins = /<pageMargins\b[^>]*>/.exec(xml)?.[0] ?? ''
    expect(margins).toContain('left="1"')
    expect(margins).toContain('top="1"')
    // The workbook still opens after the print edits.
    const reopened = await readBasicWorkbook(saved)
    expect(reopened.snapshot.sheets[0]!.cells['A1']!.value).toBe(1)
  })

  it('untouched print attributes stay verbatim when another field is edited', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
      '<pageSetup paperSize="9" orientation="portrait" scale="100" blackAndWhite="1"/>',
    )
    const saved = await saveWithPageSetup(buffer, [{ sheetName: 'Data', orientation: 'landscape' }])
    const xml = await readEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('orientation="landscape"')
    expect(xml).toContain('paperSize="9"')
    // An attribute the canonical model does not know stays verbatim.
    expect(xml).toContain('blackAndWhite="1"')
  })
})

describe('EXCEL-026 no-op and package preservation', () => {
  it('no pageSetupStates → the worksheet part stays byte-identical', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="0" showFormulas="1" workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    const mutation = await applyCellEditsToXlsx(buffer, [])
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet2.xml')
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
    expect(after.get('xl/worksheets/sheet1.xml')).toBe(before.get('xl/worksheets/sheet1.xml'))
    expect(after.get('xl/worksheets/sheet2.xml')).toBe(before.get('xl/worksheets/sheet2.xml'))
  })

  it('editing one sheet leaves the other worksheet and unrelated parts byte-identical', async () => {
    const buffer = await buildViewFixture(
      '<sheetView showGridLines="0" workbookViewId="0"/>',
      '<sheetView showGridLines="0" showFormulas="1" workbookViewId="0"/>',
    )
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
      [{ sheetName: 'Data', showGridlines: true } as never],
    )
    expect(mutation.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet2.xml')
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
    expect(after.get('xl/worksheets/sheet2.xml')).toBe(before.get('xl/worksheets/sheet2.xml'))
    // The untouched sheet KEEPS its non-default view state.
    const reopened = await readBasicWorkbook(mutation.buffer)
    expect(reopened.snapshot.sheets[1]!.view).toEqual({
      showGridlines: false,
      showFormulas: true,
    })
  })

  it('a page-setup state for an unknown sheet name fails closed (throws)', async () => {
    const buffer = await buildViewFixture(
      '<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"/>',
    )
    await expect(
      saveWithPageSetup(buffer, [{ sheetName: 'DoesNotExist', showGridlines: false }]),
    ).rejects.toThrow('Sheet "DoesNotExist" was not found')
  })
})
