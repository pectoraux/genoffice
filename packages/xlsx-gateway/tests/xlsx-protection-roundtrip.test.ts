/**
 * EXCEL-020 — Sheet/Workbook Protection round-trip tests.
 *
 * Proves the canonical protection family end-to-end at the gateway:
 *
 *   read: <sheetProtection> (per worksheet) + <workbookProtection>
 *         (workbook.xml) → WorksheetState.sheetProtection /
 *         WorkbookSnapshot.workbookProtection, with hasPassword flags for
 *         BOTH password forms so the browser can refuse an unprotect
 *         toggle up front (the gateway itself fails closed).
 *
 *   write: applyCellEditsToXlsx(sheetProtections) writes/removes the
 *         worksheet element; the new trailing workbookProtectionState
 *         argument writes/removes lockStructure in workbook.xml. Both
 *         fail closed on password-bearing elements.
 *
 *   round-trip: protect → save → readBasicWorkbook → protected state;
 *         unprotect → save → read → unprotected. A no-op save (empty
 *         families) preserves the source bytes exactly.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'
import type { SheetProtectionState } from '../src/gateway/xlsx-gateway'
import {
  parseSheetProtectionState,
  parseWorkbookProtectionState,
} from '../src/gateway/xlsx-protection'

/// Minimal single-sheet workbook whose worksheet may carry a
/// <sheetProtection> element and whose workbook.xml may carry a
/// <workbookProtection> element.
async function buildProtectionFixture(options?: {
  readonly sheetProtectionAttrs?: string
  readonly workbookProtectionAttrs?: string
}): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
  ${options?.workbookProtectionAttrs ? `<workbookProtection ${options.workbookProtectionAttrs}/>` : ''}
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c></row>
  </sheetData>
  ${options?.sheetProtectionAttrs ? `<sheetProtection ${options.sheetProtectionAttrs}/>` : ''}
</worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function worksheetXml(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  return zip.file('xl/worksheets/sheet1.xml')!.async('string')
}

async function workbookXml(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  return zip.file('xl/workbook.xml')!.async('string')
}

// ── Element parsers ─────────────────────────────────────────────────────────

describe('parseSheetProtectionState', () => {
  it('returns null without an element', () => {
    expect(parseSheetProtectionState('<worksheet><sheetData/></worksheet>')).toBeNull()
  })

  it('reads an enabled element', () => {
    const xml =
      '<worksheet><sheetData/><sheetProtection sheet="1" objects="1" scenarios="1"/></worksheet>'
    expect(parseSheetProtectionState(xml)).toEqual({ protected: true, hasPassword: false })
  })

  it('reads a disabled element as not protected', () => {
    const xml =
      '<worksheet><sheetData/><sheetProtection sheet="0" selectLockedCells="1"/></worksheet>'
    expect(parseSheetProtectionState(xml)).toEqual({ protected: false, hasPassword: false })
  })

  it('flags the legacy password hash', () => {
    const xml = '<worksheet><sheetData/><sheetProtection sheet="1" password="83AF"/></worksheet>'
    expect(parseSheetProtectionState(xml)).toEqual({ protected: true, hasPassword: true })
  })

  it('flags the modern algorithm hash', () => {
    const xml =
      '<worksheet><sheetData/><sheetProtection sheet="1" algorithmName="SHA-512" hashValue="x" saltValue="y" spinCount="100000"/></worksheet>'
    expect(parseSheetProtectionState(xml)).toEqual({ protected: true, hasPassword: true })
  })

  it('accepts the boolean true form', () => {
    const xml = '<worksheet><sheetData/><sheetProtection sheet="true"/></worksheet>'
    expect(parseSheetProtectionState(xml)).toEqual({ protected: true, hasPassword: false })
  })
})

describe('parseWorkbookProtectionState', () => {
  it('returns null without an element', () => {
    expect(parseWorkbookProtectionState('<workbook><sheets/></workbook>')).toBeNull()
  })

  it('reads a locked structure', () => {
    const xml = '<workbook><workbookProtection lockStructure="1"/><sheets/></workbook>'
    expect(parseWorkbookProtectionState(xml)).toEqual({ lockStructure: true, hasPassword: false })
  })

  it('reads an unlocked structure', () => {
    const xml = '<workbook><workbookProtection lockStructure="0"/><sheets/></workbook>'
    expect(parseWorkbookProtectionState(xml)).toEqual({ lockStructure: false, hasPassword: false })
  })

  it('flags the legacy workbook password', () => {
    const xml =
      '<workbook><workbookProtection lockStructure="1" workbookPassword="83AF"/><sheets/></workbook>'
    expect(parseWorkbookProtectionState(xml)).toEqual({ lockStructure: true, hasPassword: true })
  })

  it('flags the modern workbook hash', () => {
    const xml =
      '<workbook><workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="x"/><sheets/></workbook>'
    expect(parseWorkbookProtectionState(xml)).toEqual({ lockStructure: true, hasPassword: true })
  })
})

// ── readBasicWorkbook integration ───────────────────────────────────────────

describe('readBasicWorkbook protection state (EXCEL-020)', () => {
  it('surfaces a protected sheet and locked workbook', async () => {
    const buffer = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" objects="1" scenarios="1"',
      workbookProtectionAttrs: 'lockStructure="1"',
    })
    const { snapshot } = await readBasicWorkbook(buffer)
    expect(snapshot.sheets[0]!.sheetProtection).toEqual({
      protected: true,
      hasPassword: false,
    })
    expect(snapshot.workbookProtection).toEqual({ lockStructure: true, hasPassword: false })
  })

  it('omits both fields for an unprotected workbook', async () => {
    const buffer = await buildProtectionFixture()
    const { snapshot } = await readBasicWorkbook(buffer)
    expect(snapshot.sheets[0]!.sheetProtection).toBeUndefined()
    expect(snapshot.workbookProtection).toBeUndefined()
  })

  it('surfaces password flags for both password forms', async () => {
    const legacy = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" password="83AF"',
      workbookProtectionAttrs: 'lockStructure="1" workbookPassword="83AF"',
    })
    let snapshot = (await readBasicWorkbook(legacy)).snapshot
    expect(snapshot.sheets[0]!.sheetProtection).toEqual({ protected: true, hasPassword: true })
    expect(snapshot.workbookProtection).toEqual({ lockStructure: true, hasPassword: true })

    const modern = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" algorithmName="SHA-512" hashValue="x"',
      workbookProtectionAttrs:
        'lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="x"',
    })
    snapshot = (await readBasicWorkbook(modern)).snapshot
    expect(snapshot.sheets[0]!.sheetProtection).toEqual({ protected: true, hasPassword: true })
    expect(snapshot.workbookProtection).toEqual({ lockStructure: true, hasPassword: true })
  })
})

// ── Write path through applyCellEditsToXlsx ─────────────────────────────────

async function saveProtection(
  buffer: Buffer,
  sheetProtections: readonly SheetProtectionState[],
  workbookProtectionState: { readonly lockStructure: boolean } | null,
) {
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
    sheetProtections,
    null,
    [],
    [],
    [],
    workbookProtectionState,
  )
  return mutation
}

describe('applyCellEditsToXlsx protection family (EXCEL-020)', () => {
  it('writes sheet protection with Excel defaults and locks the structure', async () => {
    const source = await buildProtectionFixture()
    const saved = await saveProtection(source, [{ sheetName: 'Data', protected: true }], {
      lockStructure: true,
    })
    expect(await worksheetXml(saved.buffer)).toContain(
      '<sheetProtection sheet="1" objects="1" scenarios="1"/>',
    )
    expect(await workbookXml(saved.buffer)).toContain('<workbookProtection lockStructure="1"/>')
  })

  it('removes protection elements on unprotect', async () => {
    const source = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" objects="1" scenarios="1"',
      workbookProtectionAttrs: 'lockStructure="1"',
    })
    const saved = await saveProtection(source, [{ sheetName: 'Data', protected: false }], {
      lockStructure: false,
    })
    expect(await worksheetXml(saved.buffer)).not.toContain('<sheetProtection')
    expect(await workbookXml(saved.buffer)).not.toContain('<workbookProtection')
  })

  it('round-trips: protect → save → read → protected; unprotect → read → clear', async () => {
    const source = await buildProtectionFixture()
    const protectedBuffer = await saveProtection(source, [{ sheetName: 'Data', protected: true }], {
      lockStructure: true,
    })
    let snapshot = (await readBasicWorkbook(protectedBuffer.buffer)).snapshot
    expect(snapshot.sheets[0]!.sheetProtection).toEqual({ protected: true, hasPassword: false })
    expect(snapshot.workbookProtection).toEqual({ lockStructure: true, hasPassword: false })

    const unprotectedBuffer = await saveProtection(
      protectedBuffer.buffer,
      [{ sheetName: 'Data', protected: false }],
      { lockStructure: false },
    )
    snapshot = (await readBasicWorkbook(unprotectedBuffer.buffer)).snapshot
    expect(snapshot.sheets[0]!.sheetProtection).toBeUndefined()
    expect(snapshot.workbookProtection).toBeUndefined()
  })

  it('preserves the protection XML when both families are untouched', async () => {
    const source = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" objects="1" scenarios="1"',
    })
    const mutation = await saveProtection(source, [], null)
    // The worksheet part is not even mentioned by the empty plan — its
    // <sheetProtection> element survives byte-identically. (workbook.xml
    // legitimately gains fullCalcOnLoad — that is the gateway's standing
    // recalc behavior, not a protection change.)
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
    expect(await worksheetXml(mutation.buffer)).toContain(
      '<sheetProtection sheet="1" objects="1" scenarios="1"/>',
    )
    const before = new Map(mutation.beforeEntries.map((e) => [e.path, e.sha256]))
    const after = new Map(mutation.afterEntries.map((e) => [e.path, e.sha256]))
    expect(after.get('xl/worksheets/sheet1.xml')).toBe(before.get('xl/worksheets/sheet1.xml'))
  })

  it('fails closed when unprotecting a password-protected sheet', async () => {
    const source = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" password="83AF"',
    })
    await expect(
      saveProtection(source, [{ sheetName: 'Data', protected: false }], null),
    ).rejects.toThrow(/password/)
  })

  it('fails closed when unlocking a password-protected workbook structure', async () => {
    const source = await buildProtectionFixture({
      workbookProtectionAttrs: 'lockStructure="1" workbookPassword="83AF"',
    })
    await expect(saveProtection(source, [], { lockStructure: false })).rejects.toThrow(/password/)
  })

  it('re-affirming protection on a password-protected sheet keeps the element verbatim', async () => {
    const source = await buildProtectionFixture({
      sheetProtectionAttrs: 'sheet="1" password="83AF"',
    })
    const mutation = await saveProtection(source, [{ sheetName: 'Data', protected: true }], null)
    // applySheetProtection returns the element verbatim when sheet="1" is
    // already set — the password attribute survives (fail-closed state is
    // never silently dropped or rewritten).
    const xml = await worksheetXml(mutation.buffer)
    expect(xml).toContain('<sheetProtection sheet="1" password="83AF"/>')
    const before = new Map(mutation.beforeEntries.map((e) => [e.path, e.sha256]))
    const after = new Map(mutation.afterEntries.map((e) => [e.path, e.sha256]))
    expect(after.get('xl/worksheets/sheet1.xml')).toBe(before.get('xl/worksheets/sheet1.xml'))
  })

  it('throws on an unknown sheet name (fail-closed)', async () => {
    const source = await buildProtectionFixture()
    await expect(
      saveProtection(source, [{ sheetName: 'Nope', protected: true }], null),
    ).rejects.toThrow()
  })
})
