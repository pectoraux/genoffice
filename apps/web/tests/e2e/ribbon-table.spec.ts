/**
 * REAL browser E2E — Insert → Tables (EXCEL-021).
 *
 * Proves the table persistence chain end-to-end through the REAL HTTP
 * boundary:
 *
 *   open → readBasicWorkbook resolves <tableParts> through the worksheet
 *   rels → WorksheetState.tables (metadata + PRE-RESOLVED banding colors)
 *   → browser seeds its file-state refs, paints the banding into the cell
 *   matrix, registers the visual Univer table (muted plain theme), and
 *   installs the table-owned filter origin → Insert → Table journals the
 *   canonical SheetTableAddition (desktop applyAiTableAdd parity) → save
 *   emits the typed tableAdditions family → /api/office/workbooks/save →
 *   routeOffice strict validation → applyCellEditsToXlsx trailing param →
 *   applyTableAdditions writes xl/tables/tableN.xml + <tableParts> + rel
 *   + [Content_Types] override → reopen → file-native semantics.
 *
 * Delete is convert-to-range for session tables (journal splice — the
 * bytes never change); file-native tables refuse with the desktop's exact
 * message. A sheet whose filter belongs to a table refuses filter edits
 * (BeforeCommandExecute gate). No-op saves preserve the table parts
 * byte-for-byte.
 *
 * Fixtures (buildTableLedgerFixture variants): the notes-ledger shape
 * with (a) an existing SalesTable (TableStyleMedium2, header + 3 data
 * rows, autoFilter inside the table part — the worksheet carries none),
 * (b) no table at all (create-from-scratch).
 */
import { test, expect, type Page } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import {
  buildExcelTableFixture,
  buildExcelTableCreateFixture,
  readZipEntry,
  listZipEntries,
} from './fixtures'

const NAME_BOX = '[data-testid="excel-name-box"]'

/** Snapshot view of the open response's table state. */
interface TableSnapshotView {
  sheets: Array<{
    name: string
    tables?: Array<{
      name?: string
      area: { startRow: number; startColumn: number; endRow: number; endColumn: number }
      headerRowCount: number
      columns: string[]
      styleName?: string
      showRowStripes: boolean
      headerFill?: string
      headerFontColor?: string
      stripeFill?: string
    }>
  }>
}

/** Select a range (or single cell) through the Name Box. */
async function selectRange(page: Page, ref: string): Promise<void> {
  const box = page.locator(NAME_BOX)
  await box.click()
  await box.fill(ref)
  await box.press('Enter')
  await page.waitForTimeout(400)
}

/** Insert rows on the named sheet through the real Univer facade. */
async function insertRows(
  page: Page,
  sheetName: string,
  row: number,
  count: number,
): Promise<void> {
  await page.evaluate(
    ({ name, rowIndex, n }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(name: string): { insertRows(row: number, count?: number): unknown }
            }
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const wb = rt.univerAPI.getActiveWorkbook()
      if (!wb) throw new Error('no active workbook')
      const ws = wb.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      ws.insertRows(rowIndex, n)
    },
    { name: sheetName, rowIndex: row, n: count },
  )
}

/** Delete rows on the named sheet through the real Univer facade. */
async function deleteRows(
  page: Page,
  sheetName: string,
  row: number,
  count: number,
): Promise<void> {
  await page.evaluate(
    ({ name, rowIndex, n }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(name: string): { deleteRows(row: number, count?: number): unknown }
            }
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const wb = rt.univerAPI.getActiveWorkbook()
      if (!wb) throw new Error('no active workbook')
      const ws = wb.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      ws.deleteRows(rowIndex, n)
    },
    { name: sheetName, rowIndex: row, n: count },
  )
}

/** Open a fixture file and wait for the editor to finish loading it. */
async function openFixture(page: Page, bytes: Buffer, path: string): Promise<void> {
  writeFileSync(path, bytes)
  await page.setInputFiles('input[type="file"]', path)
  await expect(page.getByText(`Opened ${path.split('/').pop()!}`)).toBeVisible({
    timeout: 30_000,
  })
  await page.waitForTimeout(1500)
}

test.describe('Insert tab — Tables persist through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1: opening a workbook with a table surfaces the metadata (snapshot + delete refusal + no undo pollution)', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelTableFixture()
    writeFileSync('/tmp/e2e-ribbon-table-open.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-table-open.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-table-open.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // READ path: the snapshot carried the parsed table metadata with the
    // Medium2 palette resolved against the DEFAULT Office accents.
    const snapshot = (await openResponse.json()).snapshot as TableSnapshotView
    const table = snapshot.sheets[0].tables?.[0]
    expect(table, 'snapshot carries the table').toBeDefined()
    expect(table!.name).toBe('SalesTable')
    expect(table!.area).toEqual({ startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 })
    expect(table!.headerRowCount).toBe(1)
    expect(table!.columns).toEqual(['Item', 'Amount'])
    expect(table!.styleName).toBe('TableStyleMedium2')
    expect(table!.showRowStripes).toBe(true)
    expect(table!.headerFill).toBe('#4472C4')
    expect(table!.headerFontColor).toBe('#FFFFFF')
    expect(table!.stripeFill).toBe('#DAE3F3')

    // Importing the table must NOT create an undo entry — the workbook
    // starts clean.
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    // DELETE path (file-native): put the active cell inside the table and
    // ask Delete Table — refused with the desktop's exact message.
    await selectRange(page, 'B2')
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Insert' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: /Delete Table/i }).click()
    await expect(
      page.getByText(
        'Table "SalesTable" does not exist or was not created this session — ' +
          'tables already in the file cannot be deleted yet.',
      ),
    ).toBeVisible()
    // A refusal journals nothing.
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    expect(pageErrors).toEqual([])
  })

  test('2: create → typed wire → table part + wiring → reopen → file-native refusal', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelTableCreateFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-table-create.xlsx')

    // Select A1:B4 through the Name Box, then Insert → Table.
    await selectRange(page, 'A1:B4')
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Insert' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: /create a table/i }).click()
    await expect(page.getByText('Table created — save with ⌘S.')).toBeVisible()
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the plan must carry the typed tableAdditions family.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        tableAdditions?: Array<{
          sheetName: string
          area: { startRow: number; startColumn: number; endRow: number; endColumn: number }
          name: string
          columnNames: string[]
          style?: string
          bandedRows: boolean
        }>
      }
    }
    expect(saveBody.savePlan.tableAdditions, 'typed tableAdditions on the wire').toEqual([
      {
        sheetName: 'Ledger',
        area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
        name: 'Table1',
        columnNames: ['Item', 'Amount'],
        style: 'TableStyleMedium2',
        bandedRows: true,
      },
    ])

    // Saved XML: the table part + its worksheet wiring all exist.
    const entries = await listZipEntries(saved)
    expect(entries).toContain('xl/tables/table1.xml')
    const tablePart = await readZipEntry(saved, 'xl/tables/table1.xml')
    expect(tablePart).toContain('displayName="Table1"')
    expect(tablePart).toContain('ref="A1:B4"')
    expect(tablePart).toContain('<tableColumn id="1" name="Item"/>')
    expect(tablePart).toContain('<tableColumn id="2" name="Amount"/>')
    expect(tablePart).toContain('name="TableStyleMedium2"')
    expect(tablePart).toContain('showRowStripes="1"')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The create fixture ships no worksheet rels, so the writer's first
    // relationship is rId1.
    expect(sheet1).toContain('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    const rels = await readZipEntry(saved, 'xl/worksheets/_rels/sheet1.xml.rels')
    expect(rels).toContain('Id="rId1"')
    expect(rels).toContain('relationships/table" Target="../tables/table1.xml"')
    const contentTypes = await readZipEntry(saved, '[Content_Types].xml')
    expect(contentTypes).toContain('spreadsheetml.table+xml')
    await expect(page.getByText('Saved e2e-ribbon-table-create.xlsx')).toBeVisible({
      timeout: 15_000,
    })

    // Reopen: the saved table is FILE-NATIVE (metadata round-trips) and
    // Delete Table now refuses.
    writeFileSync('/tmp/e2e-ribbon-table-create-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-table-create-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-table-create-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = ((await reopenResponse.json()) as { snapshot: TableSnapshotView }).snapshot
    const reopenedTable = reopened.sheets[0].tables?.[0]
    expect(reopenedTable?.name).toBe('Table1')
    expect(reopenedTable?.area).toEqual({ startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 })
    expect(reopenedTable?.columns).toEqual(['Item', 'Amount'])
    await page.waitForTimeout(1000)
    await selectRange(page, 'A2')
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Insert' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: /Delete Table/i }).click()
    await expect(
      page.getByText(
        'Table "Table1" does not exist or was not created this session — ' +
          'tables already in the file cannot be deleted yet.',
      ),
    ).toBeVisible()

    expect(pageErrors).toEqual([])
  })

  test('3: delete a session table → save → nothing persists (no-op bytes)', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelTableCreateFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-table-delete.xlsx')

    // Create the table, then delete it (convert-to-range).
    await selectRange(page, 'A1:B4')
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Insert' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: /create a table/i }).click()
    await expect(page.getByText('Table created — save with ⌘S.')).toBeVisible()
    await selectRange(page, 'B2')
    await page.getByRole('button', { name: /Delete Table/i }).click()
    await expect(
      page.getByText('Table "Table1" removed — the cells stay as they are.'),
    ).toBeVisible()

    // Save: the plan carries NO tableAdditions (the journal entry was
    // spliced — an unsaved table never reaches the file).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { tableAdditions?: unknown }
    }
    expect(saveBody.savePlan.tableAdditions, 'journal splice → no tableAdditions').toBeUndefined()

    // Saved bytes: NO table part, NO worksheet wiring — the save added
    // nothing table-related. (workbook.xml legitimately gains the engine's
    // <calcPr fullCalcOnLoad="1"/> on every save — unrelated to tables, so
    // the byte comparison covers the entries a table write could touch.)
    const savedEntries = await listZipEntries(saved)
    expect(savedEntries.some((entry) => entry.startsWith('xl/tables/'))).toBe(false)
    expect(savedEntries.sort()).toEqual([...(await listZipEntries(fixture))].sort())
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).not.toContain('<tableParts')
    for (const entry of [
      'xl/worksheets/sheet1.xml',
      'xl/sharedStrings.xml',
      'xl/styles.xml',
      '[Content_Types].xml',
    ]) {
      expect(await readZipEntry(saved, entry), `${entry} preserved byte-for-byte`).toBe(
        await readZipEntry(fixture, entry),
      )
    }

    expect(pageErrors).toEqual([])
  })

  test('4: row insert inside a table → save → ref grows → reopen; header-row delete fails closed', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelTableFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-table-shift.xlsx')

    // Insert one row INSIDE the table (row index 2, between Tax and Tip)
    // through the real facade — the structural shift must grow the table.
    await insertRows(page, 'Ledger', 2, 1)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    await saveReq
    await expect(page.getByText('Saved e2e-ribbon-table-shift.xlsx')).toBeVisible({
      timeout: 15_000,
    })

    // The table part's ref and autoFilter both grew by one row.
    const tablePart = await readZipEntry(saved, 'xl/tables/table1.xml')
    expect(tablePart).toContain('ref="A1:B5"')
    expect(tablePart).toContain('<autoFilter ref="A1:B5"/>')

    // Reopen: the shifted area round-trips.
    writeFileSync('/tmp/e2e-ribbon-table-shift-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-table-shift-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-table-shift-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = ((await reopenResponse.json()) as { snapshot: TableSnapshotView }).snapshot
    expect(reopened.sheets[0].tables?.[0]?.area).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 4,
      endColumn: 1,
    })
    await page.waitForTimeout(1000)

    // FAIL-CLOSED: deleting the table's HEADER row cannot be saved — the
    // gateway refuses the anatomy change and the browser surfaces it.
    await deleteRows(page, 'Ledger', 0, 1)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect(
      page.getByText(/Save failed: .*Deleting the header row of table "SalesTable"/),
    ).toBeVisible({ timeout: 30_000 })

    expect(pageErrors).toEqual([])
  })

  test('5: a table-owned filter refuses Data → Filter (table part untouched)', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelTableFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-table-filter.xlsx')

    // Data → Filter: the sheet's filter belongs to the table (the
    // worksheet carries no <autoFilter>), so the command is refused with
    // the desktop's exact message.
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: /AutoFilter/i }).click()
    await expect(
      page.getByText(
        "This sheet's filter belongs to an Excel table — editing it cannot be saved yet.",
      ),
    ).toBeVisible()

    // The refusal cancelled the command: nothing journaled, nothing saved.
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    expect(pageErrors).toEqual([])
  })

  test('6: an unrelated edit saves with the table parts preserved byte-for-byte', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelTableFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-table-noop.xlsx')

    // Edit a cell OUTSIDE the table (A5 — row 5 is not part of the
    // SalesTable range A1:B4) through the real grid.
    await selectRange(page, 'A5')
    await page.keyboard.type('Total2')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(500)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the plan carries the cell edit but NO tableAdditions family.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { tableAdditions?: unknown; edits: unknown[] }
    }
    expect(saveBody.savePlan.tableAdditions, 'no table family on a no-table save').toBeUndefined()
    expect(saveBody.savePlan.edits.length).toBeGreaterThan(0)

    // The table part and its wiring survive the save byte-for-byte.
    expect(await readZipEntry(saved, 'xl/tables/table1.xml')).toBe(
      await readZipEntry(fixture, 'xl/tables/table1.xml'),
    )
    expect(await readZipEntry(saved, 'xl/worksheets/_rels/sheet1.xml.rels')).toBe(
      await readZipEntry(fixture, 'xl/worksheets/_rels/sheet1.xml.rels'),
    )
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>')

    expect(pageErrors).toEqual([])
  })
})
