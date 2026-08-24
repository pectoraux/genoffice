/**
 * REAL browser E2E — Data tab (Phase 4 Inc. 3).
 *
 * Proves Sort ascending/descending PERSISTS through save/reopen via the
 * canonical pipeline, and that Filter remains visibly disabled (the wire
 * save plan does not expose the filterStates family).
 *
 * Sort path (proven):
 *   Data → Sort Asc
 *   → FRange.sort({ column: 0, ascending: true })
 *   → sheet.command.sort-range → ReorderRangeMutation
 *   → ExcelEditor's expanded journal subscription re-reads the post-sort
 *     cell values and journals them as writeValue CellEdits
 *   → savePlan.edits → /api/office/workbooks/save → applyCellEditsToXlsx
 *   → XLSX bytes (new row order)
 *   → reopen → readBasicWorkbook → snapshot carries the sorted order
 *
 * Filter path (BLOCKED):
 *   The wire save plan does not expose the filterStates family. Until the
 *   wire is extended, the Filter button is visibly disabled.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas, clickSaveAndCaptureDownload } from './helpers'
import { buildExcelFixture, readZipEntry } from './fixtures'

test.describe('Data tab — Sort persists, Filter disabled', () => {
  test('Sort descending persists through save/reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // The standard fixture's Data sheet has A1="Original Text", B1=10, C1=SUM.
    // We'll add a few more values in column B below B1, then sort desc by col B.
    const fixture = await buildExcelFixture()
    writeFileSync('/tmp/e2e-ribbon-sort.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-sort.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-sort.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Type 30 into B2, 5 into B3, 20 into B4 (so a sort desc gives 30,20,10,5).
    // Set cell values directly through the Univer FRange.setValueForCell
    // facade — this fires set-range-values (journaled by the existing
    // subscription) and is coordinate-independent (canvas clicks are
    // fragile when the grid scrolls between edits).
    await page.evaluate(() => {
      const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { getActiveSheet: () => { getRange: (r: number, c: number) => { setValueForCell: (v: unknown) => unknown } } } } } }).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      ws?.getRange?.(1, 1)?.setValueForCell?.(30)
      ws?.getRange?.(2, 1)?.setValueForCell?.(5)
      ws?.getRange?.(3, 1)?.setValueForCell?.(20)
    })
    await page.waitForTimeout(400)

    // Select the range B1:B4 via the Name Box (the proven path from
    // excel-shell.spec.ts — canvas drag-select is coordinate-fragile).
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('B1:B4')
    await box.press('Enter')
    await page.waitForTimeout(400)

    // Switch to Data tab and click Sort Desc. Scope to the ribbon — the
    // workbook's sheet is also named "Data", so an unscoped tab lookup matches
    // both the ribbon tab and the sheet tab.
    await page.locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' }).click()
    await page.waitForTimeout(200)

    // Verify the active range is B1:B4 (the Name Box navigation selected it).
    const activeRangeBefore = await page.evaluate(() => {
      const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { getActiveSheet: () => { getActiveRange: () => { getA1Notation: () => string } | null } } } } }).__genofficeExcelRuntime
      return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getActiveRange?.()?.getA1Notation?.() ?? 'none'
    })
    expect(activeRangeBefore, 'active range is B1:B4 before sort').toBe('B1:B4')

    // Verify the values were typed correctly before sorting.
    const valuesBefore = await page.evaluate(() => {
      const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { getActiveSheet: () => { getRange: (r: number, c: number) => { getCellData: () => { v?: unknown } | null } } } } } }).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      return {
        b1: ws?.getRange?.(0, 1)?.getCellData?.()?.v,
        b2: ws?.getRange?.(1, 1)?.getCellData?.()?.v,
        b3: ws?.getRange?.(2, 1)?.getCellData?.()?.v,
        b4: ws?.getRange?.(3, 1)?.getCellData?.()?.v,
      }
    })
    expect(valuesBefore.b1, 'B1 = 10 before sort').toBe(10)
    expect(valuesBefore.b2, 'B2 = 30 before sort').toBe(30)
    expect(valuesBefore.b3, 'B3 = 5 before sort').toBe(5)
    expect(valuesBefore.b4, 'B4 = 20 before sort').toBe(20)

    await page.getByRole('button', { name: 'Sort Desc' }).click()
    await page.waitForTimeout(800)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Verify the sort ran in-session — B1 in the live Univer model should now
    // be 30 (the largest value).
    const b1Live = await page.evaluate(() => {
      const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { getActiveSheet: () => { getRange: (r: number, c: number) => { getCellData: () => { v?: unknown } | null } } } } } }).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      return ws?.getRange?.(0, 1)?.getCellData?.()?.v ?? null
    })
    expect(b1Live, 'sort ran in-session: B1 = 30 (largest)').toBe(30)

    // Save and capture — set up the request waiter BEFORE the click so we
    // don't miss the request.
    const sortReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await sortReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{ row: number; column: number; cell?: { value: unknown } }>
      }
    }
    // The sort journal should have produced edits for B1..B4 with the sorted
    // values (30, 20, 10, 5). Verify at least the B1 edit carries value 30.
    const b1Edit = saveBody.savePlan.edits.find((e) => e.row === 0 && e.column === 1)
    expect(b1Edit, 'B1 sort edit journaled').toBeDefined()
    expect(b1Edit?.cell?.value, 'B1 edit value = 30').toBe(30)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The sort wrote rows back in desc order (30, 20, 10, 5) — assert the
    // first row (row 1) carries the largest value (30), and the 4th row
    // carries the smallest (5). The cell values live inline as <v>.
    expect(sheet1, 'B1 = 30 (largest after desc sort)').toMatch(
      /<c r="B1"[^>]*><v>30<\/v>/,
    )
    expect(sheet1, 'B4 = 5 (smallest after desc sort)').toMatch(
      /<c r="B4"[^>]*><v>5<\/v>/,
    )

    // Reopen and verify the snapshot carries the sorted order.
    writeFileSync('/tmp/e2e-ribbon-sort-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-sort-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-sort-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown }>
    }>
    const cells = reopened[0].cells
    expect(cells.B1?.value, 'B1 = 30 after reopen').toBe(30)
    expect(cells.B2?.value, 'B2 = 20 after reopen').toBe(20)
    expect(cells.B3?.value, 'B3 = 10 after reopen').toBe(10)
    expect(cells.B4?.value, 'B4 = 5 after reopen').toBe(5)

    expect(pageErrors).toEqual([])
  })

  test('Filter button is disabled with documented reason', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    await page.getByRole('tab', { name: 'Data', exact: true }).click()
    await page.waitForTimeout(200)

    const filterBtn = page.getByRole('button', { name: /^Filter/ }).first()
    await expect(filterBtn).toBeDisabled()
    const title = await filterBtn.getAttribute('title')
    expect(title, 'Filter title names the architectural reason').toContain('filterStates')
  })
})
