/**
 * REAL browser E2E — View tab (Phase 4 Inc. 3; EXCEL-026 regression).
 *
 * Proves Gridlines and Zoom work in-session on the blank workbook, and
 * Freeze Panes PERSISTS through save/reopen via the canonical pipeline:
 *
 *   View → Freeze Panes
 *   → FWorksheet.setFreeze({ startRow, startColumn, xSplit, ySplit })
 *   → sheet.mutation.set-frozen
 *   → ExcelEditor's journal subscription captures the freeze state into
 *     the per-sheet pageSetupRef (the EXCEL-026 page-setup journal)
 *   → on Save, emit as a BrowserSheetPageSetupState
 *   → savePlan.pageSetupStates → /api/office/workbooks/save → routeOffice
 *     → applyCellEditsToXlsx(buf, edits, structuralOps, [], undefined, [],
 *       [], [], [], [], null, pageSetupStates)
 *   → applyPageSetupState writes the <pane> element into the worksheet XML
 *   → reopen → readBasicWorkbook parses <pane> → WorksheetState.freeze
 *   → loadSnapshot seeds the freeze into Univer's freeze config
 *   → freeze visible on reopen
 *
 * The full EXCEL-026 view-persistence surface (gridline persistence,
 * formula view, freeze CLEAR, page layout, byte preservation) lives in
 * ribbon-view-persistence.spec.ts.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelFixture, readZipEntry } from './fixtures'

test.describe('View tab — Gridlines/Zoom in-session, Freeze Panes persists', () => {
  test('freeze panes persists through save/reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFixture()
    writeFileSync('/tmp/e2e-ribbon-freeze.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-freeze.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-freeze.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Navigate the active cell to C3 via the Name Box, then apply Freeze
    // Panes — should freeze rows 1-2 and columns A-B. (B3 is avoided because
    // the fixture has A3:B3 merged — navigating to B3 selects the merge
    // range A3:B3, and the freeze would anchor at column A instead of B.)
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('C3')
    await box.press('Enter')
    await page.waitForTimeout(500)

    // Switch to View tab and click Freeze Panes.
    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Freeze Panes' }).click()
    await page.waitForTimeout(500)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Verify the freeze was applied in the live Univer model.
    const freezeInUniver = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getFreeze: () => {
                    startRow: number
                    startColumn: number
                    xSplit: number
                    ySplit: number
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getFreeze?.() ?? null
    })
    expect(freezeInUniver, 'Univer model carries the freeze').not.toBeNull()
    expect(freezeInUniver!.startRow, '2 frozen rows (startRow = 2)').toBe(2)
    expect(freezeInUniver!.startColumn, '2 frozen columns (startColumn = 2)').toBe(2)
    expect(freezeInUniver!.ySplit, 'ySplit = 2').toBe(2)
    expect(freezeInUniver!.xSplit, 'xSplit = 2').toBe(2)

    // Save and capture. The save plan should include pageSetupStates.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        pageSetupStates?: Array<{
          sheetName: string
          frozenRows?: number
          frozenColumns?: number
        }>
      }
    }
    expect(
      saveBody.savePlan.pageSetupStates,
      'save plan includes pageSetupStates for freeze',
    ).toBeDefined()
    const freezeState = saveBody.savePlan.pageSetupStates!.find((s) => s.sheetName === 'Data')
    expect(freezeState, 'freeze state for Data sheet').toBeDefined()
    expect(freezeState!.frozenRows, '2 frozen rows in save plan').toBe(2)
    expect(freezeState!.frozenColumns, '2 frozen columns in save plan').toBe(2)

    // Saved XLSX: the worksheet carries a <pane> with ySplit="2" xSplit="2"
    // state="frozen".
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'pane element written').toContain('<pane')
    expect(sheet1, 'ySplit=2 written').toContain('ySplit="2"')
    expect(sheet1, 'xSplit=2 written').toContain('xSplit="2"')
    expect(sheet1, 'state=frozen written').toContain('state="frozen"')

    // Reopen and verify the snapshot carries the freeze.
    writeFileSync('/tmp/e2e-ribbon-freeze-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-freeze-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-freeze-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      freeze?: { frozenRows: number; frozenColumns: number }
    }>
    const dataSheet = reopened.find((s) => s.name === 'Data')
    expect(dataSheet, 'Data sheet present after reopen').toBeDefined()
    expect(dataSheet!.freeze, 'freeze survived reopen').toBeDefined()
    expect(dataSheet!.freeze!.frozenRows, '2 frozen rows after reopen').toBe(2)
    expect(dataSheet!.freeze!.frozenColumns, '2 frozen columns after reopen').toBe(2)

    // The reopened freeze should also be reflected in the live Univer model.
    await page.waitForTimeout(1500)
    const reopenedFreeze = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getFreeze: () => {
                    startRow: number
                    startColumn: number
                    xSplit: number
                    ySplit: number
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getFreeze?.() ?? null
    })
    expect(reopenedFreeze, 'reopened Univer model carries the freeze').not.toBeNull()
    expect(reopenedFreeze!.startRow, '2 frozen rows in reopened model').toBe(2)
    expect(reopenedFreeze!.startColumn, '2 frozen columns in reopened model').toBe(2)

    expect(pageErrors).toEqual([])
  })

  test('gridlines toggle changes visibility in-session', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)

    const gridlinesBtn = page.getByRole('button', { name: 'Gridlines' })
    // Initially active (gridlines visible by default).
    await expect(gridlinesBtn).toHaveAttribute('aria-pressed', 'true')
    // Toggle off.
    await gridlinesBtn.click()
    await page.waitForTimeout(300)
    await expect(gridlinesBtn).not.toHaveAttribute('aria-pressed', 'true')
    // Toggle back on.
    await gridlinesBtn.click()
    await page.waitForTimeout(300)
    await expect(gridlinesBtn).toHaveAttribute('aria-pressed', 'true')
  })

  test('zoom in/out changes the zoom percent in-session', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)

    const initialPct = await page.locator('.excel-zoom-value').first().textContent()
    expect(initialPct).toMatch(/\d+%/)

    // Zoom in (scope to the ribbon — the status bar also has a zoom-in button).
    await page
      .locator('[data-testid="excel-ribbon"]')
      .getByRole('button', { name: 'Zoom in' })
      .click()
    await page.waitForTimeout(300)
    const afterInPct = await page.locator('.excel-zoom-value').first().textContent()
    const afterInNum = parseInt((afterInPct ?? '').replace('%', ''), 10)
    const initialNum = parseInt((initialPct ?? '').replace('%', ''), 10)
    expect(afterInNum, 'zoom in increased percent').toBeGreaterThan(initialNum)

    // Zoom out.
    await page
      .locator('[data-testid="excel-ribbon"]')
      .getByRole('button', { name: 'Zoom out' })
      .click()
    await page.waitForTimeout(300)
    const afterOutPct = await page.locator('.excel-zoom-value').first().textContent()
    const afterOutNum = parseInt((afterOutPct ?? '').replace('%', ''), 10)
    expect(afterOutNum, 'zoom out decreased percent').toBeLessThan(afterInNum)
  })
})
