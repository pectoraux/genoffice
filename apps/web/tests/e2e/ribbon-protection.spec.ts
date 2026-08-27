/**
 * REAL browser E2E — Review → Protection (EXCEL-020).
 *
 * Proves the protection persistence chain end-to-end through the REAL
 * HTTP boundary:
 *
 *   open → readBasicWorkbook parses <sheetProtection> (per worksheet) +
 *   <workbookProtection> (workbook.xml) → WorksheetState.sheetProtection /
 *   WorkbookSnapshot.workbookProtection → browser seeds its file-state
 *   refs + ribbon echo → user toggles Protect Sheet / Protect Workbook /
 *   Lock-Unlock Cell (journal-only semantics, desktop parity) → save
 *   emits the typed sheetProtections + workbookProtectionState families +
 *   protectionLocked style edits → /api/office/workbooks/save →
 *   routeOffice strict validation → applyCellEditsToXlsx args 10/trailing
 *   → applySheetProtection / applyWorkbookProtection / buildProtection →
 *   XLSX elements → reopen → same protection state.
 *
 * No browser-side OOXML. The browser only ever exchanges typed
 * protection state + journal decisions.
 *
 * Fixtures (buildProtectionLedgerFixture variants): the notes ledger
 * shape with (a) no protection elements, (b) sheet + workbook structure
 * protection, (c) password-bearing elements on both levels.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import {
  buildExcelProtectionFixture,
  buildExcelProtectedFixture,
  buildExcelPasswordFixture,
  readZipEntry,
} from './fixtures'

/** Snapshot view of the open response's protection state. */
interface SnapshotView {
  sheets: Array<{
    name: string
    sheetProtection?: { protected: boolean; hasPassword: boolean }
  }>
  workbookProtection?: { lockStructure: boolean; hasPassword: boolean }
}

test.describe('Review tab — Protection persists through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1: opening a protected workbook surfaces the state (snapshot + ribbon echo)', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelProtectedFixture()
    writeFileSync('/tmp/e2e-ribbon-protection-open.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-open.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-open.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // READ path: the snapshot carried the parsed protection state.
    const snapshot = (await openResponse.json()).snapshot as SnapshotView
    expect(snapshot.sheets[0].sheetProtection).toEqual({ protected: true, hasPassword: false })
    expect(snapshot.workbookProtection).toEqual({ lockStructure: true, hasPassword: false })

    // ECHO path: the ribbon reflects the file's protection state (the
    // label flips to Unprotect — desktop sheetProtectionEcho parity).
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await expect(page.getByRole('button', { name: 'Unprotect Sheet', exact: true })).toBeVisible()
    await expect(
      page.getByRole('button', { name: 'Unprotect Workbook', exact: true }),
    ).toBeVisible()

    // Loading protection state must NOT create an undo entry — the
    // workbook starts clean.
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    expect(pageErrors).toEqual([])
  })

  test('2: protect sheet + workbook from scratch → typed wire → XML → reopen', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelProtectionFixture()
    writeFileSync('/tmp/e2e-ribbon-protection-protect.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-protect.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-protect.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await page.waitForTimeout(200)

    // Toggle sheet protection — desktop-parity status message.
    await page.getByRole('button', { name: 'Protect Sheet', exact: true }).click()
    await expect(page.getByText(/Sheet protection will be written on save/)).toBeVisible()
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    // The label flipped with the journal echo.
    await expect(page.getByRole('button', { name: 'Unprotect Sheet', exact: true })).toBeVisible()

    // Toggle workbook structure protection.
    await page.getByRole('button', { name: 'Protect Workbook', exact: true }).click()
    await expect(
      page.getByText(/Workbook structure protection will be written on save/),
    ).toBeVisible()
    await expect(
      page.getByRole('button', { name: 'Unprotect Workbook', exact: true }),
    ).toBeVisible()

    // Save: the plan must carry BOTH typed protection families.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        sheetProtections?: Array<{ sheetName: string; protected: boolean }>
        workbookProtectionState?: { lockStructure: boolean }
      }
    }
    expect(saveBody.savePlan.sheetProtections, 'typed sheetProtections on the wire').toEqual([
      { sheetName: 'Ledger', protected: true },
    ])
    expect(
      saveBody.savePlan.workbookProtectionState,
      'typed workbookProtectionState on the wire',
    ).toEqual({ lockStructure: true })

    // Saved XLSX: the worksheet gained the protection element with
    // Excel's defaults; workbook.xml gained the structure lock.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>')
    const workbook = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbook).toContain('lockStructure="1"')

    // Reopen: the snapshot carries the protection state and the ribbon
    // echoes it.
    writeFileSync('/tmp/e2e-ribbon-protection-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot as SnapshotView
    expect(reopened.sheets[0].sheetProtection).toEqual({ protected: true, hasPassword: false })
    expect(reopened.workbookProtection).toEqual({ lockStructure: true, hasPassword: false })

    await expect(page.getByRole('button', { name: 'Unprotect Sheet', exact: true })).toBeVisible()
    await expect(
      page.getByRole('button', { name: 'Unprotect Workbook', exact: true }),
    ).toBeVisible()

    // The saved state is in the file — the reopened workbook is clean.
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 5000 })

    expect(pageErrors).toEqual([])
  })

  test('3: unprotect round-trip — element removed, reopen is clean', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelProtectedFixture()
    writeFileSync('/tmp/e2e-ribbon-protection-unprotect.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-unprotect.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-unprotect.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await page.waitForTimeout(200)

    // Unprotect the sheet AND the workbook structure.
    await page.getByRole('button', { name: 'Unprotect Sheet', exact: true }).click()
    await expect(page.getByText(/Sheet protection will be removed on save/)).toBeVisible()
    await page.getByRole('button', { name: 'Unprotect Workbook', exact: true }).click()
    await expect(
      page.getByText(/Workbook structure protection will be removed on save/),
    ).toBeVisible()

    // Toggle-back semantics (desktop recordSheetProtection): toggling
    // back to the file's ORIGINAL state DROPS the journal entry. Undo the
    // workbook unprotect — the save must then carry NO workbook state.
    await page.getByRole('button', { name: 'Protect Workbook', exact: true }).click()
    await page.waitForTimeout(300)

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        sheetProtections?: Array<{ sheetName: string; protected: boolean }>
        workbookProtectionState?: { lockStructure: boolean } | null
      }
    }
    expect(saveBody.savePlan.sheetProtections).toEqual([{ sheetName: 'Ledger', protected: false }])
    expect(saveBody.savePlan.workbookProtectionState).toBeUndefined()

    // Saved XLSX: the sheet element is GONE; the workbook structure lock
    // SURVIVED (toggle-back dropped that journal entry).
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).not.toContain('<sheetProtection')
    const workbook = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbook).toContain('lockStructure="1"')

    // Reopen: sheet unprotected, workbook still locked.
    writeFileSync('/tmp/e2e-ribbon-protection-unprotected.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-unprotected.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-unprotected.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot as SnapshotView
    expect(reopened.sheets[0].sheetProtection).toBeUndefined()
    expect(reopened.workbookProtection).toEqual({ lockStructure: true, hasPassword: false })
    await expect(page.getByRole('button', { name: 'Protect Sheet', exact: true })).toBeVisible()
    await expect(
      page.getByRole('button', { name: 'Unprotect Workbook', exact: true }),
    ).toBeVisible()

    expect(pageErrors).toEqual([])
  })

  test('4: editable vs locked cell behavior — unlock a cell, protect the sheet, reopen', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelProtectionFixture()
    writeFileSync('/tmp/e2e-ribbon-protection-locked.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-locked.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-locked.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Select A2 via the Name Box — the proven coordinate-independent path.
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A2')
    await box.press('Enter')
    await page.waitForTimeout(400)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await page.waitForTimeout(200)

    // Unlock the selected cell (canonical protectionLocked style edit).
    await page.getByRole('button', { name: 'Unlock Cell', exact: true }).click()
    await expect(page.getByText(/Cells unlocked/)).toBeVisible()

    // Protect the sheet — together with the unlocked cell this IS Excel's
    // editable-vs-locked semantics in the saved file: A2 stays editable,
    // everything else is read-only for readers that enforce protection.
    await page.getByRole('button', { name: 'Protect Sheet', exact: true }).click()
    await expect(page.getByText(/Sheet protection will be written on save/)).toBeVisible()

    // Save: the plan carries BOTH the style edit and the protection family.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits?: Array<{
          sheetName: string
          row: number
          column: number
          writeValue: boolean
          style?: { protectionLocked?: boolean }
        }>
        sheetProtections?: Array<{ sheetName: string; protected: boolean }>
      }
    }
    const lockedEdit = saveBody.savePlan.edits?.find(
      (e) => e.sheetName === 'Ledger' && e.row === 1 && e.column === 0,
    )
    expect(lockedEdit, 'A2 carries a journaled style edit').toBeDefined()
    expect(lockedEdit!.writeValue).toBe(false)
    expect(lockedEdit!.style?.protectionLocked).toBe(false)
    expect(saveBody.savePlan.sheetProtections).toEqual([{ sheetName: 'Ledger', protected: true }])

    // Saved XLSX: A2's xf carries <protection locked="0"/> and the sheet
    // is protected — the file's authorization semantics are complete.
    const styles = await readZipEntry(saved, 'xl/styles.xml')
    expect(styles).toContain('<protection locked="0"/>')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>')

    // Reopen: the protection state survives; a no-op save preserves the
    // unlocked cell's XML.
    writeFileSync('/tmp/e2e-ribbon-protection-locked-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles(
      'input[type="file"]',
      '/tmp/e2e-ribbon-protection-locked-reopened.xlsx',
    )
    await expect(page.getByText('Opened e2e-ribbon-protection-locked-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot as SnapshotView
    expect(reopened.sheets[0].sheetProtection).toEqual({ protected: true, hasPassword: false })

    // No-op save: the journal is empty (protection is in the file), so
    // the plan carries NO protection family and the styles + protection
    // XML survive untouched.
    const noopSaveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (r: number, c: number) => { setValueForCell: (v: unknown) => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI
        ?.getActiveWorkbook?.()
        ?.getActiveSheet?.()
        ?.getRange?.(0, 5)
        ?.setValueForCell?.('noop-edit')
    })
    await page.waitForTimeout(500)
    const noopSaved = await clickSaveAndCaptureDownload(page, 'Save')
    const noopReq = await noopSaveReq
    const noopBody = JSON.parse(noopReq.postData() ?? '{}') as {
      savePlan: { sheetProtections?: unknown[] }
    }
    expect(noopBody.savePlan.sheetProtections).toBeUndefined()
    const noopStyles = await readZipEntry(noopSaved, 'xl/styles.xml')
    expect(noopStyles).toContain('<protection locked="0"/>')
    const noopSheet = await readZipEntry(noopSaved, 'xl/worksheets/sheet1.xml')
    expect(noopSheet).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>')

    expect(pageErrors).toEqual([])
  })

  test('5: negative authorization — password-protected sheet/structure refuse to unprotect', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelPasswordFixture()
    writeFileSync('/tmp/e2e-ribbon-protection-password.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-password.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-password.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const openResponse = await openResponsePromise
    const snapshot = (await openResponse.json()).snapshot as SnapshotView
    // The reader surfaces BOTH password flags — the browser must know.
    expect(snapshot.sheets[0].sheetProtection).toEqual({ protected: true, hasPassword: true })
    expect(snapshot.workbookProtection).toEqual({ lockStructure: true, hasPassword: true })

    await page.waitForTimeout(1500)
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await page.waitForTimeout(200)

    // Attempting to unprotect the password-protected sheet is REFUSED up
    // front (desktop parity) — no journal entry, no dirty mark.
    await page.getByRole('button', { name: 'Unprotect Sheet', exact: true }).click()
    await expect(
      page.getByText(/protected with a password — removing its protection is not supported/),
    ).toBeVisible()
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    // Same for the password-protected workbook structure.
    await page.getByRole('button', { name: 'Unprotect Workbook', exact: true }).click()
    await expect(
      page.getByText(/workbook structure is password-protected — it cannot be changed here/),
    ).toBeVisible()
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    // The refuse-left-journal-clean invariant: make an unrelated edit,
    // save, and verify the plan carries NO protection family (the
    // password elements must ride the file untouched).
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (r: number, c: number) => { setValueForCell: (v: unknown) => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI
        ?.getActiveWorkbook?.()
        ?.getActiveSheet?.()
        ?.getRange?.(0, 5)
        ?.setValueForCell?.('password-noop')
    })
    await page.waitForTimeout(500)
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { sheetProtections?: unknown[]; workbookProtectionState?: unknown }
    }
    expect(saveBody.savePlan.sheetProtections).toBeUndefined()
    expect(saveBody.savePlan.workbookProtectionState).toBeUndefined()

    // The password-bearing elements survive the save verbatim.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<sheetProtection sheet="1" password="83AF"/>')
    const workbook = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbook).toContain('lockStructure="1" workbookPassword="83AF"')

    expect(pageErrors).toEqual([])
  })

  test('6: lock-through-ribbon — a locked cell retains its locked state (architect directive: lock/unlock BOTH through the actual ribbon)', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Phase 1 — unlock A2 through the ribbon, protect the sheet, save.
    // (Verified in depth by test 4; here we establish the unlocked state
    // we will then RE-LOCK through the ribbon.)
    const fixture = await buildExcelProtectionFixture()
    writeFileSync('/tmp/e2e-ribbon-protection-relock.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-relock.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-relock.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A2')
    await box.press('Enter')
    await page.waitForTimeout(400)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await page.waitForTimeout(200)

    await page.getByRole('button', { name: 'Unlock Cell', exact: true }).click()
    await expect(page.getByText(/Cells unlocked/)).toBeVisible()
    await page.getByRole('button', { name: 'Protect Sheet', exact: true }).click()
    await expect(page.getByText(/Sheet protection will be written on save/)).toBeVisible()

    const phase1SaveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const phase1Saved = await clickSaveAndCaptureDownload(page, 'Save')
    await phase1SaveReq
    // Unlocked marker present in the phase-1 file (A2 editable, rest locked).
    const phase1Styles = await readZipEntry(phase1Saved, 'xl/styles.xml')
    expect(phase1Styles).toContain('<protection locked="0"/>')
    const phase1Sheet = await readZipEntry(phase1Saved, 'xl/worksheets/sheet1.xml')
    expect(phase1Sheet).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>')

    // Phase 2 — reopen the unlocked+protected file, then LOCK A2 back
    // through the actual ribbon (the button the architect's directive
    // demands: "lock/unlock through the actual ribbon" — BOTH directions).
    writeFileSync('/tmp/e2e-ribbon-protection-relock-opened.xlsx', phase1Saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-relock-opened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-relock-opened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot as SnapshotView
    expect(reopened.sheets[0].sheetProtection).toEqual({ protected: true, hasPassword: false })
    await page.waitForTimeout(1500)

    // Select A2 again and LOCK it through the ribbon.
    await box.click()
    await box.fill('A2')
    await box.press('Enter')
    await page.waitForTimeout(400)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Review' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Lock Cell', exact: true }).click()
    await expect(page.getByText(/Cells locked/)).toBeVisible()

    // The save plan carries the canonical protectionLocked=true style edit.
    const lockSaveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const lockSaved = await clickSaveAndCaptureDownload(page, 'Save')
    const lockReq = await lockSaveReq
    const lockBody = JSON.parse(lockReq.postData() ?? '{}') as {
      savePlan: {
        edits?: Array<{
          sheetName: string
          row: number
          column: number
          writeValue: boolean
          style?: { protectionLocked?: boolean }
        }>
        sheetProtections?: unknown[]
      }
    }
    const lockEdit = lockBody.savePlan.edits?.find(
      (e) => e.sheetName === 'Ledger' && e.row === 1 && e.column === 0,
    )
    expect(lockEdit, 'A2 carries a journaled style edit').toBeDefined()
    expect(lockEdit!.writeValue).toBe(false)
    expect(lockEdit!.style?.protectionLocked).toBe(true)
    // The sheet protection itself was NOT re-journaled (it is already in
    // the file — the journal carries only the re-lock style edit).
    expect(lockBody.savePlan.sheetProtections).toBeUndefined()

    // Saved XLSX: A2 references a LOCKED xf — the OOXML default (absence
    // of <protection locked="0"/> on the cell's OWN xf = locked). The
    // previously-unlocked xf record may legally remain in cellXfs as an
    // orphan (append-mostly styles.xml — Excel itself leaves orphans);
    // what matters is the xf A2 actually references.
    const lockStyles = await readZipEntry(lockSaved, 'xl/styles.xml')
    const lockSheetForXf = await readZipEntry(lockSaved, 'xl/worksheets/sheet1.xml')
    const a2Match = /<c r="A2"[^>]*\bs="(\d+)"/.exec(lockSheetForXf)
    expect(a2Match, 'A2 carries an explicit style index after the re-lock edit').not.toBeNull()
    const xfsBlock = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(lockStyles)?.[1] ?? ''
    const xfRecords = xfsBlock
      .split(/<xf\b/)
      .slice(1)
      .map((body) => `<xf${body}`)
    const a2Xf = xfRecords[Number(a2Match?.[1])] ?? '(missing)'
    expect(
      a2Xf,
      'A2 references a LOCKED xf — no locked="0" on the referenced record (locked state retained)',
    ).not.toContain('locked="0"')
    const lockSheet = await readZipEntry(lockSaved, 'xl/worksheets/sheet1.xml')
    expect(lockSheet, 'sheet protection survives the re-lock save').toContain(
      '<sheetProtection sheet="1" objects="1" scenarios="1"/>',
    )

    // Reopen: the sheet is still protected, A2 is a locked cell under it —
    // "a locked cell retains its locked state" while the sheet is protected.
    writeFileSync('/tmp/e2e-ribbon-protection-relocked.xlsx', lockSaved)
    const finalReopenPromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-protection-relocked.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-protection-relocked.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const finalReopen = await finalReopenPromise
    const finalSnapshot = (await finalReopen.json()).snapshot as SnapshotView
    expect(finalSnapshot.sheets[0].sheetProtection).toEqual({ protected: true, hasPassword: false })
    // Re-check the referenced-xf chain on the final artifact too.
    const finalA2 = /<c r="A2"[^>]*\bs="(\d+)"/.exec(
      await readZipEntry(lockSaved, 'xl/worksheets/sheet1.xml'),
    )
    const finalXfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(lockStyles)?.[1] ?? ''
    const finalRecords = finalXfs
      .split(/<xf\b/)
      .slice(1)
      .map((body) => `<xf${body}`)
    expect(finalRecords[Number(finalA2?.[1])] ?? '(missing)').not.toContain('locked="0"')

    expect(pageErrors).toEqual([])
  })
})
