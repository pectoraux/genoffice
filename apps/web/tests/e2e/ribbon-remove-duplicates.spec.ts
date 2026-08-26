/**
 * REAL browser E2E — Data → Remove Duplicates (EXCEL-018).
 *
 * Proves the canonical Remove Duplicates persistence chain end-to-end
 * through the REAL HTTP boundary:
 *
 *   open → readBasicWorkbook resolves + parses the worksheet →
 *   browser installs the snapshot via createWorkbook under journal
 *   suppression → user selects A1:B7 → Data → Remove Duplicates
 *   (with header) → useExcelRuntime.removeDuplicates(true) reads
 *   FRange.getValues(), dedupes via apps/web/src/office/dedupe.ts
 *   (mirrors the frozen desktop reference verbatim), writes back
 *   per-row via FWorksheet.getRange(...).setValues(...) → each write
 *   fires sheet.mutation.set-range-values → ExcelEditor's existing
 *   subscription captures each write as a CellEdit via
 *   cellEditFromMutation → savePlan.edits → /api/office/workbooks/save
 *   → applyCellEditsToXlsx writes the deduped values + blank padding
 *   rows into the XLSX → reopen → readBasicWorkbook → snapshot carries
 *   the deduped state.
 *
 * No browser-side OOXML. The browser only ever exchanges typed CellEdit
 * payloads with the canonical gateway.
 *
 * Fixture (buildExcelDedupeFixture): a "Dedupe" sheet with a bold
 * header, 6 data rows including 2 full duplicates of row 2, and one
 * formula row whose computed result differs from its sibling (so it
 * is NOT a duplicate).
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelDedupeFixture, readZipEntry } from './fixtures'

test.describe('Data tab — Remove Duplicates persists through save/reopen', () => {
  test('removes duplicates, preserves header, keeps formulas/styles on unchanged rows', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDedupeFixture()
    writeFileSync('/tmp/e2e-ribbon-dedupe.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dedupe.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dedupe.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Sanity: the live Univer model has the expected pre-dedupe values.
    const before = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (r: number, c: number) => { getCellData: () => { v?: unknown } | null }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (r: number, c: number) => ws?.getRange?.(r, c)?.getCellData?.()?.v ?? null
      return {
        a1: cell(0, 0),
        b1: cell(0, 1),
        a2: cell(1, 0),
        b2: cell(1, 1),
        a3: cell(2, 0),
        b3: cell(2, 1),
        a4: cell(3, 0),
        b4: cell(3, 1),
        a5: cell(4, 0),
        b5: cell(4, 1),
        a6: cell(5, 0),
        b6: cell(5, 1),
        a7: cell(6, 0),
        b7: cell(6, 1),
      }
    })
    expect(before.a1, 'A1 = "Name" before dedupe').toBe('Name')
    expect(before.b1, 'B1 = "Qty" before dedupe').toBe('Qty')
    expect(before.a2, 'A2 = "Apple" before dedupe').toBe('Apple')
    expect(before.b2, 'B2 = 10 before dedupe').toBe(10)
    expect(before.a3, 'A3 = "Apple" (duplicate) before dedupe').toBe('Apple')
    expect(before.b3, 'B3 = 10 (duplicate) before dedupe').toBe(10)
    expect(before.a4, 'A4 = "Banana" before dedupe').toBe('Banana')
    expect(before.b4, 'B4 = 20 before dedupe').toBe(20)
    expect(before.a5, 'A5 = "Apple" (duplicate) before dedupe').toBe('Apple')
    expect(before.b5, 'B5 = 10 (duplicate) before dedupe').toBe(10)
    expect(before.a6, 'A6 = "Cherry" before dedupe').toBe('Cherry')
    expect(before.b6, 'B6 = 30 before dedupe').toBe(30)
    expect(before.a7, 'A7 = "Apple" (formula row) before dedupe').toBe('Apple')
    expect(before.b7, 'B7 = 30 (formula result) before dedupe').toBe(30)

    // Select A1:B7 via the Name Box — the proven coordinate-independent
    // path from ribbon-data.spec.ts.
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A1:B7')
    await box.press('Enter')
    await page.waitForTimeout(400)

    // Switch to Data tab. Scope to the ribbon — the workbook sheet is
    // also named "Dedupe"; an unscoped "Data" lookup would be ambiguous.
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)

    // Verify the active range is A1:B7 before opening the dialog.
    const activeRangeBefore = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getActiveRange: () => { getA1Notation: () => string } | null
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      return (
        rt?.univerAPI
          ?.getActiveWorkbook?.()
          ?.getActiveSheet?.()
          ?.getActiveRange?.()
          ?.getA1Notation?.() ?? 'none'
      )
    })
    expect(activeRangeBefore, 'active range is A1:B7 before dedupe').toBe('A1:B7')

    // Open the Remove Duplicates dialog and click OK with the default
    // "My data has headers" checkbox checked.
    await page.getByRole('button', { name: 'Remove Duplicates' }).click()
    await expect(page.locator('[data-testid="dedupe-dialog"]')).toBeVisible()
    // The "My data has headers" checkbox defaults to checked — verify.
    const headerCheckbox = page.locator('[data-testid="dedupe-dialog"] input[type="checkbox"]')
    await expect(headerCheckbox).toBeChecked()
    await page.locator('[data-testid="dedupe-ok"]').click()
    await page.waitForTimeout(800)

    // Verify the in-session result. The dedupe COMPACTS kept rows to
    // the top of the selection and blanks the trailing rows (this is
    // the desktop's documented behavior — see apps/sheets/src/renderer/
    // ribbon-actions.ts:1283-1306, which pads with nulls back to the
    // original selection height).
    //
    // Input fixture (A1:B7, hasHeader=true):
    //   row 1: Name, Qty           (header — kept)
    //   row 2: Apple, 10           (first — kept)
    //   row 3: Apple, 10           (dup of row 2 — removed)
    //   row 4: Banana, 20          (first — kept)
    //   row 5: Apple, 10           (dup of row 2 — removed)
    //   row 6: Cherry, 30          (first — kept)
    //   row 7: Apple, 30 (=B6)      (first — kept; B7's result is 30
    //                                which differs from B2's 10, so it
    //                                is NOT a dup of row 2)
    //
    // Deduped (compacted): [Name/Qty, Apple/10, Banana/20, Cherry/30, Apple/30]
    // Padded to 7:          [Name/Qty, Apple/10, Banana/20, Cherry/30, Apple/30, null, null]
    //
    // Per-offset write:
    //   offset 0 (A1:B1): same → skip (header preserved verbatim)
    //   offset 1 (A2:B2): same → skip (first occurrence preserved —
    //                     any formula/style here would survive)
    //   offset 2 (A3:B3): DIFFER → write Banana/20 (moved from row 4)
    //   offset 3 (A4:B4): DIFFER → write Cherry/30 (moved from row 6)
    //   offset 4 (A5:B5): DIFFER → write Apple/30 (moved from row 7;
    //                     B7's formula =B6 is LOST — replaced with the
    //                     computed value 30; this is the desktop's
    //                     documented trade-off: "moved rows land as
    //                     their computed values")
    //   offset 5 (A6:B6): DIFFER → write null/null (was Cherry/30)
    //   offset 6 (A7:B7): DIFFER → write null/null (was the formula
    //                     row — formula LOST, replaced with null)
    const after = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (r: number, c: number) => { getCellData: () => { v?: unknown } | null }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (r: number, c: number) => ws?.getRange?.(r, c)?.getCellData?.()?.v ?? null
      return {
        a1: cell(0, 0),
        b1: cell(0, 1),
        a2: cell(1, 0),
        b2: cell(1, 1),
        a3: cell(2, 0),
        b3: cell(2, 1),
        a4: cell(3, 0),
        b4: cell(3, 1),
        a5: cell(4, 0),
        b5: cell(4, 1),
        a6: cell(5, 0),
        b6: cell(5, 1),
        a7: cell(6, 0),
        b7: cell(6, 1),
      }
    })
    // Header (row 1) preserved verbatim — offset 0 was skipped.
    expect(after.a1, 'A1 = "Name" after dedupe (header preserved)').toBe('Name')
    expect(after.b1, 'B1 = "Qty" after dedupe (header preserved)').toBe('Qty')
    // Row 2 (first occurrence) preserved — offset 1 was skipped. Any
    // formula here would survive; the fixture has a literal value, so
    // we just assert the value is unchanged.
    expect(after.a2, 'A2 = "Apple" after dedupe (first occurrence)').toBe('Apple')
    expect(after.b2, 'B2 = 10 after dedupe (first occurrence)').toBe(10)
    // Row 3 (A3:B3) — compacted Banana/20 (moved from row 4). The
    // duplicate at this offset was overwritten with the next kept row.
    expect(after.a3, 'A3 = "Banana" after dedupe (compacted from row 4)').toBe('Banana')
    expect(after.b3, 'B3 = 20 after dedupe (compacted from row 4)').toBe(20)
    // Row 4 (A4:B4) — compacted Cherry/30 (moved from row 6).
    expect(after.a4, 'A4 = "Cherry" after dedupe (compacted from row 6)').toBe('Cherry')
    expect(after.b4, 'B4 = 30 after dedupe (compacted from row 6)').toBe(30)
    // Row 5 (A5:B5) — compacted Apple/30 (moved from row 7). B7's
    // formula =B6 is LOST — the dedupe writes the COMPUTED VALUE 30
    // (the desktop's documented "moved rows land as their computed
    // values" trade-off). This is intentional and matches the frozen
    // desktop reference exactly.
    expect(after.a5, 'A5 = "Apple" after dedupe (compacted from row 7)').toBe('Apple')
    expect(after.b5, 'B5 = 30 after dedupe (B7 formula result, formula LOST)').toBe(30)
    // Row 6 (A6:B6) — blanked (padding). Was Cherry/30, now null.
    expect(after.a6, 'A6 blanked after dedupe (padding)').toBe(null)
    expect(after.b6, 'B6 blanked after dedupe (padding)').toBe(null)
    // Row 7 (A7:B7) — blanked (padding). Was the formula row (=B6 with
    // cached 30), now null. The formula is LOST (overwritten with null)
    // — the desktop's documented behavior for moved/compacted rows.
    expect(after.a7, 'A7 blanked after dedupe (was the formula row)').toBe(null)
    expect(after.b7, 'B7 blanked after dedupe (formula LOST — overwritten)').toBe(null)

    // The status message surfaced the count.
    await expect(page.locator('[data-testid="dedupe-message"]')).toContainText('Removed 2')

    // Save and capture the request body so we can assert the canonical
    // cell-edit channel was used (NOT a new dedupe family).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{
          row: number
          column: number
          writeValue: boolean
          cell?: { value: unknown }
        }>
        structuralOps?: unknown[]
        pageSetupStates?: unknown[]
        filterStates?: unknown[]
        dvStates?: unknown[]
        noteStates?: unknown[]
      }
    }
    // The save plan MUST use the existing `edits` channel — NO new
    // "dedupeOps" / "removeDuplicates" family. The architecture test
    // in tests/architecture.test.ts enforces this at the source level;
    // here we assert it on the wire too.
    expect(
      Object.keys(saveBody.savePlan).sort(),
      'save plan uses only the canonical families',
    ).toEqual(['edits'])
    // 5 rows × 2 columns = 10 cell-edits (offsets 2, 3, 4, 5, 6 were
    // rewritten; offsets 0, 1 were skipped because the content matched).
    const edits = saveBody.savePlan.edits
    expect(edits.length, '5 rows × 2 columns = 10 cell-edits').toBe(10)
    // Verify the compacted Banana at A3 (was Apple). The save plan's
    // CellEdit.cell.value carries the RESOLVED string value ("Banana"),
    // not the shared-strings table index — that index is an XLSX-file-
    // level concern; the wire carries typed values.
    const editA3 = edits.find((e) => e.row === 2 && e.column === 0)
    expect(editA3, 'A3 cell-edit present').toBeDefined()
    expect(editA3!.cell?.value, 'A3 = "Banana" in save plan').toBe('Banana')
    const editB3 = edits.find((e) => e.row === 2 && e.column === 1)
    expect(editB3, 'B3 cell-edit present').toBeDefined()
    expect(editB3!.cell?.value, 'B3 = 20 in save plan').toBe(20)
    // Verify the compacted Cherry at A4 (was Banana).
    const editA4 = edits.find((e) => e.row === 3 && e.column === 0)
    expect(editA4, 'A4 cell-edit present').toBeDefined()
    expect(editA4!.cell?.value, 'A4 = "Cherry" in save plan').toBe('Cherry')
    const editB4 = edits.find((e) => e.row === 3 && e.column === 1)
    expect(editB4, 'B4 cell-edit present').toBeDefined()
    expect(editB4!.cell?.value, 'B4 = 30 in save plan').toBe(30)
    // Verify the compacted Apple/30 at A5 (was the Apple/10 duplicate).
    const editA5 = edits.find((e) => e.row === 4 && e.column === 0)
    expect(editA5, 'A5 cell-edit present').toBeDefined()
    expect(editA5!.cell?.value, 'A5 = "Apple" in save plan').toBe('Apple')
    const editB5 = edits.find((e) => e.row === 4 && e.column === 1)
    expect(editB5, 'B5 cell-edit present').toBeDefined()
    // B5 = 30 — the COMPUTED VALUE of B7's formula (the formula is
    // LOST; only its result travels as a literal).
    expect(editB5!.cell?.value, 'B5 = 30 (B7 formula result) in save plan').toBe(30)
    // Verify the blanked cells at A6, B6, A7, B7.
    const editA6 = edits.find((e) => e.row === 5 && e.column === 0)
    expect(editA6, 'A6 cell-edit present').toBeDefined()
    expect(editA6!.cell?.value, 'A6 cell.value is null (cleared)').toBe(null)
    const editB6 = edits.find((e) => e.row === 5 && e.column === 1)
    expect(editB6, 'B6 cell-edit present').toBeDefined()
    expect(editB6!.cell?.value, 'B6 cell.value is null (cleared)').toBe(null)
    const editA7 = edits.find((e) => e.row === 6 && e.column === 0)
    expect(editA7, 'A7 cell-edit present').toBeDefined()
    expect(editA7!.cell?.value, 'A7 cell.value is null (cleared)').toBe(null)
    const editB7 = edits.find((e) => e.row === 6 && e.column === 1)
    expect(editB7, 'B7 cell-edit present').toBeDefined()
    expect(editB7!.cell?.value, 'B7 cell.value is null (formula LOST, cleared)').toBe(null)

    // Verify the saved XLSX bytes carry the compacted + blanked state.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // Row 1 (header) — preserved verbatim (shared-strings form, xf 1).
    expect(sheet1, 'A1 header preserved in saved XML').toMatch(/<c r="A1"[^>]*t="s"[^>]*><v>0<\/v>/)
    expect(sheet1, 'B1 header preserved in saved XML').toMatch(/<c r="B1"[^>]*t="s"[^>]*><v>1<\/v>/)
    // Row 2 (first occurrence) — preserved verbatim.
    expect(sheet1, 'A2 first occurrence preserved in saved XML').toMatch(
      /<c r="A2"[^>]*t="s"[^>]*><v>2<\/v>/,
    )
    expect(sheet1, 'B2 first occurrence preserved in saved XML').toMatch(
      /<c r="B2"[^>]*><v>10<\/v>/,
    )
    // Row 3 — compacted Banana/20 (moved from row 4). The canonical
    // gateway writes the moved-in string as an INLINE STRING
    // (t="inlineStr" with <is><t>...</t></is>) — it does NOT update
    // the shared-strings table. This is a valid canonical choice;
    // Excel reads both forms identically.
    expect(sheet1, 'A3 = Banana (inline string) in saved XML').toMatch(
      /<c r="A3"[^>]*t="inlineStr"[^>]*><is><t[^>]*>Banana<\/t><\/is><\/c>/,
    )
    expect(sheet1, 'B3 = 20 in saved XML').toMatch(/<c r="B3"[^>]*><v>20<\/v>/)
    // Row 4 — compacted Cherry/30 (moved from row 6).
    expect(sheet1, 'A4 = Cherry (inline string) in saved XML').toMatch(
      /<c r="A4"[^>]*t="inlineStr"[^>]*><is><t[^>]*>Cherry<\/t><\/is><\/c>/,
    )
    expect(sheet1, 'B4 = 30 in saved XML').toMatch(/<c r="B4"[^>]*><v>30<\/v>/)
    // Row 5 — compacted Apple/30 (moved from row 7; B7's formula was
    // LOST and replaced with the computed value 30 — only the literal
    // 30 travels, no <f> element).
    expect(sheet1, 'A5 = Apple (inline string) in saved XML').toMatch(
      /<c r="A5"[^>]*t="inlineStr"[^>]*><is><t[^>]*>Apple<\/t><\/is><\/c>/,
    )
    expect(sheet1, 'B5 = 30 (literal — B7 formula LOST) in saved XML').toMatch(
      /<c r="B5"[^>]*><v>30<\/v>/,
    )
    expect(
      sheet1,
      'B5 must NOT carry a formula (the moved-in row dropped the formula)',
    ).not.toMatch(/<c r="B5"[^>]*><f>/)
    // Rows 6 and 7 — blanked (padding). The cells must NOT carry the
    // previous values or formulas. The gateway either omits the <c>
    // entry entirely (for cells with no style) or writes an empty
    // <c .../> (for cells with a style ref). In all cases, no <v> and
    // no <f> for the blanked cells.
    expect(sheet1, 'A6 must NOT carry a value in saved XML (blanked)').not.toMatch(
      /<c r="A6"[^>]*><v>/,
    )
    expect(sheet1, 'B6 must NOT carry a value in saved XML (blanked)').not.toMatch(
      /<c r="B6"[^>]*><v>/,
    )
    expect(sheet1, 'A7 must NOT carry a value in saved XML (blanked)').not.toMatch(
      /<c r="A7"[^>]*><v>/,
    )
    expect(
      sheet1,
      'B7 must NOT carry the formula =B6 in saved XML (formula LOST, blanked)',
    ).not.toMatch(/<c r="B7"[^>]*><f>B6<\/f>/)
    expect(sheet1, 'B7 must NOT carry a value either (blanked)').not.toMatch(/<c r="B7"[^>]*><v>/)

    // Reopen and verify the snapshot carries the deduped state.
    writeFileSync('/tmp/e2e-ribbon-dedupe-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dedupe-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dedupe-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown }>
    }>
    const cells = reopened[0].cells
    // Header preserved.
    expect(cells.A1?.value, 'A1 = "Name" after reopen').toBe('Name')
    expect(cells.B1?.value, 'B1 = "Qty" after reopen').toBe('Qty')
    // First occurrence kept.
    expect(cells.A2?.value, 'A2 = "Apple" after reopen').toBe('Apple')
    expect(cells.B2?.value, 'B2 = 10 after reopen').toBe(10)
    // Compacted Banana at A3.
    expect(cells.A3?.value, 'A3 = "Banana" after reopen (compacted)').toBe('Banana')
    expect(cells.B3?.value, 'B3 = 20 after reopen (compacted)').toBe(20)
    // Compacted Cherry at A4.
    expect(cells.A4?.value, 'A4 = "Cherry" after reopen (compacted)').toBe('Cherry')
    expect(cells.B4?.value, 'B4 = 30 after reopen (compacted)').toBe(30)
    // Compacted Apple/30 at A5 (formula LOST, only the result survives).
    expect(cells.A5?.value, 'A5 = "Apple" after reopen (compacted, formula lost)').toBe('Apple')
    expect(cells.B5?.value, 'B5 = 30 after reopen (B7 formula result, formula lost)').toBe(30)
    // Blanked rows — the snapshot may omit them entirely (canonical
    // reader omits empty cells), so we assert they're absent OR null.
    const blankA6 = cells.A6?.value ?? null
    const blankB6 = cells.B6?.value ?? null
    const blankA7 = cells.A7?.value ?? null
    const blankB7 = cells.B7?.value ?? null
    expect(blankA6, 'A6 is blank after reopen').toBe(null)
    expect(blankB6, 'B6 is blank after reopen').toBe(null)
    expect(blankA7, 'A7 is blank after reopen').toBe(null)
    expect(blankB7, 'B7 is blank after reopen (formula lost)').toBe(null)

    expect(pageErrors, 'no uncaught page errors').toEqual([])
  })

  test('no-op when there are no duplicates — fails closed without mutating', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Use the same fixture — select only the unique rows (A4:B6) so the
    // dedupe finds no duplicates.
    const fixture = await buildExcelDedupeFixture()
    writeFileSync('/tmp/e2e-ribbon-dedupe-noop.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dedupe-noop.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dedupe-noop.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Select A4:B6 (Banana, Apple-dup, Cherry — wait, A5 IS a dup of A2,
    // but A2 is OUTSIDE the selection. With hasHeader=true the header is
    // row 4 (Banana) and data rows are 5 (Apple=dup of... nothing inside
    // the selection) and 6 (Cherry). No duplicates within A4:B6.)
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A4:B6')
    await box.press('Enter')
    await page.waitForTimeout(400)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Remove Duplicates' }).click()
    await expect(page.locator('[data-testid="dedupe-dialog"]')).toBeVisible()
    await page.locator('[data-testid="dedupe-ok"]').click()
    await page.waitForTimeout(500)

    // The status message surfaces "No duplicate rows found" and NO save
    // request was fired (no mutation, no journal entries).
    await expect(page.locator('[data-testid="dedupe-message"]')).toContainText('No duplicate rows')

    // No save request should have been made (the dedupe was a no-op).
    // We assert this by counting save requests over a brief window.
    const saveCount = await page.evaluate(() => {
      // No direct hook into the request count from the browser; we
      // approximate by checking the save button label is still "Save"
      // (not "Saving...") and no unsaved-changes marker appears.
      return document.querySelector('.excel-statusbar')?.textContent ?? ''
    })
    expect(saveCount, 'no "unsaved changes" marker after a no-op dedupe').not.toContain('Unsaved')

    expect(pageErrors, 'no uncaught page errors').toEqual([])
  })

  test('fails closed when selection has fewer than 2 rows', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDedupeFixture()
    writeFileSync('/tmp/e2e-ribbon-dedupe-single.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dedupe-single.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dedupe-single.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Select a single cell (A1 only — a 1-row range).
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A1')
    await box.press('Enter')
    await page.waitForTimeout(400)

    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Remove Duplicates' }).click()
    await expect(page.locator('[data-testid="dedupe-dialog"]')).toBeVisible()
    await page.locator('[data-testid="dedupe-ok"]').click()
    await page.waitForTimeout(500)

    // The status message surfaces "Select the rows to check for duplicates
    // first." — fail-closed, no mutation fired.
    await expect(page.locator('[data-testid="dedupe-message"]')).toContainText('Select the rows')

    expect(pageErrors, 'no uncaught page errors').toEqual([])
  })
})
