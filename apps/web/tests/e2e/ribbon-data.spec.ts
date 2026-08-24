/**
 * REAL browser E2E — Data tab (Phase 4 Inc. 3, revised after architect review).
 *
 * Proves Sort ascending/descending PERSISTS through save/reopen via the
 * CANONICAL Univer sort mutation + a typed row-permutation save plan,
 * and that Filter remains visibly disabled (the wire save plan does not
 * expose the filterStates family).
 *
 * Sort path (proven — canonical, no JS read-sort-rewrite):
 *   Data → Sort Asc
 *   → FRange.sort({ column: 0, ascending: true })   [public Univer facade]
 *   → sheet.command.sort-range → ReorderRangeMutation
 *     (deepClones the ENTIRE cell record via getCellRaw —
 *      v/f/s/p/si/t travel atomically with the row)
 *   → ExcelEditor's journal subscription captures the row permutation
 *     as a `reorder-rows` structural op (range + order map)
 *   → savePlan.structuralOps → /api/office/workbooks/save
 *   → applyCellEditsToXlsx → applyStructuralOps permutes <row> blocks
 *     (renumbers r= on <row> and inner <c>, leaves cell contents
 *      untouched — styles, numfmt, formulas, hyperlinks all survive)
 *   → XLSX bytes (new row order)
 *   → reopen → readBasicWorkbook → snapshot carries the sorted order
 *
 * The previous JS-sort implementation (read values into JS, sort, write
 * back via setValueForCell) is GONE — it stripped cell metadata and was
 * flagged by the architect: "the implementation explicitly abandons
 * Univer's canonical sort mutation and instead reads cell values/formulas
 * into JavaScript, sorts them, and writes them back with setValueForCell."
 *
 * Filter path (BLOCKED):
 *   The wire save plan does not expose the filterStates family. Until the
 *   wire is extended, the Filter button is visibly disabled.
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
  buildExcelFixture,
  buildExcelSortFixture,
  buildExcelSortFormulaFixture,
  readZipEntry,
} from './fixtures'

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
      const rt = (
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
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)

    // Verify the active range is B1:B4 (the Name Box navigation selected it).
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
    expect(activeRangeBefore, 'active range is B1:B4 before sort').toBe('B1:B4')

    // Verify the values were typed correctly before sorting.
    const valuesBefore = await page.evaluate(() => {
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
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{
            kind: string
            range?: { startRow: number; endRow: number; startColumn: number; endColumn: number }
            order?: Record<string, number>
          }>
        }>
      }
    }
    // The sort journal now produces a SINGLE `reorder-rows` structural op
    // (NOT per-cell value edits — the previous JS-sort's approach is gone).
    // The op carries the row permutation Univer's ReorderRangeMutation
    // emitted; the gateway permutes <row> blocks atomically on save.
    const structuralOps = saveBody.savePlan.structuralOps
    expect(structuralOps, 'sort journaled as a structural op').toBeDefined()
    expect(structuralOps!.length, 'at least one sheet with structural ops').toBeGreaterThan(0)
    const sortSheet = structuralOps!.find((s) => s.sheetName === 'Data')
    expect(sortSheet, 'structural op targets the Data sheet').toBeDefined()
    const reorderOp = sortSheet!.ops.find((op) => op.kind === 'reorder-rows')
    expect(reorderOp, 'reorder-rows op present in the save plan').toBeDefined()
    expect(reorderOp!.order, 'reorder-rows op carries the permutation map').toBeDefined()
    // The order map is a bijection over rows 0..3 (B1:B4 = 0-based 0..3,
    // single-column range), in Univer's native DEST→SRC shape (NEW[dest]
    // = OLD[order[dest]]). The post-sort row order is 30, 20, 10, 5
    // (desc): order[0] = 1 (new row 0 takes old row 1's content, B2=30);
    // order[1] = 3 (new row 1 takes old row 3's, B4=20); order[2] = 0
    // (B1=10); order[3] = 2 (B3=5). The exact map semantics are asserted
    // in the unit test (xlsx-structure.test.ts — "matches Univer
    // ReorderRangeMutation semantics"); here we just confirm the op is
    // present and well-formed.
    expect(
      Object.keys(reorderOp!.order!).length,
      'order map covers the sort range',
    ).toBeGreaterThan(0)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The sort wrote rows back in desc order (30, 20, 10, 5) — assert the
    // first row (row 1) carries the largest value (30), and the 4th row
    // carries the smallest (5). The cell values live inline as <v>. The
    // <row> block permutation preserved the entire cell record (value,
    // style, numfmt, formula) — the gateway only renumbered r= attributes.
    expect(sheet1, 'B1 = 30 (largest after desc sort)').toMatch(/<c r="B1"[^>]*><v>30<\/v>/)
    expect(sheet1, 'B4 = 5 (smallest after desc sort)').toMatch(/<c r="B4"[^>]*><v>5<\/v>/)

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

  // ── Architect-mandated regression fixture ──────────────────────────────
  //
  // The architect's review of commit 9278962 explicitly required:
  //   "Add a regression fixture containing:
  //    * mixed styles
  //    * number formats
  //    * formulas with relative references
  //    * hyperlinks or notes
  //    Prove sort with save → reopen, checking both displayed values AND
  //    formulas/styles."
  //
  // buildExcelSortFixture (in fixtures.ts) is that fixture. This test
  // sorts it (alphabetical by column A: Apple, Banana, Cherry) and
  // asserts that styles, number formats, formula text, and worksheet-
  // level hyperlink definitions ALL survive the save/reopen round-trip.
  test('Sort preserves styles, numfmt, formulas, and hyperlinks (regression fixture)', async ({
    page,
  }) => {
    test.setTimeout(180_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Load the rich sort fixture (mixed styles + currency numfmt + relative
    // formula references + worksheet-level hyperlink).
    const fixture = await buildExcelSortFixture()
    writeFileSync('/tmp/e2e-ribbon-sort-fidelity.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-sort-fidelity.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-sort-fidelity.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Select A2:D4 via the Name Box (the data rows — row 1 is the header,
    // NOT in the sort range). The Univer FRange.sort facade takes the
    // active range's first column as the sort key (column 0 of the
    // active range = column A here).
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A2:D4')
    await box.press('Enter')
    await page.waitForTimeout(400)

    // Switch to the Data tab and click Sort Asc (alphabetical: Apple,
    // Banana, Cherry — A2 becomes "Apple", A3 becomes "Banana", A4
    // becomes "Cherry").
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Sort Asc' }).click()
    await page.waitForTimeout(800)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Verify the in-session Univer state — alphabetical sort succeeded.
    const valuesAfter = await page.evaluate(() => {
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
      return {
        a2: ws?.getRange?.(1, 0)?.getCellData?.()?.v,
        a3: ws?.getRange?.(2, 0)?.getCellData?.()?.v,
        a4: ws?.getRange?.(3, 0)?.getCellData?.()?.v,
      }
    })
    expect(valuesAfter.a2, 'A2 = Apple (alphabetical sort)').toBe('Apple')
    expect(valuesAfter.a3, 'A3 = Banana').toBe('Banana')
    expect(valuesAfter.a4, 'A4 = Cherry').toBe('Cherry')

    // Save and capture the request body — assert the reorder-rows
    // structural op is present (the canonical save representation).
    const sortReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await sortReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{
            kind: string
            range?: { startRow: number; endRow: number; startColumn: number; endColumn: number }
            order?: Record<string, number>
          }>
        }>
      }
    }
    const structuralOps = saveBody.savePlan.structuralOps
    expect(structuralOps, 'sort journaled as a structural op').toBeDefined()
    const sortSheet = structuralOps!.find((s) => s.sheetName === 'Sort')
    expect(sortSheet, 'structural op targets the Sort sheet').toBeDefined()
    const reorderOp = sortSheet!.ops.find((op) => op.kind === 'reorder-rows')
    expect(reorderOp, 'reorder-rows op present in the save plan').toBeDefined()
    expect(reorderOp!.order, 'reorder-rows op carries the permutation map').toBeDefined()

    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')

    // ── Row order: alphabetical, styles travel with the row ────────────
    // Pre-sort: A2="Banana" (italic xf 2), A3="Cherry" (regular),
    //           A4="Apple" (bold xf 4).
    // Post-sort (alphabetical): A2=Apple (bold), A3=Banana (italic),
    // A4=Cherry (regular). Univer's FormulaReorderController interceptor
    // journals the ENTIRE sorted matrix (every cell in the range, not just
    // formula cells), so the A-column string cells are also re-written as
    // journaled value edits — the gateway's cell-edit writer emits them as
    // inlineStr (t="inlineStr") while PRESERVING the style ref the row
    // permutation placed on the cell. s="4" (bold) travels to A2, s="2"
    // (italic) travels to A3, and Cherry stays unstyled.
    expect(sheet1, 'A2 = Apple with bold style (s="4") after sort').toMatch(
      /<c r="A2"[^>]*\bs="4"[^>]*t="inlineStr"[^>]*><is><t[^>]*>Apple<\/t><\/is>/,
    )
    expect(sheet1, 'A3 = Banana with italic style (s="2") after sort').toMatch(
      /<c r="A3"[^>]*\bs="2"[^>]*t="inlineStr"[^>]*><is><t[^>]*>Banana<\/t><\/is>/,
    )
    expect(sheet1, 'A4 = Cherry (regular, no style ref) after sort').toMatch(
      /<c r="A4"[^>]*t="inlineStr"[^>]*><is><t[^>]*>Cherry<\/t><\/is>/,
    )
    // A4 must NOT carry a style ref (Cherry was regular).
    const a4CellMatch = /<c r="A4"[^>]*>/.exec(sheet1)
    expect(a4CellMatch, 'A4 cell element present').not.toBeNull()
    expect(a4CellMatch![0], 'A4 has no style attribute (regular)').not.toMatch(/\bs="/)

    // ── Number format (currency, xf 3, s="3") travels with the C column ─
    // The currency numfmt style ref (s="3") lives on the C cells. The
    // post-sort C2 carries the original C4 (Apple's price = 2.00), C3
    // carries the original C2 (Banana's price = 1.5), C4 carries the
    // original C3 (Cherry's price = 3). The journaled value edits preserve
    // the style ref the row permutation placed on each cell.
    expect(sheet1, 'C2 carries currency numfmt (s="3") with Apple price 2').toMatch(
      /<c r="C2"[^>]*\bs="3"[^>]*><v>2<\/v>/,
    )
    expect(sheet1, 'C3 carries currency numfmt (s="3") with Banana price 1.5').toMatch(
      /<c r="C3"[^>]*\bs="3"[^>]*><v>1\.5<\/v>/,
    )
    expect(sheet1, 'C4 carries currency numfmt (s="3") with Cherry price 3').toMatch(
      /<c r="C4"[^>]*\bs="3"[^>]*><v>3<\/v>/,
    )

    // ── Formula references are REWRITTEN Excel-style by Univer's sort ──
    // Univer's FormulaReorderController (a sheet-interceptor on
    // ReorderRangeCommand) shifts relative references by the row offset:
    // old D4's `=B4*C4` moving to D2 (offset −2) becomes `=B2*C2` —
    // preserving the same-row relationship, exactly like Excel's sort.
    // The interceptor's SetRangeValuesMutations are journaled as formula
    // CellEdits, composing on top of the reorder-rows structural op. A
    // journaled formula edit drops the cached <v> (canonical engine
    // semantics — an edited formula has no stale cache); Excel/Univer
    // recalculate on open.
    //   - D2 = Apple's row: `=B2*C2` (20 * 2 = 40 on recalc)
    //   - D3 = Banana's row: `=B3*C3` (30 * 1.5 = 45 on recalc)
    //   - D4 = Cherry's row: `=B4*C4` (10 * 3 = 30 on recalc)
    expect(sheet1, 'D2 carries the Excel-rewritten formula B2*C2 (no stale cached v)').toMatch(
      /<c r="D2"[^>]*><f>B2\*C2<\/f><\/c>/,
    )
    expect(sheet1, 'D3 carries the Excel-rewritten formula B3*C3 (no stale cached v)').toMatch(
      /<c r="D3"[^>]*><f>B3\*C3<\/f><\/c>/,
    )
    expect(sheet1, 'D4 carries the Excel-rewritten formula B4*C4 (no stale cached v)').toMatch(
      /<c r="D4"[^>]*><f>B4\*C4<\/f><\/c>/,
    )

    // ── Worksheet-level hyperlink definition is preserved verbatim ────
    // Univer's ReorderRangeMutation does NOT move worksheet-level
    // <hyperlinks> definitions; the gateway mirrors that. The
    // hyperlink ref="A2" stays at A2 — the cell content at A2 is now
    // "Apple" (post-sort), but the hyperlink target is still the banana
    // URL. This matches Univer's live state.
    expect(sheet1, 'hyperlink ref="A2" preserved verbatim (Univer does not move it)').toMatch(
      /<hyperlink ref="A2" r:id="rId1"\/>/,
    )
    const sheet1Rels = await readZipEntry(saved, 'xl/_rels/sheet1.xml.rels')
    expect(sheet1Rels, 'banana hyperlink relationship target preserved').toContain(
      'https://example.com/banana',
    )

    // ── Reopen and verify the snapshot carries the sorted values ──────
    writeFileSync('/tmp/e2e-ribbon-sort-fidelity-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-sort-fidelity-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-sort-fidelity-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown; formula?: string }>
      styles?: Record<string, { bold?: boolean; italic?: boolean; numberFormat?: string }>
    }>
    expect(reopened[0].name, 'sheet name preserved').toBe('Sort')
    const cells = reopened[0].cells
    expect(cells.A2?.value, 'A2 = Apple after reopen').toBe('Apple')
    expect(cells.A3?.value, 'A3 = Banana after reopen').toBe('Banana')
    expect(cells.A4?.value, 'A4 = Cherry after reopen').toBe('Cherry')
    // Number values in B column traveled with the rows.
    expect(cells.B2?.value, 'B2 = 20 (Apple qty) after reopen').toBe(20)
    expect(cells.B3?.value, 'B3 = 30 (Banana qty) after reopen').toBe(30)
    expect(cells.B4?.value, 'B4 = 10 (Cherry qty) after reopen').toBe(10)
    // Currency-formatted C values.
    expect(cells.C2?.value, 'C2 = 2 (Apple price) after reopen').toBe(2)
    expect(cells.C3?.value, 'C3 = 1.5 (Banana price) after reopen').toBe(1.5)
    expect(cells.C4?.value, 'C4 = 3 (Cherry price) after reopen').toBe(3)
    // Formula text in D column survived with the Excel-style REWRITTEN
    // references (Univer's FormulaReorderController shifted the relative
    // refs by the row offset during the sort; the journal captured the
    // interceptor's formula mutations, composing on the structural op).
    // readBasicWorkbook parses formulas with the leading "=" — so the
    // reopened snapshot's formulas are "=B2*C2" etc.
    expect(cells.D2?.formula, 'D2 formula = =B2*C2 (rewritten) after reopen').toBe('=B2*C2')
    expect(cells.D3?.formula, 'D3 formula = =B3*C3 (rewritten) after reopen').toBe('=B3*C3')
    expect(cells.D4?.formula, 'D4 formula = =B4*C4 (rewritten) after reopen').toBe('=B4*C4')

    // ── Styles survived (snapshot's styles map) ────────────────────────
    const styles = reopened[0].styles
    expect(styles, 'snapshot carries styles after sort+save+reopen').toBeDefined()
    // A2 = Apple, original style was bold (xf 4) — bold travels with the row.
    expect(styles!['A2']?.bold, "A2 carries Apple's bold style after reopen").toBe(true)
    // A3 = Banana, original style was italic (xf 2) — italic travels.
    expect(styles!['A3']?.italic, "A3 carries Banana's italic style after reopen").toBe(true)
    // A4 = Cherry, original style was regular — no bold, no italic.
    expect(styles!['A4']?.bold, 'A4 has no bold (Cherry was regular)').not.toBe(true)
    expect(styles!['A4']?.italic, 'A4 has no italic (Cherry was regular)').not.toBe(true)
    // C column cells carry the currency numfmt.
    expect(styles!['C2']?.numberFormat, 'C2 carries currency numfmt after reopen').toMatch(/\$/)

    expect(pageErrors).toEqual([])
  })

  // ── Formula-semantics sort gate (architect review requirement) ─────────
  //
  // The canonical sort splits formula responsibility across the pipeline:
  //   1. UNIVER (browser): FormulaReorderController rewrites RELATIVE
  //      references by the row delta (=A3*10 moving up one row becomes
  //      =A2*10) and leaves ABSOLUTE references ($D$1) untouched; the
  //      rewrites are journaled as formula CellEdits.
  //   2. GATEWAY (server): reorder-rows permutes <row> blocks WITHOUT
  //      touching formula text — the journaled formula edits then overwrite
  //      the verbatim text with the rewritten form.
  //
  // This test proves the two stay synchronized, using the architect's exact
  // scenario (sort A2:B3-style range descending, formulas =A{row}*10) plus
  // the absolute-reference probe =A{row}+$D$1 with D1 OUTSIDE the sort range:
  //
  //   Pre-sort:  A2=10  B2==A2*10(100)   C2==A2+$D$1(15)
  //              A3=30  B3==A3*10(300)   C3==A3+$D$1(35)   D1=5
  //   Sort A2:C3 descending → rows swap (30 first).
  //   Post-sort: A2=30  B2==A2*10(300)   C2==A2+$D$1(35)
  //              A3=10  B3==A3*10(100)   C3==A3+$D$1(15)   D1=5 (untouched)
  //
  // Failure modes caught:
  //   - Browser rewrite dropped → B2 keeps verbatim =A3*10 → recalcs to
  //     10*10=100 (wrong; expected 300) and the reopened formula is "=A3*10".
  //   - Gateway rewriting formulas (overstepping) or dropping the structural
  //     op → wrong text or unmoved values in the saved XML.
  //   - Absolute ref shifted → C2 becomes =A2+$D$2 (or similar) ≠ =A2+$D$1.
  test('Sort formula semantics: relative refs rewrite, absolute refs untouched (architect gate)', async ({
    page,
  }) => {
    test.setTimeout(180_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelSortFormulaFixture()
    writeFileSync('/tmp/e2e-ribbon-sort-fx.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-sort-fx.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-sort-fx.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Select the data range A2:C3 (D1 stays outside in both axes).
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A2:C3')
    await box.press('Enter')
    await page.waitForTimeout(400)

    // Sort DESCENDING by the range's first column (A): 30 before 10 → the
    // data rows actually SWAP (old row 3 → new row 2, old row 2 → new row 3).
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Sort Desc' }).click()
    await page.waitForTimeout(1000)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // ── In-session proof: the LIVE recalculated values distinguish a
    //    rewritten formula from a verbatim one. If B2 kept the verbatim
    //    =A3*10, it would recalculate to 10*10=100; the rewritten =A2*10
    //    gives 30*10=300. Likewise C2: verbatim =A3+$D$1 → 10+5=15;
    //    rewritten =A2+$D$1 → 30+5=35.
    const live = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => { getCellData: () => Record<string, unknown> | null }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (row: number, col: number) => ws?.getRange?.(row, col)?.getCellData?.() ?? null
      return {
        a2: cell(1, 0),
        b2: cell(1, 1),
        c2: cell(1, 2),
        a3: cell(2, 0),
        b3: cell(2, 1),
        c3: cell(2, 2),
        d1: cell(0, 3),
      }
    })
    expect(live.a2?.v, 'A2 = 30 after desc sort').toBe(30)
    expect(live.a3?.v, 'A3 = 10 after desc sort').toBe(10)
    expect(live.d1?.v, 'D1 = 5 untouched (outside sort range)').toBe(5)
    // The formula-rewrite proof: rewritten text + recalculated value.
    expect(String(live.b2?.f).replace(/\s+/g, ''), 'B2 formula REWRITTEN to =A2*10').toBe('=A2*10')
    expect(
      live.b2?.v,
      'B2 recalculates to 300 (=A2*10 with A2=30; verbatim =A3*10 would give 100)',
    ).toBe(300)
    expect(
      String(live.c2?.f).replace(/\s+/g, ''),
      'C2 formula REWRITTEN to =A2+$D$1 (absolute untouched)',
    ).toBe('=A2+$D$1')
    expect(live.c2?.v, 'C2 recalculates to 35 (=A2+$D$1 → 30+5; verbatim would give 15)').toBe(35)
    expect(String(live.b3?.f).replace(/\s+/g, ''), 'B3 formula REWRITTEN to =A3*10').toBe('=A3*10')
    expect(live.b3?.v, 'B3 recalculates to 100').toBe(100)
    expect(String(live.c3?.f).replace(/\s+/g, ''), 'C3 formula REWRITTEN to =A3+$D$1').toBe(
      '=A3+$D$1',
    )
    expect(live.c3?.v, 'C3 recalculates to 15').toBe(15)

    // ── Save: the wire plan must carry BOTH halves of the split —
    //    the reorder-rows structural op AND the rewritten formula edits.
    const sortReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await sortReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{
          sheetName: string
          row: number
          column: number
          cell?: { formula?: string; value?: unknown }
        }>
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{
            kind: string
            order?: Record<string, number>
          }>
        }>
      }
    }
    const sortSheet = saveBody.savePlan.structuralOps?.find((s) => s.sheetName === 'SortFx')
    expect(sortSheet, 'structural op present for SortFx').toBeDefined()
    const reorderOp = sortSheet!.ops.find((op) => op.kind === 'reorder-rows')
    expect(reorderOp, 'reorder-rows op present').toBeDefined()
    // Descending sort of [10, 30] over 0-based rows 1..2: NEW[1]=OLD[2],
    // NEW[2]=OLD[1] → Univer's DEST→SRC map is exactly {1: 2, 2: 1}.
    expect(reorderOp!.order, 'exact permutation map (DEST→SRC)').toEqual({ 1: 2, 2: 1 })
    // The journaled formula edits carry the REWRITTEN text (wire format
    // strips the leading '='; XLSX <f> elements have no '=').
    const formulaEdit = (row: number, column: number) =>
      saveBody.savePlan.edits.find(
        (e) => e.sheetName === 'SortFx' && e.row === row && e.column === column,
      )
    expect(formulaEdit(1, 1)?.cell?.formula, 'B2 edit carries rewritten formula A2*10').toBe(
      'A2*10',
    )
    expect(
      formulaEdit(1, 2)?.cell?.formula,
      'C2 edit carries rewritten formula A2+$D$1 (absolute untouched)',
    ).toBe('A2+$D$1')
    expect(formulaEdit(2, 1)?.cell?.formula, 'B3 edit carries rewritten formula A3*10').toBe(
      'A3*10',
    )
    expect(formulaEdit(2, 2)?.cell?.formula, 'C3 edit carries rewritten formula A3+$D$1').toBe(
      'A3+$D$1',
    )

    // ── Saved XML: the composed result — the structural op moved the rows
    //    (verbatim formulas), then the formula edits overwrote the text with
    //    the rewritten forms. Whitespace-normalized comparisons guard against
    //    engine token-serialization spacing.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    const formulaOf = (xml: string, cell: string): string | null => {
      const m = new RegExp(`<c r="${cell}"[^>]*>(?:<f>([\\s\\S]*?)</f>)`).exec(xml)
      return m ? m[1]!.replace(/\s+/g, '') : null
    }
    expect(
      formulaOf(sheet1, 'B2'),
      'saved XML B2 formula = A2*10 (rewritten, NOT verbatim A3*10)',
    ).toBe('A2*10')
    expect(formulaOf(sheet1, 'B3'), 'saved XML B3 formula = A3*10').toBe('A3*10')
    expect(
      formulaOf(sheet1, 'C2'),
      'saved XML C2 formula = A2+$D$1 (relative shifted, ABSOLUTE untouched)',
    ).toBe('A2+$D$1')
    expect(formulaOf(sheet1, 'C3'), 'saved XML C3 formula = A3+$D$1').toBe('A3+$D$1')
    expect(sheet1, 'A2 = 30 (row moved)').toMatch(/<c r="A2"[^>]*><v>30<\/v>/)
    expect(sheet1, 'A3 = 10 (row moved)').toMatch(/<c r="A3"[^>]*><v>10<\/v>/)
    expect(sheet1, 'D1 = 5 untouched (outside the sort range)').toMatch(/<c r="D1"[^>]*><v>5<\/v>/)

    // ── Reopen: the saved bytes parse back with the rewritten formulas.
    writeFileSync('/tmp/e2e-ribbon-sort-fx-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-sort-fx-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-sort-fx-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown; formula?: string }>
    }>
    expect(reopened[0].name, 'sheet name preserved').toBe('SortFx')
    const cells = reopened[0].cells
    expect(cells.A2?.value, 'reopen A2 = 30').toBe(30)
    expect(cells.A3?.value, 'reopen A3 = 10').toBe(10)
    expect(cells.D1?.value, 'reopen D1 = 5 (outside range, untouched)').toBe(5)
    expect(cells.B2?.formula?.replace(/\s+/g, ''), 'reopen B2 formula = =A2*10').toBe('=A2*10')
    expect(cells.B3?.formula?.replace(/\s+/g, ''), 'reopen B3 formula = =A3*10').toBe('=A3*10')
    expect(
      cells.C2?.formula?.replace(/\s+/g, ''),
      'reopen C2 formula = =A2+$D$1 (absolute preserved)',
    ).toBe('=A2+$D$1')
    expect(cells.C3?.formula?.replace(/\s+/g, ''), 'reopen C3 formula = =A3+$D$1').toBe('=A3+$D$1')

    // ── Post-reopen recalculation: the reopened live workbook computes the
    //    SAME values as before the save — the full round-trip is
    //    semantically lossless for both reference classes.
    await page.waitForTimeout(1500)
    const recalc = await page.evaluate(() => {
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
      return {
        b2: ws?.getRange?.(1, 1)?.getCellData?.()?.v ?? null,
        c2: ws?.getRange?.(1, 2)?.getCellData?.()?.v ?? null,
        b3: ws?.getRange?.(2, 1)?.getCellData?.()?.v ?? null,
        c3: ws?.getRange?.(2, 2)?.getCellData?.()?.v ?? null,
      }
    })
    expect(
      recalc.b2,
      'post-reopen B2 recalculates to 300 (same-row relationship survived the round-trip)',
    ).toBe(300)
    expect(
      recalc.c2,
      'post-reopen C2 recalculates to 35 (absolute $D$1 still points at D1=5)',
    ).toBe(35)
    expect(recalc.b3, 'post-reopen B3 recalculates to 100').toBe(100)
    expect(recalc.c3, 'post-reopen C3 recalculates to 15').toBe(15)

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
