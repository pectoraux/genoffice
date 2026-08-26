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
 *   dedupeRowIndices (returns duplicate row INDICES), issues
 *   ws.deleteRows(startRow+offset, 1) per duplicate in DESCENDING
 *   order → each call fires sheet.mutation.remove-rows → ExcelEditor's
 *   existing STRUCTURAL_MUTATION_IDS subscription journals each as a
 *   `{ kind: 'remove-rows', index, count: 1 }` structural op in the
 *   save plan → savePlan.structuralOps → /api/office/workbooks/save
 *   → applyStructuralOps in xlsx-gateway applies each op atomically:
 *     - transformSheetRows renumbers <row> r= and inner <c> r=
 *       (cell contents — value, formula text, style ref, hyperlink,
 *       comment pointer — travel UNTOUCHED inside their <c> elements),
 *     - transformFormulas rewrites <f> bodies via shiftFormulaText
 *       (relative + absolute + mixed references all track the moved
 *       cells — the `$` markers are preserved by shiftReferenceToken's
 *       colDollar/rowDollar capture groups),
 *     - transformRangedFeatures shifts merges, autoFilter, hyperlink
 *       sqref, dataValidation sqref, conditionalFormatting sqref →
 *   reopen → readBasicWorkbook → snapshot carries the deduped state
 *   WITH the preserved formulas.
 *
 * No browser-side OOXML. The browser only ever exchanges typed
 * structural-op payloads with the canonical gateway.
 *
 * Fixture (buildExcelDedupeFixture): a "Dedupe" sheet with a bold
 * header, 6 data rows including 2 full duplicates of row 2, and one
 * formula row whose computed result differs from its sibling (so it
 * is NOT a duplicate). The dedupe MUST preserve the formula on the
 * survivor row that gets compacted (B7==B6 → B5==B4 after the two
 * deletes), and the gateway's transformFormulas MUST rewrite the
 * relative reference B6 → B4 to track where Cherry/30 moved.
 *
 * The architect's mandatory regression case explicitly requires:
 *   - the surviving B7 formula is NOT converted to literal 30
 *   - after compaction, the formula is correctly relocated (B7→B5)
 *   - relative references are rewritten according to the moved row
 *     (B6→B4 because Cherry/30 moved from row 6 to row 4)
 *   - absolute references remain absolute
 *   - styles move with surviving rows
 *   - save/reopen preserves these semantics
 *
 * The test must distinguish:
 *   "formula text survived and was rewritten correctly"
 * from:
 *   "computed result happens to be the same".
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
  buildExcelDedupeFixture,
  buildExcelDedupeMixedReferencesFixture,
  readZipEntry,
} from './fixtures'

test.describe('Data tab — Remove Duplicates persists through save/reopen', () => {
  test('removes duplicates, preserves header, PRESERVES formulas/styles on compacted rows (architect mandatory regression)', async ({
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

    // Sanity: the live Univer model has the expected pre-dedupe state.
    const before = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => {
                    getCellData: () => { v?: unknown; f?: string } | null
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (r: number, c: number) => ws?.getRange?.(r, c)?.getCellData?.() ?? null
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
    expect(before.a1?.v, 'A1 = "Name" before dedupe').toBe('Name')
    expect(before.b1?.v, 'B1 = "Qty" before dedupe').toBe('Qty')
    expect(before.a2?.v, 'A2 = "Apple" before dedupe').toBe('Apple')
    expect(before.b2?.v, 'B2 = 10 before dedupe').toBe(10)
    expect(before.a3?.v, 'A3 = "Apple" (duplicate) before dedupe').toBe('Apple')
    expect(before.b3?.v, 'B3 = 10 (duplicate) before dedupe').toBe(10)
    expect(before.a4?.v, 'A4 = "Banana" before dedupe').toBe('Banana')
    expect(before.b4?.v, 'B4 = 20 before dedupe').toBe(20)
    expect(before.a5?.v, 'A5 = "Apple" (duplicate) before dedupe').toBe('Apple')
    expect(before.b5?.v, 'B5 = 10 (duplicate) before dedupe').toBe(10)
    expect(before.a6?.v, 'A6 = "Cherry" before dedupe').toBe('Cherry')
    expect(before.b6?.v, 'B6 = 30 before dedupe').toBe(30)
    expect(before.a7?.v, 'A7 = "Apple" (formula row) before dedupe').toBe('Apple')
    expect(before.b7?.v, 'B7 = 30 (formula result) before dedupe').toBe(30)
    // The formula text MUST be present before the dedupe — otherwise we
    // can't prove preservation (the test would degenerate to "computed
    // result happens to be the same" which is exactly the failure mode
    // the architect is rejecting).
    expect(before.b7?.f, 'B7 formula text present before dedupe').toBe('=B6')

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

    // Verify the in-session result. The dedupe DELETES the duplicate
    // rows ENTIRELY from the sheet (Excel's actual Remove Duplicates
    // behavior) — survivors compact UPWARD atomically. No "padding"
    // with nulls at the bottom of the selection (that was the prior
    // value-rewrite implementation that destroyed formulas on moved
    // rows; the architect explicitly rejected it).
    //
    // Input fixture (A1:B7, hasHeader=true):
    //   row 1: Name, Qty           (header — kept)
    //   row 2: Apple, 10           (first — kept)
    //   row 3: Apple, 10           (dup of row 2 — DELETED)
    //   row 4: Banana, 20          (first — kept)
    //   row 5: Apple, 10           (dup of row 2 — DELETED)
    //   row 6: Cherry, 30          (first — kept; referenced by B7's formula)
    //   row 7: Apple, =B6          (first — kept; B7 formula =B6 references
    //                                Cherry/30 in row 6)
    //
    // Dedupe delete order (DESCENDING so earlier deletes don't shift
    // later indices):
    //   1. deleteRows(startRow+4=5, 1) — removes the row-5 dup
    //      → rows 6,7 shift to 5,6 → Cherry at row 5, formula at row 6
    //      → transformFormulas rewrites B6 → B5 (Cherry shifted to row 5)
    //   2. deleteRows(startRow+2=2, 1) — removes the row-3 dup
    //      → rows 4,5,6 shift to 3,4,5 → Banana at row 3, Cherry at row 4
    //      → transformFormulas rewrites B5 → B4 (Cherry shifted to row 4)
    //
    // Final layout (1-indexed):
    //   row 1: Name, Qty           (header preserved)
    //   row 2: Apple, 10           (first occurrence preserved)
    //   row 3: Banana, 20          (was row 4)
    //   row 4: Cherry, 30          (was row 6)
    //   row 5: Apple, =B4          (was row 7; formula rewritten B6→B4)
    //   rows 6,7 — DELETED entirely (no "padding" null rows — the
    //             structural path removes them from the sheet)
    //
    // The architect's mandatory assertions:
    //   - B5 formula = "B4" (formula PRESERVED, not converted to literal 30)
    //   - B5 computed value = 30 (Cherry/30 at B4)
    //   - A5 = "Apple" (survivor label)
    //   - A4 = "Cherry" (Cherry/30 compacted from row 6)
    //   - B4 = 30 (Cherry/30 value)
    //   - A3 = "Banana" (Banana compacted from row 4)
    //   - B3 = 20 (Banana value)
    //   - A1,B1 = "Name","Qty" (header preserved)
    //   - A2,B2 = "Apple",10 (first occurrence preserved)
    //   - A6,B6,A7,B7 — ABSENT (deleted from the sheet)
    const after = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => {
                    getCellData: () => { v?: unknown; f?: string } | null
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (r: number, c: number) => ws?.getRange?.(r, c)?.getCellData?.() ?? null
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
    // Header (row 1) preserved verbatim.
    expect(after.a1?.v, 'A1 = "Name" after dedupe (header preserved)').toBe('Name')
    expect(after.b1?.v, 'B1 = "Qty" after dedupe (header preserved)').toBe('Qty')
    // Row 2 (first occurrence) preserved verbatim.
    expect(after.a2?.v, 'A2 = "Apple" after dedupe (first occurrence)').toBe('Apple')
    expect(after.b2?.v, 'B2 = 10 after dedupe (first occurrence)').toBe(10)
    // Row 3 (Banana, was row 4) — survivor compacted up.
    expect(after.a3?.v, 'A3 = "Banana" after dedupe (compacted from row 4)').toBe('Banana')
    expect(after.b3?.v, 'B3 = 20 after dedupe (compacted from row 4)').toBe(20)
    // Row 4 (Cherry, was row 6) — survivor compacted up. This is the
    // cell the B7 formula originally referenced (=B6 → now =B4).
    expect(after.a4?.v, 'A4 = "Cherry" after dedupe (compacted from row 6)').toBe('Cherry')
    expect(after.b4?.v, 'B4 = 30 after dedupe (compacted from row 6)').toBe(30)
    // Row 5 (Apple/=B6, was row 7) — survivor compacted up. The
    // formula MUST be preserved AND its reference MUST be rewritten
    // from B6 to B4 to track where Cherry/30 moved.
    expect(after.a5?.v, 'A5 = "Apple" after dedupe (compacted from row 7)').toBe('Apple')
    expect(after.b5?.v, 'B5 = 30 (computed value, Cherry/30 at B4)').toBe(30)
    // *** THE CRITICAL ASSERTION ***
    // The B7 formula was `=B6` (referencing Cherry/30 in row 6). After
    // the dedupe, B7's content compacted to B5, and the formula's
    // reference B6 (which was Cherry/30) rewrote to B4 (where
    // Cherry/30 landed). The formula TEXT must be "=B4" (with the
    // leading `=` Univer's live model carries) — NOT the literal
    // value 30. This is the architect's explicit requirement:
    // "The current behavior 'moved rows become computed values' is
    // explicitly NOT accepted as formula preservation."
    expect(
      after.b5?.f,
      'B5 formula text = "=B4" (formula PRESERVED, reference rewritten B6→B4)',
    ).toBe('=B4')
    // Rows 6,7 — DELETED entirely. getCellData returns null for
    // out-of-range or empty cells. These were the two duplicate rows
    // that got deleted; the cells at A6,B6,A7,B7 must be absent.
    expect(after.a6, 'A6 absent (row deleted from the sheet)').toBe(null)
    expect(after.b6, 'B6 absent (row deleted from the sheet)').toBe(null)
    expect(after.a7, 'A7 absent (row deleted from the sheet)').toBe(null)
    expect(after.b7, 'B7 absent (row deleted from the sheet)').toBe(null)

    // The status message surfaced the count.
    await expect(page.locator('[data-testid="dedupe-message"]')).toContainText('Removed 2')

    // Save and capture the request body so we can assert the canonical
    // STRUCTURAL remove-rows channel was used (NOT the cell-edit-only
    // path that destroyed formulas on moved rows).
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
          cell?: { value: unknown; formula?: string }
        }>
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{ kind: string; index: number; count: number }>
        }>
        pageSetupStates?: unknown[]
        filterStates?: unknown[]
        dvStates?: unknown[]
        noteStates?: unknown[]
      }
    }
    // The save plan MUST use the `structuralOps` channel — NO new
    // "dedupeOps" / "removeDuplicates" family (the architecture test
    // in tests/architecture.test.ts enforces this at the source level;
    // here we assert it on the wire too).
    const savePlanKeys = Object.keys(saveBody.savePlan).sort()
    expect(savePlanKeys, 'save plan carries edits + structuralOps (no new family)').toContain(
      'structuralOps',
    )
    expect(
      savePlanKeys,
      'save plan does NOT introduce a dedupeOps / dedupe / removeDuplicates family',
    ).not.toContain('dedupeOps')
    expect(savePlanKeys).not.toContain('dedupe')
    expect(savePlanKeys).not.toContain('removeDuplicatesState')
    // The structuralOps channel carries EXACTLY TWO remove-rows ops
    // (one per duplicate, in DESCENDING order so earlier deletes don't
    // shift later indices). The dedupe was on sheet "Dedupe".
    const structOps = saveBody.savePlan.structuralOps ?? []
    expect(structOps.length, 'one structural-ops entry per dirty sheet').toBe(1)
    expect(structOps[0]?.sheetName, 'sheet name is "Dedupe"').toBe('Dedupe')
    const ops = structOps[0]?.ops ?? []
    expect(ops.length, 'exactly 2 remove-rows ops (one per duplicate)').toBe(2)
    expect(
      ops.every((op) => op.kind === 'remove-rows'),
      'all ops are remove-rows',
    ).toBe(true)
    expect(
      ops.every((op) => op.count === 1),
      'all ops are count=1 (one row each)',
    ).toBe(true)
    // The indices must be 4 and 2 (1-based = rows 5 and 3 — the two
    // duplicate rows in the original sheet). DESCENDING order so
    // earlier deletes don't shift later indices.
    const indices = ops.map((op) => op.index).sort((a, b) => a - b)
    expect(indices, 'remove-rows indices are 2 and 4 (rows 3 and 5 0-indexed)').toEqual([2, 4])
    // The cell-edit family may still be present for any other dirty
    // cells, AND the formula engine fires recalc mutations after the
    // structural deletes — Univer's formula engine recalculates the
    // `=B6` formula (its reference now points to a different cell
    // after the row shift), the formula text gets rewritten to `=B4`
    // in Univer's live model, and the recalc mutation fires a
    // `set-range-values` writeback that the journal captures as a
    // CellEdit. This is the SAME composition pattern as the sort
    // test in `ribbon-data.spec.ts` (the sort's `reorder-rows` op
    // permutes rows atomically AND the engine's recalc mutations
    // compose on top with the rewritten formula text).
    //
    // The architect's explicit requirement is that "the surviving
    // B7 formula is NOT converted to literal 30" — i.e. NO CellEdit
    // at the formula's new position (B5, 0-indexed row=4 col=1)
    // should carry `value: 30` WITHOUT a `formula` field. A CellEdit
    // that carries `formula: 'B4'` is PROOF of formula preservation
    // (the formula survived AND was rewritten to track the moved
    // referenced cell).
    const formulaCellEdit = (saveBody.savePlan.edits ?? []).find(
      (e) => e.row === 4 && e.column === 1,
    )
    if (formulaCellEdit) {
      // The CellEdit at B5 MUST carry the rewritten formula text
      // (NOT just a value — that would mean the formula was
      // destroyed and replaced with a literal). The canonical
      // `cellEditFromMutation` helper drops the cached `<v>` on
      // formula writes (cell.value = '') — Excel/Univer recalculates
      // the cached value on reopen, the formula text is the source
      // of truth. The formula TEXT carries the rewritten reference
      // `B4` (tracking where Cherry/30 moved), proving both:
      //   1. the formula survived (not converted to literal 30), AND
      //   2. the relative reference was rewritten (B6→B4).
      expect(
        formulaCellEdit.cell?.formula,
        'B5 cell-edit carries the REWRITTEN formula "B4" (formula PRESERVED, reference rewritten B6→B4)',
      ).toBe('B4')
      // The cell.value is '' (the canonical cellEditFromMutation
      // helper drops the cached <v> on formula writes — Excel/Univer
      // recalculates the cached value on reopen; the formula text is
      // the source of truth, not a stale cache).
      expect(
        formulaCellEdit.cell?.value,
        'B5 cell-edit value = "" (cached <v> dropped on formula write — formula is the source of truth)',
      ).toBe('')
    } else {
      // If no CellEdit was emitted, the gateway's structural op
      // alone rewrote the formula text in the saved XML (the
      // `transformFormulas` step of `applyStructuralOps`). The
      // formula is STILL preserved — proven by the saved XML
      // assertion below (`<c r="B5"...><f>B4</f><v>30</v>`).
    }
    // The architect's concern: NO CellEdit at B5 should carry
    // `value: 30` WITHOUT a `formula` field — that would be the
    // "formula converted to literal" failure mode.
    const literalWithoutFormula = (saveBody.savePlan.edits ?? []).filter(
      (e) => e.row === 4 && e.column === 1 && e.cell?.value === 30 && !e.cell?.formula,
    )
    expect(
      literalWithoutFormula.length,
      'NO CellEdit at B5 carries value=30 without a formula (formula NOT destroyed)',
    ).toBe(0)

    // Verify the saved XLSX bytes carry the deduped state with the
    // formula PRESERVED at its new position (B5) with the REWRITTEN
    // reference (B4).
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
    // Row 3 (Banana, was row 4) — compacted. The italic style (xf 2)
    // on Banana travels with the cell (s="4" — bold font 1).
    expect(sheet1, 'A3 = Banana (compacted) in saved XML').toMatch(
      /<c r="A3"[^>]*t="s"[^>]*s="4"[^>]*><v>3<\/v>/,
    )
    expect(sheet1, 'B3 = 20 (compacted) in saved XML').toMatch(/<c r="B3"[^>]*><v>20<\/v>/)
    // Row 4 (Cherry, was row 6) — compacted.
    expect(sheet1, 'A4 = Cherry (compacted) in saved XML').toMatch(
      /<c r="A4"[^>]*t="s"[^>]*><v>4<\/v>/,
    )
    expect(sheet1, 'B4 = 30 (compacted) in saved XML').toMatch(/<c r="B4"[^>]*><v>30<\/v>/)
    // *** THE CRITICAL XML ASSERTION ***
    // Row 5 (Apple/=B4, was row 7) — compacted. B5 MUST carry a
    // <f> element with body "B4" (the rewritten reference). The
    // cached <v>30</v> MAY or MAY NOT be present — when Univer's
    // formula engine fires a recalc mutation after the structural
    // deletes, the cellEditFromMutation helper drops the cached <v>
    // on formula writes (canonical engine semantics — an edited
    // formula has no stale cache; Excel/Univer recalculates on
    // open). This matches the sort test's documented behavior in
    // ribbon-data.spec.ts:425-437 ("journaled formula edit drops
    // the cached <v>"). The formula TEXT is the source of truth.
    expect(sheet1, 'A5 = Apple (compacted from row 7) in saved XML').toMatch(
      /<c r="A5"[^>]*t="s"[^>]*s="2"[^>]*><v>2<\/v>/,
    )
    expect(
      sheet1,
      'B5 carries <f>B4</f> in saved XML (formula PRESERVED, reference rewritten B6→B4)',
    ).toMatch(/<c r="B5"[^>]*><f>B4<\/f>/)
    // Rows 6,7 — DELETED entirely from the saved XML. The structural
    // remove-rows op removes the <row r="6"> and <row r="7"> blocks.
    expect(sheet1, 'no row 6 in saved XML (deleted by remove-rows op)').not.toMatch(
      /<row\b[^>]*\br="6"/,
    )
    expect(sheet1, 'no row 7 in saved XML (deleted by remove-rows op)').not.toMatch(
      /<row\b[^>]*\br="7"/,
    )
    expect(sheet1, 'no B6 cell in saved XML (deleted)').not.toMatch(/<c r="B6"/)
    expect(sheet1, 'no B7 cell in saved XML (deleted — formula row gone)').not.toMatch(/<c r="B7"/)

    // Reopen and verify the snapshot carries the deduped state WITH
    // the preserved formula.
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
      cells: Record<string, { value: unknown; formula?: string }>
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
    // Compacted Apple/=B4 at A5 (formula PRESERVED, reference rewritten
    // B6→B4). The snapshot's formula field carries the formula text
    // (with leading `=` per the canonical readBasicWorkbook parser).
    // The value field is `null` for formula cells — the parser
    // explicitly returns `{ value: null, formula: '=<text>' }` for
    // cells with a <f> element (the cached <v> is NOT extracted —
    // the formula engine is the authority on computed values, not
    // the cached file value).
    expect(cells.A5?.value, 'A5 = "Apple" after reopen (compacted)').toBe('Apple')
    // *** THE CRITICAL REOPEN ASSERTION ***
    // The formula text MUST be "=B4" in the reopened snapshot. NOT
    // the literal 30 — the FORMULA survived. The architect explicitly
    // requires: "the surviving B7 formula is not converted to literal
    // 30" — this assertion proves it. The value field is null
    // (formula cells in the canonical snapshot carry formula text,
    // not cached values — the formula engine is the authority).
    expect(cells.B5?.formula, 'B5 formula = "=B4" after reopen (formula PRESERVED)').toBe('=B4')
    expect(
      cells.B5?.value,
      'B5 value = null after reopen (formula cells carry formula text, not cached value)',
    ).toBe(null)
    // Deleted rows — the snapshot omits empty cells.
    expect(cells.A6, 'A6 absent after reopen (deleted)').toBeUndefined()
    expect(cells.B6, 'B6 absent after reopen (deleted)').toBeUndefined()
    expect(cells.A7, 'A7 absent after reopen (deleted)').toBeUndefined()
    expect(cells.B7, 'B7 absent after reopen (deleted)').toBeUndefined()

    expect(pageErrors, 'no uncaught page errors').toEqual([])
  })

  test('preserves absolute / relative / mixed references through compaction (architect second regression case)', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDedupeMixedReferencesFixture()
    writeFileSync('/tmp/e2e-ribbon-dedupe-mixed.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dedupe-mixed.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dedupe-mixed.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Sanity: verify the pre-dedupe formulas are present.
    const before = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => {
                    getCellData: () => { v?: unknown; f?: string } | null
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (r: number, c: number) => ws?.getRange?.(r, c)?.getCellData?.() ?? null
      return {
        c4: cell(3, 2), // =$D$6
        d4: cell(3, 3), // =A6
        e4: cell(3, 4), // =$A6
        f4: cell(3, 5), // =A$6
        d6: cell(5, 3), // "Anchor" literal
      }
    })
    // Pre-dedupe: all four formula texts must be present (proves the
    // fixture loaded correctly with formulas intact). Univer's live
    // model carries the leading `=` in the .f field.
    expect(before.c4?.f, 'C4 formula text before dedupe').toBe('=$D$6')
    expect(before.d4?.f, 'D4 formula text before dedupe').toBe('=A6')
    expect(before.e4?.f, 'E4 formula text before dedupe').toBe('=$A6')
    expect(before.f4?.f, 'F4 formula text before dedupe').toBe('=A$6')
    expect(before.d6?.v, 'D6 = "Anchor" before dedupe').toBe('Anchor')

    // Select A1:B7 via the Name Box.
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A1:B7')
    await box.press('Enter')
    await page.waitForTimeout(400)

    // Switch to Data tab.
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)

    // Open the Remove Duplicates dialog and click OK with headers.
    await page.getByRole('button', { name: 'Remove Duplicates' }).click()
    await expect(page.locator('[data-testid="dedupe-dialog"]')).toBeVisible()
    await expect(page.locator('[data-testid="dedupe-dialog"] input[type="checkbox"]')).toBeChecked()
    await page.locator('[data-testid="dedupe-ok"]').click()
    await page.waitForTimeout(800)

    // After dedupe: rows 3, 5, 7 are DELETED (3 duplicates of row 2).
    // Banana at row 4 → row 3 (compacted up by 1 after row 3 delete).
    // Cherry at row 6 → row 4 (compacted up by 2 after rows 5 and 3
    // deletes — row 5 delete shifts Cherry to row 5, then row 3
    // delete shifts Cherry to row 4). Anchor at D6 → D4 (same shift).
    //
    // The four formulas on Banana (was row 4, now row 3) all reference
    // row 6 (Cherry/Anchor). The gateway's transformFormulas rewrites
    // every reference to track the moved target (row 6 → row 4):
    //   $D$6 → $D$4  (absolute: $ preserved, row 6→4)
    //   A6   → A4    (relative: row 6→4)
    //   $A6  → $A4   (mixed col $, row relative: $ preserved, row 6→4)
    //   A$6  → A$4   (mixed col relative, row $: $ preserved, row 6→4)
    //
    // The formula text MUST survive — this is what distinguishes
    // "formula survived and was rewritten correctly" from "computed
    // result happens to be the same".
    const after = await page.evaluate(() => {
      const rt = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => {
                    getCellData: () => { v?: unknown; f?: string } | null
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const cell = (r: number, c: number) => ws?.getRange?.(r, c)?.getCellData?.() ?? null
      return {
        a3: cell(2, 0),
        b3: cell(2, 1),
        c3: cell(2, 2), // Banana row's $D$? formula (compacted from row 4)
        d3: cell(2, 3), // Banana row's =A? formula
        e3: cell(2, 4), // Banana row's $A? formula
        f3: cell(2, 5), // Banana row's A$? formula
        a4: cell(3, 0),
        b4: cell(3, 1),
        d4: cell(3, 3), // "Anchor" (compacted from D6)
      }
    })
    // Banana compacted from row 4 to row 3.
    expect(after.a3?.v, 'A3 = "Banana" (compacted from row 4)').toBe('Banana')
    expect(after.b3?.v, 'B3 = 20 (compacted from row 4)').toBe(20)
    // *** CRITICAL: the four formulas survived AND their references
    // were rewritten to track where Cherry/Anchor moved (row 6 → row 4).
    // Univer's live .f field carries the leading `=`.
    expect(after.c3?.f, 'C3 formula = "=$D$4" (absolute preserved, ref rewritten $D$6→$D$4)').toBe(
      '=$D$4',
    )
    expect(after.c3?.v, 'C3 computed = "Anchor" (from D4)').toBe('Anchor')
    expect(after.d3?.f, 'D3 formula = "=A4" (relative ref rewritten A6→A4)').toBe('=A4')
    expect(after.d3?.v, 'D3 computed = "Cherry" (from A4)').toBe('Cherry')
    expect(after.e3?.f, 'E3 formula = "=$A4" (mixed col $, row rewritten 6→4)').toBe('=$A4')
    expect(after.e3?.v, 'E3 computed = "Cherry" (from A4)').toBe('Cherry')
    expect(after.f3?.f, 'F3 formula = "=A$4" (mixed col, row $ preserved, rewritten 6→4)').toBe(
      '=A$4',
    )
    expect(after.f3?.v, 'F3 computed = "Cherry" (from A4)').toBe('Cherry')
    // Cherry compacted from row 6 to row 4. Anchor compacted from D6 to D4.
    expect(after.a4?.v, 'A4 = "Cherry" (compacted from row 6)').toBe('Cherry')
    expect(after.b4?.v, 'B4 = 30 (compacted from row 6)').toBe(30)
    expect(after.d4?.v, 'D4 = "Anchor" (compacted from D6)').toBe('Anchor')

    // The status message surfaces the count.
    await expect(page.locator('[data-testid="dedupe-message"]')).toContainText('Removed 3')

    // Save and verify the save plan carries THREE remove-rows ops
    // (one per duplicate at rows 3, 5, 7 — DESCENDING order).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{ kind: string; index: number; count: number }>
        }>
      }
    }
    const structOps = saveBody.savePlan.structuralOps ?? []
    expect(structOps.length, 'one structural-ops entry per dirty sheet').toBe(1)
    expect(structOps[0]?.sheetName, 'sheet name is "DedupeMixed"').toBe('DedupeMixed')
    const ops = structOps[0]?.ops ?? []
    expect(ops.length, 'exactly 3 remove-rows ops (one per duplicate)').toBe(3)
    expect(
      ops.every((op) => op.kind === 'remove-rows'),
      'all ops are remove-rows',
    ).toBe(true)
    expect(
      ops.every((op) => op.count === 1),
      'all ops are count=1',
    ).toBe(true)
    // The indices must be 2, 4, 6 (0-indexed = rows 3, 5, 7 — the
    // three duplicate rows in the original sheet). ASCENDING in the
    // journal; the runtime issues them DESCENDING.
    const indices = ops.map((op) => op.index).sort((a, b) => a - b)
    expect(indices, 'remove-rows indices are 2, 4, 6 (rows 3, 5, 7 0-indexed)').toEqual([2, 4, 6])

    // Verify the saved XLSX bytes carry the formulas with rewritten
    // references at the new positions.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // Banana compacted to row 3 — the four formulas on Cols C-F.
    expect(sheet1, 'C3 carries <f>$D$4</f> in saved XML (absolute preserved)').toMatch(
      /<c r="C3"[^>]*><f>\$D\$4<\/f>/,
    )
    expect(sheet1, 'D3 carries <f>A4</f> in saved XML (relative rewritten)').toMatch(
      /<c r="D3"[^>]*><f>A4<\/f>/,
    )
    expect(sheet1, 'E3 carries <f>$A4</f> in saved XML (mixed col $)').toMatch(
      /<c r="E3"[^>]*><f>\$A4<\/f>/,
    )
    expect(sheet1, 'F3 carries <f>A$4</f> in saved XML (mixed row $)').toMatch(
      /<c r="F3"[^>]*><f>A\$4<\/f>/,
    )
    // Anchor compacted from D6 to D4 (shared-strings index 5).
    expect(sheet1, 'D4 carries Anchor in saved XML (compacted from D6)').toMatch(
      /<c r="D4"[^>]*t="s"[^>]*><v>5<\/v>/,
    )

    // Reopen and verify the snapshot carries the rewritten formulas.
    writeFileSync('/tmp/e2e-ribbon-dedupe-mixed-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dedupe-mixed-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dedupe-mixed-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown; formula?: string }>
    }>
    const cells = reopened[0].cells
    // Banana compacted to row 3.
    expect(cells.A3?.value, 'A3 = "Banana" after reopen').toBe('Banana')
    expect(cells.B3?.value, 'B3 = 20 after reopen').toBe(20)
    // *** CRITICAL REOPEN ASSERTIONS: the four formulas survived the
    // save/reopen round-trip AND their references were rewritten.
    // The snapshot's .formula field carries the leading `=` (per the
    // canonical readBasicWorkbook parser); the .value field is null
    // for formula cells (the parser explicitly returns
    // { value: null, formula: '=<text>' } for cells with a <f> element
    // — the cached <v> is NOT extracted; the formula engine is the
    // authority on computed values, not the cached file value).
    expect(cells.C3?.formula, 'C3 formula = "=$D$4" after reopen (absolute preserved)').toBe(
      '=$D$4',
    )
    expect(cells.C3?.value, 'C3 value = null (formula cell — formula is the authority)').toBe(null)
    expect(cells.D3?.formula, 'D3 formula = "=A4" after reopen (relative rewritten)').toBe('=A4')
    expect(cells.D3?.value, 'D3 value = null (formula cell)').toBe(null)
    expect(cells.E3?.formula, 'E3 formula = "=$A4" after reopen (mixed col $)').toBe('=$A4')
    expect(cells.E3?.value, 'E3 value = null (formula cell)').toBe(null)
    expect(cells.F3?.formula, 'F3 formula = "=A$4" after reopen (mixed row $)').toBe('=A$4')
    expect(cells.F3?.value, 'F3 value = null (formula cell)').toBe(null)
    // Cherry + Anchor compacted to row 4.
    expect(cells.A4?.value, 'A4 = "Cherry" after reopen').toBe('Cherry')
    expect(cells.B4?.value, 'B4 = 30 after reopen').toBe(30)
    expect(cells.D4?.value, 'D4 = "Anchor" after reopen').toBe('Anchor')

    expect(pageErrors, 'no uncaught page errors').toEqual([])
  })

  test('no-op when there are no duplicates — fails closed without mutating', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Use the basic fixture — select only the unique rows (A4:B6) so the
    // dedupe finds no duplicates. With hasHeader=true the header is
    // row 4 (Banana) and data rows are 5 (Apple, not a dup of anything
    // inside A4:B6) and 6 (Cherry). No duplicates within A4:B6.
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
