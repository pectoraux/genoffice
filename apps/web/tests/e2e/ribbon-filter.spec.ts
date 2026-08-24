/**
 * REAL browser E2E — Data → Filter (Phase 4 Increment 4).
 *
 * Proves the AutoFilter persistence chain end-to-end through the REAL
 * HTTP boundary:
 *
 *   open → read <autoFilter> → render in the real Univer UI
 *   → user applies/toggles a filter (real Univer commands)
 *   → filter mutations mark the sheet filter-dirty
 *   → save snapshots the LIVE filter model as a typed SheetFilterState
 *   → savePlan.filterStates → /api/office/workbooks/save → routeOffice
 *   → xlsx-gateway applyFilterState → XLSX bytes
 *   → reopen → filterState + row visibility + criteria survive
 *
 * No browser-side OOXML. No JS-side XLSX rewriting. The browser only ever
 * sends typed SheetFilterState snapshots taken from Univer's live model.
 *
 * Fixture (buildExcelFilterFixture / buildExcelFilteredFixture): a 7-row
 * produce table with mixed styles (bold header, italic A2, currency numfmt
 * on C), relative formulas on D (=C{row}*2), one blank row (row 5), and a
 * worksheet hyperlink.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelFilterFixture, buildExcelFilteredFixture, readZipEntry } from './fixtures'

/** Type-only view of the exposed runtime (all public Univer facades). */
type FilterRuntime = {
  univerAPI: {
    getActiveWorkbook: () => {
      getActiveSheet: () => {
        getSheetName: () => string
        getFilter: () => {
          getRange: () => {
            getRow: () => number
            getColumn: () => number
            getHeight: () => number
            getWidth: () => number
          }
          getColumnFilterCriteria: (column: number) => {
            colId: number
            filters?: { blank?: true; filters?: string[] }
            customFilters?: {
              and?: number
              customFilters: { val: string | number; operator?: string }[]
            }
          } | null
          setColumnFilterCriteria: (
            column: number,
            criteria: {
              colId: number
              filters?: { blank?: true; filters?: string[] }
              customFilters?: {
                and?: number
                customFilters: { val: string | number; operator?: string }[]
              }
            },
          ) => unknown
          getFilteredOutRows: () => number[]
          remove: () => boolean
        } | null
        getRange: (
          row: number,
          col: number,
          numRows: number,
          numCols: number,
        ) => {
          createFilter: () => unknown
        }
      }
    }
  }
}

/**
 * Plain-data view of the live filter model (returned from page.evaluate —
 * functions cannot cross the boundary, so all facade access happens inside
 * the evaluate callback). Null when the sheet has no filter.
 */
interface LiveFilterView {
  row: number
  column: number
  height: number
  width: number
  columns: Array<{
    colId: number
    values?: string[]
    blank?: boolean
    customs?: { and?: boolean; filters: { val: string | number; operator?: string }[] }
  }>
  filteredOutRows: number[]
}

function readLiveFilter(page: import('@playwright/test').Page): Promise<LiveFilterView | null> {
  return page.evaluate(() => {
    const runtime = (window as { __genofficeExcelRuntime?: FilterRuntime }).__genofficeExcelRuntime
    const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
    const filter = ws?.getFilter?.()
    if (!filter) return null
    const range = filter.getRange()
    const columns: LiveFilterView['columns'] = []
    for (
      let column = range.getColumn();
      column < range.getColumn() + range.getWidth();
      column += 1
    ) {
      const criteria = filter.getColumnFilterCriteria(column)
      if (!criteria) continue
      columns.push({
        colId: column - range.getColumn(),
        ...(criteria.filters?.filters ? { values: [...criteria.filters.filters] } : {}),
        ...(criteria.filters?.blank ? { blank: true } : {}),
        ...(criteria.customFilters
          ? {
              customs: {
                ...(criteria.customFilters.and === 1 ? { and: true } : {}),
                filters: criteria.customFilters.customFilters.map(
                  (custom: { val: string | number; operator?: string }) => ({
                    val: custom.val,
                    ...(custom.operator !== undefined ? { operator: custom.operator } : {}),
                  }),
                ),
              },
            }
          : {}),
      })
    }
    return {
      row: range.getRow(),
      column: range.getColumn(),
      height: range.getHeight(),
      width: range.getWidth(),
      columns,
      filteredOutRows: filter.getFilteredOutRows(),
    }
  })
}

test.describe('Data tab — Filter persists through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1: opening a workbook with an existing <autoFilter> renders it in the real Univer model', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // The filtered variant carries Category="Fruit" with the Veg rows hidden.
    const fixture = await buildExcelFilteredFixture()
    writeFileSync('/tmp/e2e-ribbon-filter-open.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-open.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-open.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // The READ path: the snapshot carried the parsed filter state.
    const snapshot = (await openResponse.json()).snapshot.sheets as Array<{
      name: string
      filterState?: {
        filter: { range: unknown; columns: unknown[] } | null
        hiddenRows: number[]
      }
    }>
    expect(snapshot[0].filterState, 'snapshot carries the parsed filter state').toBeDefined()
    expect(snapshot[0].filterState!.filter!.columns).toEqual([{ colId: 0, values: ['Fruit'] }])
    expect(snapshot[0].filterState!.hiddenRows).toEqual([2, 4, 6])

    // The RENDER path: the live Univer model carries the filter.
    await page.waitForTimeout(1500)
    const live = await readLiveFilter(page)
    expect(live, 'the live Univer model has the filter').not.toBeNull()
    expect(live!.row).toBe(0)
    expect(live!.column).toBe(0)
    expect(live!.height).toBe(8)
    expect(live!.width).toBe(4)
    // The criteria round-tripped into the live model.
    expect(live!.columns).toEqual([{ colId: 0, values: ['Fruit'] }])
    // The recalculated hidden rows match the file's Veg rows (0-based 2, 4, 6).
    expect(live!.filteredOutRows).toEqual([2, 4, 6])

    expect(pageErrors).toEqual([])
  })

  test('2: applying a value filter hides the non-matching rows in-session', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFilterFixture()
    writeFileSync('/tmp/e2e-ribbon-filter-apply.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-apply.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-apply.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Select the table range and toggle the AutoFilter through the ribbon
    // (the REAL smart-toggle-filter command).
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A1:D8')
    await box.press('Enter')
    await page.waitForTimeout(400)
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page
      .getByRole('button', { name: /AutoFilter/i })
      .first()
      .click()
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Apply Category = "Fruit" through the REAL filter command — the same
    // SetSheetsFilterCriteriaCommand the UI dropdown executes (and the same
    // facade path the desktop's Advanced Filter dialog drives).
    await page.evaluate(() => {
      const runtime = (window as { __genofficeExcelRuntime?: FilterRuntime })
        .__genofficeExcelRuntime
      const ws = runtime?.univerAPI.getActiveWorkbook().getActiveSheet()
      const filter = ws?.getFilter()
      if (!filter) throw new Error('filter missing after toggle')
      filter.setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['Fruit'] } })
    })
    await page.waitForTimeout(800)

    // The live model: Fruit rows visible, Veg + blank rows filtered out.
    // Veg rows are 0-based 2 (Carrot), 4 (Pea — blank row), 6 (Kale).
    const live = await readLiveFilter(page)
    expect(live, 'the live filter model exists').not.toBeNull()
    expect(live!.columns).toEqual([{ colId: 0, values: ['Fruit'] }])
    expect(live!.filteredOutRows).toEqual([2, 4, 6])

    expect(pageErrors).toEqual([])
  })

  test('3: save + reopen — typed filter state on the wire, criteria + hidden rows + values + styles + formulas in the XLSX', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFilterFixture()
    writeFileSync('/tmp/e2e-ribbon-filter-save.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-save.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-save.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Toggle + apply Category="Fruit" (same as Test 2).
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('A1:D8')
    await box.press('Enter')
    await page.waitForTimeout(400)
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page
      .getByRole('button', { name: /AutoFilter/i })
      .first()
      .click()
    await page.waitForTimeout(600)
    await page.evaluate(() => {
      const runtime = (window as { __genofficeExcelRuntime?: FilterRuntime })
        .__genofficeExcelRuntime
      const ws = runtime?.univerAPI.getActiveWorkbook().getActiveSheet()
      const filter = ws?.getFilter()
      filter?.setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['Fruit'] } })
    })
    await page.waitForTimeout(800)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: capture the request — the plan must carry the TYPED filter state.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        filterStates?: Array<{
          sheetName: string
          filter: {
            range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
            columns: Array<{ colId: number; values?: string[] }>
          } | null
          hiddenRows: number[]
          visibilityRange: { startRow: number; endRow: number }
        }>
      }
    }
    expect(saveBody.savePlan.filterStates, 'typed filter state on the wire').toHaveLength(1)
    const state = saveBody.savePlan.filterStates![0]!
    expect(state.sheetName).toBe('Produce')
    expect(state.filter!.range).toEqual({
      startRow: 0,
      endRow: 7,
      startColumn: 0,
      endColumn: 3,
    })
    expect(state.filter!.columns).toEqual([{ colId: 0, values: ['Fruit'] }])
    expect(state.hiddenRows).toEqual([2, 4, 6])
    expect(state.visibilityRange).toEqual({ startRow: 0, endRow: 7, startColumn: 0, endColumn: 3 })

    // Saved XML: filter range + criteria + hidden rows + values + styles +
    // formulas + hyperlink ALL survive.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<autoFilter ref="A1:D8">')
    expect(sheet1).toContain(
      '<filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn>',
    )
    // The Veg rows (1-based 3, 5, 7) are hidden.
    expect(sheet1).toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(sheet1).toMatch(/<row r="5"[^>]*hidden="1"/)
    expect(sheet1).toMatch(/<row r="7"[^>]*hidden="1"/)
    // The Fruit rows stay visible with their content.
    expect(sheet1).toMatch(/<c r="A2" t="s" s="2"><v>4<\/v>/)
    expect(sheet1).toMatch(/<c r="C2" s="3"><v>10<\/v>/)
    // Formulas survive (journaled formula edits drop the stale cached v —
    // canonical semantics; Excel recalculates on open).
    expect(sheet1).toMatch(/<c r="D2"[^>]*><f>C2\*2<\/f>/)
    expect(sheet1).toMatch(/<c r="D8"[^>]*><f>C8\*2<\/f>/)
    // The worksheet hyperlink definition survives.
    expect(sheet1).toContain('<hyperlink ref="B2" r:id="rId1"/>')

    // Reopen through the REAL /open route: the filter state + hidden rows
    // + unrelated content all come back.
    writeFileSync('/tmp/e2e-ribbon-filter-save-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-save-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-save-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown; formula?: string }>
      styles?: Record<string, { bold?: boolean; italic?: boolean; numberFormat?: string }>
      filterState?: {
        filter: { range: unknown; columns: unknown[] } | null
        hiddenRows: number[]
      }
    }>
    expect(reopened[0].name).toBe('Produce')
    // Filter criteria + hidden rows survive the reopen.
    expect(reopened[0].filterState?.filter?.columns).toEqual([{ colId: 0, values: ['Fruit'] }])
    expect(reopened[0].filterState?.hiddenRows).toEqual([2, 4, 6])
    // Unrelated cell values survive.
    expect(reopened[0].cells.A1?.value).toBe('Category')
    expect(reopened[0].cells.A2?.value).toBe('Fruit')
    expect(reopened[0].cells.C2?.value).toBe(10)
    expect(reopened[0].cells.C8?.value).toBe(30)
    // Formulas survive.
    expect(reopened[0].cells.D2?.formula?.replace(/\s+/g, '')).toBe('=C2*2')
    // Styles survive.
    expect(reopened[0].styles?.A1?.bold).toBe(true)
    expect(reopened[0].styles?.A2?.italic).toBe(true)
    expect(reopened[0].styles?.C2?.numberFormat).toMatch(/\$/)

    // The reopened live model re-renders the filter (range + criteria).
    await page.waitForTimeout(1500)
    const live = await readLiveFilter(page)
    expect(live, 'reopened workbook carries the filter in the live model').not.toBeNull()
    expect(live!.columns).toEqual([{ colId: 0, values: ['Fruit'] }])
    expect(live!.filteredOutRows).toEqual([2, 4, 6])

    expect(pageErrors).toEqual([])
  })

  test('4: clearing the filter removes the autoFilter and restores the hidden rows', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFilteredFixture()
    writeFileSync('/tmp/e2e-ribbon-filter-clear.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-clear.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-clear.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Toggle the filter OFF through the ribbon (smart-toggle removes it).
    await page
      .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Data' })
      .click()
    await page.waitForTimeout(200)
    await page
      .getByRole('button', { name: /AutoFilter/i })
      .first()
      .click()
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // The live model no longer has a filter.
    const live = await readLiveFilter(page)
    expect(live, 'filter removed from the live model').toBeNull()

    // Save: the plan carries the explicit cleared state (filter: null).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        filterStates?: Array<{ filter: unknown | null; hiddenRows: number[] }>
      }
    }
    expect(saveBody.savePlan.filterStates).toHaveLength(1)
    expect(saveBody.savePlan.filterStates![0]!.filter).toBeNull()
    expect(saveBody.savePlan.filterStates![0]!.hiddenRows).toEqual([])

    // Saved XML: autoFilter GONE, every previously hidden row restored.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).not.toContain('autoFilter')
    expect(sheet1).toMatch(/<row r="3"(?![^>]*hidden)[^>]*>/)
    expect(sheet1).toMatch(/<row r="5"(?![^>]*hidden)[^>]*>/)
    expect(sheet1).toMatch(/<row r="7"(?![^>]*hidden)[^>]*>/)

    // Reopen: no filter state, all rows visible.
    writeFileSync('/tmp/e2e-ribbon-filter-clear-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-clear-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-clear-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      cells: Record<string, { value: unknown }>
      filterState?: unknown
    }>
    expect(reopened[0].filterState).toBeUndefined()
    // The previously hidden Veg rows' values are back.
    expect(reopened[0].cells.A3?.value).toBe('Veg')
    expect(reopened[0].cells.A5?.value).toBe('Veg')
    expect(reopened[0].cells.A7?.value).toBe('Veg')

    expect(pageErrors).toEqual([])
  })

  test('5: no-op save (unrelated edit only) preserves the filter XML byte-for-byte', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFilteredFixture()
    writeFileSync('/tmp/e2e-ribbon-filter-noop.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-filter-noop.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-filter-noop.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // An UNRELATED edit (G1, outside the filter range) enables Save; the
    // filter itself is never touched. Typed through the REAL facade.
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
        ?.getActiveWorkbook()
        ?.getActiveSheet()
        ?.getRange(0, 6)
        ?.setValueForCell?.('note')
    })
    await page.waitForTimeout(500)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the plan carries the cell edit but NO filter state (the sheet
    // is not filter-dirty).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { filterStates?: unknown[] }
    }
    expect(
      saveBody.savePlan.filterStates,
      'no filter state on a no-filter-change save',
    ).toBeUndefined()

    // The saved XML preserves the file's own autoFilter + hidden rows.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain(
      '<autoFilter ref="A1:D8"><filterColumn colId="0"><filters><filter val="Fruit"/></filters></filterColumn></autoFilter>',
    )
    expect(sheet1).toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(sheet1).toMatch(/<row r="5"[^>]*hidden="1"/)
    expect(sheet1).toMatch(/<row r="7"[^>]*hidden="1"/)

    expect(pageErrors).toEqual([])
  })
})
