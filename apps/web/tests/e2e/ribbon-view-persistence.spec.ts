/**
 * REAL browser E2E — EXCEL-026 View / Page Layout persistence.
 *
 * Proves the full view-state surface through the canonical pipeline
 * (browser → Vite → HTTP → routeOffice → xlsx-gateway → XLSX bytes):
 *
 *   1.  existing freeze state imports correctly
 *   2.  freeze edit survives save/reopen
 *   3.  freeze clear/reset survives save/reopen (the EXCEL-026 defect fix)
 *   4.  existing gridline state imports correctly
 *   5.  gridline toggle persists
 *   6.  gridline toggle back restores the original state
 *   7.  existing formula-view state renders + toggles correctly
 *   8.  formula-view toggle persists
 *   9.  unrelated cell edits do not corrupt view state
 *   10. a save that does not touch view state preserves the view XML
 *       byte-for-byte
 *   11. Page Layout print family (orientation/margins/size/fit) persists
 *
 * Mechanism under test:
 *   gridlines → FWorksheet.setHiddenGridlines → sheet.mutation
 *   .toggle-gridlines → pageSetupRef journal → savePlan.pageSetupStates
 *   → applyPageSetupState writes sheetView@showGridLines
 *   formula view → shell flag set + RENDER_RAW_FORMULA_KEY +
 *   CELL_CONTENT interceptor → pageSetupRef journal → sheetView@showFormulas
 *   freeze → sheet.mutation.set-frozen (0/0 journals the CLEAR) → <pane>
 */
import { test, expect, type Page } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelViewFixture, readZipEntry } from './fixtures'

const GRID = '#genoffice-web-excel canvas'

type ViewSnapshot = {
  sheets: Array<{
    name: string
    freeze?: { frozenRows: number; frozenColumns: number }
    view?: { showGridlines?: boolean; showFormulas?: boolean; showHeadings?: boolean }
    cells?: Record<string, { value?: unknown }>
  }>
}

/** Open a fixture file and return the open-API snapshot. */
async function openFixture(page: Page, bytes: Buffer, fileName: string): Promise<ViewSnapshot> {
  writeFileSync(`/tmp/${fileName}`, bytes)
  const openResponsePromise = page.waitForResponse(
    (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
  )
  await page.setInputFiles('input[type="file"]', `/tmp/${fileName}`)
  await expect(page.getByText(`Opened ${fileName}`)).toBeVisible({ timeout: 30_000 })
  const openResponse = await openResponsePromise
  expect(openResponse.status()).toBe(200)
  await page.waitForTimeout(1200)
  return ((await openResponse.json()).snapshot ?? { sheets: [] }) as ViewSnapshot
}

/** Save + return { savedBytes, savePlan } from the captured request. */
async function saveAndCapturePlan(
  page: Page,
): Promise<{ bytes: Buffer; plan: Record<string, unknown> }> {
  const saveReq = page.waitForRequest(
    (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
  )
  const bytes = await clickSaveAndCaptureDownload(page, 'Save')
  const req = await saveReq
  const body = JSON.parse(req.postData() ?? '{}') as { savePlan?: Record<string, unknown> }
  return { bytes, plan: body.savePlan ?? {} }
}

/** Read the live freeze state from the exposed runtime. */
function liveFreeze(page: Page): Promise<{ startRow: number; startColumn: number } | null> {
  return page.evaluate(() => {
    const rt = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook?: () => {
              getActiveSheet?: () => {
                getFreeze?: () => { startRow: number; startColumn: number }
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getFreeze?.() ?? null
  })
}

/** Read the live gridline visibility from the exposed runtime. */
function liveGridlinesHidden(page: Page): Promise<boolean | null> {
  return page.evaluate(() => {
    const rt = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook?: () => {
              getActiveSheet?: () => { hasHiddenGridLines?: () => boolean }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.hasHiddenGridLines?.() ?? null
  })
}

/** Screenshot a fixed clip of the grid's top-left region (row 1 cells). */
async function gridClip(page: Page): Promise<Buffer> {
  const grid = page.locator(GRID).last()
  const bbox = await grid.boundingBox()
  if (!bbox) throw new Error('grid canvas not found')
  return page.screenshot({ clip: { x: bbox.x, y: bbox.y, width: 620, height: 140 } })
}

/** Double-click a cell (0-based row/column), select-all, type, Enter. */
async function editCell(page: Page, row: number, column: number, text: string): Promise<void> {
  const grid = page.locator(GRID).last()
  const bbox = await grid.boundingBox()
  if (!bbox) throw new Error('grid canvas not found')
  const xBefore = column === 0 ? 0 : 173 + (column - 1) * 100
  const cellX = bbox.x + 46 + xBefore + (column === 0 ? 173 : 100) / 2
  const cellY = bbox.y + 20 + row * 20 + 10
  await page.mouse.dblclick(cellX, cellY)
  await page.waitForTimeout(400)
  await page.keyboard.press('Control+a')
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
  await page.waitForTimeout(700)
}

const CANONICAL_PANE =
  '<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/>'

test.describe('EXCEL-026 — view persistence (freeze/gridlines/formula view)', () => {
  test('1+2: existing freeze imports; freeze edit survives save/reopen', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // ── Scenario 1: the file's <pane> imports into the live model. ──────
    const fixture = await buildExcelViewFixture(
      `<sheetView workbookViewId="0">${CANONICAL_PANE}</sheetView>`,
    )
    const snapshot = await openFixture(page, fixture, 'e2e-view-frozen.xlsx')
    const data = snapshot.sheets.find((s) => s.name === 'Data')
    expect(data?.freeze, 'snapshot carries the file freeze').toEqual({
      frozenRows: 2,
      frozenColumns: 2,
    })
    const freeze = await liveFreeze(page)
    expect(freeze?.startRow, 'live model froze 2 rows').toBe(2)
    expect(freeze?.startColumn, 'live model froze 2 columns').toBe(2)

    // ── Scenario 2 (edit path): clear at the frozen cell, re-freeze, save. ──
    // Navigate to C3 (the frozen anchor) and toggle — this clears. Then
    // toggle again to re-freeze at C3; the journal re-emits the freeze.
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('C3')
    await box.press('Enter')
    await page.waitForTimeout(400)
    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Freeze Panes' }).click()
    await page.waitForTimeout(400)
    const cleared = await liveFreeze(page)
    expect(cleared?.startRow, 'freeze cleared in-session').toBe(-1)
    await page.getByRole('button', { name: 'Freeze Panes' }).click()
    await page.waitForTimeout(400)
    const refrozen = await liveFreeze(page)
    expect(refrozen?.startRow, 're-frozen at C3').toBe(2)

    const { bytes, plan } = await saveAndCapturePlan(page)
    const states = plan.pageSetupStates as Array<{
      sheetName: string
      frozenRows?: number
      frozenColumns?: number
    }>
    const dataState = states?.find((s) => s.sheetName === 'Data')
    expect(dataState?.frozenRows, 'save plan carries 2 frozen rows').toBe(2)
    expect(dataState?.frozenColumns, 'save plan carries 2 frozen columns').toBe(2)
    const sheet1 = await readZipEntry(bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('ySplit="2"')
    expect(sheet1).toContain('xSplit="2"')
    expect(sheet1).toContain('state="frozen"')

    // Reopen: the freeze survives.
    const reopened = await openFixture(page, bytes, 'e2e-view-frozen-saved.xlsx')
    const reopenedData = reopened.sheets.find((s) => s.name === 'Data')
    expect(reopenedData?.freeze, 'freeze survived save/reopen').toEqual({
      frozenRows: 2,
      frozenColumns: 2,
    })
    const reopenedFreeze = await liveFreeze(page)
    expect(reopenedFreeze?.startRow, 'reopened live model froze 2 rows').toBe(2)

    expect(pageErrors).toEqual([])
  })

  test('3: freeze CLEAR survives save/reopen (defect fix)', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelViewFixture(
      `<sheetView workbookViewId="0">${CANONICAL_PANE}</sheetView>`,
    )
    await openFixture(page, fixture, 'e2e-view-frozenclear.xlsx')

    // Navigate to C3 (the frozen anchor) so the toggle takes the
    // isFrozenHere CLEAR branch.
    const box = page.locator('[data-testid="excel-name-box"]')
    await box.click()
    await box.fill('C3')
    await box.press('Enter')
    await page.waitForTimeout(400)
    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)
    await page.getByRole('button', { name: 'Freeze Panes' }).click()
    await page.waitForTimeout(500)
    const cleared = await liveFreeze(page)
    expect(cleared?.startRow, 'freeze cleared in-session').toBe(-1)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // The save plan must carry the 0/0 CLEAR (the old defect dropped the
    // sheet from the plan entirely, silently keeping the file's <pane>).
    const { bytes, plan } = await saveAndCapturePlan(page)
    const states = plan.pageSetupStates as Array<{
      sheetName: string
      frozenRows?: number
      frozenColumns?: number
    }>
    const dataState = states?.find((s) => s.sheetName === 'Data')
    expect(dataState, 'cleared sheet still emits a page-setup state').toBeDefined()
    expect(dataState?.frozenRows, 'clear journaled as 0 rows').toBe(0)
    expect(dataState?.frozenColumns, 'clear journaled as 0 columns').toBe(0)
    const sheet1 = await readZipEntry(bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'pane element removed on clear').not.toContain('<pane')

    const reopened = await openFixture(page, bytes, 'e2e-view-frozenclear-saved.xlsx')
    const reopenedData = reopened.sheets.find((s) => s.name === 'Data')
    expect(reopenedData?.freeze, 'freeze stays cleared after reopen').toBeUndefined()
    const reopenedFreeze = await liveFreeze(page)
    expect(reopenedFreeze?.startRow, 'reopened live model is unfrozen').toBe(-1)

    expect(pageErrors).toEqual([])
  })

  test('4+5+6: gridlines import, toggle persists, toggle back restores', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // ── Scenario 4: the file's showGridLines="0" imports. ────────────────
    const fixture = await buildExcelViewFixture('<sheetView showGridLines="0" workbookViewId="0"/>')
    const snapshot = await openFixture(page, fixture, 'e2e-view-gridlines-import.xlsx')
    const data = snapshot.sheets.find((s) => s.name === 'Data')
    expect(data?.view?.showGridlines, 'snapshot exposes hidden gridlines').toBe(false)
    expect(await liveGridlinesHidden(page), 'live model hides gridlines').toBe(true)
    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)
    await expect(page.getByRole('button', { name: 'Gridlines' })).not.toHaveAttribute(
      'aria-pressed',
      'true',
    )
    // The OTHER sheet keeps its default view (per-sheet state).
    const other = snapshot.sheets.find((s) => s.name === 'Other')
    expect(other?.view, 'other sheet has default view').toBeUndefined()

    // ── Scenario 5: toggle ON (restore) persists as the dropped attribute. ──
    await page.getByRole('button', { name: 'Gridlines' }).click()
    await page.waitForTimeout(400)
    expect(await liveGridlinesHidden(page), 'gridlines visible in-session').toBe(false)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const firstSave = await saveAndCapturePlan(page)
    const states1 = firstSave.plan.pageSetupStates as Array<{
      sheetName: string
      showGridlines?: boolean
    }>
    const dataState1 = states1?.find((s) => s.sheetName === 'Data')
    expect(dataState1?.showGridlines, 'restore journaled as true').toBe(true)
    const sheet1First = await readZipEntry(firstSave.bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1First, 'attribute dropped to restore the default').not.toContain('showGridLines')

    // Reopen: default (visible) gridlines.
    const reopened1 = await openFixture(page, firstSave.bytes, 'e2e-view-gridlines-restored.xlsx')
    const reopenedData1 = reopened1.sheets.find((s) => s.name === 'Data')
    expect(reopenedData1?.view, 'no view state after restore').toBeUndefined()
    expect(await liveGridlinesHidden(page), 'reopened gridlines visible').toBe(false)

    // ── Scenario 6: toggle OFF then back ON returns to the original state. ──
    await page.getByRole('button', { name: 'Gridlines' }).click()
    await page.waitForTimeout(400)
    expect(await liveGridlinesHidden(page), 'gridlines hidden in-session').toBe(true)
    const secondSave = await saveAndCapturePlan(page)
    const sheet1Second = await readZipEntry(secondSave.bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1Second, 'hidden journaled as showGridLines=0').toContain('showGridLines="0"')
    // Toggle back on and save again — the XML returns to the default state
    // (attribute absent), i.e. the ORIGINAL pre-toggle state.
    await page.getByRole('button', { name: 'Gridlines' }).click()
    await page.waitForTimeout(400)
    const thirdSave = await saveAndCapturePlan(page)
    const sheet1Third = await readZipEntry(thirdSave.bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1Third, 'toggle-back removes the attribute again').not.toContain('showGridLines')

    expect(pageErrors).toEqual([])
  })

  test('7+8: formula view renders + imports; toggle persists through save/reopen', async ({
    page,
  }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // ── Scenario 7: the file's showFormulas="1" imports AND renders. ────
    const fixture = await buildExcelViewFixture('<sheetView showFormulas="1" workbookViewId="0"/>')
    const snapshot = await openFixture(page, fixture, 'e2e-view-formulas-import.xlsx')
    const data = snapshot.sheets.find((s) => s.name === 'Data')
    expect(data?.view?.showFormulas, 'snapshot exposes formula view').toBe(true)

    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)
    const formulasButton = page.getByRole('button', { name: 'Show Formulas' })
    await expect(formulasButton).toHaveAttribute('aria-pressed', 'true')

    // Render proof: the grid paints formula text (C1 = "=SUM(A1:B1)")
    // instead of the cached value (42). Toggling OFF must CHANGE the
    // pixels; toggling back ON must return to the EXACT imported render.
    const withFormulas = await gridClip(page)
    await formulasButton.click()
    await page.waitForTimeout(700)
    await expect(formulasButton).not.toHaveAttribute('aria-pressed', 'true')
    const withValues = await gridClip(page)
    expect(withFormulas.equals(withValues), 'formula view changes the rendered cells').toBe(false)
    await formulasButton.click()
    await page.waitForTimeout(700)
    const backToFormulas = await gridClip(page)
    expect(
      backToFormulas.equals(withFormulas),
      'toggling back returns the exact formula-view render',
    ).toBe(true)

    // ── Scenario 8: the toggle persists through save/reopen. ────────────
    // Current state: formula view ON (journaled by the last toggle). Save.
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const { bytes, plan } = await saveAndCapturePlan(page)
    const states = plan.pageSetupStates as Array<{
      sheetName: string
      showFormulas?: boolean
    }>
    const dataState = states?.find((s) => s.sheetName === 'Data')
    expect(dataState?.showFormulas, 'formula view journaled as true').toBe(true)
    const sheet1 = await readZipEntry(bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'showFormulas=1 written').toContain('showFormulas="1"')

    const reopened = await openFixture(page, bytes, 'e2e-view-formulas-saved.xlsx')
    const reopenedData = reopened.sheets.find((s) => s.name === 'Data')
    expect(reopenedData?.view?.showFormulas, 'formula view survived save/reopen').toBe(true)
    // The reopened sheet renders formulas from the start.
    await page.getByRole('tab', { name: 'View', exact: true }).click()
    await page.waitForTimeout(200)
    await expect(page.getByRole('button', { name: 'Show Formulas' })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    const reopenedRender = await gridClip(page)
    expect(reopenedRender.equals(withFormulas), 'reopened render matches formula view').toBe(true)

    expect(pageErrors).toEqual([])
  })

  test('9+10: unrelated edits keep view state; no-op save preserves view XML', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // A file with ALL view state: hidden gridlines + formula view + pane
    // (canonical form so the seeded freeze rewrite is byte-identical).
    const fixture = await buildExcelViewFixture(
      `<sheetView showGridLines="0" showFormulas="1" workbookViewId="0">${CANONICAL_PANE}</sheetView>`,
    )
    const FIXTURE_SHEET_VIEWS = `<sheetViews><sheetView showGridLines="0" showFormulas="1" workbookViewId="0">${CANONICAL_PANE}</sheetView></sheetViews>`
    const FIXTURE_SHEET2 = await readZipEntry(fixture, 'xl/worksheets/sheet2.xml')
    await openFixture(page, fixture, 'e2e-view-noop.xlsx')

    // ── Scenario 9: an unrelated CELL EDIT must not corrupt the view state.
    await editCell(page, 1, 0, 'Edited')
    const { bytes, plan } = await saveAndCapturePlan(page)
    const sheet1 = await readZipEntry(bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'cell edit landed').toContain('Edited')
    expect(sheet1, 'gridlines stay hidden').toContain('showGridLines="0"')
    expect(sheet1, 'formula view stays on').toContain('showFormulas="1"')
    expect(sheet1, 'freeze pane stays').toContain('state="frozen"')
    // Only the freeze (seeded journal) + the cell edit ride the save; the
    // gridlines/formula-view attributes are untouched by the merge.
    const states = plan.pageSetupStates as Array<{
      sheetName: string
      frozenRows?: number
      showGridlines?: boolean
      showFormulas?: boolean
    }>
    const dataState = states?.find((s) => s.sheetName === 'Data')
    expect(dataState?.frozenRows, 'seeded freeze re-emitted').toBe(2)
    expect(dataState?.showGridlines, 'gridlines NOT journaled (untouched)').toBeUndefined()
    expect(dataState?.showFormulas, 'formula view NOT journaled (untouched)').toBeUndefined()

    // ── Scenario 10: the <sheetViews> section is byte-verbatim through a
    // save that never touched view state (freeze re-emission included).
    const savedViews = /<sheetViews>[\s\S]*?<\/sheetViews>/.exec(sheet1)?.[0] ?? ''
    expect(savedViews, 'sheetViews section preserved byte-for-byte').toBe(FIXTURE_SHEET_VIEWS)
    // The untouched OTHER sheet part is byte-identical.
    const savedSheet2 = await readZipEntry(bytes, 'xl/worksheets/sheet2.xml')
    expect(savedSheet2, 'unrelated worksheet byte-identical').toBe(FIXTURE_SHEET2)

    // Reopen: all three view truths intact.
    const reopened = await openFixture(page, bytes, 'e2e-view-noop-saved.xlsx')
    const reopenedData = reopened.sheets.find((s) => s.name === 'Data')
    expect(reopenedData?.view?.showGridlines, 'gridlines still hidden').toBe(false)
    expect(reopenedData?.view?.showFormulas, 'formula view still on').toBe(true)
    expect(reopenedData?.freeze, 'freeze intact').toEqual({ frozenRows: 2, frozenColumns: 2 })
    expect(reopenedData?.cells?.['A2']?.value, 'cell edit survived').toBe('Edited')

    // Default-view workbook: a cell edit + save must NOT add view attributes.
    const plain = await buildExcelViewFixture()
    await openFixture(page, plain, 'e2e-view-plain.xlsx')
    await editCell(page, 1, 0, 'PlainEdit')
    const plainSave = await saveAndCapturePlan(page)
    const plainSheet1 = await readZipEntry(plainSave.bytes, 'xl/worksheets/sheet1.xml')
    expect(plainSheet1, 'no gridline normalization on a default-view file').not.toContain(
      'showGridLines',
    )
    expect(plainSheet1, 'no formula-view normalization').not.toContain('showFormulas')
    expect(plainSheet1, 'no pane invention').not.toContain('<pane')
    const plainStates = plainSave.plan.pageSetupStates as Array<{ sheetName: string }> | undefined
    expect(plainStates, 'default-view save carries no page-setup state').toBeUndefined()

    expect(pageErrors).toEqual([])
  })

  test('11: Page Layout print family persists (orientation/margins/size/fit)', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelViewFixture()
    await openFixture(page, fixture, 'e2e-view-page-layout.xlsx')

    await page.getByRole('tab', { name: 'Page Layout', exact: true }).click()
    await page.waitForTimeout(300)
    // Orientation → Landscape.
    await page.getByLabel('Orientation').selectOption('landscape')
    // Margins → Narrow.
    await page.getByLabel('Margins').selectOption('narrow')
    // Size → A4 (code 9).
    await page.getByLabel('Size').selectOption('9')
    // Scale to Fit Width → 1 page.
    await page.getByLabel('Width').selectOption('1')
    await page.waitForTimeout(300)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const { bytes, plan } = await saveAndCapturePlan(page)
    const states = plan.pageSetupStates as Array<{
      sheetName: string
      orientation?: string
      margins?: string
      paperSize?: number
      fitToWidth?: number
      fitToHeight?: number
      fitToPage?: boolean
    }>
    const dataState = states?.find((s) => s.sheetName === 'Data')
    expect(dataState?.orientation).toBe('landscape')
    expect(dataState?.margins).toBe('narrow')
    expect(dataState?.paperSize).toBe(9)
    expect(dataState?.fitToWidth).toBe(1)
    expect(dataState?.fitToHeight).toBe(0)
    expect(dataState?.fitToPage).toBe(true)

    const sheet1 = await readZipEntry(bytes, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'landscape written').toContain('orientation="landscape"')
    expect(sheet1, 'A4 written').toContain('paperSize="9"')
    // fitToWidth 1 is the OOXML default → dropped; fitToHeight 0 (auto)
    // is explicit. Fit-to-page engages through sheetPr.
    expect(sheet1, 'fitToHeight auto written').toContain('fitToHeight="0"')
    expect(sheet1, 'fitToPage engaged').toContain('<pageSetUpPr fitToPage="1"')
    const margins = /<pageMargins\b[^>]*>/.exec(sheet1)?.[0] ?? ''
    expect(margins, 'narrow margins written').toContain('left="0.25"')

    // The workbook reopens cleanly after the print edits.
    const reopened = await openFixture(page, bytes, 'e2e-view-page-layout-saved.xlsx')
    const reopenedData = reopened.sheets.find((s) => s.name === 'Data')
    expect(reopenedData?.cells?.['A1']?.value, 'cells intact after print edits').toBe(10)

    expect(pageErrors).toEqual([])
  })
})
