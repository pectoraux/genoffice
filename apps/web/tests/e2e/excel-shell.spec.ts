/**
 * REAL browser E2E — GenOffice web Sheets workspace shell (Phase 4).
 *
 * Drives the full production-shaped stack:
 *   browser → Vite proxy → HTTP /api/office/* → vercel-handler →
 *   routeOffice → @genoffice/xlsx-gateway → real XLSX bytes → browser
 *
 * Verifies the shell visually + behaviorally resembles the Electron Sheets
 * workspace: title bar, 7-tab ribbon, name box, formula bar, grid, status bar
 * with zoom, and that the new chrome does NOT regress the fidelity invariant
 * (formula bar edits commit through the existing mutation pipeline and save).
 *
 * Never calls routeOffice() directly.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas, clickSaveAndCaptureDownload } from './helpers'
import { buildExcelFixture, readZipEntry } from './fixtures'

const SHELL = '[data-testid="excel-shell"]'
const CANVAS = '#genoffice-web-excel canvas'
const RIBBON_TABS = '.excel-ribbon-tab'
const NAME_BOX = '[data-testid="excel-name-box"]'
const FORMULA_BAR = '[data-testid="excel-formula-bar"]'
const STATUSBAR = '[data-testid="excel-statusbar"]'
const ZOOM = '[data-testid="excel-zoom"]'
const THEME_TOGGLE = '[data-testid="theme-toggle"]'

/**
 * Read the active cell's A1 notation from the live Univer runtime exposed on
 * window. Used by the Name-Box navigation tests to assert the selection
 * actually moved (not just that goTo returned null).
 */
async function getActiveCellA1(page: import('@playwright/test').Page): Promise<string> {
  return page.evaluate(() => {
    const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { getActiveSheet: () => { getActiveCell: () => { getA1Notation: () => string } | null } | null } | null } } }).__genofficeExcelRuntime
    return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getActiveCell?.()?.getA1Notation?.() ?? 'none'
  })
}

/**
 * Read the active range's A1 notation (e.g. "A1:C5") from the live runtime.
 */
async function getActiveRangeA1(page: import('@playwright/test').Page): Promise<string> {
  return page.evaluate(() => {
    const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { getActiveSheet: () => { getActiveRange: () => { getA1Notation: () => string } | null } | null } | null } } }).__genofficeExcelRuntime
    return rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getActiveRange?.()?.getA1Notation?.() ?? 'none'
  })
}

test.describe('Excel workspace shell (Phase 4 parity)', () => {
  test.describe.configure({ mode: 'serial' })

  test('1. Sheets shell renders with exactly one Name Box and one Formula Bar', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await page.waitForSelector(SHELL, { timeout: 30_000 })
    await waitForGridCanvas(page)
    await expect(page.locator(SHELL)).toBeVisible()
    // ── DUPLICATE-HEADER REGRESSION (Phase 4 Inc. 2 Objective A) ──
    // With SheetsCorePreset header:false + formulaBar:false, Univer's
    // internal SpreadsheetHeader (Name Box + Formula Bar) is NOT mounted.
    // The custom .excel-formula-row is the SOLE name box + formula bar.
    // This asserts exactly one of each — if the config regresses to
    // header:true, the count would be 2 (Univer's + custom).
    await expect(page.locator(NAME_BOX)).toHaveCount(1)
    await expect(page.locator(FORMULA_BAR)).toHaveCount(1)
    await expect(page.locator(NAME_BOX)).toBeVisible()
    await expect(page.locator(FORMULA_BAR)).toBeVisible()
    // No Univer preset name-box / formula-bar elements inside the container.
    const univerHeaderCount = await page.evaluate(() => {
      const c = document.getElementById('genoffice-web-excel')
      if (!c) return -1
      return c.querySelectorAll(
        '[class*="univer-namebox"],[class*="univer-name-box"],[class*="univer-formula"],[class*="univer-header"]',
      ).length
    })
    expect(univerHeaderCount, 'no Univer preset header chrome inside the grid container').toBe(0)
    // The grid canvas has real, non-zero dimensions.
    const size = await page.evaluate(() => {
      const canvases = Array.from(document.querySelectorAll('#genoffice-web-excel canvas')) as HTMLCanvasElement[]
      let best = { w: 0, h: 0 }
      for (const c of canvases) {
        const r = c.getBoundingClientRect()
        if (r.width > best.w && r.height > best.h) best = { w: Math.round(r.width), h: Math.round(r.height) }
      }
      return best
    })
    expect(size.w, 'grid canvas width must be > 0').toBeGreaterThan(200)
    expect(size.h, 'grid canvas height must be > 0').toBeGreaterThan(100)
    expect(pageErrors).toEqual([])
  })

  test('2. Seven ribbon tabs are present', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await page.waitForSelector(RIBBON_TABS, { timeout: 30_000 })
    const labels = await page.locator(RIBBON_TABS).allTextContents()
    expect(labels.map((l) => l.trim())).toEqual([
      'Home',
      'Insert',
      'Page Layout',
      'Formulas',
      'Data',
      'Review',
      'View',
    ])
  })

  test('3. Formula bar is visible', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await page.waitForSelector(FORMULA_BAR, { timeout: 30_000 })
    await expect(page.locator(FORMULA_BAR)).toBeVisible()
  })

  test('4. Name box is visible', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await page.waitForSelector(NAME_BOX, { timeout: 30_000 })
    await expect(page.locator(NAME_BOX)).toBeVisible()
  })

  test('5. Sheet tabs are visible', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    // Univer's footer.sheetBar renders the sheet tabs inside the grid container.
    const text = await page.evaluate(() => document.body.textContent ?? '')
    expect(text).toContain('Sheet1')
  })

  test('6. Status bar and zoom controls are visible', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await page.waitForSelector(STATUSBAR, { timeout: 30_000 })
    await expect(page.locator(STATUSBAR)).toBeVisible()
    await expect(page.locator(ZOOM)).toBeVisible()
    // The zoom slider is a real range input.
    await expect(page.locator(`${ZOOM} input[type="range"]`)).toBeVisible()
    // Zoom percent label renders.
    const pct = await page.locator('.excel-zoom-value').first().textContent()
    expect(pct).toMatch(/\d+%/)
  })

  test('7. Switching ribbon tabs preserves the workbook', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    // Click through every ribbon tab and back to Home.
    for (const label of ['Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Home']) {
      await page.getByRole('tab', { name: label, exact: true }).click()
      await page.waitForTimeout(120)
    }
    // The grid canvas must still be present and painted.
    await expect(page.locator(CANVAS).last()).toBeVisible()
    // The workbook is still alive — the runtime still reports an active workbook.
    const alive = await page.evaluate(() => {
      const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => unknown } } }).__genofficeExcelRuntime
      return !!rt?.univerAPI?.getActiveWorkbook?.()
    })
    expect(alive, 'workbook must survive ribbon tab switches').toBe(true)
    expect(pageErrors).toEqual([])
  })

  test('8. Name box B5 navigation selects B5', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    // Upload a real workbook through the /api/office/workbooks/open path.
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(500)
    const box = page.locator(NAME_BOX)
    // Before: active cell is A1 (the default).
    expect(await getActiveCellA1(page)).toBe('A1')
    // Type B5 in the Name Box and press Enter.
    await box.click()
    await box.pressSequentially('B5', { delay: 30 })
    await expect(box).toHaveValue('B5')
    await box.press('Enter')
    await page.waitForTimeout(400)
    // ── CORE ACCEPTANCE: active cell must BE B5 (not just "not flagged
    // invalid"). The Univer facade's workbook.setActiveRange (the desktop's
    // proven path) must have actually moved the selection.
    expect(await getActiveCellA1(page), 'B5 navigation must move active cell to B5').toBe('B5')
    // The Name Box must NOT be flagged invalid.
    const invalidAfterValid = await box.evaluate((el) => el.classList.contains('invalid'))
    expect(invalidAfterValid, 'a valid A1 ref must not flag the name box as invalid').toBe(false)
    expect(pageErrors).toEqual([])
  })

  test('9. Name box range A1:C5 navigation selects the range', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(500)
    const box = page.locator(NAME_BOX)
    // Navigate to B5 first (so the active range is a single cell), then to
    // A1:C5 (a range) — proves range navigation works after a cell jump.
    await box.click()
    await box.fill('B5')
    await box.press('Enter')
    await page.waitForTimeout(300)
    expect(await getActiveCellA1(page)).toBe('B5')
    // Now jump to the range A1:C5.
    await box.click()
    await box.fill('A1:C5')
    await box.press('Enter')
    await page.waitForTimeout(400)
    // The active range must be A1:C5.
    expect(await getActiveRangeA1(page), 'range navigation must select A1:C5').toBe('A1:C5')
    // And the active cell (primary) within the range is A1.
    expect(await getActiveCellA1(page), 'primary cell of the range is A1').toBe('A1')
    const invalid = await box.evaluate((el) => el.classList.contains('invalid'))
    expect(invalid, 'a valid range ref must not flag the name box as invalid').toBe(false)
    expect(pageErrors).toEqual([])
  })

  test('10. Name box invalid ref returns a validation error', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(500)
    const box = page.locator(NAME_BOX)
    // Invalid cell ref.
    await box.click()
    await box.fill('!!bad')
    await box.press('Enter')
    await page.waitForTimeout(200)
    const invalidCell = await box.evaluate((el) => el.classList.contains('invalid'))
    expect(invalidCell, 'an invalid cell ref must flag the name box as invalid').toBe(true)
    // Invalid range ref.
    await box.click()
    await box.fill('A1:!!bad')
    await box.press('Enter')
    await page.waitForTimeout(200)
    const invalidRange = await box.evaluate((el) => el.classList.contains('invalid'))
    expect(invalidRange, 'an invalid range ref must flag the name box as invalid').toBe(true)
    // The active cell must NOT have moved (still A1 — the fixture's default).
    expect(await getActiveCellA1(page), 'invalid ref must not move the selection').toBe('A1')
  })

  test('11. Cell click changes the active selection (grid interactivity)', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    // With header:false the grid canvas is the first child of the container
    // and owns the full height — clicks must land on real cells (no offset
    // miscalculation from a hidden/mounted header plugin).
    expect(await getActiveCellA1(page), 'default active cell is A1').toBe('A1')
    const bbox = await page.evaluate(() => {
      const cs = Array.from(document.querySelectorAll('#genoffice-web-excel canvas')) as HTMLCanvasElement[]
      let best: { x: number; y: number; w: number; h: number } | null = null
      for (const c of cs) {
        const r = c.getBoundingClientRect()
        if (r.width > 200 && r.height > 100 && (!best || r.width > best.w)) {
          best = { x: r.x, y: r.y, w: r.width, h: r.height }
        }
      }
      return best
    })
    expect(bbox, 'grid canvas must be present').not.toBeNull()
    // Click a cell well inside the grid (offset +280, +90 ≈ column E, row 4).
    await page.mouse.click((bbox!.x ?? 0) + 280, (bbox!.y ?? 0) + 90)
    await page.waitForTimeout(400)
    const activeAfterClick = await getActiveCellA1(page)
    expect(activeAfterClick, 'a real grid click must move the active cell').not.toBe('A1')
    expect(activeAfterClick, 'active cell must be a real A1 ref').toMatch(/^[A-Z]+[0-9]+$/)
    expect(pageErrors).toEqual([])
  })

  test('12. Name box navigation works after switching sheets', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(500)
    const box = page.locator(NAME_BOX)
    // Navigate on the Data sheet (the fixture's active sheet).
    await box.click()
    await box.fill('B5')
    await box.press('Enter')
    await page.waitForTimeout(300)
    expect(await getActiveCellA1(page), 'B5 on Data sheet').toBe('B5')
    // Insert + activate a new sheet via the runtime (the sheet tabs are
    // canvas-rendered, so UI tab-clicking is coordinate-fragile; the
    // runtime path exercises the same Univer sheet-switch the UI would).
    await page.evaluate(() => {
      const rt = (window as { __genofficeExcelRuntime?: { univerAPI: { getActiveWorkbook: () => { insertSheet: () => { getSheetName: () => string }; setActiveSheet: (s: unknown) => void; getActiveSheet: () => { getSheetId: () => string } } } } }).__genofficeExcelRuntime
      const wb = rt?.univerAPI?.getActiveWorkbook?.()
      if (!wb) return
      const ns = wb.insertSheet()
      wb.setActiveSheet(ns as unknown as never)
    })
    await page.waitForTimeout(400)
    // Navigate on the new (now-active) sheet — proves navigation works
    // after the active sheet changed.
    await box.click()
    await box.fill('D10')
    await box.press('Enter')
    await page.waitForTimeout(400)
    expect(await getActiveCellA1(page), 'D10 on the new sheet after switch').toBe('D10')
    expect(pageErrors).toEqual([])
  })

  test('13. Name box navigation followed by formula bar edit saves', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(500)
    const box = page.locator(NAME_BOX)
    // Navigate to B5 via the Name Box (the keyboard focus hands back to
    // the grid via ILayoutService.focus()).
    await box.click()
    await box.fill('B5')
    await box.press('Enter')
    await page.waitForTimeout(400)
    expect(await getActiveCellA1(page), 'navigated to B5').toBe('B5')
    // Now edit the formula bar (the "navigation followed by typing/editing"
    // acceptance) and commit a value into B5.
    const bar = page.locator(FORMULA_BAR)
    await bar.click()
    await bar.fill('=SUM(B1:B1)*3')
    await bar.press('Enter')
    await page.waitForTimeout(300)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    // Save through the real HTTP boundary and capture the download.
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    expect(saved.length).toBeGreaterThan(0)
    // The saved XLSX must carry the formula committed into B5.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'formula committed after navigation must persist').toContain('SUM(B1:B1)*3')
    expect(pageErrors).toEqual([])
  })

  test('14. Formula bar can edit a formula and save it', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Upload the standard fixture (Data sheet: A1 text, B1=10, C1==SUM(B1:B1)).
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })

    // Select C1 (the formula cell) via the Name Box — the now-working
    // navigation path (workbook.setActiveRange).
    const box = page.locator(NAME_BOX)
    await box.click()
    await box.fill('C1')
    await box.press('Enter')
    await page.waitForTimeout(300)

    const bar = page.locator(FORMULA_BAR)
    await expect(bar).toBeVisible()
    // The formula bar should echo the existing formula.
    await page.waitForFunction(
      () => (document.querySelector('[data-testid="excel-formula-bar"]') as HTMLInputElement | null)?.value,
      { timeout: 10_000 },
    )
    const before = await bar.inputValue()
    expect(before).toContain('SUM(B1:B1)')

    // Edit the formula: =SUM(B1:B1)*2
    await bar.click()
    await bar.fill('=SUM(B1:B1)*2')
    await bar.press('Enter')
    await page.waitForTimeout(300)
    // The grid should now be dirty.
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save through the real HTTP boundary and capture the download.
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    expect(saved.length).toBeGreaterThan(0)

    // The saved XLSX must carry the edited formula.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('SUM(B1:B1)*2')
    expect(pageErrors).toEqual([])
  })

  test('15. Dark theme renders without broken Univer chrome', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
    // Switch to dark theme via the single theme toggle.
    await page.locator(THEME_TOGGLE).selectOption('dark')
    await page.waitForTimeout(400)
    // The <html data-theme> attribute must be "dark".
    const attr = await page.evaluate(() => document.documentElement.getAttribute('data-theme'))
    expect(attr).toBe('dark')
    // The shell + grid still render.
    await expect(page.locator(SHELL)).toBeVisible()
    await expect(page.locator(CANVAS).last()).toBeVisible()
    // A dark-theme background token should now be a dark color (not white).
    const bg = await page.evaluate(() => {
      const el = document.querySelector('.excel-shell') as HTMLElement | null
      return el ? getComputedStyle(el).backgroundColor : ''
    })
    expect(bg, 'dark shell background').not.toBe('rgb(255, 255, 255)')
    expect(pageErrors).toEqual([])
    // Capture a dark-mode screenshot for the parity comparison.
    await page.screenshot({ path: 'tests/e2e/.results/excel-dark.png', fullPage: false })
    // And a light-mode screenshot.
    await page.locator(THEME_TOGGLE).selectOption('light')
    await page.waitForTimeout(400)
    await page.screenshot({ path: 'tests/e2e/.results/excel-light.png', fullPage: false })
  })
})
