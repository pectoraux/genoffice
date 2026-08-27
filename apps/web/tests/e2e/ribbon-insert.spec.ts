/**
 * REAL browser E2E — Insert tab (Phase 4 Inc. 3; EXCEL-021 flipped Table;
 * EXCEL-022 flipped Picture; EXCEL-023 flipped Chart).
 *
 * Verifies the Insert tab's control wiring: Picture is ENABLED since
 * EXCEL-022 (the visualAdditions family is on the wire;
 * ribbon-images.spec.ts proves the full round-trip). Chart is ENABLED
 * since EXCEL-023 (the chartEdits family + visualAdditions.chart are on
 * the wire; ribbon-charts.spec.ts proves the full round-trip). Table is
 * ENABLED since EXCEL-021.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas } from './helpers'

test.describe('Insert tab — wired controls', () => {
  test('Chart, Picture and Table are enabled with documented actions', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Switch to the Insert tab.
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.waitForTimeout(200)

    // Table — ENABLED since EXCEL-021 (the tableAdditions family is on
    // the wire; the canonical round-trip is proven in ribbon-table.spec).
    const tableBtn = page.getByRole('button', { name: /^Table/, exact: false }).first()
    await expect(tableBtn).toBeEnabled()
    const tableTitle = await tableBtn.getAttribute('title')
    expect(tableTitle, 'Table title names the create action').toContain('create a table')

    // Delete Table — the convert-to-range companion command.
    const deleteBtn = page.getByRole('button', { name: /Delete Table/i })
    await expect(deleteBtn).toBeEnabled()

    // Chart — ENABLED since EXCEL-023: the chartEdits family +
    // visualAdditions.chart are on the wire (the canonical round-trip is
    // proven in ribbon-charts.spec.ts).
    const chartBtn = page.getByRole('button', { name: /^Chart/ }).first()
    await expect(chartBtn).toBeEnabled()
    const chartTitle = await chartBtn.getAttribute('title')
    expect(chartTitle, 'Chart title names the insert action').toContain(
      'insert a chart from the selected data range',
    )

    // Picture — ENABLED since EXCEL-022: the visualAdditions (image
    // embed) family is on the wire; the canonical round-trip is proven
    // in ribbon-images.spec.ts.
    const pictureBtn = page.getByRole('button', { name: /^Picture/ }).first()
    await expect(pictureBtn).toBeEnabled()
    const pictureTitle = await pictureBtn.getAttribute('title')
    expect(pictureTitle, 'Picture title names the insert action').toContain('insert an image')

    expect(pageErrors).toEqual([])
  })
})
