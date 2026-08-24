/**
 * REAL browser E2E — Insert tab (Phase 4 Inc. 3).
 *
 * Verifies the Insert tab's Table / Picture / Chart controls are VISIBLY
 * DISABLED (per spec: "a disabled button is preferable to a fake
 * feature"). The architectural reason: the wire save plan
 * (BrowserWorkbookSavePlan) does not expose the tableAdditions,
 * visualAdditions, or chartEdits families — applyCellEditsToXlsx accepts
 * them, but routeOffice's handleSaveWorkbook does not pass them through,
 * so any in-session insert would NOT survive save/reopen. Until the wire
 * is extended, the controls must remain disabled with the reason in the
 * title tooltip.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas } from './helpers'

test.describe('Insert tab — disabled-by-design controls', () => {
  test('Table, Picture, Chart are all disabled with documented reason', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Switch to the Insert tab.
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.waitForTimeout(200)

    // Table — must be disabled with a title that names the missing
    // tableAdditions save family.
    const tableBtn = page.getByRole('button', { name: /^Table/ }).first()
    await expect(tableBtn).toBeDisabled()
    const tableTitle = await tableBtn.getAttribute('title')
    expect(tableTitle, 'Table title names the architectural reason').toContain('tableAdditions')

    // Chart — must be disabled with a title naming chartEdits /
    // visualAdditions.
    const chartBtn = page.getByRole('button', { name: /^Chart/ }).first()
    await expect(chartBtn).toBeDisabled()
    const chartTitle = await chartBtn.getAttribute('title')
    expect(chartTitle, 'Chart title names the architectural reason').toMatch(
      /chartEdits|visualAdditions/,
    )

    // Picture — must be disabled with a title naming visualAdditions.
    const pictureBtn = page.getByRole('button', { name: /^Picture/ }).first()
    await expect(pictureBtn).toBeDisabled()
    const pictureTitle = await pictureBtn.getAttribute('title')
    expect(pictureTitle, 'Picture title names the architectural reason').toContain(
      'visualAdditions',
    )

    expect(pageErrors).toEqual([])
  })
})
