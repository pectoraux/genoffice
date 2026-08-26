/**
 * REAL browser E2E — Insert tab (Phase 4 Inc. 3; EXCEL-021 flipped Table).
 *
 * Verifies the Insert tab's Picture / Chart controls are VISIBLY DISABLED
 * (per spec: "a disabled button is preferable to a fake feature") — the
 * wire save plan still does not expose the visualAdditions / chartEdits
 * families. Table is ENABLED since EXCEL-021: the tableAdditions family
 * IS on the wire (ribbon-table.spec.ts proves the full round-trip), so
 * the old disabled-stub assertion flips to enabled + wired.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas } from './helpers'

test.describe('Insert tab — disabled-by-design controls', () => {
  test('Picture and Chart are disabled with documented reason; Table is enabled', async ({
    page,
  }) => {
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
