/**
 * REAL browser E2E — Home ribbon formatting persistence (Phase 4 Inc. 3).
 *
 * Proves every Home formatting command that claims to be enabled SURVIVES
 * save/reopen through the canonical pipeline:
 *
 *   browser (Univer) → Vite proxy → HTTP /api/office/workbooks/save →
 *   vercel-handler → routeOffice → applyCellEditsToXlsx → XLSX bytes →
 *   reopen → /api/office/workbooks/open → readBasicWorkbook → snapshot
 *
 * Covers the persistence gap the audit identified (number format does NOT
 * persist before Inc. 3 because the journal did not capture
 * sheet.mutation.set.numfmt), and re-validates the already-persisting
 * style families (bold, font color, fill, alignment, wrap, merge).
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas, clickSaveAndCaptureDownload } from './helpers'
import { buildExcelFormatFixture, readZipEntry } from './fixtures'

test.describe('Home ribbon persistence (real HTTP + real engine)', () => {
  test('number format persists through save/reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormatFixture()
    writeFileSync('/tmp/e2e-ribbon-numfmt.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-numfmt.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-numfmt.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Select B4 ("Plain", unformatted) and apply the "Currency" number format
    // via the Home ribbon's Number group select. This fires the numfmt facade
    // mixin's .n(pattern) → sheet.mutation.set.numfmt, which Inc. 3 journals
    // as a per-cell style.numberFormat CellEdit.
    const GRID_CANVAS_SELECTOR = '#genoffice-web-excel canvas'
    const grid = page.locator(GRID_CANVAS_SELECTOR).last()
    const bbox = await grid.boundingBox()
    expect(bbox).not.toBeNull()
    const COLUMN_A_WIDTH_PX = 173
    // B4 = column 1 (B), row 3 (0-indexed). xBefore is the horizontal offset
    // to the left edge of column B.
    const xBefore = COLUMN_A_WIDTH_PX + (1 - 1) * 100
    const b4x = bbox!.x + 46 + xBefore + 100 / 2
    const b4y = bbox!.y + 20 + 3 * 20 + 10
    await page.mouse.click(b4x, b4y)
    await page.waitForTimeout(300)

    // Switch to the Home tab if not already active (it is the default).
    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(150)

    // Apply the Currency number format (pattern '$#,##0').
    const numfmtSelect = page.getByRole('combobox', { name: 'Number format' })
    await numfmtSelect.selectOption('$#,##0')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Capture the save request and the downloaded bytes.
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
          style?: { numberFormat?: string }
        }>
      }
    }
    // The numfmt journal produced ONE style-only edit for B4 with
    // style.numberFormat === '$#,##0'.
    const numfmtEdits = saveBody.savePlan.edits.filter(
      (e) => e.row === 3 && e.column === 1 && e.style?.numberFormat === '$#,##0',
    )
    expect(numfmtEdits, 'numfmt edit must be journaled as style.numberFormat').toHaveLength(1)
    expect(numfmtEdits[0].writeValue, 'numfmt edit is style-only').toBe(false)

    // Saved XLSX: the cellXfs entry referenced by B4 carries numFmtId pointing
    // at a numFmt entry for '$#,##0'.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'B4 cell present in saved XML').toContain('r="B4"')
    const stylesXml = await readZipEntry(saved, 'xl/styles.xml')
    expect(stylesXml, 'numFmt for $#,##0 written').toContain('$#,##0')

    // Reopen and verify the snapshot carries the number format.
    writeFileSync('/tmp/e2e-ribbon-numfmt-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-numfmt-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-numfmt-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      styles?: Record<string, { numberFormat?: string }>
    }>
    // The StylesheetReader resolved B4's numFmtId back to a numberFormat
    // string on the snapshot.
    const formats = reopened[0]
    expect(formats.name).toBe('Formats')
    const b4Fmt = formats.styles?.B4
    expect(b4Fmt?.numberFormat, 'B4 numberFormat survived save/reopen').toBe('$#,##0')

    expect(pageErrors).toEqual([])
  })

  test('bold, font color, fill, alignment, wrap all persist through save/reopen', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormatFixture()
    writeFileSync('/tmp/e2e-ribbon-home.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-home.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-home.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Switch to Home tab.
    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(150)

    // Select B4 (unformatted "Plain") via a grid click.
    const GRID_CANVAS_SELECTOR = '#genoffice-web-excel canvas'
    const grid = page.locator(GRID_CANVAS_SELECTOR).last()
    const bbox = await grid.boundingBox()
    expect(bbox).not.toBeNull()
    const COLUMN_A_WIDTH_PX = 173
    const b4x = bbox!.x + 46 + COLUMN_A_WIDTH_PX + 100 / 2
    const b4y = bbox!.y + 20 + 3 * 20 + 10
    await page.mouse.click(b4x, b4y)
    await page.waitForTimeout(300)

    // Apply Bold (Ctrl+B fires SetBoldCommand → set-range-values with {s:{bl:1}}).
    await page.keyboard.press('Control+b')
    await page.waitForTimeout(200)

    // Apply font color (red #FF0000) via the Home ribbon color picker.
    const fontColorInput = page.locator('input[aria-label="Font color picker"]')
    await fontColorInput.fill('#FF0000')
    await page.waitForTimeout(200)

    // Apply fill color (yellow #FFFF00) via the Home ribbon fill picker.
    const fillColorInput = page.locator('input[aria-label="Fill color picker"]')
    await fillColorInput.fill('#FFFF00')
    await page.waitForTimeout(200)

    // Apply horizontal alignment = center via the Home ribbon.
    await page.getByRole('button', { name: 'Align center' }).click()
    await page.waitForTimeout(200)

    // Apply wrap text via the Home ribbon.
    await page.getByRole('button', { name: 'Wrap text' }).click()
    await page.waitForTimeout(200)

    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save and capture.
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
          style?: {
            bold?: boolean
            fontColor?: string
            fillColor?: string
            horizontalAlignment?: string
            wrapText?: boolean
          }
        }>
      }
    }
    const b4Edits = saveBody.savePlan.edits.filter((e) => e.row === 3 && e.column === 1)
    expect(b4Edits, 'B4 has at least one journaled edit').not.toHaveLength(0)
    // The journal merges multiple style deltas on the same cell into one
    // CellEdit (mergeCellEdit composes style fields). Assert that across all
    // B4 edits, every applied style field is present (bold, fontColor,
    // fillColor, horizontalAlignment, wrapText).
    const hasBold = b4Edits.some((e) => e.style?.bold === true)
    const hasFontColor = b4Edits.some((e) => e.style?.fontColor === '#FF0000')
    const hasFillColor = b4Edits.some((e) => e.style?.fillColor === '#FFFF00')
    const hasHAlign = b4Edits.some((e) => e.style?.horizontalAlignment === 'center')
    const hasWrap = b4Edits.some((e) => e.style?.wrapText === true)
    expect(hasBold, 'B4 edit carries bold').toBe(true)
    expect(hasFontColor, 'B4 edit carries fontColor #FF0000').toBe(true)
    expect(hasFillColor, 'B4 edit carries fillColor #FFFF00').toBe(true)
    expect(hasHAlign, 'B4 edit carries hAlign center').toBe(true)
    expect(hasWrap, 'B4 edit carries wrapText').toBe(true)

    // Saved XLSX: B4 references a bold + colored + filled + center + wrap xf.
    const stylesXml = await readZipEntry(saved, 'xl/styles.xml')
    expect(stylesXml, 'bold font written').toMatch(/<font>.*<b\/>.*<\/font>/s)
    expect(stylesXml, 'red font color written').toContain('FF0000')
    expect(stylesXml, 'yellow fill written').toContain('FFFF00')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'B4 cell present').toContain('r="B4"')

    // Reopen and verify the snapshot carries the merged format.
    writeFileSync('/tmp/e2e-ribbon-home-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-home-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-home-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      styles?: Record<string, {
        bold?: boolean
        fontColor?: string
        fillColor?: string
        horizontalAlign?: string
        wrapText?: boolean
      }>
    }>
    const b4Fmt = reopened[0].styles?.B4
    expect(b4Fmt?.bold, 'bold survived reopen').toBe(true)
    expect(b4Fmt?.fontColor?.toUpperCase(), 'fontColor survived reopen').toContain('FF0000')
    expect(b4Fmt?.fillColor?.toUpperCase(), 'fillColor survived reopen').toContain('FFFF00')
    expect(b4Fmt?.horizontalAlign, 'hAlign survived reopen').toBe('center')
    expect(b4Fmt?.wrapText, 'wrap survived reopen').toBe(true)

    expect(pageErrors).toEqual([])
  })

  test('merge & center persists through save/reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormatFixture()
    writeFileSync('/tmp/e2e-ribbon-merge.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-merge.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-merge.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Select the D4:E4 range via the Name Box (the proven path — canvas
    // drag-select is coordinate-fragile), then apply Merge & Center.
    const nb = page.locator('[data-testid="excel-name-box"]')
    await nb.click()
    await nb.fill('D4:E4')
    await nb.press('Enter')
    await page.waitForTimeout(400)

    // Apply Merge via the Home ribbon.
    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(150)
    await page.getByRole('button', { name: 'Merge & center' }).click()
    await page.waitForTimeout(400)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save and verify the merge survived in the saved XML.
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1, 'D4:E4 merge written').toMatch(/<mergeCell[^>]*ref="D4:E4"/)

    // Reopen and verify the snapshot carries the merge.
    writeFileSync('/tmp/e2e-ribbon-merge-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-merge-saved.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-merge-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      merges?: string[]
    }>
    expect(reopened[0].merges, 'D4:E4 merge survived reopen').toContain('D4:E4')
    // The original A3:B3 merge should also still be there.
    expect(reopened[0].merges, 'A3:B3 merge preserved').toContain('A3:B3')

    expect(pageErrors).toEqual([])
  })
})
