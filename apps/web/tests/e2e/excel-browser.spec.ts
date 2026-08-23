/**
 * REAL browser E2E — Excel editor.
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Univer) → Vite proxy → HTTP /api/office/workbooks/* →
 *   vercel-handler → routeOffice → @genoffice/xlsx-gateway → XLSX bytes →
 *   browser download
 *
 * Never calls routeOffice() directly — every assertion observes the real
 * HTTP boundary (network interception) or the real file bytes (zip parsing).
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute } from './helpers'
import { buildExcelFixture, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

const GRID_CANVAS_SELECTOR = '#genoffice-web-excel canvas'

test.describe('Excel browser E2E (real HTTP + real engine)', () => {
  test('upload → render → edit A1 → save → reopen → verify fidelity', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    // ── 1. Launch browser → navigate /office/excel (real login flow) ──────
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await page.waitForSelector(GRID_CANVAS_SELECTOR, { timeout: 30_000 })
    await expect(page.getByText('GenOffice Excel').first()).toBeVisible()

    // ── 2. Upload fixture.xlsx through the hidden file input ──────────────
    const fixture = await buildExcelFixture()
    const fixturePath = '/tmp/e2e-excel-fixture.xlsx'
    writeFileSync(fixturePath, fixture)

    // Capture the OPEN response that crosses the real HTTP boundary.
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)
    const openBody = await openResponse.json()

    // ── 3. Verify the snapshot the browser received over HTTP ─────────────
    const sheets = openBody.snapshot.sheets as Array<{
      id: string
      name: string
      cells: Record<string, { value: unknown; formula?: string }>
    }>
    expect(sheets.map((s) => s.name)).toEqual(['Data', 'HiddenSheet'])
    const dataCells = sheets[0].cells
    expect(dataCells.A1.value).toBe('Original Text')
    expect(dataCells.B1.value).toBe(10)
    expect(dataCells.C1.formula).toBe('=SUM(B1:B1)')
    expect(dataCells.A3.value).toBe('Merged Header')

    // ── 4. Verify cells rendered in Univer (canvas actually painted) ──────
    const painted = await page.evaluate(() => {
      const canvases = Array.from(document.querySelectorAll('#genoffice-web-excel canvas'))
      for (const c of canvases) {
        const r = c.getBoundingClientRect()
        if (r.width < 200 || r.height < 200) continue
        const ctx = (c as HTMLCanvasElement).getContext('2d')
        if (!ctx) continue
        const colors = new Set<string>()
        const strip = ctx.getImageData(60, 30, Math.floor(r.width) - 120, 80)
        for (let i = 0; i < strip.data.length; i += 40) {
          colors.add(`${strip.data[i]},${strip.data[i + 1]},${strip.data[i + 2]}`)
        }
        return colors.size
      }
      return 0
    })
    // A painted grid (headers + gridlines + text) has many distinct colors.
    expect(painted, 'Univer grid canvas should be painted, not blank').toBeGreaterThan(2)

    // Sheet tabs for both sheets should be present in the DOM.
    const tabText = await page.evaluate(() => document.body.textContent ?? '')
    expect(tabText).toContain('Data')
    expect(tabText).toContain('HiddenSheet')

    // ── 5. Edit A1 in the real Univer grid ────────────────────────────────
    // Grid geometry: row header 46px wide, column header 20px tall, column A
    // 100px wide, row 1 20px tall (all values come from loadSnapshot's config).
    const grid = page.locator(GRID_CANVAS_SELECTOR).last()
    const bbox = await grid.boundingBox()
    expect(bbox).not.toBeNull()
    const a1X = (bbox!.x ?? 0) + 46 + 50
    const a1Y = (bbox!.y ?? 0) + 20 + 10

    // Save is disabled before any edit (nothing dirty).
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeDisabled()

    await page.mouse.click(a1X, a1Y)
    await page.waitForTimeout(300)
    await page.keyboard.type('E2E Edited Cell')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(500)

    // Change-driven dirty state flipped by the mutation subscription.
    await expect(page.getByText('● Unsaved changes')).toBeVisible()
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeEnabled()

    // ── 6. Save → capture the save request payload + the download ─────────
    const saveRequestPromise = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saveResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/save') && r.request().method() === 'POST',
    )
    const downloadPromise = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()

    const saveRequest = await saveRequestPromise
    const saveResponse = await saveResponsePromise
    const download = await downloadPromise
    await expect(page.getByText('Saved e2e-excel-fixture.xlsx')).toBeVisible({ timeout: 15_000 })
    expect(saveResponse.status()).toBe(200)

    // Change-driven save: EXACTLY ONE edit — the A1 cell the user touched.
    const saveBody = JSON.parse(saveRequest.postData() ?? '{}') as {
      fileName: string
      savePlan: {
        edits: Array<{ sheetName: string; row: number; column: number; cell: { value: unknown } }>
      }
    }
    expect(saveBody.fileName).toBe('e2e-excel-fixture.xlsx')
    expect(saveBody.savePlan.edits).toEqual([
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: 'E2E Edited Cell' },
      },
    ])

    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    expect(saved.length).toBeGreaterThan(0)
    expect(saved.subarray(0, 2).toString('latin1')).toBe('PK') // real zip bytes

    // ── 7. Verify the saved XLSX bytes (engine output fidelity) ───────────
    const sheet1Xml = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // Edited A1: new value as inlineStr, original style index s="1" kept.
    const a1Match = sheet1Xml.match(/<c r="A1"[^>]*(?:\/>|>[\s\S]*?<\/c>)/)?.[0] ?? ''
    expect(a1Match).toContain('E2E Edited Cell')
    expect(a1Match).toContain('s="1"')
    // Untouched cell survived byte-for-byte.
    expect(sheet1Xml).toContain('<c r="B1"><v>10</v></c>')
    // Formula survived.
    expect(sheet1Xml).toContain('<f>SUM(B1:B1)</f>')
    // Merge / row height / column width survived.
    expect(sheet1Xml).toContain('<mergeCell ref="A3:B3"/>')
    expect(sheet1Xml).toContain('ht="30"')
    expect(sheet1Xml).toContain('width="24" customWidth="1"')
    // Hidden sheet survived in workbook.xml.
    const workbookXml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbookXml).toContain('state="hidden"')
    expect(workbookXml).toContain('HiddenSheet')
    // Styled cell definition survived in styles.xml.
    const stylesXml = await readZipEntry(saved, 'xl/styles.xml')
    expect(stylesXml).toContain('<b/>')
    expect(stylesXml).toContain('FFFFF2CC')

    // ── 8. Reopen the downloaded XLSX in the browser ──────────────────────
    const savedPath = '/tmp/e2e-excel-saved.xlsx'
    writeFileSync(savedPath, saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath)
    await expect(page.getByText('Opened e2e-excel-saved.xlsx')).toBeVisible({ timeout: 30_000 })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopenBody = await reopenResponse.json()
    const reopened = reopenBody.snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown; formula?: string }>
    }>
    // Edited cell shows the new value.
    expect(reopened[0].cells.A1.value).toBe('E2E Edited Cell')
    // Untouched cell survived.
    expect(reopened[0].cells.B1.value).toBe(10)
    // Formula survived the round-trip.
    expect(reopened[0].cells.C1.formula).toBe('=SUM(B1:B1)')
    // Hidden sheet still present.
    expect(reopened[1].name).toBe('HiddenSheet')
    expect(reopened[1].cells.A1.value).toBe('Hidden Value')

    // ── 9. No unexpected page errors during the whole flow ────────────────
    expect(pageErrors).toEqual([])
  })
})
