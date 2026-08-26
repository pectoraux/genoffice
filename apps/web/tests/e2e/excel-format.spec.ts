/**
 * REAL browser E2E — Excel cell formatting (Phase 3 Increment 9).
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Univer) → Vite proxy → HTTP /api/office/workbooks/* →
 *   vercel-handler → routeOffice → @genoffice/xlsx-gateway
 *   (readBasicWorkbook presentation + StylesheetEditor.resolveStyle) →
 *   XLSX bytes → browser download
 *
 * Covers: presentation rendering (styles/merges/heights/widths over the
 * wire + painted grid), formatting a cell through the REAL Univer bold
 * command (Ctrl+B → set-range-values mutation → style delta), the change-
 * driven save payload (style-only and value+style merged edits), saved-XML
 * fidelity, and reopen round-trips.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas } from './helpers'
import { buildExcelFormatFixture, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

interface WireFormat {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  fontFamily?: string
  fontSize?: number
  fontColor?: string
  fillColor?: string
  horizontalAlign?: string
  verticalAlign?: string
  wrapText?: boolean
}
interface WireSheet {
  id: string
  name: string
  cells: Record<string, { value: unknown; formula?: string }>
  styles?: Record<string, WireFormat>
  merges?: string[]
  rowHeights?: Record<string, number>
  colWidths?: Record<string, number>
}

const GRID_CANVAS_SELECTOR = '#genoffice-web-excel canvas'

/**
 * Grid geometry (loadSnapshot config): row header 46px, column header 20px,
 * default column 100px / row 20px — EXCEPT column A, whose custom width
 * (24 chars ≈ 173px) shifts every following column.
 */
const COLUMN_A_WIDTH_PX = 173
function cellCenter(bbox: { x: number; y: number }, row: number, column: number) {
  const xBefore = column === 0 ? 0 : COLUMN_A_WIDTH_PX + (column - 1) * 100
  return {
    x: bbox.x + 46 + xBefore + (column === 0 ? COLUMN_A_WIDTH_PX : 100) / 2,
    y: bbox.y + 20 + row * 20 + 10,
  }
}

test.describe('Excel cell formatting (real HTTP + real engine)', () => {
  test('render: styles, merges, heights, widths flow over HTTP and paint the grid', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormatFixture()
    writeFileSync('/tmp/e2e-fmt-fixture.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-fmt-fixture.xlsx')
    await expect(page.getByText('Opened e2e-fmt-fixture.xlsx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)
    const openBody = await openResponse.json()

    // ── Presentation data crosses the real HTTP boundary ──────────────────
    const sheets = openBody.snapshot.sheets as WireSheet[]
    const formats = sheets[0]
    expect(formats.name).toBe('Formats')
    expect(formats.styles?.A1?.bold).toBe(true)
    expect(formats.styles?.B1?.italic).toBe(true)
    expect(formats.styles?.C1?.underline).toBe(true)
    expect(formats.styles?.C1?.strikethrough).toBe(true)
    expect(formats.styles?.C1?.fontColor).toBe('C00000')
    expect(formats.styles?.D1?.fontSize).toBe(14)
    expect(formats.styles?.E1?.fillColor).toBe('FFD966')
    expect(formats.styles?.A2?.horizontalAlign).toBe('center')
    expect(formats.styles?.A2?.wrapText).toBe(true)
    expect(formats.styles?.B2?.horizontalAlign).toBe('right')
    expect(formats.merges).toEqual(['A3:B3'])
    expect(formats.rowHeights).toEqual({ '5': 30 })
    expect(formats.colWidths).toEqual({ A: 173 })

    // ── The Univer grid renders with formatting ────────────────────────────
    // Wait for the canvas to settle, then sample distinct colors: the yellow
    // fill (FFD966) + red font + borders + text must paint more colors than
    // a blank grid.
    await page.waitForTimeout(2500)
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
    expect(painted, 'grid canvas must paint formatted content (fill/font colors)').toBeGreaterThan(
      5,
    )

    expect(pageErrors).toEqual([])
  })

  test('format a cell via the real Univer bold command → save → reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormatFixture()
    writeFileSync('/tmp/e2e-fmt-bold.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-fmt-bold.xlsx')
    await expect(page.getByText('Opened e2e-fmt-bold.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // ── Select B4 ("Plain", unformatted) and apply Bold through the real
    //    Univer command (Ctrl+B fires SetStyleCommand → set-range-values
    //    with an {s:{bl:1}} payload — the exact mutation the editor
    //    captures). ────────────────────────────────────────────────────────
    const grid = page.locator(GRID_CANVAS_SELECTOR).last()
    const bbox = await grid.boundingBox()
    expect(bbox).not.toBeNull()
    const b4 = cellCenter(bbox!, 3, 1)
    await page.mouse.click(b4.x, b4.y)
    await page.waitForTimeout(300)
    await page.keyboard.press('Control+b')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // ── Save: the payload carries EXACTLY ONE style-only edit ─────────────
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-fmt-bold.xlsx')).toBeVisible({ timeout: 15_000 })
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{
          sheetName: string
          row: number
          column: number
          writeValue: boolean
          style?: { bold?: boolean }
        }>
      }
    }
    expect(saveBody.savePlan.edits).toHaveLength(1)
    const edit = saveBody.savePlan.edits[0]
    expect(edit.sheetName).toBe('Formats')
    expect(edit.row).toBe(3)
    expect(edit.column).toBe(1)
    expect(edit.writeValue).toBe(false)
    expect(edit.style?.bold).toBe(true)

    // ── Saved XML: B4 references a bold font ──────────────────────────────
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<c r="B4" s="1"')
    const stylesXml = await readZipEntry(saved, 'xl/styles.xml')
    expect(stylesXml).toContain('<font><b/><sz val="11"/><name val="Calibri"/></font>')
    // Untouched formatting byte-preserved.
    expect(stylesXml).toContain('<fgColor rgb="FFFFD966"/>')

    // ── Reopen: the bold format survived ──────────────────────────────────
    writeFileSync('/tmp/e2e-fmt-bold-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-fmt-bold-saved.xlsx')
    await expect(page.getByText('Opened e2e-fmt-bold-saved.xlsx')).toBeVisible({ timeout: 30_000 })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    expect(reopened[0].styles?.B4?.bold).toBe(true)
    expect(reopened[0].cells.B4?.value).toBe('Plain')
    // Everything else survived.
    expect(reopened[0].styles?.A1?.bold).toBe(true)
    expect(reopened[0].styles?.E1?.fillColor).toBe('FFD966')
    expect(reopened[0].merges).toEqual(['A3:B3'])
    expect(reopened[0].rowHeights).toEqual({ '5': 30 })
    expect(reopened[1].cells.A1?.value).toBe('Untouched')

    expect(pageErrors).toEqual([])
  })

  test('value + style edits on one cell merge into a single save edit', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormatFixture()
    writeFileSync('/tmp/e2e-fmt-compose.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-fmt-compose.xlsx')
    await expect(page.getByText('Opened e2e-fmt-compose.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // Type a value into D4 (empty), then bold it: the dirty map must merge
    // both mutations into ONE CellEdit.
    const grid = page.locator(GRID_CANVAS_SELECTOR).last()
    const bbox = await grid.boundingBox()
    expect(bbox).not.toBeNull()
    const d4 = cellCenter(bbox!, 3, 3)
    await page.mouse.click(d4.x, d4.y)
    await page.waitForTimeout(300)
    await page.keyboard.type('hello')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(300)
    // Re-select D4 (Enter moved the selection down) and bold.
    const d4again = cellCenter(bbox!, 3, 3)
    await page.mouse.click(d4again.x, d4again.y)
    await page.waitForTimeout(300)
    await page.keyboard.press('Control+b')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    await dl
    await expect(page.getByText('Saved e2e-fmt-compose.xlsx')).toBeVisible({ timeout: 15_000 })
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{
          row: number
          column: number
          writeValue: boolean
          cell: { value: unknown }
          style?: { bold?: boolean }
        }>
      }
    }
    const edits = saveBody.savePlan.edits.filter((e) => e.row === 3 && e.column === 3)
    expect(edits).toHaveLength(1)
    expect(edits[0].writeValue).toBe(true)
    expect(edits[0].cell.value).toBe('hello')
    expect(edits[0].style?.bold).toBe(true)

    expect(pageErrors).toEqual([])
  })
})
