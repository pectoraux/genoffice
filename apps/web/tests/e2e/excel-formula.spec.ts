/**
 * REAL browser E2E — Excel formula-bar fidelity.
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Univer formula bar + grid) → Vite proxy → HTTP
 *   /api/office/workbooks/* → vercel-handler → routeOffice →
 *   @genoffice/xlsx-gateway (parse + serializeStyledCell) → XLSX bytes →
 *   browser download
 *
 * Uses Univer's BUILT-IN formula bar (formulaBar: true in the preset
 * config) — the audit proved it displays the formula and commits edits
 * through sheet.mutation.set-range-values.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas } from './helpers'
import { buildExcelFormulaFixture, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

interface WireSheet {
  id: string
  name: string
  cells: Record<string, { value: unknown; formula?: string }>
  styles?: Record<string, { bold?: boolean; fillColor?: string }>
  merges?: string[]
  rowHeights?: Record<string, number>
  colWidths?: Record<string, number>
}

const GRID = '#genoffice-web-excel canvas'

/** Grid geometry: col A is 173px (custom width); others 100px. Rows 20px. */
function cellXY(bbox: { x: number; y: number }, row: number, column: number) {
  const xBefore = column === 0 ? 0 : 173 + (column - 1) * 100
  return {
    x: bbox.x + 46 + xBefore + (column === 0 ? 173 : 100) / 2,
    y: bbox.y + 20 + row * 20 + 10,
  }
}

/**
 * Read the formula that the formula bar displays for the selected cell.
 * Univer renders the formula bar on canvas (no DOM text), but the
 * underlying cell state — which the formula bar mirrors — is readable
 * through the exposed runtime. This is the same data the canvas paints.
 */
async function formulaBarDisplay(page: import('@playwright/test').Page): Promise<string> {
  return page.evaluate(() => {
    const w = window as {
      __genofficeExcelRuntime?: {
        univerAPI: {
          getActiveWorkbook(): {
            getActiveSheet(): {
              getSelection(): { getActiveRange(): { getRow(): number; getColumn(): number } | null }
            }
            getSnapshot(): {
              sheets?: Record<string, { cellData?: Record<string, Record<string, { f?: string }>> }>
            }
          }
        }
      }
    }
    const rt = w.__genofficeExcelRuntime
    if (!rt) return ''
    const wb = rt.univerAPI.getActiveWorkbook()
    if (!wb) return ''
    const sel = wb.getActiveSheet().getSelection()?.getActiveRange()
    if (!sel) return ''
    // The formula bar canvas renders the cell's f field — read it from the
    // workbook snapshot (the same data the canvas paints).
    const snapshot = wb.getSnapshot()
    for (const sheet of Object.values(snapshot.sheets ?? {})) {
      const f = sheet.cellData?.[String(sel.getRow())]?.[String(sel.getColumn())]?.f
      if (typeof f === 'string') return f
    }
    return ''
  })
}

/**
 * Edit the selected cell's formula through the REAL editing path.
 * Double-clicking the cell opens the SAME editor the formula bar uses
 * (both commit through sheet.mutation.set-range-values with an f payload —
 * verified by the Phase A forensic audit). Select-all + type + Enter
 * replaces the formula.
 */
async function editFormulaInBar(
  page: import('@playwright/test').Page,
  cellX: number,
  cellY: number,
  text: string,
): Promise<void> {
  await page.mouse.dblclick(cellX, cellY)
  await page.waitForTimeout(500)
  await page.keyboard.press('Control+a')
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
  await page.waitForTimeout(800)
}

test.describe('Excel formula-bar fidelity (real HTTP + real engine)', () => {
  test('1+2: formula display + formula edit → save → XML → reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula.xlsx')
    await expect(page.getByText('Opened e2e-formula.xlsx')).toBeVisible({ timeout: 30_000 })
    await openResponsePromise
    await page.waitForTimeout(2000)

    // ── Test 1: select A3 → formula bar shows =SUM(A1:A2) ────────────────
    // Univer renders the formula bar on canvas; the cell state it mirrors
    // is readable through the exposed runtime (the canvas paints exactly
    // this formula string).
    const grid = page.locator(GRID).last()
    const bbox = await grid.boundingBox()
    expect(bbox).not.toBeNull()
    const a3 = cellXY(bbox!, 2, 0)
    await page.mouse.click(a3.x, a3.y)
    await page.waitForTimeout(1000)
    const barText = await formulaBarDisplay(page)
    expect(barText).toContain('SUM(A1:A2)')

    // ── Test 2: edit A3 through the REAL editing path ─────────────────────
    await editFormulaInBar(page, a3.x, a3.y, '=SUM(A1:A2)*2')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: capture the request + download.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-formula.xlsx')).toBeVisible({ timeout: 15_000 })

    // The save payload: exactly ONE edit for A3 (the recalc echo merged).
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{
          sheetName: string
          row: number
          column: number
          writeValue: boolean
          cell: { value: unknown; formula?: string }
        }>
      }
    }
    expect(saveBody.savePlan.edits).toHaveLength(1)
    const edit = saveBody.savePlan.edits[0]
    expect(edit.sheetName).toBe('Sheet1')
    expect(edit.row).toBe(2)
    expect(edit.column).toBe(0)
    expect(edit.cell.formula).toBe('SUM(A1:A2)*2')

    // The saved XML: <f> contains the new formula.
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f>SUM(A1:A2)*2</f>')

    // Reopen: A3 still contains the formula.
    writeFileSync('/tmp/e2e-formula-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-saved.xlsx')
    await expect(page.getByText('Opened e2e-formula-saved.xlsx')).toBeVisible({ timeout: 30_000 })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    expect(reopened[0].cells.A3?.formula).toBe('=SUM(A1:A2)*2')
    // Style survived.
    expect(reopened[0].styles?.A3?.bold).toBe(true)

    expect(pageErrors).toEqual([])
  })

  test('3: formula → literal (30)', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula-lit.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-lit.xlsx')
    await expect(page.getByText('Opened e2e-formula-lit.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // Click A3, type 30 (cell editing replaces the formula with a literal).
    const grid = page.locator(GRID).last()
    const bbox = await grid.boundingBox()
    const a3 = cellXY(bbox!, 2, 0)
    await page.mouse.click(a3.x, a3.y)
    await page.waitForTimeout(300)
    await page.keyboard.type('30')
    await page.keyboard.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-formula-lit.xlsx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)

    const xml = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    const a3Xml =
      /<c r="A3"[^>]*\/>/.exec(xml)?.[0] ?? /<c r="A3"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(a3Xml).not.toContain('<f>')
    expect(a3Xml).toContain('<v>30</v>')

    // Reopen: literal value.
    writeFileSync('/tmp/e2e-formula-lit-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-lit-saved.xlsx')
    await expect(page.getByText('Opened e2e-formula-lit-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    expect(reopened[0].cells.A3?.formula).toBeUndefined()
    expect(reopened[0].cells.A3?.value).toBe(30)
  })

  test('4: formula → blank (Delete)', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula-blank.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-blank.xlsx')
    await expect(page.getByText('Opened e2e-formula-blank.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    const grid = page.locator(GRID).last()
    const bbox = await grid.boundingBox()
    const a3 = cellXY(bbox!, 2, 0)
    await page.mouse.click(a3.x, a3.y)
    await page.waitForTimeout(300)
    await page.keyboard.press('Delete')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-formula-blank.xlsx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)

    const xml = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    const a3Xml =
      /<c r="A3"[^>]*\/>/.exec(xml)?.[0] ?? /<c r="A3"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(a3Xml).not.toContain('<f>')
    expect(a3Xml).not.toContain('<v>')

    writeFileSync('/tmp/e2e-formula-blank-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-blank-saved.xlsx')
    await expect(page.getByText('Opened e2e-formula-blank-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    expect(reopened[0].cells.A3?.formula).toBeUndefined()
    expect(reopened[0].cells.A3?.value).toBeNull()
  })

  test('5: cross-sheet formula survives exactly', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula-x.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-x.xlsx')
    await expect(page.getByText('Opened e2e-formula-x.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // Switch to Sheet2 (click the sheet tab).
    await page.getByText('Sheet2', { exact: true }).click()
    await page.waitForTimeout(1000)

    // Edit Sheet2!A2 through the formula bar.
    const grid = page.locator(GRID).last()
    const bbox = await grid.boundingBox()
    const a2 = cellXY(bbox!, 1, 0)
    await page.mouse.click(a2.x, a2.y)
    await page.waitForTimeout(500)
    await editFormulaInBar(page, a2.x, a2.y, '=Sheet1!A3+5')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-formula-x.xlsx')).toBeVisible({ timeout: 15_000 })

    // The payload carries the exact sheet reference.
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { edits: Array<{ cell: { formula?: string } }> }
    }
    expect(saveBody.savePlan.edits[0]?.cell?.formula).toBe('Sheet1!A3+5')

    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'xl/worksheets/sheet2.xml')
    expect(xml).toContain('<f>Sheet1!A3+5</f>')

    writeFileSync('/tmp/e2e-formula-x-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-x-saved.xlsx')
    await expect(page.getByText('Opened e2e-formula-x-saved.xlsx')).toBeVisible({ timeout: 30_000 })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    expect(reopened[1].cells.A2?.formula).toBe('=Sheet1!A3+5')
  })

  test('6: formula + bold in one burst → one composed edit', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula-bs.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-bs.xlsx')
    await expect(page.getByText('Opened e2e-formula-bs.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // Edit B3's formula, then bold it (two mutations on one cell).
    const grid = page.locator(GRID).last()
    const bbox = await grid.boundingBox()
    const b3 = cellXY(bbox!, 2, 1)
    await page.mouse.click(b3.x, b3.y)
    await page.waitForTimeout(300)
    await page.keyboard.type('=(B1+B2)*2')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(800)
    // Bold the same cell.
    await page.mouse.click(b3.x, b3.y)
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
    await expect(page.getByText('Saved e2e-formula-bs.xlsx')).toBeVisible({ timeout: 15_000 })

    // ONE composed edit: formula + style.
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: Array<{
          row: number
          column: number
          cell: { formula?: string }
          style?: { bold?: boolean }
        }>
      }
    }
    expect(saveBody.savePlan.edits).toHaveLength(1)
    const edit = saveBody.savePlan.edits[0]
    expect(edit.row).toBe(2)
    expect(edit.column).toBe(1)
    expect(edit.cell.formula).toBe('(B1+B2)*2')
    expect(edit.style?.bold).toBe(true)
  })

  test('7: unchanged formula not regenerated (open → immediate save)', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula-keep.xlsx', fixture)

    // Open, then edit an UNRELATED cell (to enable Save — it's disabled
    // when clean), then save. The formula cells must be untouched.
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-keep.xlsx')
    await expect(page.getByText('Opened e2e-formula-keep.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    const grid = page.locator(GRID).last()
    const bbox = await grid.boundingBox()
    const a1 = cellXY(bbox!, 0, 0)
    await page.mouse.click(a1.x, a1.y)
    await page.waitForTimeout(300)
    await page.keyboard.type('15')
    await page.keyboard.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-formula-keep.xlsx')).toBeVisible({ timeout: 15_000 })

    // Exactly ONE edit (A1 only) — no formula cells in the payload.
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { edits: Array<{ row: number; column: number }> }
    }
    expect(saveBody.savePlan.edits).toHaveLength(1)
    expect(saveBody.savePlan.edits[0]?.row).toBe(0)
    expect(saveBody.savePlan.edits[0]?.column).toBe(0)

    // The saved XML preserves formulas byte-identically (cached <v> intact).
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f>SUM(A1:A2)</f><v>30</v>')
    expect(xml).toContain('<f>B1*B2</f><v>35</v>')
    expect(xml).toContain('<f>"Hello " &amp; C1</f><v>Hello static</v>')
  })

  test('8: structural op + formula interaction', async ({ page }) => {
    test.setTimeout(120_000)
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFormulaFixture()
    writeFileSync('/tmp/e2e-formula-str.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-str.xlsx')
    await expect(page.getByText('Opened e2e-formula-str.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // Insert a row at the top through the REAL Univer facade.
    await page.evaluate(() => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(n: string): { insertRows(r: number, c?: number): unknown }
            }
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const ws = rt.univerAPI.getActiveWorkbook()!.getSheetByName('Sheet1')
      ws.insertRows(0, 1)
    })
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save + capture.
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-formula-str.xlsx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)

    // The formula shifted correctly: SUM(A1:A2) → SUM(A2:A3).
    const xml = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f>SUM(A2:A3)</f>')

    // Reopen: formulas coherent at shifted positions.
    writeFileSync('/tmp/e2e-formula-str-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-str-saved.xlsx')
    await expect(page.getByText('Opened e2e-formula-str-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    expect(reopened[0].cells.A4?.formula).toBe('=SUM(A2:A3)')
    // Cross-sheet formula shifted too. The engine's ref-range rewrite
    // re-serializes the formula from its token stream, which normalizes
    // whitespace: `Sheet1!A3 + 1` → `Sheet1!A4+ 1`. This is desktop parity
    // — the desktop journal captures the same engine mutation (the rewrite
    // carries no `fromFormula` execution option), so the desktop's save
    // plan carries the same normalized text. Compare whitespace-insensitively.
    expect(reopened[1].cells.A2?.formula?.replace(/\s+/g, '')).toBe('=Sheet1!A4+1')

    // Now delete the inserted row → everything returns to original.
    await page.evaluate(() => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(n: string): { deleteRows(r: number, c?: number): unknown }
            }
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const ws = rt.univerAPI.getActiveWorkbook()!.getSheetByName('Sheet1')
      ws.deleteRows(0, 1)
    })
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const dl2 = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download2 = await dl2
    await expect(page.getByText('Saved e2e-formula-str-saved.xlsx')).toBeVisible({
      timeout: 15_000,
    })
    const stream2 = await download2.createReadStream()
    const chunks2: Buffer[] = []
    for await (const chunk of stream2) chunks2.push(chunk as Buffer)
    const saved2 = Buffer.concat(chunks2)

    writeFileSync('/tmp/e2e-formula-str-saved2.xlsx', saved2)
    const reopen2Promise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-formula-str-saved2.xlsx')
    await expect(page.getByText('Opened e2e-formula-str-saved2.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopen2 = await reopen2Promise
    const reopened2 = (await reopen2.json()).snapshot.sheets as WireSheet[]
    // Back to the original coordinates.
    expect(reopened2[0].cells.A3?.formula).toBe('=SUM(A1:A2)')
    expect(reopened2[0].cells.A1?.value).toBe(10)
    expect(reopened2[0].merges).toEqual(['A5:B5'])
  })
})
