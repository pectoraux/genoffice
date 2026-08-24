/**
 * REAL browser E2E — Excel row/column structural operations.
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Univer insertRows/deleteRows API) →
 *   sheet.mutation.insert-row / remove-rows CommandExecuted events →
 *   journaled structural ops + shifted dirty map → save plan with
 *   structuralOps → HTTP /api/office/workbooks/save → vercel-handler →
 *   routeOffice → applyCellEditsToXlsx (structural replay BEFORE edits) →
 *   XLSX bytes → browser download
 *
 * The Univer facade's insertRows/deleteRows are the same methods the real
 * UI drives (context menu / ribbon → command → mutation); calling them
 * through the exposed runtime exercises the identical mutation pipeline.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute, waitForGridCanvas } from './helpers'
import { buildExcelFixture, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

interface WireSheet {
  id: string
  name: string
  cells: Record<string, { value: unknown; formula?: string }>
  merges?: string[]
  rowHeights?: Record<string, number>
  colWidths?: Record<string, number>
}

const GRID_CANVAS_SELECTOR = '#genoffice-web-excel canvas'

/**
 * Execute insertRows on the named sheet through the real Univer facade
 * (fires sheet.mutation.insert-row → the editor's journaling handler).
 */
async function insertRows(
  page: import('@playwright/test').Page,
  sheetName: string,
  rowIndex: number,
  count: number,
): Promise<void> {
  await page.evaluate(
    ({ name, row, n }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(name: string): { insertRows(row: number, count?: number): unknown }
            }
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const wb = rt.univerAPI.getActiveWorkbook()
      if (!wb) throw new Error('no active workbook')
      const ws = wb.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      ws.insertRows(row, n)
    },
    { name: sheetName, row: rowIndex, n: count },
  )
}

/**
 * Execute deleteRows on the named sheet through the real Univer facade.
 */
async function deleteRows(
  page: import('@playwright/test').Page,
  sheetName: string,
  rowIndex: number,
  count: number,
): Promise<void> {
  await page.evaluate(
    ({ name, row, n }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(name: string): { deleteRows(row: number, count?: number): unknown }
            }
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const wb = rt.univerAPI.getActiveWorkbook()
      if (!wb) throw new Error('no active workbook')
      const ws = wb.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      ws.deleteRows(row, n)
    },
    { name: sheetName, row: rowIndex, n: count },
  )
}

test.describe('Excel structural operations (real HTTP + real engine)', () => {
  test('insert rows → save → payload + XML + reopen fidelity', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFixture()
    writeFileSync('/tmp/e2e-struct-insert.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-struct-insert.xlsx')
    await expect(page.getByText('Opened e2e-struct-insert.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // ── Insert 2 rows at the top through the REAL Univer API ─────────────
    await insertRows(page, 'Data', 0, 2)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // ── Save: the payload carries the structural op ──────────────────────
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-struct-insert.xlsx')).toBeVisible({ timeout: 15_000 })

    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        edits: unknown[]
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{ kind: string; index: number; count: number }>
        }>
      }
    }
    // The save plan carries exactly ONE structural op (insert-rows 0, 2).
    expect(saveBody.savePlan.structuralOps).toEqual([
      { sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] },
    ])
    // No cell edits (nothing typed into the new rows).
    expect(saveBody.savePlan.edits).toEqual([])

    // ── Saved XML: rows shifted, content preserved ───────────────────────
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The original A1 content is now at A3.
    expect(sheet1).toContain('<c r="A3"')
    // The formula shifted from C1 to C3.
    expect(sheet1).toContain('<f>SUM(B3:B3)</f>')
    // The merge shifted from A3:B3 to A5:B5.
    expect(sheet1).toContain('<mergeCell ref="A5:B5"/>')

    // ── Reopen: the structural change round-trips ────────────────────────
    writeFileSync('/tmp/e2e-struct-insert-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-struct-insert-saved.xlsx')
    await expect(page.getByText('Opened e2e-struct-insert-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    const data = reopened[0]
    expect(data.cells.A3?.value).toBe('Original Text')
    expect(data.cells.B3?.value).toBe(10)
    expect(data.cells.C3?.formula).toBe('=SUM(B3:B3)')
    expect(data.merges).toEqual(['A5:B5'])
    expect(data.rowHeights).toEqual({ '7': 30 })
    expect(data.colWidths).toEqual({ A: 173 })
    // Hidden sheet untouched.
    expect(reopened[1].cells.A1?.value).toBe('Hidden Value')

    expect(pageErrors).toEqual([])
  })

  test('delete rows → save → reopen: row removed, neighbors intact', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelFixture()
    writeFileSync('/tmp/e2e-struct-delete.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-struct-delete.xlsx')
    await expect(page.getByText('Opened e2e-struct-delete.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(2000)

    // ── Delete the first row through the REAL Univer API ─────────────────
    await deleteRows(page, 'Data', 0, 1)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // ── Save ──────────────────────────────────────────────────────────────
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-struct-delete.xlsx')).toBeVisible({ timeout: 15_000 })

    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        structuralOps?: Array<{
          sheetName: string
          ops: Array<{ kind: string; index: number; count: number }>
        }>
      }
    }
    expect(saveBody.savePlan.structuralOps).toEqual([
      { sheetName: 'Data', ops: [{ kind: 'remove-rows', index: 0, count: 1 }] },
    ])

    // ── Capture the saved bytes from the FIRST save's download ──────────
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    writeFileSync('/tmp/e2e-struct-delete-saved.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-struct-delete-saved.xlsx')
    await expect(page.getByText('Opened e2e-struct-delete-saved.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot.sheets as WireSheet[]
    const data = reopened[0]
    // Original A1 row is gone; A3 "Merged Header" is now A2.
    expect(data.cells.A2?.value).toBe('Merged Header')
    // Merge shifted from A3:B3 to A2:B2.
    expect(data.merges).toEqual(['A2:B2'])
    // Row height shifted from row 5 to row 4.
    expect(data.rowHeights).toEqual({ '4': 30 })

    expect(pageErrors).toEqual([])
  })
})
