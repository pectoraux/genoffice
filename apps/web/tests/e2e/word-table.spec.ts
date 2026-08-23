/**
 * REAL browser E2E — editable Word tables (Phase 3 Increment 7).
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Tiptap table nodes) → Vite proxy → HTTP /api/office/documents/* →
 *   vercel-handler → routeOffice → @genoffice/docx-engine (parse +
 *   generateTableModelXml) → DOCX bytes → browser download
 *
 * Covers: open, real <table> in the Tiptap DOM, docxIndex identity, cell
 * edit, save + XML inspection, reopen fidelity, add/delete row, merge/split,
 * and the table dirty-state flags (edited=false unchanged / edited=true
 * modified). Never calls routeOffice() directly.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute } from './helpers'
import { buildWordTableFixture, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

interface WireRun {
  text: string
  bold?: boolean
  italic?: boolean
}
interface WireBlock {
  docxIndex: number | null
  type: string
  text: string
  table?: {
    rows: Array<
      Array<{
        paras: string[]
        richParas?: Array<{ runs: WireRun[]; align?: string }>
        colSpan?: number
        vMerge?: 'restart' | 'continue'
        fill?: string
      }>
    >
    headerRows?: boolean[]
  }
  edited?: boolean
  hidden?: boolean
}

test.describe('Word editable tables (real HTTP + real engine)', () => {
  test('open → edit cell → save → XML fidelity → reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    // ── 1. Open a DOCX containing a 2×3 table ──────────────────────────────
    const fixture = await buildWordTableFixture()
    const fixturePath = '/tmp/e2e-table-fixture.docx'
    writeFileSync(fixturePath, fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-table-fixture.docx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)
    const blocks = (await openResponse.json()).blocks as WireBlock[]

    // The table block carries a typed payload (2 rows × 3 grid columns).
    const tableBlock = blocks.find((b) => b.type === 'table')
    expect(tableBlock, 'table block with payload').toBeDefined()
    expect(tableBlock!.docxIndex).toBe(2)
    expect(tableBlock!.table!.rows).toHaveLength(2)
    expect(tableBlock!.table!.rows[0]).toHaveLength(3)
    // vMerge restart/continue pair in column 1.
    expect(tableBlock!.table!.rows[0][0].vMerge).toBe('restart')
    expect(tableBlock!.table!.rows[1][0].vMerge).toBe('continue')
    // Rich runs in the merged cell; fill + center on cell (0,1).
    expect(tableBlock!.table!.rows[0][0].richParas?.[0].runs).toEqual([
      { text: 'Merged ' },
      { text: 'bold', bold: true },
      { text: ' and ' },
      { text: 'italic', italic: true },
      { text: ' cell' },
    ])
    expect(tableBlock!.table!.rows[0][1].fill).toBe('FFF2CC')
    expect(tableBlock!.table!.rows[0][1].richParas?.[0].align).toBe('center')

    // ── 2. A real <table> exists in the Tiptap DOM ─────────────────────────
    const tableSel = '.ProseMirror table[data-docx-index="2"]'
    await expect(page.locator(tableSel)).toBeVisible()
    // ── 3. data-docx-index preserved ───────────────────────────────────────
    await expect(page.locator(tableSel)).toHaveAttribute('data-docx-index', '2')
    // The vertical merge renders as rowspan=2 on the first cell.
    await expect(page.locator(`${tableSel} tr`).first().locator('td').first()).toHaveAttribute(
      'rowspan',
      '2',
    )
    // Cell text renders (bold/italic inside the merged cell).
    await expect(page.locator(`${tableSel} td`).first()).toContainText('Merged')
    await expect(page.locator(`${tableSel} td`).first().locator('strong')).toHaveText('bold')
    await expect(page.locator(`${tableSel} td`).first().locator('em')).toHaveText('italic')

    // ── 4. Edit a PARAGRAPH (not the table), save, and verify the untouched
    //    table is still sent edited=false while the paragraph is edited=true.
    //    (The app disables Save when the document is clean, so a paragraph
    //    edit both enables it and proves dirty-state isolation.) ────────────
    const para0 = page.locator('.ProseMirror p[data-docx-index="0"]')
    await para0.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' (edited)')
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const saveReq1 = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl1 = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req1 = await saveReq1
    const download1 = await dl1
    await expect(page.getByText('Saved e2e-table-fixture.docx')).toBeVisible({ timeout: 15_000 })
    const blocks1 = JSON.parse(req1.postData() ?? '{}').blocks as WireBlock[]
    const para1 = blocks1.find((b) => b.docxIndex === 0)
    const tbl1 = blocks1.find((b) => b.type === 'table')
    expect(para1?.edited).toBe(true)
    // ── Requirement: unchanged table → edited=false ─────────────────────────
    expect(tbl1?.edited).toBe(false)
    // Unchanged table: the original bytes are copied (kind: 'original').
    const stream1 = await download1.createReadStream()
    const chunks1: Buffer[] = []
    for await (const chunk of stream1) chunks1.push(chunk as Buffer)
    const savedUnchanged = Buffer.concat(chunks1)
    const xmlUnchanged = await readZipEntry(savedUnchanged, 'word/document.xml')
    expect(xmlUnchanged).toContain('Bottom middle')
    expect(xmlUnchanged).toContain('w:fill="FFF2CC"')
    // Reload the saved document so the later cell-edit steps start clean.
    const savedPath0 = '/tmp/e2e-table-after-para-edit.docx'
    writeFileSync(savedPath0, savedUnchanged)
    const reloadResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath0)
    await expect(page.getByText('Opened e2e-table-after-para-edit.docx')).toBeVisible({
      timeout: 30_000,
    })
    await reloadResponsePromise
    await expect(page.locator(tableSel)).toBeVisible()

    // ── 5. Edit one cell ("Bottom middle" → "EDITED cell") ────────────────
    // Triple-click selects the cell (prosemirror-tables turns a triple-click
    // inside a cell into a single-cell CellSelection), so typing replaces
    // the cell content. (Ctrl+a would select the whole document.)
    const bottomMiddle = page.locator(`${tableSel} tr`).nth(1).locator('td').nth(0)
    await bottomMiddle.click({ clickCount: 3 })
    await page.keyboard.type('EDITED cell')
    await expect(page.getByText('● Unsaved')).toBeVisible()

    // ── 6. Save → 7./8. inspect the resulting DOCX XML ─────────────────────
    const saveReq2 = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl2 = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req2 = await saveReq2
    const download2 = await dl2
    // The document was reloaded from e2e-table-after-para-edit.docx above, so
    // this save reports that file name.
    await expect(page.getByText('Saved e2e-table-after-para-edit.docx')).toBeVisible({
      timeout: 15_000,
    })
    // ── Requirement: modified table → edited=true ───────────────────────────
    const blocks2 = JSON.parse(req2.postData() ?? '{}').blocks as WireBlock[]
    const tbl2 = blocks2.find((b) => b.type === 'table')
    expect(tbl2?.edited).toBe(true)
    expect(tbl2?.table?.rows[1][1].paras[0]).toBe('EDITED cell')

    const stream = await download2.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const docXml = await readZipEntry(saved, 'word/document.xml')
    const tblXml = docXml.match(/<w:tbl>[\s\S]*<\/w:tbl>/)?.[0] ?? ''
    // Edited cell text changed.
    expect(tblXml).toContain('EDITED cell')
    // Untouched cells survived.
    expect(tblXml).toContain('Merged ')
    expect(tblXml).toContain('Top right')
    expect(tblXml).toContain('Bottom right')
    // Structure survived: vertical merge + fill + marks + alignment.
    expect(tblXml).toContain('<w:vMerge w:val="restart"/>')
    expect(/<w:vMerge\/>/.test(tblXml)).toBe(true)
    expect(tblXml).toContain('w:fill="FFF2CC"')
    expect(tblXml).toContain('<w:jc w:val="center"/>')
    expect(/<w:b\/>(?:<w:bCs\/>)?<\/w:rPr><w:t xml:space="preserve">bold<\/w:t>/.test(tblXml)).toBe(
      true,
    )
    expect(
      /<w:i\/>(?:<w:iCs\/>)?<\/w:rPr><w:t xml:space="preserve">italic<\/w:t>/.test(tblXml),
    ).toBe(true)
    // 3 grid columns.
    expect((tblXml.match(/<w:gridCol/g) ?? []).length).toBe(3)

    // ── 9./10. Reopen through /api/office/documents/open ───────────────────
    const savedPath = '/tmp/e2e-table-saved.docx'
    writeFileSync(savedPath, saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath)
    await expect(page.getByText('Opened e2e-table-saved.docx')).toBeVisible({ timeout: 30_000 })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).blocks as WireBlock[]
    const rt = reopened.find((b) => b.type === 'table')?.table
    // Table still exists with correct dimensions/content.
    expect(rt?.rows).toHaveLength(2)
    expect(rt?.rows[0]).toHaveLength(3)
    expect(rt?.rows[1][1].paras[0]).toBe('EDITED cell')
    expect(rt?.rows[0][0].vMerge).toBe('restart')
    expect(rt?.rows[1][0].vMerge).toBe('continue')
    expect(rt?.rows[0][1].fill).toBe('FFF2CC')
    expect(rt?.rows[0][0].richParas?.[0].runs).toEqual([
      { text: 'Merged ' },
      { text: 'bold', bold: true },
      { text: ' and ' },
      { text: 'italic', italic: true },
      { text: ' cell' },
    ])
    // Renders in the real editor again.
    await expect(page.locator('.ProseMirror table td').first()).toContainText('Merged')

    expect(pageErrors).toEqual([])
  })

  test('add a row and delete a row survive save + reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordTableFixture()
    writeFileSync('/tmp/e2e-table-rowops.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-table-rowops.docx')
    await expect(page.getByText('Opened e2e-table-rowops.docx')).toBeVisible({ timeout: 30_000 })
    const tableSel = '.ProseMirror table[data-docx-index="2"]'
    await expect(page.locator(tableSel)).toBeVisible()

    // ── 11. Add a row (toolbar action) and verify it survives ──────────────
    // Click into the last row's middle cell so the table toolbar appears.
    await page.locator(`${tableSel} tr`).nth(1).locator('td').nth(0).click()
    await page.getByRole('button', { name: '+ Row', exact: true }).click()
    await expect(page.locator(`${tableSel} tr`)).toHaveCount(3)
    // Type a marker into the new row's middle cell.
    await page.locator(`${tableSel} tr`).nth(2).locator('td').nth(1).click()
    await page.keyboard.type('NEW ROW CELL')

    const saveReq1 = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl1 = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req1 = await saveReq1
    const download1 = await dl1
    await expect(page.getByText('Saved e2e-table-rowops.docx')).toBeVisible({ timeout: 15_000 })
    const blocks1 = JSON.parse(req1.postData() ?? '{}').blocks as WireBlock[]
    const t1 = blocks1.find((b) => b.type === 'table')
    expect(t1?.edited).toBe(true)
    expect(t1?.table?.rows).toHaveLength(3)

    const stream1 = await download1.createReadStream()
    const chunks1: Buffer[] = []
    for await (const chunk of stream1) chunks1.push(chunk as Buffer)
    const saved1 = Buffer.concat(chunks1)
    const savedPath = '/tmp/e2e-table-rowops-saved.docx'
    writeFileSync(savedPath, saved1)

    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath)
    await expect(page.getByText('Opened e2e-table-rowops-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    await reopenResponsePromise
    const reopenedSel = '.ProseMirror table'
    await expect(page.locator(`${reopenedSel} tr`)).toHaveCount(3)
    await expect(page.locator(`${reopenedSel} td`, { hasText: 'NEW ROW CELL' })).toHaveCount(1)

    // ── 12. Delete a row and verify it disappears after reopen ─────────────
    // The save-download blurs the editor, so after the reopen the DOM caret
    // is stale (still parked in the NEW ROW CELL from before the save) while
    // PM state holds the reset selection — clicking that SAME cell fires no
    // selectionchange, so the editor never syncs and the table toolbar never
    // appears. Click a paragraph first to re-anchor the selection.
    await page.locator('.ProseMirror p').first().click()
    await page.locator(`${reopenedSel} tr`).nth(2).locator('td').nth(1).click()
    await page.getByRole('button', { name: '- Row', exact: true }).click()
    await expect(page.locator(`${reopenedSel} tr`)).toHaveCount(2)
    await expect(page.locator(`${reopenedSel} td`, { hasText: 'NEW ROW CELL' })).toHaveCount(0)

    const saveReq2 = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl2 = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await saveReq2
    const download2 = await dl2
    await expect(page.getByText('Saved e2e-table-rowops-saved.docx')).toBeVisible({
      timeout: 15_000,
    })
    const stream2 = await download2.createReadStream()
    const chunks2: Buffer[] = []
    for await (const chunk of stream2) chunks2.push(chunk as Buffer)
    const saved2 = Buffer.concat(chunks2)
    const xml2 = await readZipEntry(saved2, 'word/document.xml')
    expect(xml2).not.toContain('NEW ROW CELL')

    const savedPath2 = '/tmp/e2e-table-rowops-saved2.docx'
    writeFileSync(savedPath2, saved2)
    const reopen2 = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath2)
    await expect(page.getByText('Opened e2e-table-rowops-saved2.docx')).toBeVisible({
      timeout: 30_000,
    })
    await reopen2
    await expect(page.locator(`${reopenedSel} tr`)).toHaveCount(2)

    expect(pageErrors).toEqual([])
  })

  test('merge and split cells round-trip through the engine', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordTableFixture()
    writeFileSync('/tmp/e2e-table-merge.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-table-merge.docx')
    await expect(page.getByText('Opened e2e-table-merge.docx')).toBeVisible({ timeout: 30_000 })
    const tableSel = '.ProseMirror table[data-docx-index="2"]'
    await expect(page.locator(tableSel)).toBeVisible()

    // ── 13. Merge cells: build a real prosemirror-tables CellSelection over
    // the two plain cells of row 1 ("Filled center" + "Top right") via the
    // app's E2E hooks (window.__genofficeWordEditor + CellSelection), then
    // Merge Cells. Works against local dev and deployed builds alike. ─────
    const selected = await page.evaluate(() => {
      type PmNode = {
        type: { name: string }
        nodeSize: number
        childCount: number
        child(i: number): PmNode
      }
      const w = window as unknown as {
        __genofficeWordEditor?: {
          state: {
            doc: {
              descendants(fn: (node: PmNode, pos: number) => boolean | void): void
              nodeAt(pos: number): PmNode | null
            }
            tr: { setSelection(sel: unknown): unknown }
          }
          view: { dispatch(tr: unknown): void }
        }
        __genofficeCellSelection?: { create(doc: unknown, anchor: number, head: number): unknown }
      }
      const editor = w.__genofficeWordEditor
      const CellSelection = w.__genofficeCellSelection
      if (!editor || !CellSelection) return false
      const { state, view } = editor
      const { doc } = state
      let tablePos = -1
      doc.descendants((node, pos) => {
        if (tablePos === -1 && node.type.name === 'table') {
          tablePos = pos
          return false
        }
        return true
      })
      if (tablePos === -1) return false
      const table = doc.nodeAt(tablePos)
      if (!table) return false
      const cellPos = (ri: number, ci: number): number | null => {
        let rowOffset = tablePos + 1
        for (let r = 0; r < table.childCount; r++) {
          const row = table.child(r)
          if (r === ri) {
            let cellOffset = rowOffset + 1
            for (let c = 0; c < row.childCount; c++) {
              if (c === ci) return cellOffset
              cellOffset += row.child(c).nodeSize
            }
            return null
          }
          rowOffset += row.nodeSize
        }
        return null
      }
      // Row 0 DOM cells: [rowspanned merged cell, "Filled center", "Top right"]
      // → select DOM cell indexes 1..2 of row 0.
      const anchor = cellPos(0, 1)
      const head = cellPos(0, 2)
      if (anchor === null || head === null) return false
      view.dispatch(state.tr.setSelection(CellSelection.create(doc, anchor, head)))
      return true
    })
    expect(selected).toBe(true)
    await page.getByRole('button', { name: 'Merge Cells', exact: true }).click()
    // Row 1 now has the rowspanned cell + ONE merged cell (colspan=2).
    await expect(page.locator(`${tableSel} tr`).first().locator('td')).toHaveCount(2)
    const merged = page.locator(`${tableSel} tr`).first().locator('td').nth(1)
    await expect(merged).toHaveAttribute('colspan', '2')
    await expect(merged).toContainText('Filled center')
    await expect(merged).toContainText('Top right')

    // Save → the XML carries gridSpan=2.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-table-merge.docx')).toBeVisible({ timeout: 15_000 })
    const blocksSaved = JSON.parse(req.postData() ?? '{}').blocks as WireBlock[]
    const tblSaved = blocksSaved.find((b) => b.type === 'table')
    expect(tblSaved?.edited).toBe(true)
    expect(tblSaved?.table?.rows[0][1].colSpan).toBe(2)

    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('<w:gridSpan w:val="2"/>')

    // ── Split: reopen, click the merged cell, split it back ────────────────
    const savedPath = '/tmp/e2e-table-merge-saved.docx'
    writeFileSync(savedPath, saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath)
    await expect(page.getByText('Opened e2e-table-merge-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    await reopenResponsePromise
    const reopenedSel = '.ProseMirror table'
    // The merged cell round-tripped with colspan=2.
    await expect(page.locator(`${reopenedSel} tr`).first().locator('td')).toHaveCount(2)
    await expect(page.locator(`${reopenedSel} tr`).first().locator('td').nth(1)).toHaveAttribute(
      'colspan',
      '2',
    )

    await page.locator(`${reopenedSel} tr`).first().locator('td').nth(1).click()
    await page.getByRole('button', { name: 'Split Cell', exact: true }).click()
    await expect(page.locator(`${reopenedSel} tr`).first().locator('td')).toHaveCount(3)

    const saveReq2 = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl2 = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await saveReq2
    const download2 = await dl2
    await expect(page.getByText('Saved e2e-table-merge-saved.docx')).toBeVisible({
      timeout: 15_000,
    })
    const stream2 = await download2.createReadStream()
    const chunks2: Buffer[] = []
    for await (const chunk of stream2) chunks2.push(chunk as Buffer)
    const saved2 = Buffer.concat(chunks2)
    const xml2 = await readZipEntry(saved2, 'word/document.xml')
    expect(xml2).not.toContain('<w:gridSpan w:val="2"/>')

    expect(pageErrors).toEqual([])
  })
})
