/**
 * REAL browser E2E — Review → Notes/Comments (Phase 4 Increment 6).
 *
 * Proves the legacy-note persistence chain end-to-end through the REAL
 * HTTP boundary:
 *
 *   open → readBasicWorkbook resolves + parses the comments part →
 *   WorksheetState.notes → browser installs each note via
 *   createOrUpdateNote under journal suppression (real Univer note UI:
 *   cell markers + popup) → user creates/edits/deletes through the REAL
 *   facade → note mutations mark the sheet note-dirty → save snapshots
 *   the LIVE note model as canonical SheetNote[] →
 *   savePlan.noteStates → /api/office/workbooks/save → routeOffice →
 *   applyCellEditsToXlsx arg 13 → applySheetNotes → XLSX (comments part
 *   + VML shapes + rels + content types) → reopen → same note state.
 *
 * No browser-side OOXML. The browser only ever exchanges typed SheetNote
 * snapshots taken from Univer's live note model.
 *
 * Fixture (buildExcelNotesFixture / buildExcelNotedFixture): a ledger
 * sheet with a bold header, 4 data rows, and (noted variant) a comments
 * part with two notes wired through the worksheet rels + VML shapes.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelNotesFixture, buildExcelNotedFixture, readZipEntry } from './fixtures'

/** Plain-data view of the live note model (page.evaluate cannot serialize
 *  the runtime's functions — all facade access happens inside evaluate). */
interface LiveNoteView {
  notes: Array<{ row: number; column: number; note: string }>
}

function readLiveNotes(page: import('@playwright/test').Page): Promise<LiveNoteView> {
  return page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook: () => {
              getActiveSheet: () => {
                getNotes: () => Array<{
                  row: number
                  col: number
                  note: string
                }>
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
    const notes = (ws?.getNotes?.() ?? []).map(({ row, col, note }) => ({
      row,
      column: col,
      note,
    }))
    return { notes }
  })
}

/** Create/update a note through the REAL facade (the same FRange
 *  createOrUpdateNote the note popup commits through). Univer's note map
 *  is keyed by id and updateNote looks up BY ID when one is passed — so
 *  editing an existing note must reuse ITS id (fetch getNote() first) or
 *  the update would create a duplicate at the same position. */
async function upsertNote(
  page: import('@playwright/test').Page,
  row: number,
  column: number,
  text: string,
): Promise<void> {
  await page.evaluate(
    ({ row, column, text }) => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => {
                    createOrUpdateNote: (note: {
                      id: string
                      row: number
                      col: number
                      width: number
                      height: number
                      note: string
                    }) => unknown
                    getNote: () => { id: string } | null
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const range = ws?.getRange?.(row, column)
      if (!range) return
      const existing = range.getNote?.()
      range.createOrUpdateNote?.({
        id: existing?.id ?? `e2e-note-${row}-${column}`,
        row,
        col: column,
        width: 220,
        height: 90,
        note: text,
      })
    },
    { row, column, text },
  )
}

/** Delete a note through the REAL facade (FRange.deleteNote). */
async function deleteNote(
  page: import('@playwright/test').Page,
  row: number,
  column: number,
): Promise<void> {
  await page.evaluate(
    ({ row, column }) => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (r: number, c: number) => { deleteNote: () => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      ws?.getRange?.(row, column)?.deleteNote?.()
    },
    { row, column },
  )
}

test.describe('Review tab — Notes persist through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1+2: opening a workbook with notes installs them in the real Univer note model', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNotedFixture()
    writeFileSync('/tmp/e2e-ribbon-notes-open.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-notes-open.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-notes-open.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // READ path: the snapshot carried the parsed notes.
    const snapshot = (await openResponse.json()).snapshot.sheets as Array<{
      name: string
      notes?: Array<{ row: number; column: number; author: string; text: string }>
    }>
    expect(snapshot[0].notes).toEqual([
      { row: 1, column: 1, author: 'Reviewer', text: 'Verify the fee <amount> & tax' },
      { row: 3, column: 0, author: '', text: 'second note' },
    ])

    // IMPORT path: the live Univer note model carries both notes with the
    // desktop's "Author:\nText" convention.
    await page.waitForTimeout(1500)
    const live = await readLiveNotes(page)
    expect(live.notes).toHaveLength(2)
    expect(live.notes).toContainEqual({
      row: 1,
      column: 1,
      note: 'Reviewer:\nVerify the fee <amount> & tax',
    })
    expect(live.notes).toContainEqual({ row: 3, column: 0, note: 'second note' })

    // Loading notes must NOT create an Undo entry (journal suppression +
    // the undo filter): the workbook starts clean.
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })

    expect(pageErrors).toEqual([])
  })

  test('3+4: create a note then edit it through the real facade', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNotesFixture()
    writeFileSync('/tmp/e2e-ribbon-notes-create.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-notes-create.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-notes-create.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Create a note on B3.
    await upsertNote(page, 2, 1, 'Check this tax figure')
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    let live = await readLiveNotes(page)
    expect(live.notes).toEqual([{ row: 2, column: 1, note: 'Check this tax figure' }])

    // Edit the same note.
    await upsertNote(page, 2, 1, 'Check this tax figure — confirmed')
    await page.waitForTimeout(600)
    live = await readLiveNotes(page)
    expect(live.notes).toEqual([{ row: 2, column: 1, note: 'Check this tax figure — confirmed' }])

    expect(pageErrors).toEqual([])
  })

  test('5+6: two notes — deleting one keeps the other', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNotesFixture()
    writeFileSync('/tmp/e2e-ribbon-notes-delete.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-notes-delete.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-notes-delete.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    await upsertNote(page, 2, 1, 'note one')
    await page.waitForTimeout(300)
    await upsertNote(page, 4, 0, 'note two')
    await page.waitForTimeout(600)
    let live = await readLiveNotes(page)
    expect(live.notes).toHaveLength(2)

    // Delete ONLY the first.
    await deleteNote(page, 2, 1)
    await page.waitForTimeout(600)
    live = await readLiveNotes(page)
    expect(live.notes).toEqual([{ row: 4, column: 0, note: 'note two' }])

    expect(pageErrors).toEqual([])
  })

  test('7+8+9: save/reopen — typed noteStates on the wire, XML inspected, untouched note survives an edit', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Start from the fixture with TWO existing notes; edit ONE of them.
    const fixture = await buildExcelNotedFixture()
    writeFileSync('/tmp/e2e-ribbon-notes-save.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-notes-save.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-notes-save.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Edit the B2 note (keep the "Author:\nText" convention).
    await upsertNote(page, 1, 1, 'Reviewer:\nVerify the fee — DONE')
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the plan must carry the typed canonical notes (BOTH — the edit
    // plus the untouched A4 note).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        noteStates?: Array<{
          sheetName: string
          notes: Array<{ row: number; column: number; author: string; text: string }>
        }>
      }
    }
    expect(saveBody.savePlan.noteStates, 'typed noteStates on the wire').toHaveLength(1)
    const state = saveBody.savePlan.noteStates![0]!
    expect(state.sheetName).toBe('Ledger')
    expect(state.notes).toHaveLength(2)
    expect(state.notes).toContainEqual({
      row: 1,
      column: 1,
      author: 'Reviewer',
      text: 'Verify the fee — DONE',
    })
    // The untouched note rides along verbatim.
    expect(state.notes).toContainEqual({ row: 3, column: 0, author: '', text: 'second note' })

    // Saved XLSX: the comments part carries both notes.
    const comments = await readZipEntry(saved, 'xl/comments1.xml')
    expect(comments).toContain('<comment ref="B2" authorId="0">')
    expect(comments).toContain('Verify the fee — DONE')
    expect(comments).toContain('<comment ref="A4" authorId="1">')
    expect(comments).toContain('second note')
    // The VML drawing still carries the note shapes + the legacyDrawing
    // element survives on the worksheet.
    const vml = await readZipEntry(saved, 'xl/drawings/vmlDrawing1.vml')
    expect(vml.match(/ObjectType="Note"/g)).toHaveLength(2)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<legacyDrawing r:id="rId6"/>')
    // Unrelated cell values survive.
    expect(sheet1).toMatch(/<c r="B5"[^>]*><v>17<\/v>/)

    // Reopen: the snapshot + live model carry both notes.
    writeFileSync('/tmp/e2e-ribbon-notes-save-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-notes-save-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-notes-save-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown }>
      notes?: Array<{ row: number; column: number; author: string; text: string }>
    }>
    expect(reopened[0].notes).toHaveLength(2)
    expect(reopened[0].notes).toContainEqual({
      row: 1,
      column: 1,
      author: 'Reviewer',
      text: 'Verify the fee — DONE',
    })
    expect(reopened[0].cells.B5?.value).toBe(17)

    await page.waitForTimeout(1500)
    const live = await readLiveNotes(page)
    expect(live.notes).toHaveLength(2)
    expect(live.notes).toContainEqual({
      row: 1,
      column: 1,
      note: 'Reviewer:\nVerify the fee — DONE',
    })
    expect(live.notes).toContainEqual({ row: 3, column: 0, note: 'second note' })

    expect(pageErrors).toEqual([])
  })

  test('10: no-op save preserves the note XML byte-identically', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNotedFixture()
    writeFileSync('/tmp/e2e-ribbon-notes-noop.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-notes-noop.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-notes-noop.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // The workbook starts clean (note install was journal-suppressed) —
    // no Save possible without an edit. Make an UNRELATED cell edit.
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    r: number,
                    c: number,
                  ) => {
                    setValueForCell: (v: unknown) => unknown
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI
        ?.getActiveWorkbook?.()
        ?.getActiveSheet?.()
        ?.getRange?.(0, 5)
        ?.setValueForCell?.('note-unrelated')
    })
    await page.waitForTimeout(500)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the plan carries the cell edit but NO noteStates (the sheet is
    // not note-dirty).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { noteStates?: unknown[] }
    }
    expect(saveBody.savePlan.noteStates, 'no note state on a no-note-change save').toBeUndefined()

    // The saved XLSX preserves the comments part verbatim.
    const comments = await readZipEntry(saved, 'xl/comments1.xml')
    expect(comments).toContain(
      '<comment ref="B2" authorId="0"><text><t>Verify the fee &lt;amount&gt; &amp; tax</t></text></comment>',
    )
    expect(comments).toContain(
      '<comment ref="A4" authorId="1"><text><t>second note</t></text></comment>',
    )

    expect(pageErrors).toEqual([])
  })
})
