/**
 * REAL browser E2E — EXCEL-027 Advanced cell formatting (borders, line
 * styles, border colors, text rotation).
 *
 * Proves the advanced formatting surface end-to-end through the canonical
 * pipeline:
 *
 *   browser (Univer) → Vite proxy → HTTP /api/office/workbooks/save →
 *   routeOffice (strict WorkbookStyleEdit validation) →
 *   applyCellEditsToXlsx (StylesheetEditor copy-on-write) → XLSX bytes →
 *   reopen → /api/office/workbooks/open → StylesheetReader → snapshot →
 *   live model
 *
 * Every mandated scenario inspects the LIVE MODEL (public FRange
 * getCellStyleData), the SAVE REQUEST (the journaled canonical deltas), the
 * SAVED XLSX XML (exact border/rotation attributes), and the REOPENED
 * model — never merely the button state.
 *
 * Scenario matrix (handoff §Browser E2E):
 *   1  import existing borders ........ test '1+2+7+14'
 *   2  render borders correctly ....... test '1+2+7+14' (pixel proof)
 *   3  edit one border side ........... test '3+5+6'
 *   4  clear a border side ............ test '4'
 *   5  preserve unrelated sides ....... test '3+5+6' + '11'
 *   6  persist border style/color ..... test '3+5+6'
 *   7  import existing rotation ....... test '1+2+7+14'
 *   8  edit rotation .................. test '8+10' (positive/negative)
 *   9  clear rotation ................. test '9' (vertical + clear)
 *   10 persist rotation ............... tests '8+10' / '9'
 *   11 combined border+rotation+fmt ... test '11'
 *   12 no-op preservation ............. test '12+13+14'
 *   13 unrelated edit preserves fmt ... test '12+13+14'
 *   14 malformed/unsupported closed ... tests '1+2+7+14' / '12+13+14'
 */
import { test, expect, type Page } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelAdvancedFormatFixture, readZipEntry } from './fixtures'

const GRID = '#genoffice-web-excel canvas'

/** One live border edge as the engine reports it (IStyleData.bd). */
interface LiveEdge {
  readonly s?: number
  readonly cl?: { readonly rgb?: string }
}
interface LiveStyle {
  readonly bd?: {
    readonly t?: LiveEdge
    readonly b?: LiveEdge
    readonly l?: LiveEdge
    readonly r?: LiveEdge
  } | null
  readonly tr?: { readonly a?: number; readonly v?: number } | null
  readonly pd?: { readonly l?: number } | null
  readonly n?: { readonly pattern?: string } | null
  readonly bl?: number
  readonly bg?: { readonly rgb?: string }
  readonly ht?: number
  readonly vt?: number
  readonly tb?: number
}

/** Read a cell's live resolved style through the PUBLIC FRange facade. */
async function liveStyle(page: Page, ref: string): Promise<LiveStyle | null> {
  return page.evaluate((cellRef) => {
    const rt = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook?: () => {
              getActiveSheet?: () => {
                getRange?: (ref: string) => { getCellStyleData?: () => unknown }
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const range = rt?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getRange?.(cellRef)
    return (range?.getCellStyleData?.() as LiveStyle | null) ?? null
  }, ref)
}

/** Select a cell through the shell's name box (no pixel-coordinate math). */
async function selectCell(page: Page, ref: string): Promise<void> {
  const box = page.locator('[data-testid="excel-name-box"]')
  await box.click()
  await box.fill(ref)
  await box.press('Enter')
  await page.waitForTimeout(300)
}

/** Screenshot a fixed clip of the grid's first-row region. */
async function gridClip(page: Page): Promise<Buffer> {
  const grid = page.locator(GRID).last()
  const bbox = await grid.boundingBox()
  if (!bbox) throw new Error('grid canvas not found')
  return page.screenshot({ clip: { x: bbox.x, y: bbox.y, width: 900, height: 160 } })
}

/** A per-cell snapshot style record (index-signature access on purpose). */
type SheetStyles = Record<string, Record<string, unknown> | undefined>

/** Open a fixture and return the open-API snapshot's Data sheet styles. */
async function openFixture(page: Page, bytes: Buffer, fileName: string): Promise<SheetStyles> {
  writeFileSync(`/tmp/${fileName}`, bytes)
  const openResponsePromise = page.waitForResponse(
    (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
  )
  await page.setInputFiles('input[type="file"]', `/tmp/${fileName}`)
  await expect(page.getByText(`Opened ${fileName}`)).toBeVisible({ timeout: 30_000 })
  const openResponse = await openResponsePromise
  expect(openResponse.status()).toBe(200)
  await page.waitForTimeout(1200)
  const snapshot = (await openResponse.json()).snapshot as {
    sheets: Array<{ name: string; styles?: Record<string, Record<string, unknown>> }>
  }
  const data = snapshot.sheets.find((s) => s.name === 'Data')
  return (data?.styles ?? {}) as SheetStyles
}

/** Save + return { savedBytes, savePlan } from the captured request. */
async function saveAndCapturePlan(
  page: Page,
): Promise<{ bytes: Buffer; plan: Record<string, unknown> }> {
  const saveReq = page.waitForRequest(
    (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
  )
  const bytes = await clickSaveAndCaptureDownload(page, 'Save')
  const req = await saveReq
  const body = JSON.parse(req.postData() ?? '{}') as { savePlan?: Record<string, unknown> }
  return { bytes, plan: body.savePlan ?? {} }
}

/** The journaled style of one cell from the captured save plan. */
function plannedStyle(plan: Record<string, unknown>, row: number, column: number): unknown {
  const edits = (plan.edits as Array<{ row: number; column: number; style?: unknown }>) ?? []
  return edits.find((e) => e.row === row && e.column === column)?.style
}

test.describe('EXCEL-027 — advanced formatting (borders + rotation)', () => {
  test('1+2+7+14: borders and rotation import, render, and fail closed for unmodelable styles', async ({
    page,
  }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const styles = await openFixture(
      page,
      await buildExcelAdvancedFormatFixture(),
      'e2e-advfmt-import.xlsx',
    )

    // ── Scenario 1: the file's borders import into the snapshot. ────────
    expect(styles.A1, 'mixed sides with colors').toEqual({
      border: {
        top: { style: 'thin', color: 'C00000' },
        bottom: { style: 'medium', color: '0000FF' },
      },
    })
    expect(styles.B1, 'single modeled edge (diagonal skipped)').toEqual({
      border: { left: { style: 'dashed' } },
    })
    // ── Scenario 7: the file's rotations import (45 / 135 / 255 / indent). ──
    expect(styles.C1).toEqual({ textRotation: 45 })
    expect(styles.D1).toEqual({ textRotation: 135 })
    expect(styles.E1).toEqual({ textRotation: 'vertical' })
    expect(styles.F1).toEqual({ indent: 2 })
    // The combined cell imports every family at once.
    expect(styles.G1).toEqual({
      bold: true,
      fillColor: 'FFF2CC',
      numberFormat: '#,##0.000',
      horizontalAlign: 'center',
      verticalAlign: 'bottom',
      wrapText: true,
      textRotation: 60,
      border: {
        top: { style: 'double', color: '00B050' },
        bottom: { style: 'double', color: '00B050' },
        left: { style: 'double', color: '00B050' },
        right: { style: 'double', color: '00B050' },
      },
    })
    // ── Scenario 14 (import side): the unmodelable "wave" border is NOT
    // modeled — the snapshot carries nothing for H1.
    expect(styles.H1).toBeUndefined()

    // ── The live model carries the imported state (public read-back). ───
    const a1 = await liveStyle(page, 'A1')
    expect(a1?.bd?.t).toEqual({ s: 1, cl: { rgb: '#C00000' } })
    expect(a1?.bd?.b).toEqual({ s: 8, cl: { rgb: '#0000FF' } })
    const c1 = await liveStyle(page, 'C1')
    expect(c1?.tr).toEqual({ a: 45 })
    const e1 = await liveStyle(page, 'E1')
    expect(e1?.tr).toEqual({ a: 0, v: 1 })
    const f1 = await liveStyle(page, 'F1')
    expect(f1?.pd).toEqual({ l: 16 })
    const g1 = await liveStyle(page, 'G1')
    expect(g1?.bl).toBe(1)
    expect(g1?.bg?.rgb).toBe('#FFF2CC')
    expect(g1?.n?.pattern).toBe('#,##0.000')
    expect(g1?.tr).toEqual({ a: 60 })
    expect(g1?.bd?.l).toEqual({ s: 7, cl: { rgb: '#00B050' } })
    const h1 = await liveStyle(page, 'H1')
    expect(h1?.bd, 'wave border never seeds the live model').toBeUndefined()

    // ── Scenario 2: borders and rotation RENDER. Pixel proof — clearing
    // the border/rotation in-session must CHANGE the rendered pixels (if
    // they did not render, clearing them could not change anything).
    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(200)
    const withBorder = await gridClip(page)
    await selectCell(page, 'A1')
    await page.getByRole('combobox', { name: 'Border preset' }).selectOption('none')
    await page.waitForTimeout(700)
    const borderCleared = await gridClip(page)
    expect(
      withBorder.equals(borderCleared),
      'clearing the imported border must change the render',
    ).toBe(false)

    const stacked = await liveStyle(page, 'E1')
    expect(stacked?.tr, 'E1 still stacked (untouched)').toEqual({ a: 0, v: 1 })
    // Rotation render proof via the whole-row clip: E1's stacked text is in
    // the same 900px clip; clearing its rotation changes pixels again.
    await selectCell(page, 'E1')
    await page.getByRole('combobox', { name: 'Orientation' }).selectOption('0')
    await page.waitForTimeout(700)
    const rotationCleared = await gridClip(page)
    expect(
      borderCleared.equals(rotationCleared),
      'clearing the stacked rotation must change the render',
    ).toBe(false)
    const clearedE1 = await liveStyle(page, 'E1')
    // The engine models a cleared rotation as a zero angle (the JOURNAL
    // maps {a: 0} to the canonical 0 clear sentinel — proven in test 9).
    expect(clearedE1?.tr).toEqual({ a: 0 })

    expect(pageErrors).toEqual([])
  })

  test('3+5+6: edit one border side, keep the others, persist style and color', async ({
    page,
  }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const styles = await openFixture(
      page,
      await buildExcelAdvancedFormatFixture(),
      'e2e-advfmt-edit.xlsx',
    )
    expect(styles.A1?.border).toBeDefined()

    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(200)

    // ── Scenario 3: edit ONE side of A1 (top+bottom exist → add right). ──
    await selectCell(page, 'A1')
    await page.getByRole('combobox', { name: 'Border preset' }).selectOption('right')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const a1Live = await liveStyle(page, 'A1')
    expect(a1Live?.bd?.r, 'right edge set in-session').toEqual({ s: 1, cl: { rgb: '#000000' } })
    expect(a1Live?.bd?.t, 'top edge untouched in-session').toEqual({ s: 1, cl: { rgb: '#C00000' } })

    // ── Line style proof: a dashed left border on the plain cell I1. ────
    await page.getByRole('combobox', { name: 'Border line style' }).selectOption('dashed')
    await selectCell(page, 'I1')
    await page.getByRole('combobox', { name: 'Border preset' }).selectOption('left')
    await page.waitForTimeout(400)
    const i1Live = await liveStyle(page, 'I1')
    expect(i1Live?.bd?.l, 'dashed left edge in-session').toEqual({ s: 4, cl: { rgb: '#000000' } })

    const { bytes, plan } = await saveAndCapturePlan(page)

    // ── The save plan carries ONLY the edited edges (side isolation at
    // the journal level). A1's right edge + the engine's adjacent clear of
    // B1's shared left edge (Univer's own outline-border semantics — the
    // same mutation the desktop journals); I1's dashed left.
    const a1Style = plannedStyle(plan, 0, 0) as Record<string, unknown>
    expect(a1Style?.borderRight).toEqual({ style: 'thin', color: '#000000' })
    expect(a1Style?.borderTop, 'journal carries only the right edge for A1').toBeUndefined()
    const b1Style = plannedStyle(plan, 0, 1) as Record<string, unknown>
    expect(b1Style?.borderLeft, "engine clears the neighbor's shared edge").toBeNull()
    const i1Style = plannedStyle(plan, 0, 8) as Record<string, unknown>
    expect(i1Style?.borderLeft).toEqual({ style: 'dashed', color: '#000000' })

    // ── Scenario 5+6: the saved XML keeps A1's top/bottom (untouched
    // sides) and adds the right edge — one border entry with all three. ──
    const stylesXml = await readZipEntry(bytes, 'xl/styles.xml')
    expect(stylesXml).toContain(
      '<left/><right style="thin"><color rgb="FF000000"/></right>' +
        '<top style="thin"><color rgb="FFC00000"/></top>' +
        '<bottom style="medium"><color rgb="FF0000FF"/></bottom><diagonal/>',
    )
    // The dashed border entry for I1.
    expect(stylesXml).toContain('<left style="dashed"><color rgb="FF000000"/></left>')

    // ── Scenario 6: save/reopen — the snapshot and live model carry the
    // full three-sided border and the dashed left edge. ──────────────────
    const reopened = await openFixture(page, bytes, 'e2e-advfmt-edit-saved.xlsx')
    expect(reopened.A1).toEqual({
      border: {
        top: { style: 'thin', color: 'C00000' },
        bottom: { style: 'medium', color: '0000FF' },
        right: { style: 'thin', color: '000000' },
      },
    })
    expect(reopened.I1).toEqual({ border: { left: { style: 'dashed', color: '000000' } } })
    // B1's left edge was cleared by the engine's own semantics; its
    // modeled snapshot border is gone (the diagonal is preserved in the
    // file but deliberately not modeled).
    expect(reopened.B1?.border).toBeUndefined()
    const a1Reopened = await liveStyle(page, 'A1')
    expect(a1Reopened?.bd?.r).toEqual({ s: 1, cl: { rgb: '#000000' } })
    expect(a1Reopened?.bd?.t).toEqual({ s: 1, cl: { rgb: '#C00000' } })

    expect(pageErrors).toEqual([])
  })

  test('4: clear a border side through No Border and persist the clear', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    await openFixture(page, await buildExcelAdvancedFormatFixture(), 'e2e-advfmt-clear.xlsx')

    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(200)

    // A1 carries top thin red + bottom medium blue. No Border clears every
    // side (the only per-side clear the desktop surface offers).
    await selectCell(page, 'A1')
    await page.getByRole('combobox', { name: 'Border preset' }).selectOption('none')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const cleared = await liveStyle(page, 'A1')
    expect(cleared?.bd ?? null, 'borders cleared in-session').toBeNull()

    const { bytes, plan } = await saveAndCapturePlan(page)

    // The journal carries the four null edges (a REAL clear — the route
    // passes nulls through to the engine; the old route dropped them).
    const a1Style = plannedStyle(plan, 0, 0) as Record<string, unknown>
    expect(a1Style?.borderTop).toBeNull()
    expect(a1Style?.borderBottom).toBeNull()
    expect(a1Style?.borderLeft).toBeNull()
    expect(a1Style?.borderRight).toBeNull()

    // The saved XML: the clear dedups to the EXISTING empty border 0 (the
    // derived xf equals xf 0, so no new border entry and no s attribute on
    // A1 — the borders section is unchanged at count=5).
    const sheet1 = await readZipEntry(bytes, 'xl/worksheets/sheet1.xml')
    const a1Tag = /<c r="A1"[^>]*(?:\/>|>)/.exec(sheet1)?.[0] ?? ''
    expect(a1Tag, 'A1 present in the saved sheet').toContain('<c r="A1"')
    expect(
      /\bs="([1-9]\d*)"/.test(a1Tag),
      'A1 carries no non-default style index after the full clear',
    ).toBe(false)
    const stylesXml = await readZipEntry(bytes, 'xl/styles.xml')
    // B1's shared left edge was cleared by the engine's own outline
    // semantics — the new border entry keeps the DIAGONAL and its
    // diagonalUp/Down attributes verbatim (the untouched-edges invariant).
    expect(stylesXml).toContain(
      '<border diagonalDown="1" diagonalUp="1"><left/><right/><top/><bottom/>' +
        '<diagonal style="hair"/></border>',
    )

    // Reopen: no border anywhere on A1.
    const reopened = await openFixture(page, bytes, 'e2e-advfmt-clear-saved.xlsx')
    expect(reopened.A1).toBeUndefined()
    const reopenedLive = await liveStyle(page, 'A1')
    expect(reopenedLive?.bd ?? null, 'no border in the reopened live model').toBeNull()

    expect(pageErrors).toEqual([])
  })

  test('8+10: edit rotation (positive and negative) and persist through save/reopen', async ({
    page,
  }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    await openFixture(page, await buildExcelAdvancedFormatFixture(), 'e2e-advfmt-rot.xlsx')

    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(200)

    // ── Scenario 8: positive rotation (45° counterclockwise) on I1. ─────
    await selectCell(page, 'I1')
    await page.getByRole('combobox', { name: 'Orientation' }).selectOption('45')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const live45 = await liveStyle(page, 'I1')
    expect(live45?.tr).toEqual({ a: 45 })

    let { bytes, plan } = await saveAndCapturePlan(page)
    expect(plannedStyle(plan, 0, 8)).toEqual({ textRotation: 45 })
    let stylesXml = await readZipEntry(bytes, 'xl/styles.xml')
    expect(stylesXml).toContain('textRotation="45"')
    let reopened = await openFixture(page, bytes, 'e2e-advfmt-rot-saved.xlsx')
    expect(reopened.I1).toEqual({ textRotation: 45 })
    let live = await liveStyle(page, 'I1')
    expect(live?.tr).toEqual({ a: 45 })

    // ── Negative rotation (−45° clockwise) journals as the OOXML 135 form. ──
    await selectCell(page, 'I1')
    await page.getByRole('combobox', { name: 'Orientation' }).selectOption('-45')
    await page.waitForTimeout(500)
    ;({ bytes, plan } = await saveAndCapturePlan(page))
    expect(plannedStyle(plan, 0, 8)).toEqual({ textRotation: 135 })
    stylesXml = await readZipEntry(bytes, 'xl/styles.xml')
    expect(stylesXml).toContain('textRotation="135"')
    reopened = await openFixture(page, bytes, 'e2e-advfmt-rot-saved.xlsx')
    expect(reopened.I1).toEqual({ textRotation: 135 })
    live = await liveStyle(page, 'I1')
    expect(live?.tr).toEqual({ a: -45 })

    expect(pageErrors).toEqual([])
  })

  test('9: vertical text and clear rotation persist through save/reopen', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    await openFixture(page, await buildExcelAdvancedFormatFixture(), 'e2e-advfmt-vert.xlsx')

    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(200)

    // Vertical (stacked) text on the plain cell.
    await selectCell(page, 'I1')
    await page.getByRole('combobox', { name: 'Orientation' }).selectOption('vertical')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const stacked = await liveStyle(page, 'I1')
    expect(stacked?.tr).toEqual({ a: 0, v: 1 })

    const { bytes, plan } = await saveAndCapturePlan(page)
    expect(plannedStyle(plan, 0, 8)).toEqual({ textRotation: 255 })
    const stylesXml = await readZipEntry(bytes, 'xl/styles.xml')
    expect(stylesXml).toContain('textRotation="255"')
    const reopened = await openFixture(page, bytes, 'e2e-advfmt-vert-saved.xlsx')
    expect(reopened.I1).toEqual({ textRotation: 'vertical' })
    const stackedLive = await liveStyle(page, 'I1')
    expect(stackedLive?.tr).toEqual({ a: 0, v: 1 })

    // ── Scenario 9: clear the rotation (0 sentinel drops the attribute). ──
    await selectCell(page, 'I1')
    await page.getByRole('combobox', { name: 'Orientation' }).selectOption('0')
    await page.waitForTimeout(500)
    const { bytes: clearedBytes, plan: clearedPlan } = await saveAndCapturePlan(page)
    expect(plannedStyle(clearedPlan, 0, 8)).toEqual({ textRotation: 0 })
    const clearedReopened = await openFixture(page, clearedBytes, 'e2e-advfmt-vert-cleared.xlsx')
    expect(clearedReopened.I1).toBeUndefined()
    const clearedLive = await liveStyle(page, 'I1')
    // A cleared rotation is either dropped from the live style or the
    // engine's zero-angle form — never stacked, never angled.
    expect(clearedLive?.tr?.v, 'no stacked flag after clear').toBeUndefined()
    expect(clearedLive?.tr?.a ?? 0, 'no angle after clear').toBe(0)

    expect(pageErrors).toEqual([])
  })

  test('11: a border edit on the combined cell preserves every other family', async ({ page }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // G1 = bold + fill + number format + alignment + border (double green
    // ×4) + rotation 60. Apply a TOP border edit — nothing else may move.
    await openFixture(page, await buildExcelAdvancedFormatFixture(), 'e2e-advfmt-combo.xlsx')

    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.waitForTimeout(200)

    await selectCell(page, 'G1')
    await page.getByRole('combobox', { name: 'Border preset' }).selectOption('top')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const g1Live = await liveStyle(page, 'G1')
    expect(g1Live?.bd?.t).toEqual({ s: 1, cl: { rgb: '#000000' } })
    expect(g1Live?.bd?.b).toEqual({ s: 7, cl: { rgb: '#00B050' } })
    expect(g1Live?.bl).toBe(1)
    expect(g1Live?.n?.pattern).toBe('#,##0.000')
    expect(g1Live?.tr).toEqual({ a: 60 })

    const { bytes, plan } = await saveAndCapturePlan(page)
    const g1Style = plannedStyle(plan, 0, 6) as Record<string, unknown>
    expect(g1Style?.borderTop).toEqual({ style: 'thin', color: '#000000' })
    expect(g1Style?.bold, 'the border edit journals no font delta').toBeUndefined()
    expect(g1Style?.textRotation, 'the border edit journals no rotation delta').toBeUndefined()

    const reopened = await openFixture(page, bytes, 'e2e-advfmt-combo-saved.xlsx')
    expect(reopened.G1).toEqual({
      bold: true,
      fillColor: 'FFF2CC',
      numberFormat: '#,##0.000',
      horizontalAlign: 'center',
      verticalAlign: 'bottom',
      wrapText: true,
      textRotation: 60,
      border: {
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'double', color: '00B050' },
        left: { style: 'double', color: '00B050' },
        right: { style: 'double', color: '00B050' },
      },
    })
    const g1Reopened = await liveStyle(page, 'G1')
    expect(g1Reopened?.bd?.t).toEqual({ s: 1, cl: { rgb: '#000000' } })
    expect(g1Reopened?.n?.pattern).toBe('#,##0.000')
    expect(g1Reopened?.tr).toEqual({ a: 60 })

    expect(pageErrors).toEqual([])
  })

  test('12+13+14: no-op preservation, unrelated edits, and the unmodelable border stays closed', async ({
    page,
  }) => {
    test.setTimeout(150_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelAdvancedFormatFixture()
    const styles = await openFixture(page, fixture, 'e2e-advfmt-noop.xlsx')
    expect(styles.A1?.border).toBeDefined()

    // ── Scenario 13: edit an UNRELATED cell (I3 = row 2, col 8) — a pure
    // value edit through the formula bar. ─────────────────────────────────
    await selectCell(page, 'I3')
    const bar = page.locator('[data-testid="excel-formula-bar"]')
    await bar.click()
    await bar.fill('unrelated')
    await bar.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const { bytes, plan } = await saveAndCapturePlan(page)
    // Only the value edit journals — no border/rotation deltas anywhere.
    const edits = (plan.edits as Array<{ row: number; column: number; style?: unknown }>) ?? []
    expect(edits.filter((e) => e.style !== undefined)).toEqual([])
    expect(edits.filter((e) => e.row === 2 && e.column === 8)).toHaveLength(1)

    // ── Scenario 12: no-op preservation — a value-only save never touches
    // styles.xml (byte-identical: no new xf, no new border) and never
    // touches the sibling sheet. ──────────────────────────────────────────
    const fixtureStyles = await readZipEntry(fixture, 'xl/styles.xml')
    const savedStyles = await readZipEntry(bytes, 'xl/styles.xml')
    expect(savedStyles, 'styles.xml byte-identical through a value-only save').toBe(fixtureStyles)
    const fixtureSheet2 = await readZipEntry(fixture, 'xl/worksheets/sheet2.xml')
    const savedSheet2 = await readZipEntry(bytes, 'xl/worksheets/sheet2.xml')
    expect(savedSheet2, 'untouched sibling sheet byte-identical').toBe(fixtureSheet2)

    // ── Scenario 14: the unmodelable "wave" border stays closed AND
    // byte-preserved (it is inside the untouched styles.xml above — assert
    // it explicitly for the record). ─────────────────────────────────────
    expect(savedStyles).toContain('<left style="wave"/>')

    // Reopen: every advanced format survives the unrelated edit intact.
    const reopened = await openFixture(page, bytes, 'e2e-advfmt-noop-saved.xlsx')
    expect(reopened.A1).toEqual({
      border: {
        top: { style: 'thin', color: 'C00000' },
        bottom: { style: 'medium', color: '0000FF' },
      },
    })
    expect(reopened.C1).toEqual({ textRotation: 45 })
    expect(reopened.E1).toEqual({ textRotation: 'vertical' })
    expect((reopened.G1?.border as Record<string, unknown> | undefined)?.left).toEqual({
      style: 'double',
      color: '00B050',
    })
    expect(reopened.G1?.textRotation).toBe(60)
    expect(reopened.H1, 'wave border still not modeled after save/reopen').toBeUndefined()

    expect(pageErrors).toEqual([])
  })
})
