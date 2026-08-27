/**
 * REAL browser E2E — Insert → Chart / chart edit (EXCEL-023).
 *
 * Proves the chart persistence chain end-to-end through the REAL HTTP
 * boundary:
 *
 *   open → readBasicWorkbook resolves the drawing relationship chain →
 *   WorksheetState.charts (both locators + canonical ChartVisualState) →
 *   browser seeds its chart store and floats the SVG frames over the
 *   grid through the PUBLIC registerComponent + addFloatDomToRange
 *   facades (journal-free) → user selects/moves/resizes/edits/deletes →
 *   save snapshots the LIVE state → chartEdits (semantic, keyed by
 *   chartPath) + visualEdits (geometry, keyed by the anchor locator) +
 *   visualAdditions.chart (session creations) → /api/office/workbooks/save
 *   → routeOffice strict validation → applyCellEditsToXlsx →
 *   applyChartEdit / applyVisualEdits / applyVisualAdditions rewrite the
 *   chart + drawing XML (delete cascades the chart part, its rels, and
 *   the content-type override) → reopen → file-native chart state.
 *
 * Unsupported charts (3-D plots, absolute anchors) fail closed: never
 * surfaced, never relocated, and a no-op save preserves their bytes
 * byte-for-byte.
 */
import { test, expect, type Page } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import {
  buildExcelChartFixture,
  buildExcelChartsFixture,
  buildExcel3DChartFixture,
  buildExcelOneCellChartFixture,
  buildExcelAbsoluteChartFixture,
  buildExcelChartDataOnlyFixture,
  readZipEntry,
  readZipEntryBytes,
  listZipEntries,
} from './fixtures'

const OPEN_INPUT = 'input[accept=".xlsx,.csv,.xls"]'
const NAME_BOX = 'input[data-testid="excel-name-box"]'

/** Snapshot view of the open response's chart state. */
interface ChartSnapshotView {
  sheets: Array<{
    name: string
    charts?: Array<{
      drawingPath: string
      drawingIndex: number
      chartPath: string
      anchorType: 'two-cell' | 'one-cell'
      anchor: Record<string, number>
      widthPx?: number
      heightPx?: number
      chart: {
        chartTypes: string[]
        title: string
        series: Array<{
          name: string
          categories: string[]
          values: number[]
          valuesRef?: string
          categoriesRef?: string
        }>
      }
    }>
  }>
}

async function openFixture(page: Page, bytes: Buffer, path: string): Promise<ChartSnapshotView> {
  writeFileSync(path, bytes)
  const openResponsePromise = page.waitForResponse(
    (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
  )
  await page.setInputFiles(OPEN_INPUT, path)
  await expect(page.getByText(`Opened ${path.split('/').pop()!}`)).toBeVisible({
    timeout: 30_000,
  })
  const openResponse = await openResponsePromise
  expect(openResponse.status()).toBe(200)
  await page.waitForTimeout(1200)
  return ((await openResponse.json()).snapshot ?? {}) as ChartSnapshotView
}

/** The floating chart frames currently rendered over the grid. */
async function chartFrames(
  page: Page,
): Promise<
  Array<{ key: string; title: string; selected: boolean; width: number; height: number }>
> {
  return page.evaluate(() => {
    const frames = Array.from(document.querySelectorAll<HTMLElement>('.chart-frame'))
    return frames.map((frame) => {
      const parent = frame.parentElement
      return {
        key: frame.dataset.chartKey ?? '',
        title: frame.querySelector('.chart-title')?.textContent ?? '',
        selected: frame.classList.contains('chart-selected'),
        width: parent?.clientWidth ?? frame.clientWidth,
        height: parent?.clientHeight ?? frame.clientHeight,
      }
    })
  })
}

async function selectRange(page: Page, ref: string): Promise<void> {
  const box = page.locator(NAME_BOX)
  await box.click()
  await box.fill(ref)
  await box.press('Enter')
  await page.waitForTimeout(400)
}

/** Drags a chart frame by (dx, dy) pixels — real mouse events. */
async function dragChart(page: Page, key: string, dx: number, dy: number): Promise<void> {
  const frame = page.locator(`.chart-frame[data-chart-key="${key}"]`)
  await expect(frame).toBeVisible()
  const box = await frame.boundingBox()
  expect(box).not.toBeNull()
  if (box === null) return
  const cx = box.x + box.width / 2
  const cy = box.y + box.height / 2
  await page.mouse.move(cx, cy)
  await page.mouse.down()
  await page.mouse.move(cx + dx, cy + dy, { steps: 8 })
  await page.mouse.up()
  await page.waitForTimeout(500)
}

/** Drags a resize handle by (dx, dy) pixels — real mouse events. */
async function resizeChart(
  page: Page,
  key: string,
  corner: string,
  dx: number,
  dy: number,
): Promise<void> {
  const frame = page.locator(`.chart-frame[data-chart-key="${key}"]`)
  await expect(frame).toBeVisible()
  const handle = page.locator(
    `.chart-frame[data-chart-key="${key}"] .chart-handle.handle-${corner}`,
  )
  await expect(handle).toBeVisible()
  const box = await handle.boundingBox()
  expect(box).not.toBeNull()
  if (box === null) return
  const cx = box.x + box.width / 2
  const cy = box.y + box.height / 2
  await page.mouse.move(cx, cy)
  await page.mouse.down()
  await page.mouse.move(cx + dx, cy + dy, { steps: 8 })
  await page.mouse.up()
  await page.waitForTimeout(500)
}

/** Saves and returns the mutated XLSX bytes (download capture). */
async function saveWorkbookBytes(page: Page): Promise<Buffer> {
  return clickSaveAndCaptureDownload(page, 'Save')
}

/** Types a trivial cell edit to mark the workbook dirty — the EXCEL-022
 *  ribbon-images precedent: byte-preservation proof rides a cell-edit
 *  save (Save is disabled while the workbook has no unsaved work). Row 1
 *  is chart-free in every chart fixture (charts anchor from row 2). */
async function markDirtyWithCellEdit(page: Page): Promise<void> {
  await page.locator('#genoffice-web-excel').click({ position: { x: 320, y: 8 } })
  await page.keyboard.type('noop')
  await page.keyboard.press('Enter')
  await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
}

test.describe('Insert tab — Charts persist through the canonical pipeline', () => {
  test.setTimeout(300_000)

  test.beforeEach(async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)
  })

  test('1-4: opening a workbook with a supported chart renders it with stable identity and typed source ranges', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    const bytes = await buildExcelChartFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-1.xlsx')

    // Typed state on the open response: chart type, series, source
    // ranges, dimensions (anchor), both locators.
    const chart = snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    expect(chart?.chartPath).toBe('xl/charts/chart1.xml')
    expect(chart?.drawingPath).toBe('xl/drawings/drawing1.xml')
    expect(chart?.drawingIndex).toBe(0)
    expect(chart?.anchorType).toBe('two-cell')
    expect(chart?.anchor.fromColumn).toBe(1)
    expect(chart?.anchor.toRow).toBe(18)
    expect(chart?.chart.chartTypes).toEqual(['barChart'])
    expect(chart?.chart.title).toBe('Sales & Cost')
    expect(chart?.chart.series).toHaveLength(2)
    expect(chart?.chart.series[0]?.name).toBe('Revenue')
    expect(chart?.chart.series[0]?.values).toEqual([10, 20, 30])
    expect(chart?.chart.series[0]?.valuesRef).toBe('Data!$B$2:$B$4')
    expect(chart?.chart.series[0]?.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(chart?.chart.series[1]?.valuesRef).toBe('Data!$C$2:$C$4')

    // The chart RENDERS as a floating SVG frame over the grid, keyed by
    // its canonical locator id — the identity that survives the whole
    // editor lifecycle.
    await page.waitForTimeout(600)
    const frames = await chartFrames(page)
    expect(frames).toHaveLength(1)
    expect(frames[0]?.key).toBe('file-chart:xl/drawings/drawing1.xml#0')
    expect(frames[0]?.title).toBe('Sales & Cost')
    expect(frames[0]?.width).toBeGreaterThan(100)
    expect(frames[0]?.height).toBeGreaterThan(100)
    // SVG bars rendered (two series × three categories).
    const bars = await page.locator('.chart-frame .chart-svg rect').count()
    expect(bars).toBeGreaterThanOrEqual(6)
    expect(pageErrors).toEqual([])
  })

  test('5-6: moving and resizing a chart journal final-state geometry through visualEdits', async ({
    page,
  }) => {
    const bytes = await buildExcelChartFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-2.xlsx')
    const key = snapshot.sheets[0]?.charts?.[0]
    expect(key).toBeDefined()

    // MOVE: drag right and down by ~120px.
    await dragChart(page, 'file-chart:xl/drawings/drawing1.xml#0', 120, 60)
    // RESIZE: pull the SE handle out by ~80px.
    await resizeChart(page, 'file-chart:xl/drawings/drawing1.xml#0', 'se', 80, 60)

    const saved = await saveWorkbookBytes(page)
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).not.toBeNull()
    // The from marker moved right/down from (col 1, row 2)…
    const fromCol = /<xdr:from><xdr:col>(\d+)<\/xdr:col>/.exec(drawing ?? '')
    const fromRow = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(drawing ?? '')
    expect(fromCol && Number(fromCol[1])).toBeGreaterThanOrEqual(2)
    expect(fromRow && Number(fromRow[1])).toBeGreaterThanOrEqual(4)
    // …and the to marker grew past the fixture's (col 8, row 18).
    const toCol = /<xdr:to><xdr:col>(\d+)<\/xdr:col>/.exec(drawing ?? '')
    const toRow = /<xdr:to>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(drawing ?? '')
    expect(toCol && Number(toCol[1])).toBeGreaterThanOrEqual(9)
    expect(toRow && Number(toRow[1])).toBeGreaterThanOrEqual(19)

    // Reopen: the persisted geometry is the file-native state.
    const reopened = await openFixture(page, saved, '/tmp/chart-e2e-2-re.xlsx')
    const moved = reopened.sheets[0]?.charts?.[0]
    expect(moved?.anchor.fromColumn).toBeGreaterThanOrEqual(2)
    expect(moved?.anchor.fromRow).toBeGreaterThanOrEqual(4)
    expect(moved?.anchor.toColumn).toBeGreaterThanOrEqual(9)
    expect(moved?.anchor.toRow).toBeGreaterThanOrEqual(19)
  })

  test('7: changing a supported chart property journals through chartEdits and round-trips', async ({
    page,
  }) => {
    const bytes = await buildExcelChartFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-3.xlsx')
    expect(snapshot.sheets[0]?.charts?.[0]?.chartPath).toBe('xl/charts/chart1.xml')

    // Select the chart → the frame becomes selected.
    const frame = page.locator(
      '.chart-frame[data-chart-key="file-chart:xl/drawings/drawing1.xml#0"]',
    )
    await frame.click()
    await page.waitForTimeout(400)
    expect(await frame.getAttribute('class')).toContain('chart-selected')

    // Chart Design pane: rename the title and convert column → line.
    const panel = page.locator('[data-testid="chart-panel"]')
    await expect(panel).toBeVisible()
    const titleInput = panel.locator('[data-testid="chart-panel-title"]')
    await titleInput.fill('Renamed Revenue')
    await titleInput.press('Enter')
    await page.waitForTimeout(300)
    await panel.locator('[data-testid="chart-panel-convert"]').selectOption('line')
    await page.waitForTimeout(300)
    // The on-screen chart previews the pending edit (overlay semantics).
    await expect(frame.locator('.chart-title')).toHaveText('Renamed Revenue')

    const saved = await saveWorkbookBytes(page)
    const chartXml = await readZipEntry(saved, 'xl/charts/chart1.xml')
    expect(chartXml).not.toBeNull()
    expect(chartXml).toContain('Renamed Revenue')
    expect(chartXml).toContain('<c:lineChart>')
    expect(chartXml).not.toContain('<c:barChart>')

    // Reopen: file-native state reflects the edit.
    const reopened = await openFixture(page, saved, '/tmp/chart-e2e-3-re.xlsx')
    const edited = reopened.sheets[0]?.charts?.[0]
    expect(edited?.chart.chartTypes).toEqual(['lineChart'])
    expect(edited?.chart.title).toBe('Renamed Revenue')
    expect(edited?.chart.series).toHaveLength(2)
  })

  test('8-10: creating a chart from selected data persists and reopens file-native', async ({
    page,
  }) => {
    // A data-only fixture: the cells A1:C4 exist but no chart anchors.
    const dataBytes = await buildExcelChartDataOnlyFixture()
    const snapshot = await openFixture(page, dataBytes, '/tmp/chart-e2e-4.xlsx')
    expect(snapshot.sheets[0]?.charts ?? []).toHaveLength(0)

    // Select the data range and insert a column chart through the panel.
    await selectRange(page, 'A1:C4')
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.waitForTimeout(300)
    await page.getByRole('button', { name: /^Chart/ }).click()
    const createPanel = page.locator('[data-testid="chart-panel-create"]')
    await expect(createPanel).toBeVisible()
    await expect(createPanel.locator('[data-testid="chart-panel-parsed"]')).toContainText(
      '2 series, 3 categories',
    )
    await createPanel.locator('[data-testid="chart-panel-type"]').selectOption('column')
    await createPanel.locator('[data-testid="chart-panel-insert"]').click()
    await page.waitForTimeout(600)

    // The session chart renders as a floating frame keyed by its session id
    // (desktop title rule: the series name for a single series, "Chart
    // Title" for multiple — visual-actions.ts parity).
    const frames = await chartFrames(page)
    expect(frames).toHaveLength(1)
    expect(frames[0]?.key).toContain('session-chart-')
    expect(frames[0]?.title).toBe('Chart Title')

    const saved = await saveWorkbookBytes(page)
    // The chart part exists with the typed series, and the drawing holds
    // a chart graphicFrame anchor.
    const entries = await listZipEntries(saved)
    const chartPart = entries.find((entry) => /^xl\/charts\/chart\d+\.xml$/.test(entry))
    expect(chartPart).toBeDefined()
    const chartXml = await readZipEntry(saved, chartPart ?? 'xl/charts/chart1.xml')
    expect(chartXml).toContain('<c:barChart>')
    expect(chartXml).toContain('Revenue')
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).toContain('<xdr:graphicFrame')
    expect(drawing).toContain('r:id=')
    const rels = await readZipEntry(saved, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).toContain('/chart')
    const contentTypes = await readZipEntry(saved, '[Content_Types].xml')
    expect(contentTypes).toContain(chartPart ?? '')

    // Reopen: the created chart is file-native with BOTH locators.
    const reopened = await openFixture(page, saved, '/tmp/chart-e2e-4-re.xlsx')
    const created = reopened.sheets[0]?.charts?.[0]
    expect(created).toBeDefined()
    expect(created?.chartPath).toBe(chartPart)
    expect(created?.chart.chartTypes).toEqual(['barChart'])
    expect(created?.chart.series[0]?.values).toEqual([10, 20, 30])
    expect(created?.chart.series[1]?.name).toBe('Cost')
    await page.waitForTimeout(600)
    expect(await chartFrames(page)).toHaveLength(1)
  })

  test('11-13: deleting a chart cascades its part while the unrelated chart survives', async ({
    page,
  }) => {
    const bytes = await buildExcelChartsFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-5.xlsx')
    expect(snapshot.sheets[0]?.charts ?? []).toHaveLength(2)

    // Delete the FIRST chart through its frame's delete button.
    const frame = page.locator(
      '.chart-frame[data-chart-key="file-chart:xl/drawings/drawing1.xml#0"]',
    )
    await frame.click()
    await page.waitForTimeout(300)
    await frame.locator('[data-testid="chart-delete"]').click()
    await page.waitForTimeout(600)
    expect(await chartFrames(page)).toHaveLength(1)

    const saved = await saveWorkbookBytes(page)
    // The deleted chart's part, rels, and content-type override are gone…
    const entries = await listZipEntries(saved)
    expect(entries).not.toContain('xl/charts/chart1.xml')
    expect(entries).not.toContain('xl/charts/_rels/chart1.xml.rels')
    const rels = await readZipEntry(saved, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).not.toContain('rIdChart1')
    const contentTypes = await readZipEntry(saved, '[Content_Types].xml')
    expect(contentTypes).not.toContain('chart1.xml')
    // …while the UNRELATED chart survives byte-identical.
    const surviving = await readZipEntryBytes(saved, 'xl/charts/chart2.xml')
    const original = await readZipEntryBytes(bytes, 'xl/charts/chart2.xml')
    expect(surviving?.equals(original ?? Buffer.alloc(0))).toBe(true)
    // Its relationship still resolves.
    expect(rels).toContain('rIdChart2')

    // Reopen: one file-native chart remains.
    const reopened = await openFixture(page, saved, '/tmp/chart-e2e-5-re.xlsx')
    expect(reopened.sheets[0]?.charts ?? []).toHaveLength(1)
    expect(reopened.sheets[0]?.charts?.[0]?.chartPath).toBe('xl/charts/chart2.xml')
  })

  test('14: a no-op save preserves the chart parts byte-for-byte', async ({ page }) => {
    const bytes = await buildExcelChartsFixture()
    await openFixture(page, bytes, '/tmp/chart-e2e-6.xlsx')

    // A cell-edit save with NO chart interaction (the chart families stay
    // empty — the journal holds nothing).
    await markDirtyWithCellEdit(page)
    const saved = await saveWorkbookBytes(page)
    for (const path of [
      'xl/charts/chart1.xml',
      'xl/charts/chart2.xml',
      'xl/drawings/drawing1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
    ]) {
      const before = await readZipEntryBytes(bytes, path)
      const after = await readZipEntryBytes(saved, path)
      expect(after).not.toBeNull()
      expect(before?.equals(after ?? Buffer.alloc(0))).toBe(true)
    }
  })

  test('15: unsupported chart structures fail closed — omitted, never relocated, bytes preserved', async ({
    page,
  }) => {
    const bytes = await buildExcel3DChartFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-7.xlsx')

    // Only the supported bar chart surfaces; the 3-D chart is omitted
    // (its anchor still counted — index 1 for the survivor).
    const charts = snapshot.sheets[0]?.charts ?? []
    expect(charts).toHaveLength(1)
    expect(charts[0]?.chartPath).toBe('xl/charts/chart1.xml')
    expect(charts[0]?.drawingIndex).toBe(1)
    await page.waitForTimeout(500)
    expect(await chartFrames(page)).toHaveLength(1)

    // A cell-edit save with NO chart interaction preserves the
    // unsupported chart's bytes exactly — nothing silently flattened
    // or relocated.
    await markDirtyWithCellEdit(page)
    const saved = await saveWorkbookBytes(page)
    const unsupportedBefore = await readZipEntryBytes(bytes, 'xl/charts/chart2.xml')
    const unsupportedAfter = await readZipEntryBytes(saved, 'xl/charts/chart2.xml')
    expect(unsupportedAfter?.equals(unsupportedBefore ?? Buffer.alloc(0))).toBe(true)
    const drawingBefore = await readZipEntryBytes(bytes, 'xl/drawings/drawing1.xml')
    const drawingAfter = await readZipEntryBytes(saved, 'xl/drawings/drawing1.xml')
    expect(drawingAfter?.equals(drawingBefore ?? Buffer.alloc(0))).toBe(true)
  })

  test('absolute-anchored charts are omitted, never relocated (fail closed)', async ({ page }) => {
    const bytes = await buildExcelAbsoluteChartFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-8.xlsx')

    const charts = snapshot.sheets[0]?.charts ?? []
    expect(charts).toHaveLength(1)
    expect(charts[0]?.chartPath).toBe('xl/charts/chart2.xml')
    expect(charts[0]?.drawingIndex).toBe(1)
    await page.waitForTimeout(500)
    expect(await chartFrames(page)).toHaveLength(1)

    // The absolute chart's drawing XML (including its fixed pos/ext)
    // survives a cell-edit save byte-for-byte.
    await markDirtyWithCellEdit(page)
    const saved = await saveWorkbookBytes(page)
    const drawingBefore = await readZipEntryBytes(bytes, 'xl/drawings/drawing1.xml')
    const drawingAfter = await readZipEntryBytes(saved, 'xl/drawings/drawing1.xml')
    expect(drawingAfter?.equals(drawingBefore ?? Buffer.alloc(0))).toBe(true)
    const absChartBefore = await readZipEntryBytes(bytes, 'xl/charts/chart1.xml')
    const absChartAfter = await readZipEntryBytes(saved, 'xl/charts/chart1.xml')
    expect(absChartAfter?.equals(absChartBefore ?? Buffer.alloc(0))).toBe(true)
  })

  test('a one-cell chart moves but refuses resize (fail closed)', async ({ page }) => {
    const bytes = await buildExcelOneCellChartFixture()
    const snapshot = await openFixture(page, bytes, '/tmp/chart-e2e-9.xlsx')
    const chart = snapshot.sheets[0]?.charts?.[0]
    expect(chart?.anchorType).toBe('one-cell')
    expect(chart?.widthPx).toBe(50)
    expect(chart?.heightPx).toBe(30)

    // Selecting shows NO resize handles (resize is not representable —
    // only the from marker would be rewritten).
    const frame = page.locator(
      '.chart-frame[data-chart-key="file-chart:xl/drawings/drawing1.xml#0"]',
    )
    await frame.click()
    await page.waitForTimeout(300)
    expect(await frame.locator('.chart-handle').count()).toBe(0)

    // Moving IS supported: the from marker follows.
    await dragChart(page, 'file-chart:xl/drawings/drawing1.xml#0', 100, 40)
    const saved = await saveWorkbookBytes(page)
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).toContain('<xdr:oneCellAnchor>')
    const fromCol = /<xdr:from><xdr:col>(\d+)<\/xdr:col>/.exec(drawing ?? '')
    expect(fromCol && Number(fromCol[1])).toBeGreaterThanOrEqual(2)
  })
})
