/**
 * REAL browser E2E — Home → Conditional Formatting (EXCEL-024).
 *
 * Proves the conditional-formatting persistence chain end-to-end through
 * the REAL HTTP boundary:
 *
 *   open → read <conditionalFormatting> (+ dxf resolution in styles.xml)
 *   → WorksheetState.cfRules (canonical CfWireRule wire shape)
 *   → browser installs the rules in the real Univer CF model
 *     (sheet.mutation.add-conditional-rule under journal suppression)
 *   → grid renders the rule effects through Univer's CF pipeline
 *   → user creates/edits/deletes rules through the REAL Univer facade
 *     (the same builder the CF panel drives) → CF mutations mark the
 *     sheet CF-dirty
 *   → save snapshots the LIVE CF model as canonical CfWireRules
 *   → savePlan.cfStates → /api/office/workbooks/save → routeOffice
 *     (strict wire validation) → xlsx-gateway applyCfRules → XLSX bytes
 *   → reopen → rule state reconstructed + reinstalled
 *
 * Also proves the fail-closed surfaces: a sheet whose CF cannot be
 * represented (x14 extension, time period) opens with NO rules installed
 * and REFUSES rule edits; a no-CF-dirty save preserves the CF XML
 * byte-for-byte; structural row/column operations shift the live rule
 * ranges in lockstep with the canonical replay.
 *
 * No browser-side OOXML. The browser only ever exchanges typed CfWireRule
 * snapshots taken from Univer's live model.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelCfFixture, buildExcelCfLockedFixture, readZipEntry } from './fixtures'

/** Plain-data view of the live CF model. */
interface LiveCfRule {
  cfId: string
  stopIfTrue: boolean
  ranges: Array<{
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
  }>
  rule: Record<string, unknown>
}

function readLiveCf(
  page: import('@playwright/test').Page,
  sheetName: string,
): Promise<LiveCfRule[]> {
  return page.evaluate((name: string) => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook: () => {
              getSheetByName: (sheet: string) => {
                getConditionalFormattingRules: () => Array<{
                  cfId: string
                  stopIfTrue?: boolean
                  ranges: Array<{
                    startRow: number
                    endRow: number
                    startColumn: number
                    endColumn: number
                  }>
                  rule: Record<string, unknown>
                }>
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.(name)
    return (ws?.getConditionalFormattingRules?.() ?? []).map((live) => ({
      cfId: live.cfId,
      stopIfTrue: live.stopIfTrue === true,
      ranges: live.ranges.map((range) => ({
        startRow: range.startRow,
        endRow: range.endRow,
        startColumn: range.startColumn,
        endColumn: range.endColumn,
      })),
      rule: live.rule as Record<string, unknown>,
    }))
  }, sheetName)
}

/** Create a number rule through the REAL Univer facade (the same builder
 *  the CF panel drives — the web shell has no CF UI of its own). */
async function createNumberRule(
  page: import('@playwright/test').Page,
  sheetName: string,
  anchor: string,
): Promise<void> {
  await page.evaluate(
    ({ name, anchorCell }) => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  newConditionalFormattingRule: () => {
                    whenNumberLessThan: (value: number) => {
                      setBackground: (color: string) => {
                        setRanges: (ranges: unknown[]) => { build: () => unknown }
                      }
                      setRanges: (ranges: unknown[]) => {
                        setBackground: (color: string) => {
                          setRanges: (ranges: unknown[]) => { build: () => unknown }
                        }
                        build: () => unknown
                      }
                    }
                  }
                  addConditionalFormattingRule: (rule: unknown) => unknown
                  getRange: (
                    row: number,
                    col: number,
                    rows: number,
                    cols: number,
                  ) => {
                    getRange: () => unknown
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const api = runtime?.univerAPI
      const ws = api?.getActiveWorkbook?.()?.getSheetByName?.(name)
      if (!ws || !api) return
      void anchorCell
      const built = ws
        .newConditionalFormattingRule()
        .whenNumberLessThan(50)
        .setBackground('rgb(255,199,206)')
        .setRanges([ws.getRange(1, 11, 5, 1).getRange()])
        .build()
      ws.addConditionalFormattingRule(built)
    },
    { name: sheetName, anchorCell: anchor },
  )
}

test.describe('Home tab — Conditional Formatting persists through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1: opening a workbook with existing rules installs and renders them', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-open.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-open.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-open.xlsx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // READ path: the snapshot carried the parsed rules (12 on Data, 1 on
    // Other; nothing locked).
    const sheets = (await openResponse.json()).snapshot.sheets as Array<{
      name: string
      cfRules?: Array<{ ranges: unknown[]; rule: Record<string, unknown> }>
      cfLocked?: boolean
    }>
    expect(sheets[0]!.cfRules).toHaveLength(12)
    expect(sheets[0]!.cfLocked).toBeUndefined()
    expect(sheets[1]!.cfRules).toHaveLength(1)

    // IMPORT path: the live Univer model carries every rule.
    await page.waitForTimeout(1500)
    const live = await readLiveCf(page, 'Data')
    expect(live).toHaveLength(12)
    const greater = live.find(
      (r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan',
    )
    expect(greater?.ranges).toEqual([{ startRow: 1, endRow: 5, startColumn: 1, endColumn: 1 }])
    expect(greater?.rule.value).toBe(50)
    // The dxf style arrived PRE-RESOLVED (bold + font color + fill).
    const style = greater?.rule.style as Record<string, unknown>
    expect(style.bl).toBe(1)
    expect((style.cl as { rgb: string }).rgb).toBe('#9C0006')
    expect((style.bg as { rgb: string }).rgb).toBe('#FFC7CE')
    // Visual families survived the wire.
    expect(live.some((r) => r.rule.type === 'colorScale')).toBe(true)
    expect(live.some((r) => r.rule.type === 'dataBar')).toBe(true)
    expect(live.some((r) => r.rule.type === 'iconSet')).toBe(true)
    // stopIfTrue survived.
    expect(live.some((r) => r.stopIfTrue)).toBe(true)

    // RENDER: the CF pipeline paints (the pink fill + red font of the
    // greaterThan rule produce more distinct canvas colors than a blank
    // grid strip).
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
    expect(painted, 'grid canvas must paint conditional formatting').toBeGreaterThan(5)

    expect(pageErrors).toEqual([])
  })

  test('2: create a rule through the real facade → save → XML → reopen', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-create.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-create.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-create.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Create L2:L6 "number lessThan 50" through the REAL facade.
    await createNumberRule(page, 'Data', 'L2')
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The new rule serialized with the interned dxf (pink fill).
    expect(sheet1).toContain('operator="lessThan"')
    expect(sheet1).toContain('<formula>50</formula>')
    expect(sheet1).toContain('sqref="L2:L6"')
    // The imported rules survived the rewrite (siblings kept).
    expect(sheet1).toContain('operator="greaterThan"')
    expect(sheet1).toContain('type="colorScale"')
    expect(sheet1).toContain('type="dataBar"')
    expect(sheet1).toContain('iconSet="3TrafficLights1"')
    expect(sheet1).toContain('stopIfTrue="1"')
    // The dxfs landed in styles.xml.
    const styles = await readZipEntry(saved, 'xl/styles.xml')
    expect(styles).toContain('<dxfs')
    expect(styles).toContain('FFFFC7CE')

    // REOPEN: the saved bytes restore the full rule set (12 imported + 1
    // created = 13). The lessThan-50 rule re-reads as a FORMULA rule —
    // the reader's blank-divergence compensation (Excel paints blanks for
    // lessThan, Univer's native condition skips them; ≤20k covered cells
    // swap to the equivalent formula — desktop parity).
    writeFileSync('/tmp/e2e-ribbon-cf-reopen.xlsx', saved)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-reopen.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-reopen.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)
    const live = await readLiveCf(page, 'Data')
    expect(live).toHaveLength(13)
    const created = live.filter((r) => r.rule.subType === 'formula' && r.rule.value === '=L2<50')
    expect(created).toHaveLength(1)
    expect(created[0]!.ranges).toEqual([{ startRow: 1, endRow: 5, startColumn: 11, endColumn: 11 }])

    expect(pageErrors).toEqual([])
  })

  test('3: edit one rule while preserving every sibling', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-edit.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-edit.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-edit.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Edit ONE rule through the REAL facade: change the greaterThan
    // operand from 50 to 70 (setConditionalFormattingRule — the panel's
    // own edit path).
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  getConditionalFormattingRules: () => Array<{
                    cfId: string
                    ranges: unknown[]
                    stopIfTrue?: boolean
                    rule: Record<string, unknown>
                  }>
                  setConditionalFormattingRule: (cfId: string, rule: unknown) => unknown
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')
      if (!ws) return
      const target = ws
        .getConditionalFormattingRules()
        .find((r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan')
      if (!target) return
      ws.setConditionalFormattingRule(target.cfId, {
        ...target,
        rule: { ...target.rule, value: 70 },
      })
    })
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<formula>70</formula>')
    expect(sheet1).not.toContain('<formula>50</formula>')
    // Every sibling family survived the single-rule edit.
    expect(sheet1).toContain('type="containsText"')
    expect(sheet1).toContain('type="duplicateValues"')
    expect(sheet1).toContain('type="top10"')
    expect(sheet1).toContain('type="aboveAverage"')
    expect(sheet1).toContain('type="expression"')
    expect(sheet1).toContain('type="containsBlanks"')
    // Priority ordering: the canonical writer assigns each rule a distinct
    // sequential priority (1..12 — a permutation with no gaps or repeats).
    const priorities = [...sheet1.matchAll(/\spriority="(\d+)"/g)]
      .map((m) => Number(m[1]))
      .sort((a, b) => a - b)
    expect(priorities).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])

    // Reopen and verify the edit landed.
    writeFileSync('/tmp/e2e-ribbon-cf-edit-reopen.xlsx', saved)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-edit-reopen.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-edit-reopen.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)
    const live = await readLiveCf(page, 'Data')
    expect(live).toHaveLength(12)
    const edited = live.find(
      (r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan',
    )
    expect(edited?.rule.value).toBe(70)

    expect(pageErrors).toEqual([])
  })

  test('4: delete one rule while preserving siblings', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-delete.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-delete.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-delete.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Delete the containsText rule through the REAL facade.
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  getConditionalFormattingRules: () => Array<{
                    cfId: string
                    rule: Record<string, unknown>
                  }>
                  deleteConditionalFormattingRule: (cfId: string) => unknown
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')
      if (!ws) return
      const target = ws
        .getConditionalFormattingRules()
        .find((r) => r.rule.subType === 'text' && r.rule.operator === 'containsText')
      if (target) ws.deleteConditionalFormattingRule(target.cfId)
    })
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).not.toContain('type="containsText"')
    // The other 11 rules survive.
    expect(sheet1).toContain('operator="greaterThan"')
    expect(sheet1).toContain('type="colorScale"')
    expect(sheet1).toContain('type="dataBar"')
    expect(sheet1).toContain('iconSet="3TrafficLights1"')
    expect(sheet1).toContain('type="duplicateValues"')
    expect(sheet1).toContain('type="top10"')
    expect(sheet1).toContain('type="aboveAverage"')
    expect(sheet1).toContain('type="expression"')
    expect(sheet1).toContain('type="containsBlanks"')
    expect(sheet1).toContain('stopIfTrue="1"')

    expect(pageErrors).toEqual([])
  })

  test('5: row insertion transforms live rule ranges and persists them', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-insert.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-insert.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-insert.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Insert 2 rows at row 2 through the REAL Univer facade (the same
    // method the context menu drives) → sheet.mutation.insert-row →
    // Univer's NATIVE CF transforms rewrite the live rule ranges (the
    // CF-UI plugin's ConditionalFormattingFormulaRefRangeController +
    // FormulaRefRangeService extend ranges whose top edge sits at the
    // insertion point) and split formula rules with re-anchored formulas.
    // Those native transform mutations journal the sheet CF-dirty, so the
    // save writes the LIVE snapshot — file and live model stay in
    // lockstep (EXCEL-024 §8 resolution; verified against the installed
    // package source and empirically in this test).
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  insertRows: (row: number, count?: number) => unknown
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')?.insertRows?.(1, 2)
    })
    await page.waitForTimeout(800)

    // The live greaterThan rule extended B2:B6 → B2:B8 (insert at the
    // range's top edge extends it, Excel's own boundary behavior).
    const live = await readLiveCf(page, 'Data')
    const greater = live.find(
      (r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan',
    )
    expect(greater?.ranges).toEqual([{ startRow: 1, endRow: 7, startColumn: 1, endColumn: 1 }])

    await page.keyboard.press('Escape')
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The live snapshot won: the extended range + the split formula rules
    // are exactly what the canonical writer serialized.
    expect(sheet1).toContain('sqref="B2:B8"')
    expect(sheet1).toContain('sqref="K2:K8 M2:M8"')
    expect(sheet1).toContain('type="expression"')
    expect(sheet1).toContain('type="colorScale"')
    expect(sheet1).toContain('stopIfTrue="1"')

    // Reopen: the transformed ranges survive (14 rules — the two formula
    // rules split into per-segment rules with re-anchored formulas).
    writeFileSync('/tmp/e2e-ribbon-cf-insert-reopen.xlsx', saved)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-insert-reopen.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-insert-reopen.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)
    const reopened = await readLiveCf(page, 'Data')
    expect(reopened).toHaveLength(14)
    const reopenedGreater = reopened.find(
      (r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan',
    )
    expect(reopenedGreater?.ranges).toEqual([
      { startRow: 1, endRow: 7, startColumn: 1, endColumn: 1 },
    ])

    expect(pageErrors).toEqual([])
  })

  test('6: column insertion shifts live rule ranges (column structural edit)', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-col.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-col.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-col.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Insert a column at C (index 2) → the colorScale range (C2:C6, first
    // column = the insertion edge) EXTENDS to C2:D6; the later columns
    // shift right (Univer's native transform — the same boundary rule
    // Excel applies).
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  insertColumns: (col: number, count?: number) => unknown
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')?.insertColumns?.(2, 1)
    })
    await page.waitForTimeout(800)
    const live = await readLiveCf(page, 'Data')
    const scale = live.find((r) => r.rule.type === 'colorScale')
    expect(scale?.ranges).toEqual([{ startRow: 1, endRow: 5, startColumn: 2, endColumn: 3 }])

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('sqref="C2:D6"')
    // The pre-insertion columns are untouched.
    expect(sheet1).toContain('sqref="B2:B6"')

    expect(pageErrors).toEqual([])
  })

  test('7: unsupported CF fails closed (x14 + timePeriod lock the sheet, edits refused)', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfLockedFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-locked.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-locked.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-locked.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // The workbook opens; the unrepresentable sheet carries NO rules plus
    // the fail-closed lock marker; the sibling sheet still parses.
    const sheets = (await openResponse.json()).snapshot.sheets as Array<{
      name: string
      cfRules?: unknown[]
      cfLocked?: boolean
    }>
    expect(sheets[0]!.cfRules).toBeUndefined()
    expect(sheets[0]!.cfLocked).toBe(true)
    expect(sheets[1]!.cfRules).toHaveLength(1)
    expect(sheets[1]!.cfLocked).toBeUndefined()

    await page.waitForTimeout(1500)
    // Live model: nothing installed on the locked sheet.
    const lockedLive = await readLiveCf(page, 'Data')
    expect(lockedLive).toHaveLength(0)
    const otherLive = await readLiveCf(page, 'Other')
    expect(otherLive).toHaveLength(1)

    // Attempting to create a rule on the locked sheet is REFUSED by the
    // BeforeCommandExecute gate (the facade routes through
    // sheet.command.add-conditional-rule, which the gate cancels).
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  newConditionalFormattingRule: () => {
                    whenNumberGreaterThan: (value: number) => {
                      setRanges: (ranges: unknown[]) => { build: () => unknown }
                    }
                  }
                  addConditionalFormattingRule: (rule: unknown) => unknown
                  getRange: (
                    row: number,
                    col: number,
                    rows: number,
                    cols: number,
                  ) => {
                    getRange: () => unknown
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const api = runtime?.univerAPI
      const ws = api?.getActiveWorkbook?.()?.getSheetByName?.('Data')
      if (!ws || !api) return
      const built = ws
        .newConditionalFormattingRule()
        .whenNumberGreaterThan(1)
        .setRanges([ws.getRange(0, 0, 3, 1).getRange()])
        .build()
      ws.addConditionalFormattingRule(built)
    })
    await page.waitForTimeout(600)
    const afterAttempt = await readLiveCf(page, 'Data')
    expect(afterAttempt).toHaveLength(0)
    await expect(
      page.getByText('This sheet has conditional formatting that cannot be edited yet'),
    ).toBeVisible()

    // A no-op save (nothing became CF-dirty) preserves the unrepresentable
    // CF XML byte-for-byte — the x14 twin and the timePeriod rule stay.
    const cellEdit = page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    row: number,
                    col: number,
                  ) => {
                    setValue: (value: unknown) => unknown
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
        ?.getRange?.(7, 0)
        ?.setValue?.(99)
    })
    await cellEdit
    await page.waitForTimeout(600)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    const originalSheet1 = await readZipEntry(fixture, 'xl/worksheets/sheet1.xml')
    // Everything from the CF sections onward is byte-identical (the cell
    // edit only rewrote the value + injected a dimension).
    const fromConditional = (xml: string): string =>
      xml.slice(xml.indexOf('<conditionalFormatting'))
    expect(fromConditional(sheet1)).toBe(fromConditional(originalSheet1))

    expect(pageErrors).toEqual([])
  })

  test('8: a no-CF-dirty save preserves the CF XML byte-for-byte', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-noop.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-noop.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-noop.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // A plain cell edit — no CF interaction, so no sheet is CF-dirty and
    // applyCfRules never runs.
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    row: number,
                    col: number,
                  ) => {
                    setValue: (value: unknown) => unknown
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
        ?.getRange?.(7, 0)
        ?.setValue?.(42)
    })
    await page.waitForTimeout(600)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')

    const savedSheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    const originalSheet1 = await readZipEntry(fixture, 'xl/worksheets/sheet1.xml')
    const fromConditional = (xml: string): string =>
      xml.slice(xml.indexOf('<conditionalFormatting'))
    expect(fromConditional(savedSheet1)).toBe(fromConditional(originalSheet1))

    // The unrelated sheet is untouched ENTIRELY (no cell edits there).
    const savedSheet2 = await readZipEntry(saved, 'xl/worksheets/sheet2.xml')
    const originalSheet2 = await readZipEntry(fixture, 'xl/worksheets/sheet2.xml')
    expect(savedSheet2).toBe(originalSheet2)

    // styles.xml keeps its dxfs untouched.
    const savedStyles = await readZipEntry(saved, 'xl/styles.xml')
    const originalStyles = await readZipEntry(fixture, 'xl/styles.xml')
    expect(savedStyles).toBe(originalStyles)

    expect(pageErrors).toEqual([])
  })

  test('9: the Home ribbon opens the real Univer conditional-formatting panel', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-panel.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-panel.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-panel.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // The Home tab carries the Conditional Formatting button; clicking it
    // opens Univer's own manage-rules sidebar — the panel's own UI text
    // ("Manage Entire Sheet Rules") plus the imported rule list appear.
    await expect(
      page.getByRole('button', { name: /^Conditional Formatting/ }).first(),
    ).toBeVisible()
    await page
      .getByRole('button', { name: /^Conditional Formatting/ })
      .first()
      .click()
    await expect(page.getByText('Manage Entire Sheet Rules')).toBeVisible({ timeout: 10_000 })
    // The imported rules are listed with their ranges.
    await expect(page.getByText('B2:B6').first()).toBeVisible()
    await expect(page.getByText('Color Scale').first()).toBeVisible()

    expect(pageErrors).toEqual([])
  })

  test('10: row deletion shrinks live rule ranges and persists them', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelCfFixture()
    writeFileSync('/tmp/e2e-ribbon-cf-delrow.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-delrow.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-delrow.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Delete rows 5-6 (0-based 4-5) through the REAL facade → Univer's
    // native transform shrinks the B2:B6 rule to B2:B4. (Rows 5-6 carry no
    // CF formula references — the gateway's frozen fail-closed rule aborts
    // deletions that would #REF! a rule formula, e.g. the containsText
    // rule's A2 anchor in rows 2-3.)
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  deleteRows: (row: number, count?: number) => unknown
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')?.deleteRows?.(4, 2)
    })
    await page.waitForTimeout(800)
    const live = await readLiveCf(page, 'Data')
    const greater = live.find(
      (r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan',
    )
    expect(greater?.ranges).toEqual([{ startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }])

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('sqref="B2:B4"')
    expect(sheet1).toContain('type="colorScale"')

    // Reopen: the shrunk range survives.
    writeFileSync('/tmp/e2e-ribbon-cf-delrow-reopen.xlsx', saved)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-cf-delrow-reopen.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-cf-delrow-reopen.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)
    const reopened = await readLiveCf(page, 'Data')
    const reopenedGreater = reopened.find(
      (r) => r.rule.subType === 'number' && r.rule.operator === 'greaterThan',
    )
    expect(reopenedGreater?.ranges).toEqual([
      { startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 },
    ])

    expect(pageErrors).toEqual([])
  })
})
