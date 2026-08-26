/**
 * REAL browser E2E — Data → Data Validation (Phase 4 Increment 5).
 *
 * Proves the data-validation persistence chain end-to-end through the REAL
 * HTTP boundary:
 *
 *   open → read <dataValidations> → WorksheetState.dvRules
 *   → browser installs the rules in the real Univer model
 *     (data-validation.mutation.addRule under journal suppression)
 *   → user creates/edits validation through the REAL Univer facade/panel
 *   → DV mutations mark the sheet DV-dirty
 *   → save snapshots the LIVE validation model as canonical DvWireRules
 *   → savePlan.dvStates → /api/office/workbooks/save → routeOffice
 *   → xlsx-gateway applyDvRules → XLSX bytes
 *   → reopen → validation state reconstructed + reinstalled
 *
 * No browser-side OOXML. The browser only ever exchanges typed DvWireRule
 * snapshots taken from Univer's live model.
 *
 * Fixture (buildExcelDvFixture / buildExcelDvExistingFixture): a survey
 * sheet with a header row, 5 mixed data rows (incl. one blank), currency
 * numfmt, a hyperlink; the "existing" variant carries whole-between /
 * list / custom validations.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import { buildExcelDvExistingFixture, buildExcelDvFixture, readZipEntry } from './fixtures'

/** Plain-data view of the live DV model (page.evaluate cannot serialize
 *  the runtime's functions — all facade access happens inside evaluate). */
interface LiveDvView {
  rules: Array<{
    type: string
    operator?: string
    formula1?: string
    formula2?: string
    allowBlank?: boolean
    showDropDown?: boolean
    showErrorMessage?: boolean
    errorTitle?: string
    error?: string
    ranges: Array<{
      startRow: number
      endRow: number
      startColumn: number
      endColumn: number
    }>
  }>
}

function readLiveDv(page: import('@playwright/test').Page): Promise<LiveDvView> {
  return page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook: () => {
              getActiveSheet: () => {
                getDataValidations: () => Array<{
                  rule: {
                    type: string
                    operator?: string
                    formula1?: string
                    formula2?: string
                    allowBlank?: boolean
                    showDropDown?: boolean
                    showErrorMessage?: boolean
                    errorTitle?: string
                    error?: string
                    ranges?: Array<{
                      startRow: number
                      endRow: number
                      startColumn: number
                      endColumn: number
                    }>
                  }
                }>
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
    const rules = (ws?.getDataValidations?.() ?? []).map(({ rule }) => {
      const { ranges, ...rest } = rule
      return {
        ...rest,
        ranges: (ranges ?? []).map((range) => ({
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
        })),
      }
    })
    return { rules }
  })
}

/** Create a whole-number between validation through the REAL facade. */
async function createWholeBetween(page: import('@playwright/test').Page): Promise<void> {
  await page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            newDataValidation: () => {
              requireNumberBetween: (
                start: number,
                end: number,
                isInteger?: boolean,
              ) => {
                setOptions: (options: Record<string, unknown>) => { build: () => unknown }
              }
            }
            getActiveWorkbook: () => {
              getActiveSheet: () => {
                getRange: (
                  row: number,
                  col: number,
                  nRows: number,
                  nCols: number,
                ) => {
                  setDataValidation: (rule: unknown) => unknown
                }
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const api = runtime?.univerAPI
    const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.()
    const rule = api
      ?.newDataValidation?.()
      ?.requireNumberBetween(1, 100, true)
      ?.setOptions({
        showErrorMessage: true,
        errorTitle: 'Bad count',
        error: 'Enter 1-100',
      })
      ?.build()
    ws?.getRange?.(1, 0, 5, 1)?.setDataValidation?.(rule)
  })
}

/** Create a list validation through the REAL facade. */
async function createList(page: import('@playwright/test').Page): Promise<void> {
  await page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            newDataValidation: () => {
              requireValueInList: (
                values: string[],
                multiple?: boolean,
                showDropdown?: boolean,
              ) => { build: () => unknown }
            }
            getActiveWorkbook: () => {
              getActiveSheet: () => {
                getRange: (
                  row: number,
                  col: number,
                  nRows: number,
                  nCols: number,
                ) => {
                  setDataValidation: (rule: unknown) => unknown
                }
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const api = runtime?.univerAPI
    const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.()
    const rule = api
      ?.newDataValidation?.()
      ?.requireValueInList?.(['Fruit', 'Vegetable', 'Grain'])
      ?.build()
    ws?.getRange?.(1, 1, 5, 1)?.setDataValidation?.(rule)
  })
}

/** Create a custom-formula validation through the REAL facade. */
async function createCustom(page: import('@playwright/test').Page): Promise<void> {
  await page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            newDataValidation: () => {
              requireFormulaSatisfied: (formula: string) => { build: () => unknown }
            }
            getActiveWorkbook: () => {
              getActiveSheet: () => {
                getRange: (
                  row: number,
                  col: number,
                  nRows: number,
                  nCols: number,
                ) => {
                  setDataValidation: (rule: unknown) => unknown
                }
              }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const api = runtime?.univerAPI
    const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.()
    // The builder keeps the formula VERBATIM (its docs show '=A1>2'), so the
    // canonical '='-prefixed form must be passed explicitly — matching the
    // file-import transform in toUniverDvRule.
    const rule = api?.newDataValidation?.()?.requireFormulaSatisfied?.('=ISNUMBER(C2)')?.build()
    ws?.getRange?.(1, 2, 5, 1)?.setDataValidation?.(rule)
  })
}

test.describe('Data tab — Data Validation persists through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1: opening a workbook with existing validations installs them in the real Univer model', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDvExistingFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-open.xlsx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-open.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-open.xlsx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)

    // READ path: the snapshot carried the parsed rules.
    const snapshot = (await openResponse.json()).snapshot.sheets as Array<{
      name: string
      dvRules?: Array<{ ranges: unknown[]; rule: Record<string, unknown> }>
    }>
    expect(snapshot[0].dvRules).toHaveLength(3)

    // IMPORT path: the live Univer model carries all three rules with the
    // desktop transforms applied (list literal unquoted; custom '='-prefixed).
    await page.waitForTimeout(1500)
    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(3)
    const whole = live.rules.find((r) => r.type === 'whole')
    expect(whole?.operator).toBe('between')
    expect(whole?.formula1).toBe('1')
    expect(whole?.formula2).toBe('100')
    expect(whole?.allowBlank).toBe(true)
    expect(whole?.showErrorMessage).toBe(true)
    expect(whole?.errorTitle).toBe('Bad count')
    expect(whole?.error).toBe('Enter 1-100')
    expect(whole?.ranges).toEqual([{ startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }])
    const list = live.rules.find((r) => r.type === 'list')
    expect(list?.formula1).toBe('Fruit,Vegetable,Grain')
    expect(list?.showDropDown).toBe(true)
    const custom = live.rules.find((r) => r.type === 'custom')
    expect(custom?.formula1).toBe('=ISNUMBER(C2)')

    expect(pageErrors).toEqual([])
  })

  test('2: a whole-number validation rejects invalid input in-session', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDvFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-whole.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-whole.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-whole.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Create A2:A6 whole between 1..100 through the REAL facade.
    await createWholeBetween(page)
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // The live model carries the rule.
    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(1)
    expect(live.rules[0]!.type).toBe('whole')
    expect(live.rules[0]!.operator).toBe('between')
    expect(live.rules[0]!.ranges).toEqual([
      { startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 },
    ])

    // The real validator flags the out-of-range A3 (120) — the validation
    // status cache stores DataValidationStatus (string enum: "valid" /
    // "invalid" / "validating"); VALID cells are DELETED from the cache
    // (getValue → null/undefined), only violations persist.
    const status = await page.evaluate(async () => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getValidatorStatusAsync: () => Promise<{
                    getValue: (row: number, col: number) => string | null | undefined
                  }>
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      const matrix = await ws?.getValidatorStatusAsync?.()
      return {
        a2: matrix?.getValue?.(1, 0),
        a3: matrix?.getValue?.(2, 0),
        a4: matrix?.getValue?.(3, 0),
      }
    })
    // In-range cells (A2=5, A4=50) are valid → absent from the cache.
    expect(status.a2 ?? null).toBeNull()
    expect(status.a4 ?? null).toBeNull()
    // A3=120 violates whole-between 1..100 → INVALID.
    expect(status.a3).toBe('invalid')

    expect(pageErrors).toEqual([])
  })

  test('3: a list validation provides the dropdown values in-session', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDvFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-list.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-list.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-list.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Create B2:B6 list Fruit/Vegetable/Grain through the REAL facade.
    await createList(page)
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(1)
    expect(live.rules[0]!.type).toBe('list')
    // The panel builder serializes values as a JSON array (Univer's own
    // list model form); the file import path carries the CSV literal. The
    // dropdown parses both.
    expect(live.rules[0]!.formula1).toBe('["Fruit","Vegetable","Grain"]')
    // The dropdown is ON (the builder defaults showDropdown=true).
    expect(live.rules[0]!.showDropDown).toBe(true)
    expect(live.rules[0]!.ranges).toEqual([
      { startRow: 1, endRow: 5, startColumn: 1, endColumn: 1 },
    ])

    expect(pageErrors).toEqual([])
  })

  test('4: save/reopen — typed dvStates on the wire, rules + criteria in the XLSX', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDvFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-save.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-save.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-save.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Create all three validations (whole/list/custom) on different ranges.
    await createWholeBetween(page)
    await page.waitForTimeout(300)
    await createList(page)
    await page.waitForTimeout(300)
    await createCustom(page)
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the plan must carry the typed canonical rules.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        dvStates?: Array<{
          sheetName: string
          rules: Array<{
            ranges: Array<{
              startRow: number
              endRow: number
              startColumn: number
              endColumn: number
            }>
            rule: { type: string; operator?: string; formula1?: string; formula2?: string }
          }>
        }>
      }
    }
    expect(saveBody.savePlan.dvStates, 'typed dvStates on the wire').toHaveLength(1)
    const state = saveBody.savePlan.dvStates![0]!
    expect(state.sheetName).toBe('Survey')
    expect(state.rules).toHaveLength(3)
    const whole = state.rules.find((r) => r.rule.type === 'whole')!
    expect(whole.rule.operator).toBe('between')
    expect(whole.rule.formula1).toBe('1')
    expect(whole.rule.formula2).toBe('100')
    expect(whole.ranges).toEqual([{ startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }])
    const list = state.rules.find((r) => r.rule.type === 'list')!
    // The builder's JSON form reaches the wire verbatim (the gateway
    // normalizes it to the quoted CSV literal at serialization).
    expect(list.rule.formula1).toBe('["Fruit","Vegetable","Grain"]')
    const custom = state.rules.find((r) => r.rule.type === 'custom')!
    expect(custom.rule.formula1).toBe('=ISNUMBER(C2)')

    // Saved XML: all three validations with their formulas + attributes.
    // (The builder's whole rule carries only the error options — allowBlank
    // stays unset, which serializes as the OOXML default allowBlank="0".)
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<dataValidations count="3">')
    expect(sheet1).toContain('type="whole" showErrorMessage="1"')
    expect(sheet1).toContain('errorTitle="Bad count" error="Enter 1-100" sqref="A2:A6"')
    expect(sheet1).toContain('<formula1>1</formula1><formula2>100</formula2>')
    expect(sheet1).toContain('type="list" sqref="B2:B6"')
    expect(sheet1).toContain('<formula1>"Fruit,Vegetable,Grain"</formula1>')
    expect(sheet1).toContain('type="custom" sqref="C2:C6"')
    expect(sheet1).toContain('<formula1>ISNUMBER(C2)</formula1>')
    // Unrelated content survives.
    expect(sheet1).toMatch(/<c r="A3"[^>]*><v>120<\/v>/)
    expect(sheet1).toContain('<hyperlink ref="B2" r:id="rId1"/>')

    // Reopen: the snapshot carries the rules and the live model reinstalls them.
    writeFileSync('/tmp/e2e-ribbon-dv-save-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-save-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-save-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      name: string
      cells: Record<string, { value: unknown }>
      styles?: Record<string, { bold?: boolean; numberFormat?: string }>
      dvRules?: Array<{ ranges: unknown[]; rule: Record<string, unknown> }>
    }>
    expect(reopened[0].dvRules).toHaveLength(3)
    expect(reopened[0].cells.A2?.value).toBe(5)
    expect(reopened[0].styles?.A1?.bold).toBe(true)
    expect(reopened[0].styles?.C2?.numberFormat).toMatch(/\$/)

    await page.waitForTimeout(1500)
    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(3)
    expect(live.rules.find((r) => r.type === 'whole')?.formula1).toBe('1')
    // The file's quoted CSV literal unquotes on install (desktop parity).
    expect(live.rules.find((r) => r.type === 'list')?.formula1).toBe('Fruit,Vegetable,Grain')
    expect(live.rules.find((r) => r.type === 'custom')?.formula1).toBe('=ISNUMBER(C2)')

    expect(pageErrors).toEqual([])
  })

  test('5: multiple-rule isolation — editing one rule keeps the others intact', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Start from the fixture with all three rules installed.
    const fixture = await buildExcelDvExistingFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-isolate.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-isolate.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-isolate.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    // Edit ONLY the whole-number rule's bounds (1..100 → 1..50) through the
    // REAL facade: setDataValidation on the same range replaces its rule.
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              newDataValidation: () => {
                requireNumberBetween: (
                  start: number,
                  end: number,
                  isInteger?: boolean,
                ) => { build: () => unknown }
              }
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (
                    row: number,
                    col: number,
                    nRows: number,
                    nCols: number,
                  ) => {
                    setDataValidation: (rule: unknown) => unknown
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const api = runtime?.univerAPI
      const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.()
      const rule = api?.newDataValidation?.()?.requireNumberBetween?.(1, 50, true)?.build()
      ws?.getRange?.(1, 0, 5, 1)?.setDataValidation?.(rule)
    })
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // Save: the snapshot carries THREE rules — the edited whole rule (new
    // bounds) plus the untouched list and custom rules.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: {
        dvStates?: Array<{
          rules: Array<{ rule: { type: string; formula1?: string; formula2?: string } }>
        }>
      }
    }
    const rules = saveBody.savePlan.dvStates![0]!.rules
    expect(rules).toHaveLength(3)
    const whole = rules.find((r) => r.rule.type === 'whole')!
    expect(whole.rule.formula2).toBe('50')
    // The untouched rules ride along verbatim.
    expect(rules.find((r) => r.rule.type === 'list')?.rule.formula1).toBe('Fruit,Vegetable,Grain')
    expect(rules.find((r) => r.rule.type === 'custom')?.rule.formula1).toBe('=ISNUMBER(C2)')

    // Saved XML proves the isolation.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('<formula2>50</formula2>')
    expect(sheet1).toContain('<formula1>"Fruit,Vegetable,Grain"</formula1>')
    expect(sheet1).toContain('<formula1>ISNUMBER(C2)</formula1>')

    // Reopen: all three survive.
    writeFileSync('/tmp/e2e-ribbon-dv-isolate-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-isolate-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-isolate-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await reopenResponsePromise
    await page.waitForTimeout(1500)
    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(3)
    expect(live.rules.find((r) => r.type === 'whole')?.formula2).toBe('50')

    expect(pageErrors).toEqual([])
  })

  test('6: clearing validation removes <dataValidations> and stops rejecting input', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDvExistingFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-clear.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-clear.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-clear.xlsx')).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)

    // Clear ALL validation: setDataValidation(null) on the whole span.
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
                    nRows: number,
                    nCols: number,
                  ) => {
                    setDataValidation: (rule: unknown) => unknown
                  }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      const ws = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
      ws?.getRange?.(1, 0, 5, 3)?.setDataValidation?.(null)
    })
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    // The live model has no rules left.
    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(0)

    // Save: the plan carries the explicit cleared state (empty rules).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { dvStates?: Array<{ rules: unknown[] }> }
    }
    expect(saveBody.savePlan.dvStates).toHaveLength(1)
    expect(saveBody.savePlan.dvStates![0]!.rules).toEqual([])

    // Saved XML: <dataValidations> removed.
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).not.toContain('dataValidations')
    // Unrelated content untouched.
    expect(sheet1).toMatch(/<c r="A3"[^>]*><v>120<\/v>/)

    // Reopen: no validation state; invalid input is no longer flagged.
    writeFileSync('/tmp/e2e-ribbon-dv-clear-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-clear-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-clear-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).snapshot.sheets as Array<{
      dvRules?: unknown[]
      cells: Record<string, { value: unknown }>
    }>
    expect(reopened[0].dvRules).toBeUndefined()
    expect(reopened[0].cells.A2?.value).toBe(5)

    await page.waitForTimeout(1500)
    const liveAfter = await readLiveDv(page)
    expect(liveAfter.rules).toHaveLength(0)

    expect(pageErrors).toEqual([])
  })

  test('7: a custom formula validation survives save/reopen', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelDvFixture()
    writeFileSync('/tmp/e2e-ribbon-dv-custom.xlsx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-custom.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-custom.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)

    await createCustom(page)
    await page.waitForTimeout(600)
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    await saveReq

    // The XML carries the custom formula (the '=' the facade added is
    // stripped on write — OOXML custom formulas carry the bare body).
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(sheet1).toContain('type="custom" sqref="C2:C6"')
    expect(sheet1).toContain('<formula1>ISNUMBER(C2)</formula1>')

    // Reopen: the custom rule is back in the live model with the '=' prefix.
    writeFileSync('/tmp/e2e-ribbon-dv-custom-reopened.xlsx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-ribbon-dv-custom-reopened.xlsx')
    await expect(page.getByText('Opened e2e-ribbon-dv-custom-reopened.xlsx')).toBeVisible({
      timeout: 30_000,
    })
    await reopenResponsePromise
    await page.waitForTimeout(1500)
    const live = await readLiveDv(page)
    expect(live.rules).toHaveLength(1)
    expect(live.rules[0]!.type).toBe('custom')
    expect(live.rules[0]!.formula1).toBe('=ISNUMBER(C2)')

    expect(pageErrors).toEqual([])
  })
})
