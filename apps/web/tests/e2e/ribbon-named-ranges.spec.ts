/**
 * REAL browser E2E — Formulas → Name Manager / Name Box (EXCEL-025).
 *
 * Proves the defined-names persistence chain end-to-end through the REAL
 * HTTP boundary:
 *
 *   open → read <definedNames> (model/preserve split per (name, scope),
 *   _xlnm + hidden skipped) → WorkbookSnapshot.definedNames → browser
 *   installs the names in the real Univer defined-name model (public
 *   builder facade + the same-name sibling path, under journal
 *   suppression) → Name Manager lists them; Name Box resolves them with
 *   Excel scope precedence; formulas consume them
 *   → user creates/edits/deletes through the REAL facade → the engine's
 *   set/remove-defined-name mutations mark the workbook names-dirty
 *   → save snapshots the LIVE model as the canonical DefinedNamesState
 *   (names + preserveNames) → savePlan.definedNamesState →
 *   /api/office/workbooks/save → routeOffice (strict wire validation) →
 *   xlsx-gateway applyDefinedNamesState → XLSX bytes → reopen
 *
 * Also proves the preservation invariant (editing one name never drops
 * its siblings, the print titles, the hidden names, or the
 * reader-preserved unmodelable names), the SAME-NAME CROSS-SCOPE pair
 * (workbook + worksheet definitions of one name coexist, resolve, and
 * edit independently — the architect's blocker correction), the
 * fail-closed GENUINE same-scope duplicate, the split-save (names never
 * ride the same save as structural ops), the no-op byte preservation,
 * the structural-op reference shift, and the namesLocked fail-closed
 * surface.
 *
 * No browser-side OOXML. The browser only ever exchanges typed
 * DefinedNamesState snapshots taken from Univer's live model.
 */
import { test, expect } from '@playwright/test'
import { writeFileSync } from 'node:fs'
import {
  loginAsDemoOwner,
  gotoHashRoute,
  waitForGridCanvas,
  clickSaveAndCaptureDownload,
} from './helpers'
import {
  buildExcelNamesFixture,
  buildExcelNamesLockedFixture,
  buildExcelNamesCollisionFixture,
  buildExcelNamesSameNameFixture,
  readZipEntry,
} from './fixtures'

/** Plain-data view of the live defined-name model. */
interface LiveName {
  name: string
  ref: string
  localSheetId: string | undefined
}

function readLiveNames(page: import('@playwright/test').Page): Promise<LiveName[]> {
  return page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook: () => {
              getDefinedNames: () => Array<{
                getName: () => string
                getFormulaOrRefString: () => string
                getLocalSheetId: () => string | undefined
              }>
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const wb = runtime?.univerAPI?.getActiveWorkbook?.()
    return (wb?.getDefinedNames?.() ?? []).map((defined) => ({
      name: defined.getName(),
      ref: defined.getFormulaOrRefString(),
      localSheetId: defined.getLocalSheetId(),
    }))
  })
}

/** The live computed value of one cell (formula results included). */
function readCellValue(
  page: import('@playwright/test').Page,
  sheetName: string,
  row: number,
  col: number,
): Promise<unknown> {
  return page.evaluate(
    ({ name, r, c }) => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getSheetByName: (sheet: string) => {
                  getRange: (row: number, col: number) => { getValue: () => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      return (
        runtime?.univerAPI
          ?.getActiveWorkbook?.()
          ?.getSheetByName?.(name)
          ?.getRange?.(r, c)
          ?.getValue?.() ?? null
      )
    },
    { name: sheetName, r: row, c: col },
  )
}

/** The active sheet name (scope resolution assertions). */
function readActiveSheetName(page: import('@playwright/test').Page): Promise<string> {
  return page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook: () => { getActiveSheet: () => { getSheetName: () => string } }
          }
        }
      }
    ).__genofficeExcelRuntime
    return runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getSheetName?.() ?? ''
  })
}
/** The active cell A1 label (the Name Box echo source). */
function readActiveCell(page: import('@playwright/test').Page): Promise<string> {
  return page.evaluate(() => {
    const runtime = (
      window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook: () => {
              getActiveRange: () => { getRow: () => number; getColumn: () => number }
            }
          }
        }
      }
    ).__genofficeExcelRuntime
    const range = runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveRange?.()
    if (!range) return ''
    const col = range.getColumn()
    const letters = String.fromCharCode(65 + col)
    return `${letters}${range.getRow() + 1}`
  })
}

async function openFixture(
  page: import('@playwright/test').Page,
  fixture: Buffer,
  fileName: string,
): Promise<void> {
  writeFileSync(`/tmp/${fileName}`, fixture)
  await page.setInputFiles('input[type="file"]', `/tmp/${fileName}`)
  await expect(page.getByText(`Opened ${fileName}`)).toBeVisible({ timeout: 30_000 })
  await page.waitForTimeout(1500)
}

/** Open the Name Manager through the real ribbon (Formulas tab). */
async function openNameManager(page: import('@playwright/test').Page): Promise<void> {
  await page
    .locator('[data-testid="excel-ribbon"] .excel-ribbon-tab', { hasText: 'Formulas' })
    .click()
  await page.waitForTimeout(200)
  await expect(page.getByRole('button', { name: /^Name Manager/ }).first()).toBeVisible()
  await page
    .getByRole('button', { name: /^Name Manager/ })
    .first()
    .click()
  await expect(page.getByTestId('name-manager-dialog')).toBeVisible({ timeout: 10_000 })
}

test.describe('Formulas tab — Named Ranges persist through the canonical pipeline', () => {
  test('1: opening a workbook with names installs them; Name Box resolves the workbook-scoped name', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-open.xlsx')
    expect((await openResponsePromise).status()).toBe(200)

    // READ path: the snapshot carries the model/preserve split. Print_Titles
    // (built-in) and SecretRate (hidden) never appear — the writer keeps
    // them by rule.
    const snapshot = (await (await openResponsePromise).json()).snapshot as {
      definedNames?: {
        names: Array<{ name: string; formula: string; sheetIndex?: number }>
        preserveNames: string[]
      }
      namesLocked?: boolean
    }
    expect(snapshot.namesLocked).toBeUndefined()
    const names = snapshot.definedNames?.names ?? []
    const byName = new Map(names.map((n) => [n.name, n]))
    expect(byName.get('GlobalTotal')).toEqual({ name: 'GlobalTotal', formula: 'Data!$A$1:$A$5' })
    expect(byName.get('LocalTotal')).toEqual({
      name: 'LocalTotal',
      formula: 'Other!$B$2:$B$4',
      sheetIndex: 1,
    })
    expect(snapshot.definedNames?.preserveNames).toEqual(['A1'])

    // IMPORT path: the live model carries the modeled names with the
    // workbook-scope sentinel / sheet ids.
    const live = await readLiveNames(page)
    const liveByName = new Map(live.map((n) => [n.name, n]))
    expect(liveByName.get('GlobalTotal')?.ref).toBe('Data!$A$1:$A$5')
    expect(liveByName.get('GlobalTotal')?.localSheetId).toBe('AllDefaultWorkbook')
    expect(liveByName.get('LocalTotal')?.ref).toBe('Other!$B$2:$B$4')
    expect(liveByName.get('LocalTotal')?.localSheetId).not.toBe('AllDefaultWorkbook')
    // Exactly the two modeled names installed.
    expect(live).toHaveLength(2)

    // NAME BOX: type the workbook-scoped name + Enter → jump to Data!A1
    // (the selection moves and the echo follows).
    const nameBox = page.getByTestId('excel-name-box')
    await nameBox.click()
    await nameBox.fill('GlobalTotal')
    await nameBox.press('Enter')
    await page.waitForTimeout(800)
    expect(await readActiveCell(page)).toBe('A1')
    // Case-insensitive resolution.
    await nameBox.click()
    await nameBox.fill('globaltotal')
    await nameBox.press('Enter')
    await page.waitForTimeout(800)
    expect(await readActiveCell(page)).toBe('A1')

    expect(pageErrors).toEqual([])
  })

  test('2: Name Box resolves the sheet-scoped name and switches sheets', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-scope.xlsx')

    const nameBox = page.getByTestId('excel-name-box')
    await nameBox.click()
    await nameBox.fill('LocalTotal')
    await nameBox.press('Enter')
    await page.waitForTimeout(1000)
    // The jump switched to the Other sheet and selected B2.
    expect(await readActiveCell(page)).toBe('B2')
    const activeSheet = await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => { getActiveSheet: () => { getSheetName: () => string } }
            }
          }
        }
      ).__genofficeExcelRuntime
      return runtime?.univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()?.getSheetName?.() ?? ''
    })
    expect(activeSheet).toBe('Other')

    // A name whose ref is a FORMULA is not a jump target — the input stays
    // an error and the selection does not move. (None of the fixture names
    // are formulas, so use a garbage input to prove the refusal path.)
    const before = await readActiveCell(page)
    await nameBox.click()
    await nameBox.fill('not-a-name-or-address!')
    await nameBox.press('Enter')
    await page.waitForTimeout(400)
    expect(await readActiveCell(page)).toBe(before)

    expect(pageErrors).toEqual([])
  })

  test('3: Name Manager opens and lists the file names with their scopes', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-list.xlsx')

    await openNameManager(page)
    const rows = page.getByTestId('name-manager-row')
    await expect(rows).toHaveCount(2)
    // GlobalTotal — workbook scope
    await expect(rows.filter({ hasText: 'GlobalTotal' })).toHaveText(/Workbook/)
    await expect(rows.filter({ hasText: 'GlobalTotal' })).toHaveText(/Data!\$A\$1:\$A\$5/)
    // LocalTotal — scoped to the Other sheet
    await expect(rows.filter({ hasText: 'LocalTotal' })).toHaveText(/Other/)
    // The #REF! duplicate loser and the unsaveable A1 never list.
    await expect(rows.filter({ hasText: 'A1' })).toHaveCount(0)

    expect(pageErrors).toEqual([])
  })

  test('4: create a name through the Name Manager → save → XML → reopen', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-create.xlsx')

    await openNameManager(page)
    await page.getByTestId('name-manager-name').fill('NewRate')
    await page.getByTestId('name-manager-ref').fill('=Data!$C$1:$C$5')
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const workbookXml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbookXml).toContain('<definedName name="NewRate">Data!$C$1:$C$5</definedName>')
    // Every modeled sibling survived the rewrite.
    expect(workbookXml).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>')
    expect(workbookXml).toContain(
      '<definedName name="LocalTotal" localSheetId="1">Other!$B$2:$B$4</definedName>',
    )
    // Built-in + hidden + preserved names stay byte-verbatim.
    expect(workbookXml).toContain(
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>',
    )
    expect(workbookXml).toContain(
      '<definedName name="SecretRate" hidden="1">Data!$C$1</definedName>',
    )
    expect(workbookXml).toContain('<definedName name="A1">Data!$A$1</definedName>')

    // REOPEN: the saved bytes restore the full model (2 + NewRate).
    await openFixture(page, saved, 'e2e-ribbon-names-reopen-create.xlsx')
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'NewRate')).toHaveLength(1)
    expect(live.filter((n) => n.name === 'NewRate')[0]?.ref).toBe('Data!$C$1:$C$5')
    expect(live).toHaveLength(3)

    expect(pageErrors).toEqual([])
  })

  test('5: edit one name while preserving every sibling and built-in', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-edit.xlsx')

    await openNameManager(page)
    await page.getByTestId('name-manager-row').filter({ hasText: 'GlobalTotal' }).click()
    await page.getByTestId('name-manager-ref').fill('=Data!$A$1:$A$3')
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const workbookXml = await readZipEntry(saved, 'xl/workbook.xml')
    // The edit landed.
    expect(workbookXml).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$3</definedName>')
    // THE CRITICAL INVARIANT: every other name survived byte-verbatim —
    // the untouched sibling, the built-in print titles, the hidden name,
    // and the preserved unmodelable.
    expect(workbookXml).toContain(
      '<definedName name="LocalTotal" localSheetId="1">Other!$B$2:$B$4</definedName>',
    )
    expect(workbookXml).toContain(
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>',
    )
    expect(workbookXml).toContain(
      '<definedName name="SecretRate" hidden="1">Data!$C$1</definedName>',
    )
    expect(workbookXml).toContain('<definedName name="A1">Data!$A$1</definedName>')

    // REOPEN: the edit persisted; the model is complete.
    await openFixture(page, saved, 'e2e-ribbon-names-reopen-edit.xlsx')
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'GlobalTotal')[0]?.ref).toBe('Data!$A$1:$A$3')
    expect(live).toHaveLength(2)

    expect(pageErrors).toEqual([])
  })

  test('6: delete one name while preserving every sibling and built-in', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-delete.xlsx')

    await openNameManager(page)
    await page
      .getByTestId('name-manager-row')
      .filter({ hasText: 'LocalTotal' })
      .getByRole('button', { name: /Delete LocalTotal/ })
      .click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const workbookXml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbookXml).not.toContain('name="LocalTotal"')
    // Siblings + built-ins + preserved names survived the deletion.
    expect(workbookXml).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>')
    expect(workbookXml).toContain('name="_xlnm.Print_Titles"')
    expect(workbookXml).toContain('name="SecretRate"')
    expect(workbookXml).toContain('name="A1"')

    // REOPEN: the model lost exactly the deleted name.
    await openFixture(page, saved, 'e2e-ribbon-names-reopen-delete.xlsx')
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'LocalTotal')).toHaveLength(0)
    expect(live).toHaveLength(1)

    expect(pageErrors).toEqual([])
  })

  test('7: a formula consuming a name resolves live and survives save/reopen', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-formula.xlsx')
    await page.waitForTimeout(2500)

    // B7 = SUM(GlobalTotal) = 10+20+30+40+50 = 150 — the formula engine
    // resolved the name from the installed model.
    expect(await readCellValue(page, 'Data', 6, 1)).toBe(150)

    // A plain cell edit (the save button requires unsaved changes) — far
    // from the name or the formula, so neither is touched.
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (row: number, col: number) => { setValue: (v: unknown) => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI
        ?.getActiveWorkbook?.()
        ?.getActiveSheet?.()
        ?.getRange?.(8, 3)
        ?.setValue?.(42)
    })
    await page.waitForTimeout(600)

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const sheet1 = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    // The formula cell survived as a FORMULA (never flattened to a value).
    expect(sheet1).toContain('<f>SUM(GlobalTotal)</f>')
    // The name it consumes survived too.
    const workbookXml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbookXml).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>')

    // REOPEN: the formula still resolves.
    await openFixture(page, saved, 'e2e-ribbon-names-reopen-formula.xlsx')
    await page.waitForTimeout(2500)
    expect(await readCellValue(page, 'Data', 6, 1)).toBe(150)

    expect(pageErrors).toEqual([])
  })

  test('8: a no-name-dirty save preserves the definedNames XML byte-for-byte', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-noop.xlsx')

    // A plain cell edit — no name interaction, so the workbook is NOT
    // names-dirty and applyDefinedNamesState never runs.
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (row: number, col: number) => { setValue: (v: unknown) => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI
        ?.getActiveWorkbook?.()
        ?.getActiveSheet?.()
        ?.getRange?.(8, 0)
        ?.setValue?.(42)
    })
    await page.waitForTimeout(600)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')

    const savedWorkbook = await readZipEntry(saved, 'xl/workbook.xml')
    const originalWorkbook = await readZipEntry(fixture, 'xl/workbook.xml')
    // The whole <definedNames> section survived byte-for-byte (the only
    // workbook.xml delta is the documented fullCalcOnLoad marker).
    const section = (xml: string): string =>
      xml.slice(xml.indexOf('<definedNames'), xml.indexOf('</definedNames>'))
    expect(section(savedWorkbook)).toBe(section(originalWorkbook))

    expect(pageErrors).toEqual([])
  })

  test('9: a structural edit shifts the name reference and the split-save persists it', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-structural.xlsx')

    // Insert 2 rows at the top of Data through the REAL Univer facade —
    // the same API the context menu drives. Univer's
    // UpdateDefinedNameController rewrites the live name refs and fires
    // set-defined-name mutations (the journal catches them → names dirty).
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
      runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')?.insertRows?.(0, 2)
    })
    await page.waitForTimeout(1000)
    // The live model shifted.
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'GlobalTotal')[0]?.ref).toBe('Data!$A$3:$A$7')

    // Save — the browser must split (names + structural never ride one
    // save); phase 1 replays the structure, phase 2 writes the shifted
    // name against the shifted bytes.
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const workbookXml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(workbookXml).toContain('<definedName name="GlobalTotal">Data!$A$3:$A$7</definedName>')
    // The built-in print titles survived the split (kept verbatim in both
    // phases — the phase-1 structural shift does not touch them because
    // they are whole-row refs: $1:$1 shifts to $3:$3).
    expect(workbookXml).toContain('name="_xlnm.Print_Titles"')
    // The unrelated sheet-scoped name survived untouched.
    expect(workbookXml).toContain(
      '<definedName name="LocalTotal" localSheetId="1">Other!$B$2:$B$4</definedName>',
    )

    // REOPEN: the shifted reference is the file truth.
    await openFixture(page, saved, 'e2e-ribbon-names-reopen-structural.xlsx')
    const reopened = await readLiveNames(page)
    expect(reopened.filter((n) => n.name === 'GlobalTotal')[0]?.ref).toBe('Data!$A$3:$A$7')

    expect(pageErrors).toEqual([])
  })

  test('10: an unparseable section opens namesLocked and refuses edits (fail closed)', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesLockedFixture()
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    await openFixture(page, fixture, 'e2e-ribbon-names-locked.xlsx')
    const snapshot = (await (await openResponsePromise).json()).snapshot as {
      definedNames?: unknown
      namesLocked?: boolean
    }
    expect(snapshot.namesLocked).toBe(true)
    expect(snapshot.definedNames).toBeUndefined()
    // Nothing installed — the model never saw every entry.
    expect(await readLiveNames(page)).toEqual([])

    // The Name Manager opens with the lock message and refuses every
    // action (the inputs are disabled + the delete buttons disabled).
    await openNameManager(page)
    await expect(page.getByTestId('name-manager-locked')).toBeVisible()
    await expect(page.getByTestId('name-manager-name')).toBeDisabled()
    await expect(page.getByTestId('name-manager-apply')).toBeDisabled()

    // A no-op save (plain cell edit) preserves the section byte-for-byte.
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()
    await page.evaluate(() => {
      const runtime = (
        window as {
          __genofficeExcelRuntime?: {
            univerAPI: {
              getActiveWorkbook: () => {
                getActiveSheet: () => {
                  getRange: (row: number, col: number) => { setValue: (v: unknown) => unknown }
                }
              }
            }
          }
        }
      ).__genofficeExcelRuntime
      runtime?.univerAPI
        ?.getActiveWorkbook?.()
        ?.getActiveSheet?.()
        ?.getRange?.(4, 0)
        ?.setValue?.(42)
    })
    await page.waitForTimeout(600)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const savedWorkbook = await readZipEntry(saved, 'xl/workbook.xml')
    const originalWorkbook = await readZipEntry(fixture, 'xl/workbook.xml')
    const section = (xml: string): string =>
      xml.slice(xml.indexOf('<definedNames'), xml.indexOf('</definedNames>'))
    expect(section(savedWorkbook)).toBe(section(originalWorkbook))

    expect(pageErrors).toEqual([])
  })

  test('11: a GENUINE same-scope duplicate fails the save closed', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesCollisionFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-collision.xlsx')

    // The fixture carries Excel_Version TWICE at workbook scope (live +
    // #REF! residue — which Excel itself never writes). The reader models
    // the live winner and preserves the #REF! loser.
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'Excel_Version')).toHaveLength(1)
    expect(live.filter((n) => n.name === 'Excel_Version')[0]?.ref).toBe('Data!$D$1')

    // Editing the OTHER name still fails the save: the declarative
    // snapshot carries the modeled Excel_Version while the preserve list
    // carries its same-scope twin — the canonical writer's collision
    // guard rejects the combination (fail-closed for TRUE same-scope
    // duplicates; a cross-scope pair, by contrast, saves cleanly — see
    // the same-name scenarios below).
    await openNameManager(page)
    await page.getByTestId('name-manager-row').filter({ hasText: 'GlobalTotal' }).click()
    await page.getByTestId('name-manager-ref').fill('=Data!$A$1:$A$3')
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect(page.getByText('Save failed: The name "Excel_Version" also exists')).toBeVisible({
      timeout: 15_000,
    })

    expect(pageErrors).toEqual([])
  })

  test('12: a same-name workbook + sheet-scoped pair opens, installs, and lists both', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/workbooks/open') && r.request().method() === 'POST',
    )
    const fixture = await buildExcelNamesSameNameFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-pair-open.xlsx')
    expect((await openResponsePromise).status()).toBe(200)

    // READ: the snapshot models BOTH definitions — the architect's
    // blocker correction (the same name at workbook and sheet scope is
    // two legitimate entries; nothing is preserved, nothing is lost).
    const snapshot = (await (await openResponsePromise).json()).snapshot as {
      definedNames?: {
        names: Array<{ name: string; formula: string; sheetIndex?: number }>
        preserveNames: string[]
      }
    }
    expect(snapshot.definedNames?.names).toEqual([
      { name: 'GlobalTotal', formula: 'Data!$A$1:$A$5' },
      { name: 'Total', formula: 'Data!$B$2:$B$4' },
      { name: 'Total', formula: 'Data!$C$7:$C$9', sheetIndex: 0 },
    ])
    expect(snapshot.definedNames?.preserveNames).toEqual([])

    // IMPORT: the live engine model holds BOTH Totals — the engine's
    // defined-name service is id-keyed; the sheet-scoped twin rides the
    // public sibling param path.
    const live = await readLiveNames(page)
    const totals = live.filter((n) => n.name === 'Total')
    expect(totals).toHaveLength(2)
    expect(totals.find((n) => n.localSheetId === 'AllDefaultWorkbook')?.ref).toBe('Data!$B$2:$B$4')
    const scoped = totals.find((n) => n.localSheetId !== 'AllDefaultWorkbook')
    expect(scoped?.localSheetId).toBeTruthy()
    expect(scoped?.ref).toBe('Data!$C$7:$C$9')
    expect(live.filter((n) => n.name === 'GlobalTotal')).toHaveLength(1)
    expect(live).toHaveLength(3)

    // LIST: the Name Manager shows both definitions with their true
    // scopes (Workbook vs the Data sheet).
    await openNameManager(page)
    const rows = page.getByTestId('name-manager-row')
    await expect(rows).toHaveCount(3)
    await expect(rows.filter({ hasText: 'Data!$B$2:$B$4' })).toHaveText(/Workbook/)
    await expect(rows.filter({ hasText: 'Data!$C$7:$C$9' })).toHaveText(/Data/)
    await expect(rows.filter({ hasText: 'GlobalTotal' })).toHaveCount(1)

    expect(pageErrors).toEqual([])
  })

  test('13: the Name Box resolves the pair with Excel scope precedence', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesSameNameFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-pair-resolve.xlsx')

    const nameBox = page.getByTestId('excel-name-box')

    // ON DATA (the scoped sheet): the sheet-scoped definition shadows the
    // workbook one — 'Total' jumps to C7.
    await nameBox.click()
    await nameBox.fill('Total')
    await nameBox.press('Enter')
    await page.waitForTimeout(800)
    expect(await readActiveCell(page)).toBe('C7')
    expect(await readActiveSheetName(page)).toBe('Data')
    // Case-insensitive resolution picks the same scoped winner.
    await nameBox.click()
    await nameBox.fill('total')
    await nameBox.press('Enter')
    await page.waitForTimeout(800)
    expect(await readActiveCell(page)).toBe('C7')

    // Switch to the Other sheet (via a plain address jump), then 'Total'
    // resolves through the WORKBOOK definition — B2 on Data.
    await nameBox.click()
    await nameBox.fill('Other!B2')
    await nameBox.press('Enter')
    await page.waitForTimeout(800)
    expect(await readActiveSheetName(page)).toBe('Other')
    await nameBox.click()
    await nameBox.fill('Total')
    await nameBox.press('Enter')
    await page.waitForTimeout(1000)
    expect(await readActiveSheetName(page)).toBe('Data')
    expect(await readActiveCell(page)).toBe('B2')

    expect(pageErrors).toEqual([])
  })

  test('14: editing one same-name definition preserves the other (both directions)', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesSameNameFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-pair-edit.xlsx')

    // Direction 1: edit the WORKBOOK-scoped Total.
    await openNameManager(page)
    await page.getByTestId('name-manager-row').filter({ hasText: 'Data!$B$2:$B$4' }).click()
    await page.getByTestId('name-manager-ref').fill('=Data!$B$2:$B$6')
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    const saved1 = await clickSaveAndCaptureDownload(page, 'Save')
    const xml1 = await readZipEntry(saved1, 'xl/workbook.xml')
    expect(xml1).toContain('<definedName name="Total">Data!$B$2:$B$6</definedName>')
    // THE BLOCKER INVARIANT: the sheet-scoped twin survives byte-verbatim.
    expect(xml1).toContain(
      '<definedName name="Total" localSheetId="0">Data!$C$7:$C$9</definedName>',
    )
    expect(xml1).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>')
    expect(xml1).toContain('name="_xlnm.Print_Titles"')

    // REOPEN: both definitions are the file truth.
    await openFixture(page, saved1, 'e2e-ribbon-names-pair-reopen-edit.xlsx')
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'Total')).toHaveLength(2)
    expect(
      live.find((n) => n.name === 'Total' && n.localSheetId === 'AllDefaultWorkbook')?.ref,
    ).toBe('Data!$B$2:$B$6')
    expect(
      live.find((n) => n.name === 'Total' && n.localSheetId !== 'AllDefaultWorkbook')?.ref,
    ).toBe('Data!$C$7:$C$9')

    // Direction 2: edit the SHEET-SCOPED Total on the reopened file.
    await openNameManager(page)
    await page.getByTestId('name-manager-row').filter({ hasText: 'Data!$C$7:$C$9' }).click()
    await page.getByTestId('name-manager-ref').fill('=Data!$C$7:$C$10')
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    const saved2 = await clickSaveAndCaptureDownload(page, 'Save')
    const xml2 = await readZipEntry(saved2, 'xl/workbook.xml')
    expect(xml2).toContain(
      '<definedName name="Total" localSheetId="0">Data!$C$7:$C$10</definedName>',
    )
    // The workbook twin is untouched by the scoped edit.
    expect(xml2).toContain('<definedName name="Total">Data!$B$2:$B$6</definedName>')

    expect(pageErrors).toEqual([])
  })

  test('15: deleting one same-name definition preserves the other', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesSameNameFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-pair-delete.xlsx')

    // Delete the SHEET-SCOPED twin through the real dialog.
    await openNameManager(page)
    await page
      .getByTestId('name-manager-row')
      .filter({ hasText: 'Data!$C$7:$C$9' })
      .getByRole('button', { name: /Delete Total/ })
      .click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const xml = await readZipEntry(saved, 'xl/workbook.xml')
    // The scoped twin is gone; the workbook twin survives byte-verbatim.
    expect(xml).not.toContain('<definedName name="Total" localSheetId="0">')
    expect(xml).toContain('<definedName name="Total">Data!$B$2:$B$4</definedName>')
    expect(xml).toContain('<definedName name="GlobalTotal">Data!$A$1:$A$5</definedName>')
    expect(xml).toContain('name="_xlnm.Print_Titles"')

    // REOPEN: exactly one Total remains, the workbook one.
    await openFixture(page, saved, 'e2e-ribbon-names-pair-reopen-delete.xlsx')
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'Total')).toHaveLength(1)
    expect(live.filter((n) => n.name === 'Total')[0]?.localSheetId).toBe('AllDefaultWorkbook')

    expect(pageErrors).toEqual([])
  })

  test('16: a structural edit shifts BOTH same-name definitions (split-save)', async ({ page }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesSameNameFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-pair-structural.xlsx')

    // Insert 2 rows at the top of Data through the REAL Univer facade.
    // UpdateDefinedNameController rewrites EVERY id-keyed entry — both
    // Totals shift (B2:B4 → B4:B6, C7:C9 → C9:C11).
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
      runtime?.univerAPI?.getActiveWorkbook?.()?.getSheetByName?.('Data')?.insertRows?.(0, 2)
    })
    await page.waitForTimeout(1000)
    const live = await readLiveNames(page)
    expect(
      live.find((n) => n.name === 'Total' && n.localSheetId === 'AllDefaultWorkbook')?.ref,
    ).toBe('Data!$B$4:$B$6')
    expect(
      live.find((n) => n.name === 'Total' && n.localSheetId !== 'AllDefaultWorkbook')?.ref,
    ).toBe('Data!$C$9:$C$11')

    // Save — the split-save writes BOTH shifted definitions.
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const xml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(xml).toContain('<definedName name="Total">Data!$B$4:$B$6</definedName>')
    expect(xml).toContain(
      '<definedName name="Total" localSheetId="0">Data!$C$9:$C$11</definedName>',
    )
    expect(xml).toContain('<definedName name="GlobalTotal">Data!$A$3:$A$7</definedName>')

    // REOPEN: both shifted references are the file truth.
    await openFixture(page, saved, 'e2e-ribbon-names-pair-reopen-structural.xlsx')
    const reopened = await readLiveNames(page)
    expect(
      reopened.find((n) => n.name === 'Total' && n.localSheetId === 'AllDefaultWorkbook')?.ref,
    ).toBe('Data!$B$4:$B$6')
    expect(
      reopened.find((n) => n.name === 'Total' && n.localSheetId !== 'AllDefaultWorkbook')?.ref,
    ).toBe('Data!$C$9:$C$11')

    expect(pageErrors).toEqual([])
  })

  test('17: creating a same-name pair at a NEW scope works; the same scope is refused', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelNamesSameNameFixture()
    await openFixture(page, fixture, 'e2e-ribbon-names-pair-create.xlsx')

    await openNameManager(page)
    // 'Total' already exists at workbook + Data scope. Creating it at the
    // OTHER sheet's scope is a legal Excel pair — the scope-aware create
    // path installs it (the engine's scope-blind builder would reject it).
    await page.getByTestId('name-manager-name').fill('Total')
    await page.getByTestId('name-manager-ref').fill('=Other!$B$2')
    await page.getByTestId('name-manager-scope').selectOption({ label: 'Other' })
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('Defined names updated')).toBeVisible({ timeout: 5_000 })
    const rows = page.getByTestId('name-manager-row')
    await expect(rows.filter({ hasText: 'Other!$B$2' })).toHaveCount(1)

    // Creating 'Total' at WORKBOOK scope (already taken) is refused with
    // the scope-aware message — no collision, no silent overwrite.
    await page.getByTestId('name-manager-name').fill('Total')
    await page.getByTestId('name-manager-ref').fill('=Data!$D$1')
    await page.getByTestId('name-manager-scope').selectOption({ label: 'Workbook' })
    await page.getByTestId('name-manager-apply').click()
    await expect(page.getByText('already defined at this scope')).toBeVisible({ timeout: 5_000 })

    await page
      .getByTestId('name-manager-dialog')
      .getByRole('button', { name: 'Close', exact: true })
      .click()

    // Save: THREE Total definitions ride the declarative snapshot.
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const xml = await readZipEntry(saved, 'xl/workbook.xml')
    expect(xml).toContain('<definedName name="Total">Data!$B$2:$B$4</definedName>')
    expect(xml).toContain('<definedName name="Total" localSheetId="0">Data!$C$7:$C$9</definedName>')
    expect(xml).toContain('<definedName name="Total" localSheetId="1">Other!$B$2</definedName>')

    // REOPEN: all three are the file truth.
    await openFixture(page, saved, 'e2e-ribbon-names-pair-reopen-create.xlsx')
    const live = await readLiveNames(page)
    expect(live.filter((n) => n.name === 'Total')).toHaveLength(3)

    expect(pageErrors).toEqual([])
  })
})
