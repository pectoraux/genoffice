/**
 * REAL browser E2E — Insert → Picture / image edit (EXCEL-022).
 *
 * Proves the image persistence chain end-to-end through the REAL HTTP
 * boundary:
 *
 *   open → readBasicWorkbook resolves the drawing relationship chain →
 *   WorksheetState.images (typed metadata + inline media) → browser seeds
 *   its file-image refs and installs the pictures into the REAL Univer
 *   over-grid model under journal suppression (locator ids) → user
 *   moves/resizes/deletes through the Univer facade → save snapshots the
 *   LIVE anchors → visualEdits/visualAdditions families →
 *   /api/office/workbooks/save → routeOffice strict validation →
 *   applyCellEditsToXlsx trailing params → applyVisualEdits /
 *   applyVisualAdditions rewrite the drawing XML (delete cascades the
 *   image relationship + media part only when unreferenced) → reopen →
 *   file-native image state.
 *
 * One-cell anchors fail closed: the browser reverts a refused move and a
 * no-op save preserves the drawing parts byte-for-byte.
 *
 * Fixtures: buildExcelImageFixture variants (PNG/JPEG, one/two images,
 * multi-sheet, oneCellAnchor).
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
  buildExcelImageFixture,
  buildExcelImagesFixture,
  buildExcelMultiSheetImageFixture,
  buildExcelJpegImageFixture,
  buildExcelOneCellImageFixture,
  buildExcelAbsoluteImageFixture,
  readZipEntry,
  readZipEntryBytes,
  listZipEntries,
  buildSolidPng,
} from './fixtures'

const OPEN_INPUT = 'input[accept=".xlsx,.csv,.xls"]'
const IMAGE_INPUT = 'input[data-testid="excel-image-input"]'

/** Snapshot view of the open response's image state. */
interface ImageSnapshotView {
  sheets: Array<{
    name: string
    images?: Array<{
      drawingPath: string
      drawingIndex: number
      anchorType: 'two-cell' | 'one-cell'
      anchor: Record<string, number>
      mediaType: string
      dataUrl: string
      name?: string
    }>
  }>
}

/** Live over-grid image view (read through the runtime's facade). */
interface LiveImageView {
  id: string
  source: string
  fromColumn: number
  fromRow: number
  toColumn: number
  toRow: number
}

async function openFixture(page: Page, bytes: Buffer, path: string): Promise<ImageSnapshotView> {
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
  return ((await openResponse.json()).snapshot ?? {}) as ImageSnapshotView
}

/** Reads the live over-grid images of a sheet through the exposed runtime. */
async function liveImages(page: Page, sheetName: string): Promise<LiveImageView[]> {
  return page.evaluate(
    async ({ name }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(sheetName: string): {
                getImages(): Array<{
                  getId(): string
                  // PUBLIC geometry read — same surface the app code uses
                  // (toBuilder().buildAsync()); no private internals.
                  toBuilder(): {
                    buildAsync(): Promise<{
                      source?: string
                      sheetTransform?: {
                        from?: { column?: number; row?: number }
                        to?: { column?: number; row?: number }
                      }
                    }>
                  }
                }>
              }
            } | null
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const wb = rt.univerAPI.getActiveWorkbook()
      if (!wb) throw new Error('no active workbook')
      const ws = wb.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      const views: Array<{
        id: string
        source: string
        fromColumn: number
        fromRow: number
        toColumn: number
        toRow: number
      }> = []
      for (const image of ws.getImages()) {
        const data = await image.toBuilder().buildAsync()
        views.push({
          id: image.getId(),
          source: data.source ?? '',
          fromColumn: data.sheetTransform?.from?.column ?? -1,
          fromRow: data.sheetTransform?.from?.row ?? -1,
          toColumn: data.sheetTransform?.to?.column ?? -1,
          toRow: data.sheetTransform?.to?.row ?? -1,
        })
      }
      return views
    },
    { name: sheetName },
  )
}

/** Moves a live image through the real Univer facade (user drag parity). */
async function moveImage(
  page: Page,
  sheetName: string,
  id: string,
  row: number,
  column: number,
): Promise<boolean> {
  return page.evaluate(
    ({ name, id: imageId, row: targetRow, column: targetColumn }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(sheetName: string): {
                getImages(): Array<{
                  getId(): string
                  setPositionAsync(
                    row: number,
                    column: number,
                    rowOffset?: number,
                    columnOffset?: number,
                  ): Promise<boolean>
                }>
              }
            } | null
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const ws = rt.univerAPI.getActiveWorkbook()?.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      const image = ws.getImages().find((entry) => entry.getId() === imageId)
      if (!image) throw new Error('image not found')
      return image.setPositionAsync(targetRow, targetColumn)
    },
    { name: sheetName, id, row, column },
  )
}

/** Resizes a live image through the real Univer facade. */
async function resizeImage(
  page: Page,
  sheetName: string,
  id: string,
  width: number,
  height: number,
): Promise<boolean> {
  return page.evaluate(
    ({ name, id: imageId, width: targetWidth, height: targetHeight }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(sheetName: string): {
                getImages(): Array<{
                  getId(): string
                  setSizeAsync(width: number, height: number): Promise<boolean>
                }>
              }
            } | null
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const ws = rt.univerAPI.getActiveWorkbook()?.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      const image = ws.getImages().find((entry) => entry.getId() === imageId)
      if (!image) throw new Error('image not found')
      return image.setSizeAsync(targetWidth, targetHeight)
    },
    { name: sheetName, id, width, height },
  )
}

/** Deletes a live image through the real Univer facade. */
async function removeImage(page: Page, sheetName: string, id: string): Promise<boolean> {
  return page.evaluate(
    ({ name, id: imageId }) => {
      const w = window as {
        __genofficeExcelRuntime?: {
          univerAPI: {
            getActiveWorkbook(): {
              getSheetByName(sheetName: string): {
                getImages(): Array<{ getId(): string; remove(): boolean }>
              }
            } | null
          }
        }
      }
      const rt = w.__genofficeExcelRuntime
      if (!rt) throw new Error('runtime not exposed')
      const ws = rt.univerAPI.getActiveWorkbook()?.getSheetByName(name)
      if (!ws) throw new Error(`sheet ${name} not found`)
      const image = ws.getImages().find((entry) => entry.getId() === imageId)
      if (!image) throw new Error('image not found')
      return image.remove()
    },
    { name: sheetName, id },
  )
}

test.describe('Insert tab — Images persist through the canonical pipeline', () => {
  test.setTimeout(240_000)

  test('1-3: opening a workbook with an image surfaces typed metadata and renders it in the correct sheet', async ({
    page,
  }) => {
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const snapshot = await openFixture(
      page,
      await buildExcelImageFixture(),
      '/tmp/e2e-ribbon-images-open.xlsx',
    )

    // READ path: the snapshot carried the typed image metadata (locator,
    // anchor, inline media, name).
    const image = snapshot.sheets[0]?.images?.[0]
    expect(image, 'snapshot carries the image').toBeDefined()
    expect(image!.drawingPath).toBe('xl/drawings/drawing1.xml')
    expect(image!.drawingIndex).toBe(0)
    expect(image!.anchorType).toBe('two-cell')
    expect(image!.anchor).toEqual({
      fromRow: 2,
      fromColumn: 1,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 12,
      toColumn: 6,
      toRowOffset: 0,
      toColumnOffset: 0,
    })
    expect(image!.mediaType).toBe('image/png')
    expect(image!.dataUrl.startsWith('data:image/png;base64,')).toBe(true)
    expect(image!.name).toBe('Red dot')

    // RENDER path: the image installed into the REAL Univer over-grid
    // model on the correct sheet, with the locator as its drawing id and
    // the wire bytes as its source.
    const live = await liveImages(page, 'Data')
    expect(live).toHaveLength(1)
    expect(live[0]!.id).toBe('file-img:xl/drawings/drawing1.xml#0')
    expect(live[0]!.source.startsWith('data:image/png;base64,')).toBe(true)
    expect(live[0]!.fromColumn).toBe(1)
    expect(live[0]!.fromRow).toBe(2)
    expect(live[0]!.toColumn).toBe(6)
    expect(live[0]!.toRow).toBe(12)

    // Loading the image must NOT dirty the workbook (no undo pollution).
    await expect(page.getByText('● Unsaved changes')).toBeHidden({ timeout: 3000 })
    expect(pageErrors).toEqual([])
  })

  test('4: moving an image persists only the anchor geometry', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelImageFixture()
    const snapshot = await openFixture(page, fixture, '/tmp/e2e-ribbon-images-move.xlsx')
    const id = `file-img:${snapshot.sheets[0]!.images![0]!.drawingPath}#${snapshot.sheets[0]!.images![0]!.drawingIndex}`

    expect(await moveImage(page, 'Data', id, 5, 3)).toBe(true)
    await page.waitForTimeout(300)
    const moved = await liveImages(page, 'Data')
    expect(moved[0]!.fromRow).toBe(5)
    expect(moved[0]!.fromColumn).toBe(3)

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).toContain('<xdr:row>5</xdr:row>')
    expect(drawing).toContain('<xdr:col>3</xdr:col>')
    // Only the anchor changed: the media bytes and the relationship are
    // byte-identical, and no second anchor appeared.
    expect(await readZipEntryBytes(saved, 'xl/media/image1.png')).toEqual(
      await readZipEntryBytes(fixture, 'xl/media/image1.png'),
    )
    expect((drawing.match(/<xdr:twoCellAnchor/g) ?? []).length).toBe(1)
    const rels = await readZipEntry(saved, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).toContain('rId1')
  })

  test('5: resizing an image persists the new dimensions', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const snapshot = await openFixture(
      page,
      await buildExcelImageFixture(),
      '/tmp/e2e-ribbon-images-resize.xlsx',
    )
    const id = `file-img:${snapshot.sheets[0]!.images![0]!.drawingPath}#${snapshot.sheets[0]!.images![0]!.drawingIndex}`

    expect(await resizeImage(page, 'Data', id, 120, 60)).toBe(true)
    await page.waitForTimeout(300)

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    // The to-marker moved with the new size (from stays at B3).
    expect(drawing).toContain('<xdr:from><xdr:col>1</xdr:col>')
    const to = /<xdr:to><xdr:col>(\d+)<\/xdr:col>/.exec(drawing)?.[1]
    expect(Number(to)).toBeGreaterThan(1)
    expect(await readZipEntryBytes(saved, 'xl/media/image1.png')).not.toBeNull()
  })

  test('6: inserting a picture creates the full canonical part set', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelImageFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-images-insert.xlsx')

    const png = buildSolidPng(24, 12, [250, 200, 30])
    const path = '/tmp/e2e-insert-picture.png'
    writeFileSync(path, png)
    await page.setInputFiles(IMAGE_INPUT, path)
    await expect(page.getByText('Picture inserted')).toBeVisible({ timeout: 15_000 })

    // The image installed into the live model with a session id.
    const live = await liveImages(page, 'Data')
    expect(live).toHaveLength(2)
    expect(live.some((entry) => entry.id.startsWith('added-img-'))).toBe(true)

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const entries = await listZipEntries(saved)
    expect(entries).toContain('xl/media/image2.png')
    expect(await readZipEntryBytes(saved, 'xl/media/image2.png')).toEqual(png)
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect((drawing.match(/<xdr:twoCellAnchor/g) ?? []).length).toBe(2)
    const rels = await readZipEntry(saved, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).toContain('image2.png')
    expect(rels).toContain('relationships/image')
    const contentTypes = await readZipEntry(saved, '[Content_Types].xml')
    expect(contentTypes).toContain('Extension="png"')
    // The pre-existing image survived untouched.
    expect(await readZipEntryBytes(saved, 'xl/media/image1.png')).toEqual(
      await readZipEntryBytes(fixture, 'xl/media/image1.png'),
    )
  })

  test('7: deleting an image cascades the relationship and the media part', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const snapshot = await openFixture(
      page,
      await buildExcelImageFixture(),
      '/tmp/e2e-ribbon-images-delete.xlsx',
    )
    const id = `file-img:${snapshot.sheets[0]!.images![0]!.drawingPath}#${snapshot.sheets[0]!.images![0]!.drawingIndex}`

    expect(await removeImage(page, 'Data', id)).toBe(true)
    await page.waitForTimeout(300)
    expect(await liveImages(page, 'Data')).toHaveLength(0)

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const entries = await listZipEntries(saved)
    expect(entries).not.toContain('xl/media/image1.png')
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).not.toContain('<xdr:pic>')
    const rels = await readZipEntry(saved, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).not.toContain('rId1')
    const contentTypes = await readZipEntry(saved, '[Content_Types].xml')
    expect(contentTypes).not.toContain('Extension="png"')
  })

  test('8: operating on one image never corrupts the other', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelImagesFixture()
    const snapshot = await openFixture(page, fixture, '/tmp/e2e-ribbon-images-isolation.xlsx')
    const images = snapshot.sheets[0]!.images!
    const firstId = `file-img:${images[0]!.drawingPath}#${images[0]!.drawingIndex}`

    // Move the FIRST image; the second stays exactly where it was.
    expect(await moveImage(page, 'Data', firstId, 8, 2)).toBe(true)
    await page.waitForTimeout(300)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).toContain('<xdr:row>8</xdr:row>')
    // The second anchor is verbatim (from col 8 row 4).
    expect(drawing).toContain('<xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row>')
    expect((drawing.match(/<xdr:twoCellAnchor/g) ?? []).length).toBe(2)
    expect(await readZipEntryBytes(saved, 'xl/media/image2.png')).toEqual(
      await readZipEntryBytes(fixture, 'xl/media/image2.png'),
    )

    // Now DELETE the second image in the same session — the first (moved)
    // image and its media survive; only the second media part goes.
    const secondId = `file-img:${images[1]!.drawingPath}#${images[1]!.drawingIndex}`
    expect(await removeImage(page, 'Data', secondId)).toBe(true)
    await page.waitForTimeout(300)
    const saved2 = await clickSaveAndCaptureDownload(page, 'Save')
    const entries = await listZipEntries(saved2)
    expect(entries).toContain('xl/media/image1.png')
    expect(entries).not.toContain('xl/media/image2.png')
    const drawing2 = await readZipEntry(saved2, 'xl/drawings/drawing1.xml')
    expect((drawing2.match(/<xdr:twoCellAnchor/g) ?? []).length).toBe(1)
    expect(drawing2).toContain('<xdr:row>8</xdr:row>')
  })

  test('9: images survive save and reopen with file-native semantics', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const snapshot = await openFixture(
      page,
      await buildExcelImageFixture(),
      '/tmp/e2e-ribbon-images-reopen.xlsx',
    )
    const id = `file-img:${snapshot.sheets[0]!.images![0]!.drawingPath}#${snapshot.sheets[0]!.images![0]!.drawingIndex}`
    expect(await moveImage(page, 'Data', id, 6, 4)).toBe(true)
    await page.waitForTimeout(300)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')

    // Reopen the saved bytes: the image re-surfaces with the MOVED anchor.
    const reopened = await openFixture(page, saved, '/tmp/e2e-ribbon-images-reopened.xlsx')
    const image = reopened.sheets[0]?.images?.[0]
    expect(image).toBeDefined()
    expect(image!.drawingPath).toBe('xl/drawings/drawing1.xml')
    expect(image!.anchor.fromRow).toBe(6)
    expect(image!.anchor.fromColumn).toBe(4)
    expect(image!.mediaType).toBe('image/png')
    const live = await liveImages(page, 'Data')
    expect(live).toHaveLength(1)
    expect(live[0]!.fromRow).toBe(6)
    expect(live[0]!.fromColumn).toBe(4)
  })

  test('10: a no-op image save preserves media, drawings, rels, and content types byte-for-byte', async ({
    page,
  }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelImagesFixture()
    await openFixture(page, fixture, '/tmp/e2e-ribbon-images-noop.xlsx')

    // Dirty the workbook with a plain cell edit (the Save button enables
    // on dirty) — the IMAGE state stays untouched, which is the no-op
    // under test. The save request must carry NO visual families.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    await page.keyboard.press('Control+Home')
    await page.waitForTimeout(200)
    await page.locator('#genoffice-web-excel').click({ position: { x: 320, y: 8 } })
    await page.keyboard.type('noop')
    await page.keyboard.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { visualEdits?: unknown; visualAdditions?: unknown }
    }
    expect(saveBody.savePlan.visualEdits, 'no image interaction → no visualEdits').toBeUndefined()
    expect(
      saveBody.savePlan.visualAdditions,
      'no image interaction → no visualAdditions',
    ).toBeUndefined()

    for (const path of [
      'xl/media/image1.png',
      'xl/media/image2.png',
      'xl/drawings/drawing1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
      'xl/worksheets/_rels/sheet1.xml.rels',
      '[Content_Types].xml',
    ]) {
      expect(await readZipEntryBytes(saved, path), `${path} preserved byte-for-byte`).toEqual(
        await readZipEntryBytes(fixture, path),
      )
    }
    // The worksheet's drawing wiring survives the cell edit verbatim.
    const worksheet = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<drawing r:id="rIdDrw1"/>')
  })

  test('11: the worksheet → drawing → image → media relationship chain stays valid through move and insert', async ({
    page,
  }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const snapshot = await openFixture(
      page,
      await buildExcelImageFixture(),
      '/tmp/e2e-ribbon-images-chain.xlsx',
    )
    const id = `file-img:${snapshot.sheets[0]!.images![0]!.drawingPath}#${snapshot.sheets[0]!.images![0]!.drawingIndex}`
    expect(await moveImage(page, 'Data', id, 3, 2)).toBe(true)
    await page.waitForTimeout(300)

    const png = buildSolidPng(10, 10, [30, 30, 120])
    writeFileSync('/tmp/e2e-chain-second.png', png)
    await page.setInputFiles(IMAGE_INPUT, '/tmp/e2e-chain-second.png')
    await expect(page.getByText('Picture inserted')).toBeVisible({ timeout: 15_000 })

    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    // worksheet → drawing relationship
    const sheetRels = await readZipEntry(saved, 'xl/worksheets/_rels/sheet1.xml.rels')
    const drawingRelId = /Id="(rIdDrw1)"/.exec(sheetRels)?.[1]
    expect(drawingRelId).toBeDefined()
    const worksheet = await readZipEntry(saved, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain(`<drawing r:id="${drawingRelId}"/>`)
    // drawing part exists with two pictures
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect((drawing.match(/<xdr:pic>/g) ?? []).length).toBe(2)
    // drawing rels → media parts resolve
    const rels = await readZipEntry(saved, 'xl/drawings/_rels/drawing1.xml.rels')
    const targets = [...rels.matchAll(/Target="\.\.\/media\/(image\d+\.png)"/g)].map((m) => m[1]!)
    expect(targets).toHaveLength(2)
    for (const target of targets) {
      expect(await readZipEntryBytes(saved, `xl/media/${target}`)).not.toBeNull()
    }
    // both embed ids referenced by the drawing exist in the rels
    for (const embed of [...drawing.matchAll(/r:embed="(rId\d+)"/g)].map((m) => m[1]!)) {
      expect(rels).toContain(`Id="${embed}"`)
    }
  })

  test('12: a one-cell anchored image renders but refuses edits (fail closed)', async ({
    page,
  }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelOneCellImageFixture()
    const snapshot = await openFixture(page, fixture, '/tmp/e2e-ribbon-images-onecell.xlsx')
    const image = snapshot.sheets[0]?.images?.[0]
    expect(image?.anchorType).toBe('one-cell')

    // The picture still renders (read-only display).
    const live = await liveImages(page, 'Data')
    expect(live).toHaveLength(1)
    const id = live[0]!.id

    // A move attempt is refused: the image snaps back and the status
    // explains why — the journal never records the change.
    expect(await moveImage(page, 'Data', id, 9, 9)).toBe(true)
    await expect(
      page.getByText(
        'This image uses a one-cell or absolute anchor — moving or resizing it is not supported yet.',
      ),
    ).toBeVisible({ timeout: 10_000 })
    await page.waitForTimeout(800)
    const reverted = await liveImages(page, 'Data')
    expect(reverted[0]!.fromRow).toBe(2)
    expect(reverted[0]!.fromColumn).toBe(1)

    // The refused edit never journaled — dirty the workbook with a plain
    // cell edit (image state untouched) to enable Save, then prove the
    // save carries NO visualEdits: the drawing XML stays byte-identical.
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/workbooks/save') && r.method() === 'POST',
    )
    await page.keyboard.press('Control+Home')
    await page.waitForTimeout(200)
    await page.locator('#genoffice-web-excel').click({ position: { x: 320, y: 8 } })
    await page.keyboard.type('ro')
    await page.keyboard.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    const req = await saveReq
    const saveBody = JSON.parse(req.postData() ?? '{}') as {
      savePlan: { visualEdits?: unknown }
    }
    expect(saveBody.savePlan.visualEdits, 'refused edit → no visualEdits').toBeUndefined()
    expect(await readZipEntryBytes(saved, 'xl/drawings/drawing1.xml')).toEqual(
      await readZipEntryBytes(fixture, 'xl/drawings/drawing1.xml'),
    )
    expect(await readZipEntryBytes(saved, 'xl/media/image1.png')).toEqual(
      await readZipEntryBytes(fixture, 'xl/media/image1.png'),
    )
  })

  test('13: multiple sheets carry their images independently; JPEG imports through the writer-supported set', async ({
    page,
  }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const snapshot = await openFixture(
      page,
      await buildExcelMultiSheetImageFixture(),
      '/tmp/e2e-ribbon-images-multisheet.xlsx',
    )
    expect(snapshot.sheets[0]?.images).toHaveLength(1)
    expect(snapshot.sheets[1]?.images).toHaveLength(1)
    const first = await liveImages(page, 'First')
    const second = await liveImages(page, 'Second')
    expect(first).toHaveLength(1)
    expect(second).toHaveLength(1)
    // Deleting the SECOND sheet's image keeps the first sheet's picture
    // and the SHARED media part (both sheets referenced it).
    expect(await removeImage(page, 'Second', second[0]!.id)).toBe(true)
    await page.waitForTimeout(300)
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    expect(await readZipEntryBytes(saved, 'xl/media/image1.png')).not.toBeNull()
    const drawing2 = await readZipEntry(saved, 'xl/drawings/drawing2.xml')
    expect(drawing2).not.toContain('<xdr:pic>')
    const stillThere = await liveImages(page, 'First')
    expect(stillThere).toHaveLength(1)
  })

  test('14: a JPEG image imports and round-trips byte-for-byte', async ({ page }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    const fixture = await buildExcelJpegImageFixture()
    const snapshot = await openFixture(page, fixture, '/tmp/e2e-ribbon-images-jpeg.xlsx')
    const image = snapshot.sheets[0]?.images?.[0]
    expect(image?.mediaType).toBe('image/jpeg')
    expect(image?.dataUrl.startsWith('data:image/jpeg;base64,')).toBe(true)
    const live = await liveImages(page, 'Data')
    expect(live).toHaveLength(1)

    // Dirty via a cell edit (image state untouched), then save.
    await page.keyboard.press('Control+Home')
    await page.waitForTimeout(200)
    await page.locator('#genoffice-web-excel').click({ position: { x: 320, y: 8 } })
    await page.keyboard.type('jpeg')
    await page.keyboard.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    expect(await readZipEntryBytes(saved, 'xl/media/image1.jpeg')).toEqual(
      await readZipEntryBytes(fixture, 'xl/media/image1.jpeg'),
    )
    const drawingSaved = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    const drawingFixture = await readZipEntry(fixture, 'xl/drawings/drawing1.xml')
    expect(drawingSaved).toEqual(drawingFixture)
    const contentTypes = await readZipEntry(saved, '[Content_Types].xml')
    expect(contentTypes).toContain('Extension="jpeg"')
  })

  test('15: an absolute-anchored image is omitted, never relocated (fail closed)', async ({
    page,
  }) => {
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/excel')
    await waitForGridCanvas(page)

    // Architect review (PR #20, blocker 2): absolute geometry cannot be
    // represented in the two-cell wire model. The reader must fail closed
    // — omit the picture from the browser model entirely and leave the
    // file untouched. A zero-marker approximation would silently
    // relocate the picture to (0,0), which is exactly what must NOT
    // happen.
    const fixture = await buildExcelAbsoluteImageFixture()
    const snapshot = await openFixture(page, fixture, '/tmp/e2e-ribbon-images-absolute.xlsx')
    expect(snapshot.sheets[0]?.images ?? []).toHaveLength(0)

    // Nothing renders over the grid.
    const live = await liveImages(page, 'Data')
    expect(live).toHaveLength(0)

    // Dirty via a cell edit (image state untouched), then save: the
    // absolute anchor XML and its media bytes survive byte-for-byte.
    await page.keyboard.press('Control+Home')
    await page.waitForTimeout(200)
    await page.locator('#genoffice-web-excel').click({ position: { x: 320, y: 8 } })
    await page.keyboard.type('abs')
    await page.keyboard.press('Enter')
    await expect(page.getByText('● Unsaved changes')).toBeVisible({ timeout: 10_000 })
    const saved = await clickSaveAndCaptureDownload(page, 'Save')
    expect(await readZipEntryBytes(saved, 'xl/drawings/drawing1.xml')).toEqual(
      await readZipEntryBytes(fixture, 'xl/drawings/drawing1.xml'),
    )
    expect(await readZipEntryBytes(saved, 'xl/media/image1.png')).toEqual(
      await readZipEntryBytes(fixture, 'xl/media/image1.png'),
    )
    const drawing = await readZipEntry(saved, 'xl/drawings/drawing1.xml')
    expect(drawing).toContain('<xdr:pos x="47625" y="9525"/>')
  })
})
