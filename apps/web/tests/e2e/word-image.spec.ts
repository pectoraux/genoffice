/**
 * REAL browser E2E — editable Word images (Phase 3 Increment 8).
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Tiptap image nodes) → Vite proxy → HTTP /api/office/documents/* →
 *   vercel-handler → routeOffice → @genoffice/docx-engine
 *   (parse + patchImageParagraphXml/applyImageWrap/NewImage embed) → DOCX bytes
 *
 * Covers: render, resize, rotate, flip, wrap change, insertion, deletion and
 * dirty-state isolation. Never calls routeOffice() directly, and uses no
 * Vite-only helper modules (works against deployed builds too).
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute } from './helpers'
import { buildWordImageFixture, buildSolidPng, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

interface WireImage {
  imageDataUrl: string | null
  widthPx?: number
  heightPx?: number
  crop?: { l: number; t: number; r: number; b: number }
  wrap?: string
  rotDeg?: number
  flipH?: boolean
  flipV?: boolean
}
interface WireBlock {
  docxIndex: number | null
  type: string
  text: string
  image?: WireImage
  newImage?: { base64: string; mime: string; widthPx: number; heightPx: number }
  edited?: boolean
  hidden?: boolean
}

const IMG = '.ProseMirror img[data-docx-image]'

test.describe('Word editable images (real HTTP + real engine)', () => {
  test('render: real <img>, docxIndex, sizes, floating + wrap metadata', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-fixture.docx', fixture)
    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-fixture.docx')
    await expect(page.getByText('Opened e2e-img-fixture.docx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)
    const blocks = (await openResponse.json()).blocks as WireBlock[]
    const images = blocks.filter((b) => b.type === 'image' && b.image)
    expect(images).toHaveLength(7)

    // Actual <img> elements render, one per image block, in document order.
    await expect(page.locator(IMG)).toHaveCount(7)
    // docxIndex survives as the schema-backed attribute.
    for (const b of images) {
      await expect(page.locator(`${IMG}[data-docx-index="${b.docxIndex}"]`)).toHaveCount(1)
    }
    // Width/height are represented (fixture sizes: 64, 80, 240×120, 100×60, 70, 70, 90).
    await expect(page.locator(`${IMG}[data-docx-index="3"]`)).toHaveAttribute('data-width', '240')
    await expect(page.locator(`${IMG}[data-docx-index="3"]`)).toHaveAttribute('data-height', '120')
    // Floating image metadata survives (offsets + position bases).
    const floating = page.locator(`${IMG}[data-docx-index="2"]`)
    await expect(floating).toHaveAttribute('data-wrap', 'square-left')
    await expect(floating).toHaveAttribute('data-offset-x', '200000')
    await expect(floating).toHaveAttribute('data-offset-y', '100000')
    await expect(floating).toHaveAttribute('data-pos-h-rel', 'column')
    await expect(floating).toHaveAttribute('data-pos-v-rel', 'paragraph')
    // Non-default wrap metadata on multiple images.
    await expect(page.locator(`${IMG}[data-docx-index="4"]`)).toHaveAttribute(
      'data-wrap',
      'topBottom',
    )
    // Rotation/flip/crop metadata render.
    await expect(page.locator(`${IMG}[data-docx-index="5"]`)).toHaveAttribute('data-rot', '90')
    await expect(page.locator(`${IMG}[data-docx-index="6"]`)).toHaveAttribute('data-flip-h', '1')
    await expect(page.locator(`${IMG}[data-docx-index="6"]`)).toHaveAttribute('data-flip-v', '1')
    await expect(page.locator(`${IMG}[data-docx-index="7"]`)).toHaveAttribute(
      'data-crop',
      '0.1,0.2,0.1,0.2',
    )
    // Pixels actually render (data-URL src on every image).
    const srcCount = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('.ProseMirror img[data-docx-image]')).filter(
        (el) => (el.getAttribute('src') ?? '').startsWith('data:image/png;base64,'),
      ).length
    })
    expect(srcCount).toBe(7)

    expect(pageErrors).toEqual([])
  })

  test('resize: dimensions change and survive save + reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-resize.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-resize.docx')
    await expect(page.getByText('Opened e2e-img-resize.docx')).toBeVisible({ timeout: 30_000 })
    await expect(page.locator(IMG)).toHaveCount(7)

    // Select the big image (240×120) — clicking an atom node selects it.
    await page.locator(`${IMG}[data-docx-index="3"]`).click()
    await expect(page.getByTestId('image-toolbar')).toBeVisible()
    // Resize to 120 preserving the aspect ratio: 120×60.
    await page.getByTestId('image-width').fill('120')
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-img-resize.docx')).toBeVisible({ timeout: 15_000 })
    const blocksSaved = JSON.parse(req.postData() ?? '{}').blocks as WireBlock[]
    expect(blocksSaved.find((b) => b.docxIndex === 3)?.edited).toBe(true)
    expect(blocksSaved.find((b) => b.docxIndex === 3)?.image?.widthPx).toBe(120)

    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('cx="1143000" cy="571500"')

    // Reopen: dimensions changed, unrelated content survived.
    writeFileSync('/tmp/e2e-img-resize-saved.docx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-resize-saved.docx')
    await expect(page.getByText('Opened e2e-img-resize-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    await reopenResponsePromise
    await expect(page.locator(IMG)).toHaveCount(7)
    await expect(page.locator(`${IMG}[data-docx-index="3"]`)).toHaveAttribute('data-width', '120')
    await expect(page.locator(`${IMG}[data-docx-index="3"]`)).toHaveAttribute('data-height', '60')
    await expect(page.locator(`${IMG}[data-docx-index="5"]`)).toHaveAttribute('data-rot', '90')
    await expect(page.getByText('First paragraph.')).toBeVisible()
    await expect(page.getByText('Last paragraph.')).toBeVisible()

    expect(pageErrors).toEqual([])
  })

  test('rotate 90°: canonical rotation metadata changes in the saved XML', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-rotate.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-rotate.docx')
    await expect(page.getByText('Opened e2e-img-rotate.docx')).toBeVisible({ timeout: 30_000 })

    // Rotate the inline image (idx1, rot 0 → 90).
    await page.locator(`${IMG}[data-docx-index="1"]`).click()
    await expect(page.getByTestId('image-toolbar')).toBeVisible()
    await page.getByTestId('image-rotate').click()
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-img-rotate.docx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    // Canonical rotation attribute: 90° * 60000.
    expect(xml).toContain('rot="5400000"')

    expect(pageErrors).toEqual([])
  })

  test('flip horizontal + vertical round-trip through save/reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-flip.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-flip.docx')
    await expect(page.getByText('Opened e2e-img-flip.docx')).toBeVisible({ timeout: 30_000 })

    // Flip the inline image (idx1: no flips → both flips).
    await page.locator(`${IMG}[data-docx-index="1"]`).click()
    await expect(page.getByTestId('image-toolbar')).toBeVisible()
    await page.getByTestId('image-flip-h').click()
    await page.getByTestId('image-flip-v').click()
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-img-flip.docx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)

    // Reopen and verify the flips round-trip.
    writeFileSync('/tmp/e2e-img-flip-saved.docx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-flip-saved.docx')
    await expect(page.getByText('Opened e2e-img-flip-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).blocks as WireBlock[]
    const r1 = reopened.find((b) => b.docxIndex === 1)?.image
    expect(r1?.flipH).toBe(true)
    expect(r1?.flipV).toBe(true)
    // The pre-flipped image keeps its flips.
    expect(reopened.find((b) => b.docxIndex === 6)?.image?.flipH).toBe(true)

    expect(pageErrors).toEqual([])
  })

  test('wrap change inline → square-left round-trips', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-wrap.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-wrap.docx')
    await expect(page.getByText('Opened e2e-img-wrap.docx')).toBeVisible({ timeout: 30_000 })

    // Change the inline image (idx1) to square-left.
    await page.locator(`${IMG}[data-docx-index="1"]`).click()
    await expect(page.getByTestId('image-toolbar')).toBeVisible()
    await page.getByTestId('image-wrap').selectOption('square-left')
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-img-wrap.docx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect((xml.match(/<wp:wrapSquare wrapText="bothSides"\/>/g) ?? []).length).toBe(2)

    // Reopen: the saved image has the expected canonical wrap mode.
    writeFileSync('/tmp/e2e-img-wrap-saved.docx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-wrap-saved.docx')
    await expect(page.getByText('Opened e2e-img-wrap-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).blocks as WireBlock[]
    expect(reopened.find((b) => b.docxIndex === 1)?.image?.wrap).toBe('square-left')

    expect(pageErrors).toEqual([])
  })

  test('insert a new PNG: media part + drawing + reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-insert.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-insert.docx')
    await expect(page.getByText('Opened e2e-img-insert.docx')).toBeVisible({ timeout: 30_000 })
    await expect(page.locator(IMG)).toHaveCount(7)

    // Insert a new PNG through the Image action.
    const newPng = buildSolidPng(12, 12, [90, 90, 240])
    writeFileSync('/tmp/e2e-img-new.png', newPng)
    await page.getByRole('button', { name: 'Image', exact: true }).click()
    await page.setInputFiles(
      'input[accept="image/png,image/jpeg,image/gif"]',
      '/tmp/e2e-img-new.png',
    )
    await expect(page.locator(IMG)).toHaveCount(8)
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    const download = await dl
    await expect(page.getByText('Saved e2e-img-insert.docx')).toBeVisible({ timeout: 15_000 })
    const blocksSaved = JSON.parse(req.postData() ?? '{}').blocks as WireBlock[]
    const newBlock = blocksSaved.find((b) => b.docxIndex === null && b.type === 'image')
    expect(newBlock?.newImage?.mime).toBe('image/png')
    expect(newBlock?.edited).toBe(true)

    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    // Media part exists; a drawing references it.
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(8)
    const rels = await readZipEntry(saved, 'word/_rels/document.xml.rels')
    expect(rels).toContain('aidocs1.png')
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/media/aidocs1.png')).not.toBeNull()

    // Reopen: the new image is exposed; the originals survive.
    writeFileSync('/tmp/e2e-img-insert-saved.docx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-insert-saved.docx')
    await expect(page.getByText('Opened e2e-img-insert-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).blocks as WireBlock[]
    expect(reopened.filter((b) => b.type === 'image' && b.image)).toHaveLength(8)
    expect(reopened.find((b) => b.docxIndex === 5)?.image?.rotDeg).toBe(90)
    await expect(page.locator(IMG)).toHaveCount(8)

    expect(pageErrors).toEqual([])
  })

  test('delete an image: gone after save/reopen, neighbors intact', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-delete.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-delete.docx')
    await expect(page.getByText('Opened e2e-img-delete.docx')).toBeVisible({ timeout: 30_000 })
    await expect(page.locator(IMG)).toHaveCount(7)

    // Delete the inline image (idx1).
    await page.locator(`${IMG}[data-docx-index="1"]`).click()
    await expect(page.getByTestId('image-toolbar')).toBeVisible()
    await page.getByTestId('image-delete').click()
    await expect(page.locator(IMG)).toHaveCount(6)
    await expect(page.getByText('● Unsaved')).toBeVisible()

    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const download = await dl
    await expect(page.getByText('Saved e2e-img-delete.docx')).toBeVisible({ timeout: 15_000 })
    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(6)
    expect(xml).toContain('First paragraph.')
    expect(xml).toContain('Last paragraph.')

    // Reopen: image gone, neighbors intact.
    writeFileSync('/tmp/e2e-img-delete-saved.docx', saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-delete-saved.docx')
    await expect(page.getByText('Opened e2e-img-delete-saved.docx')).toBeVisible({
      timeout: 30_000,
    })
    const reopenResponse = await reopenResponsePromise
    const reopened = (await reopenResponse.json()).blocks as WireBlock[]
    expect(reopened.filter((b) => b.type === 'image' && b.image)).toHaveLength(6)
    await expect(page.getByText('First paragraph.')).toBeVisible()
    await expect(page.getByText('Last paragraph.')).toBeVisible()
    // Rotated/flipped/cropped images keep their identity (renumbered).
    const rot = reopened.find((b) => b.image?.rotDeg === 90)
    expect(rot?.image).toBeDefined()

    expect(pageErrors).toEqual([])
  })

  test('dirty-state isolation: one edited image, others byte-preserved', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordImageFixture()
    writeFileSync('/tmp/e2e-img-dirty.docx', fixture)
    await page.setInputFiles('input[type="file"]', '/tmp/e2e-img-dirty.docx')
    await expect(page.getByText('Opened e2e-img-dirty.docx')).toBeVisible({ timeout: 30_000 })
    await expect(page.locator(IMG)).toHaveCount(7)

    // Modify ONE image: rotate the inline image (idx1).
    await page.locator(`${IMG}[data-docx-index="1"]`).click()
    await expect(page.getByTestId('image-toolbar')).toBeVisible()
    await page.getByTestId('image-rotate').click()
    await expect(page.getByText('● Unsaved')).toBeVisible()

    // Assert the ACTUAL save payload (not just UI state).
    const saveReq = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    const req = await saveReq
    await dl
    await expect(page.getByText('Saved e2e-img-dirty.docx')).toBeVisible({ timeout: 15_000 })
    const blocks = JSON.parse(req.postData() ?? '{}').blocks as WireBlock[]
    const imageBlocks = blocks.filter((b) => b.type === 'image')
    // Modified image → edited=true with the new rotation.
    const editedImg = imageBlocks.find((b) => b.docxIndex === 1)
    expect(editedImg?.edited).toBe(true)
    expect(editedImg?.image?.rotDeg).toBe(90)
    // Untouched images → edited=false (byte-preserved originals).
    for (const b of imageBlocks) {
      if (b.docxIndex !== 1) {
        expect(b.edited, `image docxIndex=${b.docxIndex} must be edited=false`).toBe(false)
      }
    }
    // Paragraphs untouched too.
    const para = blocks.find((b) => b.docxIndex === 0)
    expect(para?.edited).toBe(false)

    expect(pageErrors).toEqual([])
  })
})
