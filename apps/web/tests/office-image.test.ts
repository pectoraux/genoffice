/**
 * Unit tests — editable image wire contract (Phase 3 Increment 8).
 *
 * Tests the office-routes image payload: serialization on open, canonical
 * patch paths on save (patchImageParagraphXml / applyImageWrap / NewImage
 * embed), byte preservation of unchanged images, dirty classification, and
 * the runtime validation error shape for malformed image payloads.
 *
 * These exercise the pure routeOffice function directly (node environment);
 * the real browser → HTTP → engine → bytes path is covered by the Playwright
 * suite (tests/e2e/word-image.spec.ts).
 */
import { describe, expect, it } from 'vitest'
import { routeOffice } from '@contractor/core/api'
import { buildWordImageFixture, buildSolidPng, readZipEntry } from './e2e/fixtures'

interface WireImage {
  imageDataUrl: string | null
  widthPx?: number
  heightPx?: number
  crop?: { l: number; t: number; r: number; b: number }
  align?: string
  wrap?: string
  offsetXEmu?: number
  offsetYEmu?: number
  posH?: string
  posV?: string
  posHRel?: string
  posVRel?: string
  rotDeg?: number
  flipH?: boolean
  flipV?: boolean
}
interface WireBlock {
  docxIndex: number | null
  type: string
  text: string
  image?: WireImage
  newImage?: {
    base64: string
    mime: string
    widthPx: number
    heightPx: number
    align?: string
    wrap?: string
    rotDeg?: number
    flipH?: boolean
    flipV?: boolean
  }
  edited?: boolean
  hidden?: boolean
}

const b64 = (b: Buffer) => b.toString('base64')

async function openDoc(bytes: Buffer): Promise<WireBlock[]> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/open',
    body: { fileName: 'fixture.docx', fileBytes: b64(bytes) },
  })
  expect(res?.status).toBe(200)
  return (res?.body as { blocks: WireBlock[] }).blocks
}

async function saveDoc(bytes: Buffer, blocks: WireBlock[]): Promise<Buffer> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/save',
    body: { fileName: 'fixture.docx', fileBytes: b64(bytes), blocks },
  })
  if (res?.status !== 200) {
    throw new Error(`save failed: ${res?.status} ${JSON.stringify(res?.body).slice(0, 300)}`)
  }
  return Buffer.from((res?.body as { fileBytes: string }).fileBytes, 'base64')
}

/** Expect a validation error (400) and return its body. */
async function expectValidation(body: unknown): Promise<{ error: string; message: string }> {
  const res = await routeOffice({
    method: 'POST',
    path: '/office/documents/save',
    body,
  })
  expect(res?.status).toBe(400)
  const err = res?.body as { error: string; message: string }
  expect(err.error).toBe('validation')
  expect(typeof err.message).toBe('string')
  return err
}

describe('Word image wire contract', () => {
  it('open serializes the typed image payload for all 7 variants', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const images = blocks.filter((b) => b.type === 'image')
    expect(images).toHaveLength(7)
    // idx1: inline, 64×64
    expect(images[0].docxIndex).toBe(1)
    expect(images[0].image?.widthPx).toBe(64)
    expect(images[0].image?.heightPx).toBe(64)
    expect(images[0].image?.wrap).toBeUndefined()
    expect(images[0].image?.imageDataUrl).toMatch(/^data:image\/png;base64,/)
    // idx2: floating square-left with offsets + position bases
    expect(images[1].image?.wrap).toBe('square-left')
    expect(images[1].image?.offsetXEmu).toBe(200000)
    expect(images[1].image?.offsetYEmu).toBe(100000)
    expect(images[1].image?.posHRel).toBe('column')
    expect(images[1].image?.posVRel).toBe('paragraph')
    // idx3: non-default size
    expect(images[2].image?.widthPx).toBe(240)
    expect(images[2].image?.heightPx).toBe(120)
    // idx4: topBottom wrap
    expect(images[3].image?.wrap).toBe('topBottom')
    // idx5: rotation
    expect(images[4].image?.rotDeg).toBe(90)
    // idx6: flips
    expect(images[5].image?.flipH).toBe(true)
    expect(images[5].image?.flipV).toBe(true)
    // idx7: crop
    expect(images[6].image?.crop).toEqual({ l: 0.1, t: 0.2, r: 0.1, b: 0.2 })
  })

  it('unchanged images are byte-preserved through the original path', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    // The fixture's original drawing bytes survive verbatim: the rotated
    // image keeps its authored rot, the cropped one its srcRect, the big
    // one its exact extent.
    expect(xml).toContain('rot="5400000"')
    expect(xml).toContain('<a:srcRect l="10000" t="20000" r="10000" b="20000"/>')
    expect(xml).toContain('cx="2286000"')
    expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>')
  })

  it('resize edit patches extent only; other images byte-preserved', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const big = edited.find((b) => b.docxIndex === 3)!
    big.edited = true
    big.image!.widthPx = 120
    big.image!.heightPx = 60
    const saved = await saveDoc(bytes, edited)
    const xml = await readZipEntry(saved, 'word/document.xml')
    // New extent (120px = 1143000 EMU, 60px = 571500 EMU).
    expect(xml).toContain('cx="1143000" cy="571500"')
    // Untouched images keep their bytes.
    expect(xml).toContain('rot="5400000"')
    expect(xml).toContain('<a:srcRect l="10000" t="20000" r="10000" b="20000"/>')
    // Reopen: dimensions changed.
    const reopened = await openDoc(saved)
    const r3 = reopened.find((b) => b.docxIndex === 3)
    expect(r3?.image?.widthPx).toBe(120)
    expect(r3?.image?.heightPx).toBe(60)
  })

  it('rotate edit rewrites the canonical rot attribute', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.rotDeg = 90
    const saved = await saveDoc(bytes, edited)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('rot="5400000"')
    const reopened = await openDoc(saved)
    expect(reopened.find((b) => b.docxIndex === 1)?.image?.rotDeg).toBe(90)
  })

  it('flip edits round-trip', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.flipH = true
    img.image!.flipV = true
    const saved = await saveDoc(bytes, edited)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect(/flipH="1"/.test(xml)).toBe(true)
    expect(/flipV="1"/.test(xml)).toBe(true)
    const reopened = await openDoc(saved)
    const r = reopened.find((b) => b.docxIndex === 1)?.image
    expect(r?.flipH).toBe(true)
    expect(r?.flipV).toBe(true)
  })

  it('wrap change inline → square-left applies the canonical anchor', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.wrap = 'square-left'
    const saved = await saveDoc(bytes, edited)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(/<wp:anchor[\s>]/.test(xml)).toBe(true)
    const reopened = await openDoc(saved)
    expect(reopened.find((b) => b.docxIndex === 1)?.image?.wrap).toBe('square-left')
  })

  it('crop edits add/rewrite/remove the a:srcRect', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    // 1. add a crop to an uncropped image
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.crop = { l: 0.25, t: 0, r: 0.25, b: 0 }
    let saved = await saveDoc(bytes, edited)
    let xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('<a:srcRect l="25000" t="0" r="25000" b="0"/>')
    const reopened = await openDoc(saved)
    expect(reopened.find((b) => b.docxIndex === 1)?.image?.crop).toEqual({
      l: 0.25,
      t: 0,
      r: 0.25,
      b: 0,
    })
    // 2. rewrite an existing crop
    const edited2 = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img7 = edited2.find((b) => b.docxIndex === 7)!
    img7.edited = true
    img7.image!.crop = { l: 0.05, t: 0.05, r: 0.05, b: 0.05 }
    saved = await saveDoc(bytes, edited2)
    xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('<a:srcRect l="5000" t="5000" r="5000" b="5000"/>')
    // 3. remove a crop (all-zero)
    const edited3 = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img7b = edited3.find((b) => b.docxIndex === 7)!
    img7b.edited = true
    img7b.image!.crop = { l: 0, t: 0, r: 0, b: 0 }
    saved = await saveDoc(bytes, edited3)
    xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).not.toContain('<a:srcRect')
  })

  it('new image embeds media + relationship + drawing', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const png = buildSolidPng(12, 12, [90, 90, 240])
    const withNew = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    withNew.push({
      docxIndex: null,
      type: 'image',
      text: '',
      edited: true,
      newImage: { base64: png.toString('base64'), mime: 'image/png', widthPx: 48, heightPx: 48 },
    })
    const saved = await saveDoc(bytes, withNew)
    const xml = await readZipEntry(saved, 'word/document.xml')
    // 8 drawings now; the new one is inline with the given extent.
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(8)
    expect(xml).toContain('cx="457200" cy="457200"')
    const rels = await readZipEntry(saved, 'word/_rels/document.xml.rels')
    expect(rels).toContain('aidocs1.png')
    const reopened = await openDoc(saved)
    expect(reopened.filter((b) => b.type === 'image' && b.image)).toHaveLength(8)
    // Original images survive.
    expect(reopened.find((b) => b.docxIndex === 5)?.image?.rotDeg).toBe(90)
  })

  it('deleting an image omits it from the save blocks', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const reduced = blocks.filter((b) => b.docxIndex !== 1)
    const saved = await saveDoc(bytes, reduced)
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(6)
    // Neighboring content survives.
    expect(xml).toContain('First paragraph.')
    expect(xml).toContain('Last paragraph.')
    const reopened = await openDoc(saved)
    // docxIndex renumbers on reparse: the old floating image (idx2) is now
    // idx1, and only 6 images remain.
    expect(reopened.filter((b) => b.type === 'image')).toHaveLength(6)
    expect(reopened.find((b) => b.docxIndex === 1)?.image?.wrap).toBe('square-left')
    expect(reopened.find((b) => b.docxIndex === 2)?.image?.widthPx).toBe(240)
  })
})

describe('Word image wire validation (malformed payloads → 400)', () => {
  it('rejects an invalid wrap value', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    ;(img.image as unknown as Record<string, unknown>).wrap = 'diagonal'
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('wrap')
  })

  it('rejects malformed crop values', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.crop = { l: 1.5, t: 0, r: 0, b: 0 }
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('crop')
  })

  it('rejects invalid dimensions', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.widthPx = 0
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('widthPx')
  })

  it('rejects invalid rotation', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.rotDeg = 450
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('rotDeg')
  })

  it('rejects an invalid image MIME on insertion', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const withNew = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    withNew.push({
      docxIndex: null,
      type: 'image',
      text: '',
      edited: true,
      newImage: {
        base64: buildSolidPng(12, 12, [1, 2, 3]).toString('base64'),
        mime: 'image/webp',
        widthPx: 48,
        heightPx: 48,
      },
    })
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: withNew,
    })
    expect(err.message).toContain('mime')
  })

  it('rejects invalid base64 on insertion', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const withNew = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    withNew.push({
      docxIndex: null,
      type: 'image',
      text: '',
      edited: true,
      newImage: {
        base64: 'not!valid!base64!!with<script>',
        mime: 'image/png',
        widthPx: 48,
        heightPx: 48,
      },
    })
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: withNew,
    })
    expect(err.message).toContain('base64')
  })

  it('rejects an imageDataUrl with a foreign scheme', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image!.imageDataUrl = 'javascript:alert(1)'
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('imageDataUrl')
  })

  it('rejects an image payload on a non-image block', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    edited[0].image = edited[1].image
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('image')
  })

  it('rejects an edited image without a payload', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    delete img.image
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('image payload')
  })

  it('rejects newImage on an original (docxIndex != null) block', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.newImage = {
      base64: buildSolidPng(12, 12, [1, 2, 3]).toString('base64'),
      mime: 'image/png',
      widthPx: 48,
      heightPx: 48,
    }
    await expectValidation({ fileName: 'fixture.docx', fileBytes: b64(bytes), blocks: edited })
  })
})
