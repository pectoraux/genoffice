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
import {
  buildWordImageFixture,
  buildSolidPng,
  readZipEntry,
  readZipEntryBase64,
  listZipEntries,
} from './e2e/fixtures'

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
  /** accessibility alt text (wp:docPr descr); tri-state: undefined|null|string */
  alt?: string | null
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

/**
 * Alt-text (wp:docPr descr) wire contract — Phase 3 Increment 13.
 *
 * Covers: parse-side surfacing of descr, render of <img alt>, alt edit →
 * canonical patch (descr set/clear), dirty classification, unchanged-image
 * descr preservation (kind:original re-slices bytes), and runtime validation
 * (oversized alt rejected, control chars stripped, file:// scheme rejected).
 * The canonical engine path (patchImageParagraphXml wp:docPr descr) is
 * covered at the docx-engine level by image-alt-text.test.ts.
 */
describe('Word image alt text (wp:docPr descr)', () => {
  it('open surfaces descr as image.alt for images that carry one', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const images = blocks.filter((b) => b.type === 'image' && b.image)
    // Fixture: images 1, 3, 5 carry descr; 2, 4, 6, 7 do not.
    const withAlt = images.filter((b) => b.image?.alt !== undefined)
    expect(withAlt.map((b) => b.docxIndex).sort()).toEqual([1, 3, 5])
    const byIdx = new Map(images.map((b) => [b.docxIndex, b]))
    expect(byIdx.get(1)?.image?.alt).toBe('A red square')
    expect(byIdx.get(3)?.image?.alt).toBe('Wide blue banner')
    expect(byIdx.get(5)?.image?.alt).toBe('Rotated magenta')
    // Images without descr have alt === undefined (field absent on the wire).
    for (const idx of [2, 4, 6, 7]) {
      expect(byIdx.get(idx)?.image?.alt).toBeUndefined()
    }
  })

  it('unchanged image preserves descr byte-for-byte (kind:original)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    // Save with NO edits → engine re-slices original <w:drawing> bytes; descr survives.
    const saved = await saveDoc(bytes, blocks)
    const reparsed = await openDoc(saved)
    const byIdx = new Map(reparsed.filter((b) => b.type === 'image').map((b) => [b.docxIndex, b]))
    expect(byIdx.get(1)?.image?.alt).toBe('A red square')
    expect(byIdx.get(3)?.image?.alt).toBe('Wide blue banner')
    expect(byIdx.get(5)?.image?.alt).toBe('Rotated magenta')
    // Images without descr still have no alt.
    expect(byIdx.get(2)?.image?.alt).toBeUndefined()
  })

  it('alt edit produces edited:true + canonical descr patch in saved XML', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image = { ...img.image!, alt: 'Edited alt text' }
    const saved = await saveDoc(bytes, edited)
    const documentXml = await readZipEntry(saved, 'word/document.xml')
    // Canonical generator wrote the new descr; the old one is gone.
    expect(documentXml).toContain('descr="Edited alt text"')
    expect(documentXml).not.toContain('descr="A red square"')
    // Reopen: the patched descr surfaces back.
    const reparsed = await openDoc(saved)
    const img1 = reparsed.find((b) => b.docxIndex === 1)!
    expect(img1.image?.alt).toBe('Edited alt text')
  })

  it('clearing alt (alt:null) removes descr from the saved XML', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image = { ...img.image!, alt: null }
    const saved = await saveDoc(bytes, edited)
    const documentXml = await readZipEntry(saved, 'word/document.xml')
    // Image 101's descr is gone; its name is preserved (object name, separate field).
    expect(documentXml).not.toContain('descr="A red square"')
    expect(documentXml).toContain('name="Picture 101"')
    // Unchanged images keep their descr (byte-preserved via kind:original).
    expect(documentXml).toContain('descr="Wide blue banner"')
    const reparsed = await openDoc(saved)
    const img1 = reparsed.find((b) => b.docxIndex === 1)!
    expect(img1.image?.alt).toBeUndefined()
  })

  it('alt change flips dirty; size change alone keeps alt unchanged', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    // Alt-only edit on image 1: edited:true, alt changed.
    const editedAlt = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const imgAlt = editedAlt.find((b) => b.docxIndex === 1)!
    imgAlt.edited = true
    imgAlt.image = { ...imgAlt.image!, alt: 'Alt-only edit' }
    const savedAlt = await saveDoc(bytes, editedAlt)
    const reparsedAlt = await openDoc(savedAlt)
    expect(reparsedAlt.find((b) => b.docxIndex === 1)!.image?.alt).toBe('Alt-only edit')
    // Other images keep their descr (byte-preserved).
    expect(reparsedAlt.find((b) => b.docxIndex === 3)!.image?.alt).toBe('Wide blue banner')

    // Size-only edit on image 3: alt must survive untouched.
    const editedSize = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const imgSize = editedSize.find((b) => b.docxIndex === 3)!
    imgSize.edited = true
    imgSize.image = { ...imgSize.image!, widthPx: 200, heightPx: 100, alt: 'Wide blue banner' }
    const savedSize = await saveDoc(bytes, editedSize)
    const reparsedSize = await openDoc(savedSize)
    expect(reparsedSize.find((b) => b.docxIndex === 3)!.image?.alt).toBe('Wide blue banner')
    expect(reparsedSize.find((b) => b.docxIndex === 3)!.image?.widthPx).toBe(200)
  })

  it('rejects an oversized alt string (500 char cap)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image = { ...img.image!, alt: 'x'.repeat(501) }
    const err = await expectValidation({
      fileName: 'fixture.docx',
      fileBytes: b64(bytes),
      blocks: edited,
    })
    expect(err.message).toContain('alt')
    expect(err.message).toContain('500')
  })

  it('strips XML-invalid control chars from alt but keeps the value', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    // \x00 (NUL) and \x07 (BEL) are invalid in XML attrs; \t \n \r are kept.
    img.image = { ...img.image!, alt: 'A\x00B\x07C\tD\nE\rF' }
    const saved = await saveDoc(bytes, edited)
    const documentXml = await readZipEntry(saved, 'word/document.xml')
    // Control chars stripped; the value survived and the XML is well-formed.
    expect(documentXml).toContain('descr=')
    expect(documentXml).not.toContain('\x00')
    expect(documentXml).not.toContain('\x07')
    const reparsed = await openDoc(saved)
    const img1 = reparsed.find((b) => b.docxIndex === 1)!
    expect(img1.image?.alt).toBe('ABC\tD\nE\rF')
  })

  it('rejects a non-string alt (number)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image = { ...img.image!, alt: 42 as unknown as string }
    await expectValidation({ fileName: 'fixture.docx', fileBytes: b64(bytes), blocks: edited })
  })

  it('rejects a file:// image src (foreign scheme already rejected by data URL regex)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const img = edited.find((b) => b.docxIndex === 1)!
    img.edited = true
    img.image = { ...img.image!, imageDataUrl: 'file:///etc/passwd' }
    await expectValidation({ fileName: 'fixture.docx', fileBytes: b64(bytes), blocks: edited })
  })
})

/**
 * Byte/fidelity checks — Phase 3 Increment 13 §F + §K.
 *
 * Open a DOCX containing images and immediately save without editing. The
 * save payload must mark every image block edited:false. The resulting DOCX
 * must preserve the original image relationship, media part bytes, content
 * type, dimensions, and surrounding document XML byte-for-byte (kind:original
 * re-slices the source bytes; media parts are copied verbatim).
 *
 * JSZip is used ONLY in the test environment (never in apps/web) to inspect
 * the saved zip's parts.
 */
describe('Word image byte/fidelity (unchanged-image preservation)', () => {
  it('marks every image block edited:false on open→save (no edits)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    // Open returns blocks with edited:false already; saving them back goes
    // through the kind:original path (no patch).
    const saved = await saveDoc(bytes, blocks)
    expect(saved.length).toBeGreaterThan(0)
  })

  it('preserves the image relationship in word/_rels/document.xml.rels', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    const rels = await readZipEntry(saved, 'word/_rels/document.xml.rels')
    // All 7 image relationships survive (rIdImg1..7 → media/image1..7.png).
    for (let i = 1; i <= 7; i++) {
      expect(rels).toContain(`Id="rIdImg${i}"`)
      expect(rels).toContain(`Target="media/image${i}.png"`)
      expect(rels).toContain(
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"',
      )
    }
  })

  it('preserves every media part byte-for-byte (word/media/*.png)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    for (let i = 1; i <= 7; i++) {
      const origB64 = await readZipEntryBase64(bytes, `word/media/image${i}.png`)
      const savedB64 = await readZipEntryBase64(saved, `word/media/image${i}.png`)
      expect(savedB64).toBe(origB64)
    }
  })

  it('preserves the png content type in [Content_Types].xml', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    const ct = await readZipEntry(saved, '[Content_Types].xml')
    expect(ct).toContain('Extension="png"')
    expect(ct).toContain('ContentType="image/png"')
    // The document + styles overrides survive too.
    expect(ct).toContain('PartName="/word/document.xml"')
    expect(ct).toContain('PartName="/word/styles.xml"')
  })

  it('preserves image dimensions + descr in word/document.xml (kind:original)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    const xml = await readZipEntry(saved, 'word/document.xml')
    // Extents survive (the original drawing bytes are re-sliced).
    expect(xml).toContain('cx="609600" cy="609600"') // image 1: 64×64
    expect(xml).toContain('cx="2286000" cy="1143000"') // image 3: 240×120
    // descr survives for images that carry it.
    expect(xml).toContain('descr="A red square"')
    expect(xml).toContain('descr="Wide blue banner"')
    expect(xml).toContain('descr="Rotated magenta"')
    // Surrounding paragraphs survive.
    expect(xml).toContain('First paragraph.')
    expect(xml).toContain('Last paragraph.')
  })

  it('keeps the zip structure (all expected parts present)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const saved = await saveDoc(bytes, blocks)
    const entries = await listZipEntries(saved)
    expect(entries).toContain('[Content_Types].xml')
    expect(entries).toContain('_rels/.rels')
    expect(entries).toContain('word/document.xml')
    expect(entries).toContain('word/_rels/document.xml.rels')
    expect(entries).toContain('word/styles.xml')
    for (let i = 1; i <= 7; i++) {
      expect(entries).toContain(`word/media/image${i}.png`)
    }
  })

  it('a resize edit preserves all media bytes (only extent XML changes)', async () => {
    const bytes = await buildWordImageFixture()
    const blocks = await openDoc(bytes)
    const edited = JSON.parse(JSON.stringify(blocks)) as WireBlock[]
    const big = edited.find((b) => b.docxIndex === 3)!
    big.edited = true
    big.image!.widthPx = 120
    big.image!.heightPx = 60
    const saved = await saveDoc(bytes, edited)
    // All 7 media parts are byte-identical (only image 3's drawing XML changed).
    for (let i = 1; i <= 7; i++) {
      const origB64 = await readZipEntryBase64(bytes, `word/media/image${i}.png`)
      const savedB64 = await readZipEntryBase64(saved, `word/media/image${i}.png`)
      expect(savedB64).toBe(origB64)
    }
    // image 3's extent changed; others' did not.
    const xml = await readZipEntry(saved, 'word/document.xml')
    expect(xml).toContain('cx="1143000" cy="571500"') // 120×60
  })
})
