import { describe, expect, it } from 'vitest'
import { parseDocx, patchImageParagraphXml, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Image paragraph XML with a wp:docPr carrying a descr (alt text). The blip
 * references rId10 (buildDocx wires media/image1.png when withImage: true).
 */
const IMAGE_WITH_DESCR_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="1" name="Picture 1" descr="A red square"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

/** Image with name but NO descr — name is the object name, NOT alt text */
const IMAGE_WITH_NAME_ONLY_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="2" name="Logo"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

/** Image with neither descr nor name */
const IMAGE_NO_DOCPR_ALT_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="3"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

/** Image with descr carrying XML-special characters (must be escaped on patch) */
const IMAGE_WITH_SPECIAL_CHARS_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="4" name="Pic &lt;4&gt;" descr="A &lt;tag&gt; &amp; quote"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

describe('image alt text parsing (wp:docPr descr/name)', () => {
  it('surfaces descr as imageAlt', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: IMAGE_WITH_DESCR_XML, withImage: true }))
    const img = doc.blocks[0]
    expect(img.type).toBe('image')
    expect(img.imageAlt).toBe('A red square')
  })

  it('leaves imageAlt undefined when only name is present (name is not alt text)', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: IMAGE_WITH_NAME_ONLY_XML, withImage: true }),
    )
    expect(doc.blocks[0].imageAlt).toBeUndefined()
  })

  it('leaves imageAlt undefined when neither descr nor name is present', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: IMAGE_NO_DOCPR_ALT_XML, withImage: true }),
    )
    expect(doc.blocks[0].imageAlt).toBeUndefined()
  })

  it('decodes XML entities in descr', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: IMAGE_WITH_SPECIAL_CHARS_XML, withImage: true }),
    )
    expect(doc.blocks[0].imageAlt).toBe('A <tag> & quote')
  })
})

describe('patchImageParagraphXml — alt text', () => {
  const BASE_XML = IMAGE_WITH_DESCR_XML

  it('sets descr when alt is a non-empty string', () => {
    const out = patchImageParagraphXml(BASE_XML, { alt: 'Updated description' })
    expect(out).toContain('descr="Updated description"')
    // the existing descr is replaced, not duplicated
    expect(out.match(/descr=/g)?.length).toBe(1)
    // name is preserved (not the alt-text target)
    expect(out).toContain('name="Picture 1"')
  })

  it('removes descr when alt is null', () => {
    const out = patchImageParagraphXml(BASE_XML, { alt: null })
    expect(out).not.toContain('descr=')
    // name preserved
    expect(out).toContain('name="Picture 1"')
  })

  it('removes descr when alt is empty string', () => {
    const out = patchImageParagraphXml(BASE_XML, { alt: '' })
    expect(out).not.toContain('descr=')
  })

  it('keeps descr unchanged when alt is undefined', () => {
    const out = patchImageParagraphXml(BASE_XML, { widthPx: 200, heightPx: 200 })
    expect(out).toContain('descr="A red square"')
  })

  it('escapes XML-special characters in the new descr', () => {
    const out = patchImageParagraphXml(BASE_XML, { alt: 'A <tag> & "quote"' })
    expect(out).toContain('descr="A &lt;tag&gt; &amp; &quot;quote&quot;"')
  })

  it('adds descr to a docPr that had none', () => {
    const out = patchImageParagraphXml(IMAGE_NO_DOCPR_ALT_XML, { alt: 'Newly added' })
    expect(out).toContain('descr="Newly added"')
  })

  it('preserves other wp:docPr attributes (id, name) when patching descr', () => {
    const out = patchImageParagraphXml(BASE_XML, { alt: 'Changed' })
    expect(out).toContain('id="1"')
    expect(out).toContain('name="Picture 1"')
    expect(out).toContain('descr="Changed"')
  })
})

describe('image alt text round-trip through saveDocx', () => {
  it('preserves descr on an unchanged image (kind:original re-slices bytes)', async () => {
    const bytes = await buildDocx({ bodyXml: IMAGE_WITH_DESCR_XML, withImage: true })
    const doc = await parseDocx(bytes)
    // Save with NO edits → the engine re-slices the original <w:drawing> bytes
    // (kind:'original'); descr survives byte-for-byte.
    const saved = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageAlt).toBe('A red square')
  })

  it('patches descr when the image is edited', async () => {
    const bytes = await buildDocx({ bodyXml: IMAGE_WITH_DESCR_XML, withImage: true })
    const doc = await parseDocx(bytes)
    const original = doc.blocks[0]
    const patchedXml = patchImageParagraphXml(original.originalXml!, { alt: 'Patched alt' })
    const saved = await saveDocx(doc, [{ kind: 'xml', xml: patchedXml, docxIndex: 0 }])
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageAlt).toBe('Patched alt')
  })

  it('removes descr when alt is cleared', async () => {
    const bytes = await buildDocx({ bodyXml: IMAGE_WITH_DESCR_XML, withImage: true })
    const doc = await parseDocx(bytes)
    const original = doc.blocks[0]
    const patchedXml = patchImageParagraphXml(original.originalXml!, { alt: null })
    const saved = await saveDocx(doc, [{ kind: 'xml', xml: patchedXml, docxIndex: 0 }])
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].imageAlt).toBeUndefined()
  })
})
