/**
 * Image conversion between the typed wire model (SerializedImage) and the
 * browser editor's DOM representation (<img data-docx-image> with typed
 * data-* attributes).
 *
 * PURITY: this module performs NO OOXML generation whatsoever. All drawing
 * XML (wp:anchor, wp:inline, pic:pic, a:xfrm, relationships, media parts)
 * is produced exclusively by the canonical docx-engine
 * (patchImageParagraphXml / applyImageWrap / the NewImage embed path) on
 * the server. The browser only maps typed fields ↔ DOM attributes.
 */

import type {
  SerializedImage,
  SerializedImageRect,
  SerializedImageWrap,
} from '../api/office-client'

/** Escape HTML special characters + quotes in attribute values. */
function escapeAttr(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/** Serialize a rect to a data attribute value ("l,t,r,b"). */
function rectToAttr(rect: SerializedImageRect | undefined): string | undefined {
  if (!rect) return undefined
  return `${rect.l},${rect.t},${rect.r},${rect.b}`
}

/** Parse a "l,t,r,b" data attribute back into a rect. */
export function rectFromAttr(
  value: string | null,
): { l: number; t: number; r: number; b: number } | undefined {
  if (value === null) return undefined
  const parts = value.split(',').map((p) => Number(p))
  if (parts.length !== 4 || parts.some((p) => !Number.isFinite(p))) return undefined
  return { l: parts[0], t: parts[1], r: parts[2], b: parts[3] }
}

/**
 * SerializedImage → <img> HTML for the Tiptap editor. docxIndex rides on
 * the data-docx-index attribute (schema-backed on the image node); every
 * editable property is a typed data-* attribute.
 */
export function imageToHtml(image: SerializedImage, docxIndex: number | null): string {
  const attrs: string[] = ['data-docx-image="true"']
  if (docxIndex !== null) attrs.push(`data-docx-index="${docxIndex}"`)
  if (image.imageDataUrl) attrs.push(`src="${escapeAttr(image.imageDataUrl)}"`)
  else attrs.push('src=""')
  if (image.widthPx !== undefined) attrs.push(`data-width="${image.widthPx}"`)
  if (image.heightPx !== undefined) attrs.push(`data-height="${image.heightPx}"`)
  const crop = rectToAttr(image.crop)
  if (crop) attrs.push(`data-crop="${crop}"`)
  const fillRect = rectToAttr(image.fillRect)
  if (fillRect) attrs.push(`data-fill-rect="${fillRect}"`)
  if (image.align) attrs.push(`data-align="${image.align}"`)
  const wrap = image.wrap ?? 'inline'
  attrs.push(`data-wrap="${wrap}"`)
  if (image.offsetXEmu !== undefined) attrs.push(`data-offset-x="${image.offsetXEmu}"`)
  if (image.offsetYEmu !== undefined) attrs.push(`data-offset-y="${image.offsetYEmu}"`)
  if (image.posH) attrs.push(`data-pos-h="${image.posH}"`)
  if (image.posV) attrs.push(`data-pos-v="${image.posV}"`)
  if (image.posHRel) attrs.push(`data-pos-h-rel="${image.posHRel}"`)
  if (image.posVRel) attrs.push(`data-pos-v-rel="${image.posVRel}"`)
  if (image.rotDeg !== undefined) attrs.push(`data-rot="${image.rotDeg}"`)
  if (image.flipH) attrs.push('data-flip-h="1"')
  if (image.flipV) attrs.push('data-flip-v="1"')
  const w = image.widthPx ?? 0
  const h = image.heightPx ?? 0
  attrs.push(`width="${w}" height="${h}"`)
  return `<img ${attrs.join(' ')} alt="Document image" draggable="false" />`
}

export interface ImageNodeAttrs {
  docxIndex: number | null
  src: string
  widthPx: number | null
  heightPx: number | null
  crop: { l: number; t: number; r: number; b: number } | null
  fillRect: { l: number; t: number; r: number; b: number } | null
  align: 'left' | 'center' | 'right' | null
  wrap: SerializedImageWrap
  offsetXEmu: number | null
  offsetYEmu: number | null
  posH: 'left' | 'center' | 'right' | null
  posV: 'top' | 'center' | 'bottom' | null
  posHRel: string | null
  posVRel: string | null
  rotDeg: number
  flipH: boolean
  flipV: boolean
}

/** Read the typed image attributes off a rendered <img> DOM element. */
export function imageAttrsFromElement(el: HTMLImageElement): ImageNodeAttrs {
  const num = (name: string): number | null => {
    const raw = el.getAttribute(name)
    if (raw === null) return null
    const v = Number(raw)
    return Number.isFinite(v) ? v : null
  }
  const wrapRaw = el.getAttribute('data-wrap')
  const wrap = (wrapRaw && isWrap(wrapRaw) ? wrapRaw : 'inline') as SerializedImageWrap
  return {
    docxIndex: (() => {
      const raw = el.getAttribute('data-docx-index')
      if (raw === null) return null
      const v = parseInt(raw, 10)
      return Number.isInteger(v) ? v : null
    })(),
    src: el.getAttribute('src') ?? '',
    widthPx: num('data-width'),
    heightPx: num('data-height'),
    crop: rectFromAttr(el.getAttribute('data-crop')) ?? null,
    fillRect: rectFromAttr(el.getAttribute('data-fill-rect')) ?? null,
    align: (el.getAttribute('data-align') as 'left' | 'center' | 'right' | null) ?? null,
    wrap,
    offsetXEmu: num('data-offset-x'),
    offsetYEmu: num('data-offset-y'),
    posH: (el.getAttribute('data-pos-h') as 'left' | 'center' | 'right' | null) ?? null,
    posV: (el.getAttribute('data-pos-v') as 'top' | 'center' | 'bottom' | null) ?? null,
    posHRel: el.getAttribute('data-pos-h-rel'),
    posVRel: el.getAttribute('data-pos-v-rel'),
    rotDeg: num('data-rot') ?? 0,
    flipH: el.getAttribute('data-flip-h') === '1',
    flipV: el.getAttribute('data-flip-v') === '1',
  }
}

const WRAPS: readonly string[] = [
  'inline',
  'square-left',
  'square-right',
  'tight-left',
  'tight-right',
  'through-left',
  'through-right',
  'topBottom',
  'behind',
  'front',
]

function isWrap(v: string): v is SerializedImageWrap {
  return WRAPS.includes(v)
}

/** ImageNodeAttrs → wire SerializedImage (the edited state for the save path). */
export function imageAttrsToWire(a: ImageNodeAttrs): SerializedImage {
  return {
    imageDataUrl: a.src !== '' ? a.src : null,
    ...(a.widthPx !== null ? { widthPx: Math.round(a.widthPx) } : {}),
    ...(a.heightPx !== null ? { heightPx: Math.round(a.heightPx) } : {}),
    ...(a.crop ? { crop: a.crop } : {}),
    ...(a.fillRect ? { fillRect: a.fillRect } : {}),
    ...(a.align ? { align: a.align } : {}),
    ...(a.wrap ? { wrap: a.wrap } : {}),
    ...(a.offsetXEmu !== null ? { offsetXEmu: Math.round(a.offsetXEmu) } : {}),
    ...(a.offsetYEmu !== null ? { offsetYEmu: Math.round(a.offsetYEmu) } : {}),
    ...(a.posH ? { posH: a.posH } : {}),
    ...(a.posV ? { posV: a.posV } : {}),
    ...(a.posHRel ? { posHRel: a.posHRel as SerializedImage['posHRel'] } : {}),
    ...(a.posVRel ? { posVRel: a.posVRel as SerializedImage['posVRel'] } : {}),
    ...(a.rotDeg ? { rotDeg: Math.round(a.rotDeg) % 360 } : {}),
    ...(a.flipH ? { flipH: true } : {}),
    ...(a.flipV ? { flipV: true } : {}),
  }
}

/**
 * Extract the new-image embedding spec from node attrs (editor-inserted
 * images): the data URL supplies base64 + mime, attrs supply the rest.
 * Returns null when the src is not a data URL of an allowed type.
 */
export function newImageFromAttrs(a: ImageNodeAttrs): {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  widthPx: number
  heightPx: number
  align?: 'left' | 'center' | 'right'
  wrap?: SerializedImageWrap
  rotDeg?: number
  flipH?: boolean
  flipV?: boolean
} | null {
  const m = /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/.exec(a.src)
  if (!m || !a.widthPx || !a.heightPx) return null
  return {
    base64: m[2],
    mime: m[1] as 'image/png' | 'image/jpeg' | 'image/gif',
    widthPx: Math.round(a.widthPx),
    heightPx: Math.round(a.heightPx),
    ...(a.align ? { align: a.align } : {}),
    ...(a.wrap && a.wrap !== 'inline' ? { wrap: a.wrap } : {}),
    ...(a.rotDeg ? { rotDeg: Math.round(a.rotDeg) % 360 } : {}),
    ...(a.flipH ? { flipH: true } : {}),
    ...(a.flipV ? { flipV: true } : {}),
  }
}

/**
 * Fingerprint of an image's browser-EDITABLE state. Computed identically at
 * load time (from the wire image) and at save time (from the node attrs),
 * covering every browser-editable image property. Bytes (imageDataUrl),
 * echo fields (posHRel/posVRel/fillRect) and source identity are excluded —
 * they do not participate in the edited classification.
 */
export function imageFingerprint(image: SerializedImage): string {
  return JSON.stringify({
    w: image.widthPx ?? null,
    h: image.heightPx ?? null,
    a: image.align ?? null,
    wr: image.wrap ?? 'inline',
    ox: image.offsetXEmu ?? null,
    oy: image.offsetYEmu ?? null,
    ph: image.posH ?? null,
    pv: image.posV ?? null,
    r: image.rotDeg ?? 0,
    fh: image.flipH === true,
    fv: image.flipV === true,
    c: image.crop ? [image.crop.l, image.crop.t, image.crop.r, image.crop.b] : [0, 0, 0, 0],
  })
}

/** Fingerprint over node attrs (same fields as imageFingerprint). */
export function imageAttrsFingerprint(a: ImageNodeAttrs): string {
  return JSON.stringify({
    w: a.widthPx,
    h: a.heightPx,
    a: a.align,
    wr: a.wrap,
    ox: a.offsetXEmu,
    oy: a.offsetYEmu,
    ph: a.posH,
    pv: a.posV,
    r: a.rotDeg,
    fh: a.flipH,
    fv: a.flipV,
    c: a.crop ? [a.crop.l, a.crop.t, a.crop.r, a.crop.b] : [0, 0, 0, 0],
  })
}
