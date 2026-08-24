/**
 * Tiptap image node with schema-backed DOCX source identity.
 *
 * An atomic block node (not directly pixel-editable inside Tiptap — the
 * editing UI updates typed attributes). All canonical image properties are
 * schema-backed attributes rendered as data-* on the <img>, so they survive
 * setContent() → attribute edits → getHTML() round-trips.
 *
 * The node renders actual pixels from the data-URL src; floating wrap modes
 * approximate Word layout (float/behind/front) while the canonical OOXML
 * semantics are preserved on save by the engine.
 */

import { Node, mergeAttributes } from '@tiptap/core'
import type { SerializedImageWrap } from '../api/office-client'

const IMAGE_WRAPS = [
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
] as const

/** Shared docxIndex attribute (mirrors the paragraph/table extensions). */
const docxIndexAttribute = {
  docxIndex: {
    default: null,
    parseHTML: (element: HTMLElement) => {
      const raw = element.getAttribute('data-docx-index')
      return raw !== null ? parseInt(raw, 10) : null
    },
    renderHTML: (attributes: { docxIndex: number | null }) => {
      if (attributes.docxIndex === null || attributes.docxIndex === undefined) {
        return {}
      }
      return { 'data-docx-index': String(attributes.docxIndex) }
    },
  },
}

/** CSS approximating Word layout per wrap mode (preview only). */
function wrapStyle(wrap: string | null | undefined): Record<string, string> {
  switch (wrap) {
    case 'square-left':
    case 'tight-left':
    case 'through-left':
      return { float: 'left', margin: '4px 12px 4px 0' }
    case 'square-right':
    case 'tight-right':
    case 'through-right':
      return { float: 'right', margin: '4px 0 4px 12px' }
    case 'topBottom':
      return { display: 'block', margin: '8px auto', clear: 'both' }
    case 'behind':
      return { position: 'relative', zIndex: '-1', display: 'block', margin: '8px auto' }
    case 'front':
      return { position: 'relative', zIndex: '1', display: 'block', margin: '8px auto' }
    default:
      return { display: 'inline-block', verticalAlign: 'bottom' }
  }
}

/** Transform for rotation + flips (preview approximation). */
function transformStyle(rot: number | null | undefined, flipH: boolean, flipV: boolean): string {
  const parts: string[] = []
  if (rot) parts.push(`rotate(${rot}deg)`)
  if (flipH) parts.push('scaleX(-1)')
  if (flipV) parts.push('scaleY(-1)')
  return parts.join(' ')
}

export interface DocxImageOptions {
  HTMLAttributes: Record<string, unknown>
}

export const DocxImage = Node.create<DocxImageOptions>({
  name: 'docxImage',

  group: 'block',

  atom: true,

  draggable: true,

  addOptions() {
    return {
      HTMLAttributes: {},
    }
  },

  addAttributes() {
    return {
      ...docxIndexAttribute,
      src: {
        default: '',
        parseHTML: (element: HTMLElement) => element.getAttribute('src') ?? '',
        renderHTML: (attributes: { src: string }) => ({ src: attributes.src || '' }),
      },
      widthPx: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-width')
          if (raw === null) return null
          const v = Number(raw)
          return Number.isFinite(v) ? v : null
        },
        renderHTML: (attributes: { widthPx: number | null }) => {
          if (attributes.widthPx === null || attributes.widthPx === undefined) return {}
          return { 'data-width': String(attributes.widthPx), width: String(attributes.widthPx) }
        },
      },
      heightPx: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-height')
          if (raw === null) return null
          const v = Number(raw)
          return Number.isFinite(v) ? v : null
        },
        renderHTML: (attributes: { heightPx: number | null }) => {
          if (attributes.heightPx === null || attributes.heightPx === undefined) return {}
          return { 'data-height': String(attributes.heightPx), height: String(attributes.heightPx) }
        },
      },
      crop: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-crop')
          if (raw === null) return null
          const parts = raw.split(',').map(Number)
          if (parts.length !== 4 || parts.some((p) => !Number.isFinite(p))) return null
          return { l: parts[0], t: parts[1], r: parts[2], b: parts[3] }
        },
        renderHTML: (attributes: {
          crop: { l: number; t: number; r: number; b: number } | null
        }) => {
          if (!attributes.crop) return {}
          const c = attributes.crop
          return { 'data-crop': `${c.l},${c.t},${c.r},${c.b}` }
        },
      },
      fillRect: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-fill-rect')
          if (raw === null) return null
          const parts = raw.split(',').map(Number)
          if (parts.length !== 4 || parts.some((p) => !Number.isFinite(p))) return null
          return { l: parts[0], t: parts[1], r: parts[2], b: parts[3] }
        },
        renderHTML: (attributes: {
          fillRect: { l: number; t: number; r: number; b: number } | null
        }) => {
          if (!attributes.fillRect) return {}
          const c = attributes.fillRect
          return { 'data-fill-rect': `${c.l},${c.t},${c.r},${c.b}` }
        },
      },
      align: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          element.getAttribute('data-align') as 'left' | 'center' | 'right' | null,
        renderHTML: (attributes: { align: string | null }) => {
          if (!attributes.align) return {}
          return { 'data-align': attributes.align }
        },
      },
      wrap: {
        default: 'inline' as SerializedImageWrap,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-wrap')
          return (
            raw && (IMAGE_WRAPS as readonly string[]).includes(raw) ? raw : 'inline'
          ) as SerializedImageWrap
        },
        renderHTML: (attributes: { wrap: string }) => ({ 'data-wrap': attributes.wrap }),
      },
      offsetXEmu: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-offset-x')
          if (raw === null) return null
          const v = Number(raw)
          return Number.isFinite(v) ? v : null
        },
        renderHTML: (attributes: { offsetXEmu: number | null }) => {
          if (attributes.offsetXEmu === null || attributes.offsetXEmu === undefined) return {}
          return { 'data-offset-x': String(attributes.offsetXEmu) }
        },
      },
      offsetYEmu: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-offset-y')
          if (raw === null) return null
          const v = Number(raw)
          return Number.isFinite(v) ? v : null
        },
        renderHTML: (attributes: { offsetYEmu: number | null }) => {
          if (attributes.offsetYEmu === null || attributes.offsetYEmu === undefined) return {}
          return { 'data-offset-y': String(attributes.offsetYEmu) }
        },
      },
      posH: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          element.getAttribute('data-pos-h') as 'left' | 'center' | 'right' | null,
        renderHTML: (attributes: { posH: string | null }) => {
          if (!attributes.posH) return {}
          return { 'data-pos-h': attributes.posH }
        },
      },
      posV: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          element.getAttribute('data-pos-v') as 'top' | 'center' | 'bottom' | null,
        renderHTML: (attributes: { posV: string | null }) => {
          if (!attributes.posV) return {}
          return { 'data-pos-v': attributes.posV }
        },
      },
      posHRel: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-pos-h-rel'),
        renderHTML: (attributes: { posHRel: string | null }) => {
          if (!attributes.posHRel) return {}
          return { 'data-pos-h-rel': attributes.posHRel }
        },
      },
      posVRel: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-pos-v-rel'),
        renderHTML: (attributes: { posVRel: string | null }) => {
          if (!attributes.posVRel) return {}
          return { 'data-pos-v-rel': attributes.posVRel }
        },
      },
      rotDeg: {
        default: 0,
        parseHTML: (element: HTMLElement) => {
          const raw = element.getAttribute('data-rot')
          if (raw === null) return 0
          const v = Number(raw)
          return Number.isFinite(v) ? v : 0
        },
        renderHTML: (attributes: { rotDeg: number }) => ({
          'data-rot': String(attributes.rotDeg ?? 0),
        }),
      },
      flipH: {
        default: false,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-flip-h') === '1',
        renderHTML: (attributes: { flipH: boolean }) => {
          return attributes.flipH ? { 'data-flip-h': '1' } : {}
        },
      },
      flipV: {
        default: false,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-flip-v') === '1',
        renderHTML: (attributes: { flipV: boolean }) => {
          return attributes.flipV ? { 'data-flip-v': '1' } : {}
        },
      },
      alt: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('alt'),
        renderHTML: (attributes: { alt: string | null }) => {
          // null (absent) falls back to "Document image" in imageToHtml; the
          // Tiptap node mirrors the <img alt> attribute so screen readers
          // announce the picture's accessibility text.
          if (attributes.alt === null || attributes.alt === undefined) return {}
          return { alt: attributes.alt }
        },
      },
    }
  },

  parseHTML() {
    return [{ tag: 'img[data-docx-image]' }]
  },

  renderHTML({ node, HTMLAttributes }) {
    const wrap = String(node.attrs['wrap'] ?? 'inline')
    const rot = Number(node.attrs['rotDeg'] ?? 0)
    const flipH = node.attrs['flipH'] === true
    const flipV = node.attrs['flipV'] === true
    // Crop preview: object-fit window over the source picture
    // (approximation; the saved a:srcRect is canonical).
    const crop = node.attrs['crop'] as { l: number; t: number; r: number; b: number } | null
    const align = node.attrs['align'] as string | null
    const transform = transformStyle(rot, flipH, flipV)
    const style = [
      ...Object.entries(wrapStyle(wrap)).map(([k, v]) => `${k}: ${v}`),
      ...(align ? [`text-align: ${align}`] : []),
      ...(crop ? ['object-fit: cover'] : []),
      ...(crop
        ? [
            `object-position: ${((crop.l / Math.max(1 - crop.l - crop.r, 0.001)) * 100).toFixed(1)}% ${((crop.t / Math.max(1 - crop.t - crop.b, 0.001)) * 100).toFixed(1)}%`,
          ]
        : []),
      ...(transform ? [`transform: ${transform}`] : []),
      'max-width: 100%',
    ].join('; ')
    const attrs = mergeAttributes(this.options.HTMLAttributes, HTMLAttributes, {
      'data-docx-image': 'true',
      style,
    })
    return ['img', attrs]
  },

  addCommands() {
    return {
      insertDocxImage:
        (attributes: Record<string, unknown>) =>
        ({ commands }) => {
          return commands.insertContent({ type: this.name, attrs: attributes })
        },
    }
  },
})

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docxImage: {
      insertDocxImage: (attributes: Record<string, unknown>) => ReturnType
    }
  }
}
