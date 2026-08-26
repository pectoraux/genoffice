/**
 * Tiptap extensions for schema-backed DOCX source identity.
 *
 * Phase 3 Increment 6: replaces arbitrary HTML attributes (data-docx-index,
 * data-passthrough, data-passthrough-type) with proper Tiptap node attributes
 * that survive setContent() → editing → getHTML() round-trips.
 *
 * These extensions extend the built-in StarterKit nodes (Paragraph, Heading,
 * ListItem) with a `docxIndex` attribute, and add a custom `PassthroughBlock`
 * node for tables/images/embedded content that must be preserved without
 * destruction.
 */

import Paragraph from '@tiptap/extension-paragraph'
import Heading from '@tiptap/extension-heading'
import ListItem from '@tiptap/extension-list-item'
import { Node, mergeAttributes } from '@tiptap/core'

// ── docxIndex attribute ──────────────────────────────────────────────────

/**
 * Shared attribute definitions for the docxIndex node attribute and the
 * paragraph alignment. Stored as global attributes so they can be applied to
 * any node type.
 *
 * The attributes render as `data-docx-index` / `data-align` in HTML (matching
 * the existing convention), but are schema-backed — Tiptap will preserve
 * them across setContent() → getHTML() round-trips.
 */
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

/** Paragraph alignment (w:jc) — set on paragraphs inside table cells. */
const alignAttribute = {
  align: {
    default: null,
    parseHTML: (element: HTMLElement) => element.getAttribute('data-align'),
    renderHTML: (attributes: { align: string | null }) => {
      if (!attributes.align) return {}
      return { 'data-align': attributes.align, style: `text-align: ${attributes.align}` }
    },
  },
}

/** Paragraph style id (w:pStyle) — set on paragraphs inside table cells. */
const styleIdAttribute = {
  styleId: {
    default: null,
    parseHTML: (element: HTMLElement) => element.getAttribute('data-style'),
    renderHTML: (attributes: { styleId: string | null }) => {
      if (!attributes.styleId) return {}
      return { 'data-style': attributes.styleId }
    },
  },
}

// ── Paragraph with docxIndex ─────────────────────────────────────────────

export const DocxParagraph = Paragraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...docxIndexAttribute,
      ...alignAttribute,
      ...styleIdAttribute,
    }
  },
})

// ── Heading with docxIndex ───────────────────────────────────────────────

export const DocxHeading = Heading.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...docxIndexAttribute,
    }
  },
})

// ── ListItem with docxIndex ──────────────────────────────────────────────

export const DocxListItem = ListItem.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...docxIndexAttribute,
    }
  },
})

// ── PassthroughBlock node ────────────────────────────────────────────────

/**
 * Custom Tiptap node for preserving DOCX blocks that the browser editor
 * cannot edit (tables, images, charts, SmartArt, OLE objects).
 *
 * The node is atomic (non-editable) — it renders as a labeled block in
 * the editor but preserves the original docxIndex so the server can copy
 * the original bytes byte-identically on save.
 *
 * Attributes:
 *   - docxIndex: number | null — the original <w:body> child index
 *   - passthroughType: string — 'table', 'image', 'passthrough', etc.
 *   - sourceText: string — the text content (for display)
 */
export interface PassthroughBlockOptions {
  HTMLAttributes: Record<string, unknown>
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    passthroughBlock: {
      insertPassthroughBlock: (attributes: {
        docxIndex?: number | null
        passthroughType?: string
        sourceText?: string
      }) => ReturnType
    }
  }
}

export const PassthroughBlock = Node.create<PassthroughBlockOptions>({
  name: 'passthroughBlock',

  group: 'block',

  atom: true,

  addOptions() {
    return {
      HTMLAttributes: {},
    }
  },

  addAttributes() {
    return {
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
      passthroughType: {
        default: 'passthrough',
        parseHTML: (element: HTMLElement) =>
          element.getAttribute('data-passthrough-type') ?? 'passthrough',
        renderHTML: (attributes: { passthroughType: string }) => ({
          'data-passthrough-type': attributes.passthroughType ?? 'passthrough',
        }),
      },
      sourceText: {
        default: '',
        parseHTML: (element: HTMLElement) => element.textContent ?? '',
        renderHTML: () => ({}), // rendered as text content, not an attribute
      },
    }
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-passthrough="true"]',
      },
    ]
  },

  renderHTML({ node, HTMLAttributes }) {
    const attrs = mergeAttributes(this.options.HTMLAttributes, HTMLAttributes, {
      'data-passthrough': 'true',
    })
    const ptType = node.attrs['passthroughType'] ?? 'passthrough'
    const text = node.attrs['sourceText'] ?? ''
    return ['div', attrs, text || `[${ptType} — edit in desktop app]`]
  },

  addCommands() {
    return {
      insertPassthroughBlock:
        (attributes) =>
        ({ commands }) => {
          return commands.insertContent({
            type: this.name,
            attrs: attributes,
          })
        },
    }
  },
})
