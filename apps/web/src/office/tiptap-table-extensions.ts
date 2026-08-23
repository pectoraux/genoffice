/**
 * Tiptap table extensions with schema-backed DOCX source identity.
 *
 * The stock @tiptap/extension-table family provides the editing machinery
 * (row/column insert/delete, header toggle, merge/split); this module adds
 * the `docxIndex` attribute to the table node (same pattern as
 * DocxParagraph/DocxHeading/DocxListItem) plus the cell-level DOCX
 * formatting attributes (fill/vAlign) so they survive the schema round-trip
 * (unknown HTML attributes are stripped by the parser).
 */

import { Table } from '@tiptap/extension-table'
import TableRow from '@tiptap/extension-table-row'
import TableCell from '@tiptap/extension-table-cell'
import TableHeader from '@tiptap/extension-table-header'
import { CellSelection } from '@tiptap/pm/tables'

// E2E hook: expose the prosemirror-tables CellSelection constructor so
// Playwright can drive real cell selections on the app editor (exposed as
// window.__genofficeWordEditor). Synthetic mouse drags do not reliably
// produce cell selections in headless Chromium, and importing a test-host
// module only works on the Vite dev server — this hook works identically
// against local dev and deployed production builds.
const e2eWindow = window as { __genofficeCellSelection?: unknown }
e2eWindow.__genofficeCellSelection = CellSelection

/** Shared docxIndex attribute definition (mirrors the paragraph extension). */
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

/** OOXML fill hex ("RRGGBB" or "AARRGGBB") → CSS #RRGGBB. */
function fillToCss(fill: string): string {
  const hex = fill.length === 8 ? fill.slice(2) : fill
  return `#${hex}`
}

/** Cell-level DOCX attributes shared by td and th nodes. */
function cellDocxAttributes(): Record<string, unknown> {
  return {
    fill: {
      default: null,
      parseHTML: (element: HTMLElement) => element.getAttribute('data-fill'),
      renderHTML: (attributes: { fill: string | null }) => {
        if (!attributes.fill) return {}
        return {
          'data-fill': attributes.fill,
          style: `background-color: ${fillToCss(attributes.fill)}`,
        }
      },
    },
    valign: {
      default: null,
      parseHTML: (element: HTMLElement) => element.getAttribute('data-valign'),
      renderHTML: (attributes: { valign: string | null }) => {
        if (!attributes.valign || attributes.valign === 'top') return {}
        return {
          'data-valign': attributes.valign,
          style: `vertical-align: ${attributes.valign}`,
        }
      },
    },
  }
}

/** Table node carrying the DOCX source identity. */
export const DocxTable = Table.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...docxIndexAttribute,
    }
  },
})

export const DocxTableRow = TableRow

/** Table cell (td) with DOCX fill/vAlign attributes. */
export const DocxTableCell = TableCell.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...cellDocxAttributes(),
    }
  },
})

/** Table header cell (th) with DOCX fill/vAlign attributes. */
export const DocxTableHeader = TableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      ...cellDocxAttributes(),
    }
  },
})
