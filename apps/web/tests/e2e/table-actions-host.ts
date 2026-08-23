/**
 * Browser-side host for table-selection actions in the Playwright E2E tests.
 *
 * Served through the Vite dev server (imported via its URL from
 * page.evaluate) so the tests can drive a REAL prosemirror-tables
 * CellSelection on the REAL app editor (exposed as
 * window.__genofficeWordEditor by WordEditor). Synthetic mouse drags do not
 * reliably produce cell selections in headless Chromium.
 *
 * Test infrastructure only — never imported by the app.
 */
import { CellSelection } from '@tiptap/pm/tables'
import type { Node as ProsemirrorNode } from '@tiptap/pm/model'

/**
 * Minimal recursive ProseMirror node shape: just enough to walk the table
 * grid (rows → cells) without typing the full editor.
 */
interface MinimalNode {
  nodeSize: number
  childCount: number
  child(index: number): MinimalNode
}

interface MinimalEditor {
  state: {
    doc: {
      descendants(
        fn: (node: MinimalNode & { type: { name: string } }, pos: number) => boolean | void,
      ): void
      nodeAt(pos: number): MinimalNode | null
    }
    tr: { setSelection(sel: unknown): unknown }
  }
  view: { dispatch(tr: unknown): void }
}

/**
 * Select the rectangular cell range from (anchorR, anchorC) to (headR, headC)
 * in the document's FIRST table. Row/column indexes count DOM cells (the
 * rows as rendered — spans do not create extra rows).
 */
export function selectCellRange(
  anchorR: number,
  anchorC: number,
  headR: number,
  headC: number,
): boolean {
  const w = window as { __genofficeWordEditor?: MinimalEditor }
  const editor = w.__genofficeWordEditor
  if (!editor) return false
  const { state, view } = editor
  const { doc } = state
  let tablePos = -1
  doc.descendants((node, pos) => {
    if (tablePos === -1 && node.type.name === 'table') {
      tablePos = pos
      return false
    }
    return true
  })
  if (tablePos === -1) return false
  const table = doc.nodeAt(tablePos)
  if (!table) return false
  const cellPos = (ri: number, ci: number): number | null => {
    let rowOffset = tablePos + 1
    for (let r = 0; r < table.childCount; r++) {
      const row = table.child(r)
      if (r === ri) {
        let cellOffset = rowOffset + 1
        for (let c = 0; c < row.childCount; c++) {
          if (c === ci) return cellOffset
          cellOffset += row.child(c).nodeSize
        }
        return null
      }
      rowOffset += row.nodeSize
    }
    return null
  }
  const anchor = cellPos(anchorR, anchorC)
  const head = cellPos(headR, headC)
  if (anchor === null || head === null) return false
  // At runtime `doc` is a real ProseMirror Node (the structural type above
  // only avoids importing the full editor type into this test host).
  const sel = CellSelection.create(doc as ProsemirrorNode, anchor, head)
  view.dispatch(state.tr.setSelection(sel))
  return true
}
