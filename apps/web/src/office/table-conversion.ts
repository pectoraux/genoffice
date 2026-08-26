/**
 * Table conversion between the typed wire model (SerializedTable) and the
 * browser editor's DOM representation (HTML tables with colspan/rowspan).
 *
 * PURITY: this module performs NO DOCX XML generation whatsoever. It only
 * maps between the typed table model and HTML/DOM structures — the OOXML is
 * produced exclusively by the canonical docx-engine generator
 * (generateTableModelXml) on the server. The browser never builds w:tbl XML.
 *
 * Mapping rules (OOXML grid ↔ HTML table):
 *   - OOXML gridSpan (colSpan)        ↔ HTML colspan
 *   - OOXML vMerge restart + continue ↔ HTML rowspan on the restart cell;
 *                                        the continue cells do not exist in
 *                                        the HTML row
 *   - trPr <w:tblHeader/>             ↔ <th> cells for that row
 *   - cell fill / vAlign              ↔ data-fill / data-valign attributes
 *   - cell paragraph alignment        ↔ data-align on the cell's <p>
 */

import type { SerializedRun } from '../api/office-client'
import type {
  SerializedTable,
  SerializedTableCell,
  SerializedTableParagraph,
} from '../api/office-client'

/** Escape HTML special characters in text. */
function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Convert a SerializedRun to inline HTML with marks (same rules as paragraphs). */
function runToHtml(run: SerializedRun): string {
  let html = escapeHtml(run.text)
  if (run.link) {
    html = `<a href="${escapeHtml(run.link.href)}">${html}</a>`
  }
  if (run.bold) html = `<strong>${html}</strong>`
  if (run.italic) html = `<em>${html}</em>`
  if (run.underline) html = `<u>${html}</u>`
  if (run.strike) html = `<s>${html}</s>`
  return html
}

/** Rich paragraphs of a cell → <p> HTML (data-align carries w:jc). */
function richParasToHtml(paras: readonly SerializedTableParagraph[]): string {
  return paras
    .map((p) => {
      const alignAttr = p.align ? ` data-align="${p.align}"` : ''
      const inner = p.runs.map(runToHtml).join('')
      return `<p${alignAttr}>${inner}</p>`
    })
    .join('')
}

/**
 * Compute the grid slot (start column) of every cell in every row.
 * A vMerge=continue cell occupies its slot in the grid just like any other
 * cell — the OOXML rows always cover the full grid.
 */
export function gridSlotsOf(table: SerializedTable): number[][] {
  return table.rows.map((row) => {
    const slots: number[] = []
    let col = 0
    for (const cell of row) {
      slots.push(col)
      col += cell.colSpan ?? 1
    }
    return slots
  })
}

/** Total grid column count of a table (max row width in grid columns). */
function gridColumnCount(table: SerializedTable): number {
  let max = 0
  for (const row of table.rows) {
    const width = row.reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0)
    max = Math.max(max, width)
  }
  return max
}

/**
 * How many rows does the vertical merge starting at (row, slot) span?
 * Counts consecutive vMerge=continue cells with the same colSpan below.
 */
function verticalMergeSpan(
  table: SerializedTable,
  slots: number[][],
  rowIndex: number,
  slot: number,
  colSpan: number,
): number {
  let span = 1
  for (let r = rowIndex + 1; r < table.rows.length; r++) {
    const row = table.rows[r]
    const rowSlots = slots[r]
    let matched = false
    for (let c = 0; c < row.length; c++) {
      if (rowSlots[c] === slot) {
        const cell = row[c]
        if (cell.vMerge === 'continue' && (cell.colSpan ?? 1) === colSpan) matched = true
        break
      }
    }
    if (!matched) break
    span++
  }
  return span
}

/**
 * SerializedTable → HTML for the Tiptap editor.
 * docxIndex is carried as data-docx-index (schema-backed on the table node).
 */
export function tableToHtml(table: SerializedTable, docxIndex: number | null): string {
  const slots = gridSlotsOf(table)
  const indexAttr = docxIndex !== null ? ` data-docx-index="${docxIndex}"` : ''
  const rowsHtml = table.rows
    .map((row, ri) => {
      const isHeader = table.headerRows?.[ri] === true
      const cellsHtml = row
        .map((cell, ci) => {
          // vMerge continue cells are implicit in HTML (the restart cell's
          // rowspan covers them).
          if (cell.vMerge === 'continue') return ''
          const slot = slots[ri][ci]
          const colSpan = cell.colSpan ?? 1
          const attrs: string[] = []
          if (colSpan > 1) attrs.push(`colspan="${colSpan}"`)
          if (cell.vMerge === 'restart') {
            const rowSpan = verticalMergeSpan(table, slots, ri, slot, colSpan)
            if (rowSpan > 1) attrs.push(`rowspan="${rowSpan}"`)
          }
          if (cell.fill) attrs.push(`data-fill="${cell.fill}"`)
          if (cell.vAlign && cell.vAlign !== 'top') attrs.push(`data-valign="${cell.vAlign}"`)
          const attrStr = attrs.length > 0 ? ` ${attrs.join(' ')}` : ''
          const tag = isHeader ? 'th' : 'td'
          const content = cell.richParas?.length ? richParasToHtml(cell.richParas) : '<p></p>'
          return `<${tag}${attrStr}>${content}</${tag}>`
        })
        .join('')
      return `<tr>${cellsHtml}</tr>`
    })
    .join('')
  return `<table${indexAttr}><tbody>${rowsHtml}</tbody></table>`
}

// ── HTML → SerializedTable ──────────────────────────────────────────────────

interface DomCell {
  row: number
  col: number
  rowSpan: number
  colSpan: number
  header: boolean
  element: HTMLTableCellElement
}

/**
 * Expand an HTML table's rowspan/colspan cells onto an explicit grid:
 * grid[r][c] = the DomCell covering row r, column c (null = hole).
 */
function expandHtmlGrid(rows: readonly HTMLTableRowElement[]): {
  grid: Array<Array<DomCell | null>>
  colCount: number
} {
  const grid: Array<Array<DomCell | null>> = []
  let colCount = 0
  rows.forEach((tr, r) => {
    if (!grid[r]) grid[r] = []
    let c = 0
    for (const td of Array.from(tr.cells)) {
      while (grid[r][c] !== undefined && grid[r][c] !== null) c++
      const rowSpan = Math.max(1, td.rowSpan || 1)
      const colSpan = Math.max(1, td.colSpan || 1)
      const cell: DomCell = {
        row: r,
        col: c,
        rowSpan,
        colSpan,
        header: td.tagName.toLowerCase() === 'th',
        element: td,
      }
      for (let rr = r; rr < r + rowSpan; rr++) {
        if (!grid[rr]) grid[rr] = []
        for (let cc = c; cc < c + colSpan; cc++) grid[rr][cc] = cell
      }
      colCount = Math.max(colCount, c + colSpan)
      c += colSpan
    }
  })
  return { grid, colCount }
}

/** Parse one cell's <p> elements into wire paragraphs (runs via parseRuns). */
function cellParagraphs(td: HTMLTableCellElement): {
  paras: string[]
  richParas: SerializedTableParagraph[]
} {
  const parseRuns = parseRunsRef()
  const paras: string[] = []
  const richParas: SerializedTableParagraph[] = []
  const pElements = Array.from(td.querySelectorAll(':scope > p'))
  for (const p of pElements) {
    const runs = parseRuns(p)
    paras.push(p.textContent ?? '')
    const align = p.getAttribute('data-align') ?? undefined
    richParas.push({
      runs,
      ...(align ? { align: align as SerializedTableParagraph['align'] } : {}),
    })
  }
  if (paras.length === 0) {
    paras.push('')
    richParas.push({ runs: [] })
  }
  return { paras, richParas }
}

/**
 * HTML table element → SerializedTable.
 *
 * `preserveFrom` (the table as loaded from the server, when this is an
 * existing table) contributes the byte-preservation fields: cell rawTcPr /
 * borders / color / bold are copied from the original cell at the same grid
 * slot while the grid dimensions match, and table-level properties echo
 * through (recomputed column widths when the column count changed).
 */
export function tableFromHtml(
  tableEl: HTMLTableElement,
  preserveFrom?: SerializedTable,
): SerializedTable {
  const domRows = Array.from(tableEl.querySelectorAll('tr'))
  const { grid, colCount } = expandHtmlGrid(domRows)

  // Preserved cells by grid slot (only meaningful while dimensions match).
  const sameDims = preserveFrom
    ? grid.length === preserveFrom.rows.length && colCount === gridColumnCount(preserveFrom)
    : false
  let preserved: Map<string, SerializedTableCell> | null = null
  if (preserveFrom && sameDims) {
    preserved = new Map()
    const slots = gridSlotsOf(preserveFrom)
    preserveFrom.rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        preserved!.set(`${ri}:${slots[ri][ci]}`, cell)
      })
    })
  }

  const rows: SerializedTableCell[][] = []
  const headerRows: boolean[] = []
  for (let r = 0; r < grid.length; r++) {
    const rowCells: SerializedTableCell[] = []
    let headerCount = 0
    let normalCount = 0
    let c = 0
    while (c < colCount) {
      const cellData = grid[r][c]
      if (!cellData) {
        c++
        continue
      }
      const span = cellData.colSpan
      if (cellData.row === r) {
        // A cell that starts in this row (normal or vMerge restart).
        const { paras, richParas } = cellParagraphs(cellData.element)
        const preservedCell = preserved?.get(`${r}:${c}`)
        const fill = cellData.element.getAttribute('data-fill') ?? preservedCell?.fill
        const vAlignRaw = cellData.element.getAttribute('data-valign')
        const vAlign = vAlignRaw ?? preservedCell?.vAlign
        rowCells.push({
          paras,
          richParas,
          ...(span > 1 ? { colSpan: span } : {}),
          ...(cellData.rowSpan > 1 ? { vMerge: 'restart' as const } : {}),
          ...(fill ? { fill } : {}),
          ...(preservedCell?.color ? { color: preservedCell.color } : {}),
          ...(preservedCell?.bold ? { bold: preservedCell.bold } : {}),
          ...(vAlign && vAlign !== 'top' ? { vAlign: vAlign as 'top' | 'center' | 'bottom' } : {}),
          ...(preservedCell?.borders ? { borders: preservedCell.borders } : {}),
          ...(preservedCell?.rawTcPr ? { rawTcPr: preservedCell.rawTcPr } : {}),
        })
        if (cellData.header) headerCount++
        else normalCount++
      } else {
        // Slot covered by a rowspan from above → vMerge continue cell with
        // the restart cell's width.
        const preservedCell = preserved?.get(`${r}:${c}`)
        rowCells.push({
          paras: [''],
          richParas: [{ runs: [] }],
          ...(span > 1 ? { colSpan: span } : {}),
          vMerge: 'continue' as const,
          ...(preservedCell?.fill ? { fill: preservedCell.fill } : {}),
          ...(preservedCell?.borders ? { borders: preservedCell.borders } : {}),
          ...(preservedCell?.rawTcPr ? { rawTcPr: preservedCell.rawTcPr } : {}),
        })
      }
      c += span
    }
    if (rowCells.length > 0) {
      rows.push(rowCells)
      headerRows.push(headerCount > 0 && normalCount === 0)
    }
  }

  // Table-level properties: echo from the loaded table when the column count
  // is unchanged; otherwise recompute even column widths.
  const src = preserveFrom
  const colsUnchanged = !!src && (src.colWidthsPct?.length ?? -1) === colCount
  const colWidthsPct = colsUnchanged
    ? src!.colWidthsPct
    : Array.from({ length: colCount }, () => 100 / colCount)
  const out: SerializedTable = {
    rows,
    ...(colWidthsPct ? { colWidthsPct } : {}),
    ...(colsUnchanged && src!.colWidthsTwips ? { colWidthsTwips: src!.colWidthsTwips } : {}),
    ...(src?.widthPct !== undefined ? { widthPct: src.widthPct } : {}),
    ...(src?.autoLayout ? { autoLayout: src.autoLayout } : {}),
    ...(src?.cellMarTwips ? { cellMarTwips: src.cellMarTwips } : {}),
    ...(src?.borders ? { borders: src.borders } : {}),
    ...(src?.align ? { align: src.align } : {}),
    ...(src?.indentTwips !== undefined ? { indentTwips: src.indentTwips } : {}),
    ...(src?.tblStyleId !== undefined ? { tblStyleId: src.tblStyleId } : {}),
    ...(src?.bidiVisual ? { bidiVisual: src.bidiVisual } : {}),
    ...(headerRows.some(Boolean) ? { headerRows } : {}),
  }
  // Row heights / raw trPr: keep per-row while the row count matches;
  // new/removed rows drop theirs (defaults apply).
  const extra: {
    rowHeightsTwips?: readonly (number | null)[]
    rowHeightRules?: readonly ('atLeast' | 'exact' | null)[]
    rawTrPrs?: readonly (string | null)[]
  } = {}
  if (src && src.rows.length === rows.length) {
    if (src.rowHeightsTwips) extra.rowHeightsTwips = src.rowHeightsTwips
    if (src.rowHeightRules) extra.rowHeightRules = src.rowHeightRules
    if (src.rawTrPrs) extra.rawTrPrs = src.rawTrPrs
  }
  return { ...out, ...extra }
}

/**
 * Fingerprint of a table's EDITABLE surface (grid structure + cell content +
 * editable cell props). Computed identically at load time (from the wire
 * table) and at save time (from the reconstructed table), so unchanged
 * tables compare equal regardless of byte-preservation echo fields.
 */
export function tableGridFingerprint(table: SerializedTable): string {
  const grid = table.rows.map((row) =>
    row.map((cell) => ({
      s: cell.colSpan ?? 1,
      v: cell.vMerge ?? null,
      f: cell.fill ?? null,
      va: cell.vAlign ?? null,
      p: (
        cell.richParas ??
        (cell.paras.map((t) => ({ runs: [{ text: t }] })) as SerializedTableParagraph[])
      ).map((p) => ({
        a: p.align ?? null,
        r: p.runs.map((r: SerializedRun) => [
          r.text,
          r.bold === true,
          r.italic === true,
          r.underline === true,
          r.strike === true,
          r.link?.href ?? null,
        ]),
      })),
    })),
  )
  return JSON.stringify({
    grid,
    headerRows: table.headerRows ?? [],
  })
}

// parseRuns is injected lazily to avoid a static import cycle with the
// screens layer (parse-runs itself has no dependencies).
let parseRunsFn: ((element: Element) => SerializedRun[]) | null = null
export function setTableParseRuns(fn: (element: Element) => SerializedRun[]): void {
  parseRunsFn = fn
}
function parseRunsRef(): (element: Element) => SerializedRun[] {
  if (!parseRunsFn) throw new Error('setTableParseRuns was not called')
  return parseRunsFn
}
