/**
 * EXCEL-021 — Excel table banding paint (browser port of the desktop's
 * applyTableBanding).
 *
 * Approximates Excel table styles for cells that carry no explicit fill of
 * their own: the header band + bold header font, row stripes (the FIRST
 * data row is striped — Excel's firstRowStripe), the totals band,
 * first/last column emphasis, column stripes, the whole-table fill, and
 * the border-drawn families' frame (top rule MEDIUM, header rule THIN,
 * bottom rule MEDIUM).
 *
 * Purity: this module is a pure value transformation over a SPARSE cell
 * matrix (row → column → cell, the same shape Univer's createWorkbook
 * consumes). It imports NOTHING — the table shape is structurally
 * compatible with the gateway's SheetTableInfo (pre-resolved colors) and
 * the cell shape with Univer's ICellData, so callers pass their values
 * straight through. Explicit cell fills always WIN over banding.
 */

export interface TableBandingCell {
  s?: TableBandingStyle | undefined
  [key: string]: unknown
}

export interface TableBandingStyle {
  bg?: { rgb: string } | undefined
  cl?: { rgb: string } | undefined
  bl?: number | undefined
  bd?:
    | {
        t?: { s: string; cl: { rgb: string } } | undefined
        b?: { s: string; cl: { rgb: string } } | undefined
        l?: { s: string; cl: { rgb: string } } | undefined
        r?: { s: string; cl: { rgb: string } } | undefined
      }
    | undefined
  [key: string]: unknown
}

export interface TableBandingInfo {
  readonly area: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  }
  readonly headerRowCount: number
  readonly totalsRowCount?: number | undefined
  readonly showRowStripes: boolean
  readonly showColumnStripes: boolean
  readonly styleName?: string | undefined
  readonly headerFill?: string | undefined
  readonly headerFontColor?: string | undefined
  readonly stripeFill?: string | undefined
  readonly secondRowStripeFill?: string | undefined
  readonly columnStripeFill?: string | undefined
  readonly secondColumnStripeFill?: string | undefined
  readonly wholeTableFill?: string | undefined
  readonly firstColumnFill?: string | undefined
  readonly lastColumnFill?: string | undefined
  readonly totalRowFill?: string | undefined
  readonly totalRowFontColor?: string | undefined
  readonly firstHeaderCellFontColor?: string | undefined
  readonly borderColor?: string | undefined
}

export type TableBandingMatrix = Record<number, Record<number, TableBandingCell>>

/// Paints every table's banding into the sparse matrix IN PLACE. Cells the
/// snapshot did not materialize get created so banded regions render even
/// over empty cells (the desktop paints its whole viewport matrix the same
/// way).
export function applyTableBandingToMatrix(
  matrix: TableBandingMatrix,
  tables: readonly TableBandingInfo[],
): void {
  for (const table of tables) {
    const { area } = table
    // A name-less tableStyleInfo is Excel's style "None": paint nothing.
    if (
      table.styleName === undefined &&
      table.headerFill === undefined &&
      table.headerFontColor === undefined &&
      table.stripeFill === undefined
    ) {
      continue
    }
    // Colors are resolved gateway-side from the workbook's real theme
    // accents (Light/Medium/Dark variant rules) or the file's custom
    // <tableStyle> dxfs; the literals are a last-resort fallback.
    const headerFill = table.headerFill
    const headerFont = table.headerFontColor ?? '#FFFFFF'
    const stripeFill = table.stripeFill ?? '#D9E1F2'
    const dataStartRow = area.startRow + table.headerRowCount
    const totalsStartRow = area.endRow - (table.totalsRowCount ?? 0) + 1
    for (let row = area.startRow; row <= area.endRow; row += 1) {
      const isHeader = row < dataStartRow
      const isTotals = !isHeader && row >= totalsStartRow
      // Excel's firstRowStripe covers the FIRST data row (ref: Medium9
      // shades data row 1 with #B8CCE4), then alternates with
      // secondRowStripe.
      const rowParity = (row - dataStartRow) % 2
      const isStripe = !isHeader && !isTotals && table.showRowStripes && rowParity === 0
      const secondStripeFill =
        !isHeader && !isTotals && table.showRowStripes && rowParity === 1
          ? table.secondRowStripeFill
          : undefined
      for (let column = area.startColumn; column <= area.endColumn; column += 1) {
        const rowData = (matrix[row] ??= {})
        const cell = (rowData[column] ??= {})
        let style = (cell.s ?? {}) as TableBandingStyle
        if (table.borderColor !== undefined) {
          const edges: NonNullable<TableBandingStyle['bd']> = {}
          if (row === area.startRow) {
            edges.t = { s: 'medium', cl: { rgb: table.borderColor } }
          }
          if (isHeader && row === dataStartRow - 1) {
            edges.b = { s: 'thin', cl: { rgb: table.borderColor } }
          }
          if (row === area.endRow) {
            edges.b = { s: 'medium', cl: { rgb: table.borderColor } }
          }
          if (edges.t !== undefined || edges.b !== undefined) {
            cell.s = { ...style, bd: { ...(style.bd ?? {}), ...edges } }
            style = cell.s
          }
        }
        if (isHeader) {
          const fontColor =
            column === area.startColumn && table.firstHeaderCellFontColor !== undefined
              ? table.firstHeaderCellFontColor
              : headerFill !== undefined
                ? headerFont
                : (table.headerFontColor ?? '#333333')
          if (style.bg !== undefined) {
            // Baked header fill: keep it, but a default-black font still
            // takes the style's header font (Excel lets table-style text
            // win over the automatic color).
            const cellFont = (style.cl as { rgb?: string } | undefined)?.rgb
            if (headerFill !== undefined && (cellFont === undefined || cellFont === '#000000')) {
              cell.s = { ...style, cl: { rgb: fontColor }, bl: 1 }
            }
            continue
          }
          cell.s = {
            ...style,
            ...(headerFill !== undefined ? { bg: { rgb: headerFill } } : {}),
            cl: { rgb: fontColor },
            bl: 1,
          }
          continue
        }
        if (style.bg !== undefined) continue
        if (isTotals) {
          cell.s = {
            ...style,
            ...(table.totalRowFill !== undefined ? { bg: { rgb: table.totalRowFill } } : {}),
            ...(table.totalRowFontColor !== undefined
              ? { cl: { rgb: table.totalRowFontColor } }
              : {}),
            bl: 1,
          }
          continue
        }
        // Band precedence below header/totals: first/last column emphasis,
        // then row stripes, then column stripes, then the whole-table fill.
        const isFirstColumn = column === area.startColumn && table.firstColumnFill !== undefined
        const isLastColumn = column === area.endColumn && table.lastColumnFill !== undefined
        const columnStripeFill = table.showColumnStripes
          ? (column - area.startColumn) % 2 === 0
            ? table.columnStripeFill
            : table.secondColumnStripeFill
          : undefined
        const fill = isFirstColumn
          ? table.firstColumnFill
          : isLastColumn
            ? table.lastColumnFill
            : isStripe
              ? stripeFill
              : (secondStripeFill ?? columnStripeFill ?? table.wholeTableFill)
        if (fill !== undefined) cell.s = { ...style, bg: { rgb: fill } }
      }
    }
  }
}
