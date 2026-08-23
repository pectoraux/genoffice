import { useCallback, useEffect, useRef, useState } from 'react'
import type {
  BooleanNumber,
  ICellData,
  IObjectArrayPrimitiveType,
  IObjectMatrixPrimitiveType,
  IRange,
  IStyleData,
  IWorksheetData,
} from '@univerjs/core'
import type {
  CellEdit,
  CellFormatState,
  CellState,
  WorkbookSnapshot,
  WorksheetState,
} from '@genoffice/xlsx-gateway'
import { createBrowserUniver, type BrowserUniverRuntime } from '../office/create-browser-univer'
import {
  createWorkbookHandle,
  openWorkbook,
  saveWorkbook,
  readFileBytes,
  type OfficeWorkbookHandle,
} from '../api/office-client'
import { parseAddress, parseRange, columnIndex } from '../office/cell-address'
import { styles } from '../styles'

/**
 * ExcelEditor — real XLSX I/O backed by the GenOffice office API.
 *
 * Flow:
 *   New: creates a blank Univer workbook
 *   Open: user picks .xlsx → upload bytes → API reads workbook → load snapshot into Univer
 *   Save: emit only the cells the user touched (tracked via Univer's cell-change
 *         mutation events) → API applies the patch → download saved .xlsx
 *   Save As: same as Save but with a prompted filename
 *
 * The browser is a thin client — the actual XLSX parsing/mutation happens
 * server-side via @genoffice/xlsx-gateway. The editor only tracks which
 * cells changed since the last save (change-driven mutation tracking,
 * replacing the previous full-grid scan).
 */

const SET_RANGE_VALUES_MUTATION_ID = 'sheet.mutation.set-range-values'

/**
 * Build a stable per-cell key for the dirty map. Includes the sheet name so
 * edits to identically-positioned cells across two sheets don't collide.
 */
function dirtyKey(sheetName: string, row: number, column: number): string {
  return `${sheetName}:${row}:${column}`
}

/**
 * Extract a CellEdit from a Univer mutation cell payload.
 *
 * The mutation payload for `sheet.mutation.set-range-values` is a sparse
 * matrix `{ [row]: { [col]: ICellData } }`. Each cell may carry:
 *   - `v`: the cell value (string | number | boolean | null)
 *   - `f`: the formula string (with or without leading `=`)
 *   - `p`: rich-text body (we treat it as plain text for now; the engine
 *          round-trips the rich-text runs separately)
 *
 * Returns null when the cell carries no recognizable value (e.g. a
 * style-only mutation).
 */
function cellEditFromMutation(
  sheetName: string,
  row: number,
  column: number,
  cell: unknown,
): CellEdit | null {
  if (cell === null || cell === undefined) {
    // Clearing the cell.
    return {
      sheetName,
      row,
      column,
      writeValue: true,
      cell: { value: null },
    }
  }
  if (typeof cell !== 'object') return null
  const data = cell as { v?: unknown; f?: unknown }
  const formulaRaw = typeof data.f === 'string' ? data.f : undefined
  const formula = formulaRaw
    ? formulaRaw.startsWith('=')
      ? formulaRaw
      : `=${formulaRaw}`
    : undefined
  if (formula) {
    return {
      sheetName,
      row,
      column,
      writeValue: true,
      cell: { value: '', formula: formula.slice(1) },
    }
  }
  const v = data.v
  if (v === undefined) return null // style-only mutation; not a value edit
  if (
    v === null ||
    typeof v === 'string' ||
    typeof v === 'number' ||
    typeof v === 'boolean'
  ) {
    const cellState: CellState = { value: v }
    return { sheetName, row, column, writeValue: true, cell: cellState }
  }
  return null
}

/**
 * Convert a CellFormatState (xlsx-gateway's resolved cell format) into a
 * Univer IStyleData. Returns `null` when the format carries no styling
 * the loader knows how to apply — the caller then skips the setStyle call
 * rather than emitting an empty style that would clobber existing formats.
 *
 * Fields Univer cannot faithfully represent at the FRange API level are
 * documented inline; they round-trip server-side via the file's own XML
 * because the engine patches only the cells we mark dirty (style-only
 * mutations never reach the save plan today — only value edits do).
 */
function formatToUniverStyle(fmt: CellFormatState): IStyleData | null {
  const out: IStyleData = {}
  if (fmt.bold) out.bl = 1
  if (fmt.italic) out.it = 1
  // Univer's `ul` is `{ s: 1 }` (underline single). Other styles (double,
  // dash) cannot be expressed via FRange; the file's own XML carries them.
  if (fmt.underline) out.ul = { s: 1 }
  if (fmt.strikethrough) out.st = { s: 1 }
  if (fmt.fontFamily) out.ff = fmt.fontFamily
  if (typeof fmt.fontSize === 'number') out.fs = fmt.fontSize
  if (fmt.fontColor) out.cl = { rgb: fmt.fontColor.startsWith('#') ? fmt.fontColor : `#${fmt.fontColor}` }
  if (fmt.fillColor) out.bg = { rgb: fmt.fillColor.startsWith('#') ? fmt.fillColor : `#${fmt.fillColor}` }
  if (fmt.horizontalAlign) {
    out.ht = fmt.horizontalAlign === 'left' ? 1 : fmt.horizontalAlign === 'center' ? 2 : 3
  }
  if (fmt.verticalAlign) {
    out.vt = fmt.verticalAlign === 'top' ? 1 : fmt.verticalAlign === 'center' ? 2 : 3
  }
  if (fmt.wrapText) out.tb = 1
  // textRotation, indent, border, numberFormat: documented as not yet
  // applied through the FRange API in this increment; the engine
  // preserves them in the file's own XML for cells that are not
  // re-emitted by a value edit.
  const hasAny = Object.keys(out).length > 0
  return hasAny ? out : null
}

/**
 * Build the Univer cellData matrix for a sheet from the snapshot's cells.
 * Cells are addressed by A1 notation in the snapshot; we convert to
 * `{ [row]: { [col]: ICellData } }` for Univer's snapshot shape.
 */
function buildCellDataMatrix(
  cells: Readonly<Record<string, CellState>>,
): IObjectMatrixPrimitiveType<ICellData> {
  const matrix: IObjectMatrixPrimitiveType<ICellData> = {}
  for (const [addr, cell] of Object.entries(cells)) {
    let coords: { row: number; column: number }
    try {
      coords = parseAddress(addr)
    } catch {
      continue // skip malformed addresses rather than failing the whole load
    }
    const rowData = matrix[coords.row] ?? {}
    const formula = cell.formula
    const cellData: ICellData = {
      v: cell.value ?? '',
      ...(formula ? { f: formula.startsWith('=') ? formula.slice(1) : formula } : {}),
    }
    rowData[coords.column] = cellData
    matrix[coords.row] = rowData
  }
  return matrix
}

/**
 * Build the Univer mergeData (IRange[]) from the snapshot's merge address
 * list. Each entry is an inclusive row/column range.
 */
function buildMergeData(merges: readonly string[] | undefined): IRange[] {
  if (!merges || merges.length === 0) return []
  const out: IRange[] = []
  for (const addr of merges) {
    try {
      const b = parseRange(addr)
      out.push({
        startRow: b.startRow,
        endRow: b.endRow,
        startColumn: b.startColumn,
        endColumn: b.endColumn,
      })
    } catch {
      // skip malformed merge ranges
    }
  }
  return out
}

/**
 * Build the Univer rowData (with row heights) from the snapshot's
 * rowHeights map. Heights are stored in points in the snapshot; Univer
 * expects pixels (1 point ≈ 4/3 pixels at 96 DPI).
 */
function buildRowData(
  rowHeights: Readonly<Record<string, number>> | undefined,
): IObjectArrayPrimitiveType<{ h?: number }> {
  if (!rowHeights) return {}
  const out: IObjectArrayPrimitiveType<{ h?: number }> = {}
  for (const [rowKey, points] of Object.entries(rowHeights)) {
    const row = Number(rowKey) - 1 // snapshot keys are 1-based
    if (!Number.isInteger(row) || row < 0) continue
    out[row] = { h: Math.round(points * 4 / 3) }
  }
  return out
}

/**
 * Build the Univer columnData (with column widths) from the snapshot's
 * colWidths map. Widths are stored in pixels in the snapshot, keyed by
 * column label (A, B, ...).
 */
function buildColumnData(
  colWidths: Readonly<Record<string, number>> | undefined,
): IObjectArrayPrimitiveType<{ w?: number }> {
  if (!colWidths) return {}
  const out: IObjectArrayPrimitiveType<{ w?: number }> = {}
  for (const [label, width] of Object.entries(colWidths)) {
    let col: number
    try {
      col = columnIndex(label)
    } catch {
      continue
    }
    out[col] = { w: width }
  }
  return out
}

/** Apply per-cell styles after the workbook has been created. */
function applyCellStyles(
  wb: ReturnType<BrowserUniverRuntime['univerAPI']['createWorkbook']>,
  sheet: WorksheetState,
): void {
  if (!sheet.styles) return
  const ws = wb.getSheetByName(sheet.name)
  if (!ws) return
  for (const [addr, fmt] of Object.entries(sheet.styles)) {
    const univerStyle = formatToUniverStyle(fmt)
    if (!univerStyle) continue
    try {
      // Univer's FRange doesn't expose setStyle directly; use setValue with
      // an ICellData that carries the style via the `s` field.
      ws.getRange(addr).setValue({ s: univerStyle })
    } catch {
      // setStyle may reject some style combinations; the file's XML
      // preserves the canonical format regardless.
    }
  }
}

export function ExcelEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const runtimeRef = useRef<BrowserUniverRuntime | null>(null)
  const handleRef = useRef<OfficeWorkbookHandle | null>(null)
  /**
   * Change-driven dirty map. Keyed by `${sheetName}:${row}:${col}` so the
   * latest edit to a cell wins (overwriting earlier journal entries for
   * the same cell). Emptied after each successful save.
   */
  const dirtyCellsRef = useRef<Map<string, CellEdit>>(new Map())
  const [dirty, setDirty] = useState(false)
  const [status, setStatus] = useState<string>('Ready')
  const [fileName, setFileName] = useState<string>('workbook.xlsx')
  const fileInputRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    if (!containerRef.current) return
    const runtime = createBrowserUniver('genoffice-web-excel')
    runtimeRef.current = runtime
    runtime.univerAPI.createWorkbook({
      id: 'genoffice-web-workbook',
      name: 'Workbook',
      sheets: { sheet1: { name: 'Sheet1' } },
    })

    const sub = subscribeToCellMutations(runtime, dirtyCellsRef, () => setDirty(true))
    return () => {
      sub.dispose()
      runtime.univer.dispose()
      runtimeRef.current = null
    }
  }, [])

  /** Load a WorkbookSnapshot (from the API) into a fresh Univer workbook. */
  const loadSnapshot = useCallback((snapshot: WorkbookSnapshot) => {
    const rt = runtimeRef.current
    if (!rt) return
    // Dispose the old workbook and create a fresh one with the snapshot's sheets.
    try { rt.univer.dispose() } catch { /* */ }
    const fresh = createBrowserUniver('genoffice-web-excel')
    runtimeRef.current = fresh

    const sheetsConfig: Record<string, IWorksheetData> = {}
    for (const sheet of snapshot.sheets) {
      sheetsConfig[sheet.id] = {
        id: sheet.id,
        name: sheet.name,
        tabColor: '',
        hidden: (sheet as { hidden?: boolean }).hidden ? 1 : 0 as BooleanNumber,
        freeze: { startRow: -1, startColumn: -1, xSplit: 0, ySplit: 0 },
        rowCount: 1000,
        columnCount: 26,
        zoomRatio: 100,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 100,
        defaultRowHeight: 20,
        cellData: buildCellDataMatrix(sheet.cells),
        mergeData: buildMergeData(sheet.merges),
        rowData: buildRowData(sheet.rowHeights),
        columnData: buildColumnData(sheet.colWidths),
        rowHeader: { width: 0, hidden: 0 as BooleanNumber },
        columnHeader: { height: 0, hidden: 0 as BooleanNumber },
        showGridlines: 1 as BooleanNumber,
        rightToLeft: 0 as BooleanNumber,
      }
    }
    const newWb = fresh.univerAPI.createWorkbook({
      id: 'genoffice-web-workbook',
      name: fileName.replace(/\.[^.]+$/, ''),
      sheets: sheetsConfig,
    })
    // Apply per-cell styles after creation (the FRange API is the only
    // path Univer exposes for arbitrary IStyleData on existing cells).
    for (const sheet of snapshot.sheets) {
      applyCellStyles(newWb, sheet)
    }
    // Reset the dirty map — the snapshot is the new baseline.
    dirtyCellsRef.current.clear()
    // Re-subscribe on the fresh runtime (the old subscription died with the
    // previous Univer instance).
    subscribeToCellMutations(fresh, dirtyCellsRef, () => setDirty(true))
    setDirty(false)
  }, [fileName])

  /** Open a file from the user's local filesystem. */
  const handleOpenFile = useCallback(async (file: File) => {
    setStatus('Opening...')
    try {
      const bytes = await readFileBytes(file)
      const res = await openWorkbook({ fileName: file.name, fileBytes: bytes })
      handleRef.current = createWorkbookHandle(file.name, bytes)
      setFileName(file.name)
      loadSnapshot(res.snapshot)
      setStatus(`Opened ${file.name}`)
      setDirty(false)
    } catch (e) {
      setStatus(`Open failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }, [loadSnapshot])

  /** Save (or Save As) the workbook through the API. */
  const handleSave = useCallback(async (saveAs: boolean) => {
    const handle = handleRef.current
    if (!handle) {
      // New workbook — need to build from scratch. For now, show a message.
      setStatus('Nothing to save — open a file first')
      return
    }
    setStatus('Saving...')
    try {
      // Change-driven save: emit ONLY the cells the user touched since the
      // last save (or open). No full-grid scan.
      const edits = Array.from(dirtyCellsRef.current.values())
      let nextFileName = handle.fileName
      if (saveAs) {
        const newName = window.prompt('Save as:', nextFileName)
        if (!newName) { setStatus('Save cancelled'); return }
        nextFileName = newName.endsWith('.xlsx') ? newName : `${newName}.xlsx`
      }
      const savedBytes = await saveWorkbook({
        fileName: nextFileName,
        fileBytes: handle.sourceBytes,
        savePlan: { edits },
      })
      // The handle now reflects the persisted state: the new source bytes
      // and the incremented revision.
      handleRef.current = {
        fileName: nextFileName,
        sourceBytes: savedBytes,
        revision: handle.revision + 1,
      }
      setFileName(nextFileName)
      // Clear the dirty map — every edit it held is now in the saved bytes.
      dirtyCellsRef.current.clear()
      // Offer the saved file as a download.
      const blob = new Blob([savedBytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = nextFileName
      a.click()
      URL.revokeObjectURL(url)
      setStatus(`Saved ${nextFileName}`)
      setDirty(false)
    } catch (e) {
      setStatus(`Save failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }, [])

  return (
    <div style={{ height: 'calc(100vh - 64px)', display: 'flex', flexDirection: 'column', background: '#f5f6f8' }}>
      <header style={{ minHeight: 56, display: 'flex', alignItems: 'center', gap: 10, padding: '0 16px', background: '#fff', borderBottom: '1px solid #d8dde6' }}>
        <button onClick={() => onRoute('/office')} style={{ ...styles.button, padding: '7px 12px' }}>← Office</button>
        <strong style={{ flex: 1 }}>GenOffice Excel — {fileName}</strong>
        <span style={{ opacity: 0.65 }}>{dirty ? '● Unsaved changes' : '✓ Saved'}</span>
        <button style={styles.button} onClick={() => fileInputRef.current?.click()}>Open</button>
        <button style={styles.button} onClick={() => handleSave(false)} disabled={!dirty}>Save</button>
        <button style={styles.button} onClick={() => handleSave(true)}>Save As</button>
        <input
          ref={fileInputRef}
          hidden
          type="file"
          accept=".xlsx,.csv,.xls"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleOpenFile(f); e.target.value = '' }}
        />
      </header>
      {status !== 'Ready' && (
        <div style={{ padding: '4px 16px', background: '#e8f0fe', fontSize: 13, color: '#1a56c4' }}>{status}</div>
      )}
      <div id="genoffice-web-excel" ref={containerRef} style={{ flex: 1, minHeight: 0 }} />
    </div>
  )
}

/**
 * Subscribe to Univer's cell-value mutation events. Each event carries a
 * sparse cellValue matrix (`{ [row]: { [col]: ICellData } }`); we project
 * it into the per-cell dirty map keyed by `${sheetName}:${row}:${col}` so
 * the latest edit wins.
 *
 * Returns a disposable handle. The caller must dispose it when the Univer
 * instance is torn down.
 */
function subscribeToCellMutations(
  runtime: BrowserUniverRuntime,
  dirtyRef: React.MutableRefObject<Map<string, CellEdit>>,
  onDirty: () => void,
): { dispose(): void } {
  const sub = runtime.univerAPI.addEvent(
    runtime.univerAPI.Event.CommandExecuted,
    (event) => {
      if (event.id !== SET_RANGE_VALUES_MUTATION_ID) return
      const params = event.params as
        | { subUnitId?: string; cellValue?: unknown; unitId?: string }
        | undefined
      if (!params?.subUnitId) return
      const cellValue = params.cellValue
      if (typeof cellValue !== 'object' || cellValue === null) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      const matrix = cellValue as Record<string, unknown>
      let touched = false
      for (const [rowKey, rowValue] of Object.entries(matrix)) {
        const row = Number(rowKey)
        if (!Number.isInteger(row) || row < 0) continue
        if (typeof rowValue !== 'object' || rowValue === null) continue
        for (const [colKey, cell] of Object.entries(rowValue as Record<string, unknown>)) {
          const column = Number(colKey)
          if (!Number.isInteger(column) || column < 0) continue
          const edit = cellEditFromMutation(sheetName, row, column, cell)
          if (!edit) continue
          dirtyRef.current.set(dirtyKey(sheetName, row, column), edit)
          touched = true
        }
      }
      if (touched) onDirty()
    },
  )
  return { dispose: () => sub.dispose() }
}
