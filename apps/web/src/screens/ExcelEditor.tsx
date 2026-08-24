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
import { cellEditFromMutation, mergeCellEdit } from '../office/cell-mutation-merge'
import { Ribbon } from './excel/Ribbon'
import { NameBox } from './excel/NameBox'
import { FormulaBar } from './excel/FormulaBar'
import { StatusBar } from './excel/StatusBar'
import { ThemeToggle } from './excel/ThemeToggle'
import { useExcelRuntime } from './excel/useExcelRuntime'
import type { ThemeState } from '../theme'
import type { SessionInfo } from '../api/client'

/**
 * ExcelEditor — GenOffice web Sheets workspace.
 *
 * Visual structure mirrors the Electron Sheets shell (title bar → ribbon →
 * name box + formula bar → worksheet → status bar) with HTTP replacing
 * Electron/native I/O. The browser remains a thin client: every value/style
 * edit commits through Univer's FRange facade (fires set-range-values),
 * journaled by the existing subscription through cell-mutation-merge.ts; the
 * browser never serializes the workbook and never runs a second engine.
 */

const SET_RANGE_VALUES_MUTATION_ID = 'sheet.mutation.set-range-values'

/** Structural mutation IDs (insert/remove row/column). */
const STRUCTURAL_MUTATION_IDS = new Set([
  'sheet.mutation.insert-row',
  'sheet.mutation.remove-rows',
  'sheet.mutation.insert-col',
  'sheet.mutation.remove-col',
])

interface JournaledStructuralOp {
  readonly kind: 'insert-rows' | 'remove-rows' | 'insert-cols' | 'remove-cols'
  readonly index: number
  readonly count: number
}

function shiftIndex(index: number, boundary: number, delta: number): number | null {
  if (index < boundary) return index
  if (delta > 0) return index + delta
  if (index < boundary - delta) return null
  return index + delta
}

const WORKBOOK_UNIT_ID = 'genoffice-web-workbook'

function dirtyKey(sheetName: string, row: number, column: number): string {
  return `${sheetName}:${row}:${column}`
}

function formatToUniverStyle(fmt: CellFormatState): IStyleData | null {
  const out: IStyleData = {}
  if (fmt.bold) out.bl = 1
  if (fmt.italic) out.it = 1
  if (fmt.underline) out.ul = { s: 1 }
  if (fmt.strikethrough) out.st = { s: 1 }
  if (fmt.fontFamily) out.ff = fmt.fontFamily
  if (typeof fmt.fontSize === 'number') out.fs = fmt.fontSize
  if (fmt.fontColor)
    out.cl = { rgb: fmt.fontColor.startsWith('#') ? fmt.fontColor : `#${fmt.fontColor}` }
  if (fmt.fillColor)
    out.bg = { rgb: fmt.fillColor.startsWith('#') ? fmt.fillColor : `#${fmt.fillColor}` }
  if (fmt.horizontalAlign) {
    out.ht = fmt.horizontalAlign === 'left' ? 1 : fmt.horizontalAlign === 'center' ? 2 : 3
  }
  if (fmt.verticalAlign) {
    out.vt = fmt.verticalAlign === 'top' ? 1 : fmt.verticalAlign === 'center' ? 2 : 3
  }
  if (fmt.wrapText) out.tb = 1
  const hasAny = Object.keys(out).length > 0
  return hasAny ? out : null
}

function buildCellDataMatrix(
  cells: Readonly<Record<string, CellState>>,
): IObjectMatrixPrimitiveType<ICellData> {
  const matrix: IObjectMatrixPrimitiveType<ICellData> = {}
  for (const [addr, cell] of Object.entries(cells)) {
    let coords: { row: number; column: number }
    try {
      coords = parseAddress(addr)
    } catch {
      continue
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

function buildRowData(
  rowHeights: Readonly<Record<string, number>> | undefined,
): IObjectArrayPrimitiveType<{ h?: number }> {
  if (!rowHeights) return {}
  const out: IObjectArrayPrimitiveType<{ h?: number }> = {}
  for (const [rowKey, points] of Object.entries(rowHeights)) {
    const row = Number(rowKey) - 1
    if (!Number.isInteger(row) || row < 0) continue
    out[row] = { h: Math.round((points * 4) / 3) }
  }
  return out
}

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
      ws.getRange(addr).setValue({ s: univerStyle })
    } catch {
      // setStyle may reject some style combinations; the file's XML
      // preserves the canonical format regardless.
    }
  }
}

export interface ExcelEditorProps {
  onRoute: (route: string) => void
  onLogout: () => Promise<void>
  session: SessionInfo | null
  theme: ThemeState
}

export function ExcelEditor({ onRoute, onLogout, session, theme }: ExcelEditorProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [runtime, setRuntime] = useState<BrowserUniverRuntime | null>(null)
  const runtimeRef = useRef<BrowserUniverRuntime | null>(null)
  const handleRef = useRef<OfficeWorkbookHandle | null>(null)
  const dirtyCellsRef = useRef<Map<string, CellEdit>>(new Map())
  const structuralOpsRef = useRef<Map<string, JournaledStructuralOp[]>>(new Map())
  const [dirty, setDirty] = useState(false)
  const [status, setStatus] = useState<string>('Ready')
  const [fileName, setFileName] = useState<string>('workbook.xlsx')
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const fileNameRef = useRef(fileName)
  fileNameRef.current = fileName

  // ── Mount: create the Univer runtime (config parity in create-browser-univer),
  //    seed a blank workbook, wire the cell-mutation journal subscription. ──
  useEffect(() => {
    if (!containerRef.current) return
    const rt = createBrowserUniver('genoffice-web-excel')
    runtimeRef.current = rt
    setRuntime(rt)
    // The blank workbook's sheet MUST carry an explicit `id` (and the full
    // set of structural defaults) — otherwise Univer auto-generates a sheet
    // id ('sheet-01') that is NOT registered in the workbook's internal
    // sheet-id lookup map. That breaks FWorkbook.setActiveRange (the
    // desktop's proven Name-Box jump path in data-tools-actions.ts:154),
    // which internally calls workbook.getSheetBySheetId(sheetId) and throws
    // "No active sheet found" when the lookup returns null. The desktop
    // never hits this because loadSnapshotIntoUniver always creates sheets
    // with explicit ids. Matching that shape here makes the blank workbook
    // navigable (B5, A1:C5) before a real file is opened — same shape
    // loadSnapshot uses below for opened workbooks.
    rt.univerAPI.createWorkbook({
      id: WORKBOOK_UNIT_ID,
      name: 'Workbook',
      sheetOrder: ['sheet1'],
      sheets: {
        sheet1: {
          id: 'sheet1',
          name: 'Sheet1',
          tabColor: '',
          hidden: 0 as BooleanNumber,
          freeze: { startRow: -1, startColumn: -1, xSplit: 0, ySplit: 0 },
          rowCount: 1000,
          columnCount: 26,
          zoomRatio: 1,
          scrollTop: 0,
          scrollLeft: 0,
          defaultColumnWidth: 100,
          defaultRowHeight: 20,
          rowHeader: { width: 46, hidden: 0 as BooleanNumber },
          columnHeader: { height: 20, hidden: 0 as BooleanNumber },
          showGridlines: 1 as BooleanNumber,
          rightToLeft: 0 as BooleanNumber,
        },
      },
    })
    const sub = subscribeToCellMutations(rt, dirtyCellsRef, structuralOpsRef, () =>
      setDirty(true),
    )
    const w = window as { __genofficeExcelRuntime?: unknown }
    w.__genofficeExcelRuntime = rt
    return () => {
      sub.dispose()
      delete w.__genofficeExcelRuntime
      rt.univer.dispose()
      runtimeRef.current = null
      setRuntime(null)
    }
  }, [])

  // ── Theme sync: mirror the resolved <html data-theme> into Univer's canvas. ──
  useEffect(() => {
    const rt = runtimeRef.current
    if (!rt) return
    rt.themeService.setDarkMode(theme.effective === 'dark')
  }, [theme.effective, runtime])

  const api = useExcelRuntime(runtime)

  const loadSnapshot = useCallback((snapshot: WorkbookSnapshot) => {
    const rt = runtimeRef.current
    if (!rt) return
    const active = rt.univerAPI.getActiveWorkbook()
    if (active) {
      try {
        rt.univerAPI.disposeUnit(active.getId())
      } catch {
        /* already gone */
      }
    }
    const sheetsConfig: Record<string, IWorksheetData> = {}
    for (const sheet of snapshot.sheets) {
      sheetsConfig[sheet.id] = {
        id: sheet.id,
        name: sheet.name,
        tabColor: '',
        hidden: (sheet as { hidden?: boolean }).hidden ? 1 : (0 as BooleanNumber),
        freeze: { startRow: -1, startColumn: -1, xSplit: 0, ySplit: 0 },
        rowCount: 1000,
        columnCount: 26,
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 100,
        defaultRowHeight: 20,
        cellData: buildCellDataMatrix(sheet.cells),
        mergeData: buildMergeData(sheet.merges),
        rowData: buildRowData(sheet.rowHeights),
        columnData: buildColumnData(sheet.colWidths),
        rowHeader: { width: 46, hidden: 0 as BooleanNumber },
        columnHeader: { height: 20, hidden: 0 as BooleanNumber },
        showGridlines: 1 as BooleanNumber,
        rightToLeft: 0 as BooleanNumber,
      }
    }
    const newWb = rt.univerAPI.createWorkbook({
      id: WORKBOOK_UNIT_ID,
      name: fileNameRef.current.replace(/\.[^.]+$/, ''),
      sheets: sheetsConfig,
    })
    for (const sheet of snapshot.sheets) {
      applyCellStyles(newWb, sheet)
    }
    dirtyCellsRef.current.clear()
    structuralOpsRef.current.clear()
    setDirty(false)
  }, [])

  const handleOpenFile = useCallback(
    async (file: File) => {
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
    },
    [loadSnapshot],
  )

  const handleSave = useCallback(async (saveAs: boolean) => {
    const handle = handleRef.current
    if (!handle) {
      setStatus('Nothing to save — open a file first')
      return
    }
    setStatus('Saving...')
    try {
      const edits = Array.from(dirtyCellsRef.current.values())
      const structuralOps = Array.from(structuralOpsRef.current.entries())
        .map(([sheetName, ops]) => ({ sheetName, ops }))
        .filter((s) => s.ops.length > 0)
      let nextFileName = handle.fileName
      if (saveAs) {
        const newName = window.prompt('Save as:', nextFileName)
        if (!newName) {
          setStatus('Save cancelled')
          return
        }
        nextFileName = newName.endsWith('.xlsx') ? newName : `${newName}.xlsx`
      }
      const savedBytes = await saveWorkbook({
        fileName: nextFileName,
        fileBytes: handle.sourceBytes,
        savePlan: {
          edits,
          ...(structuralOps.length > 0 ? { structuralOps } : {}),
        },
      })
      handleRef.current = {
        fileName: nextFileName,
        sourceBytes: savedBytes,
        revision: handle.revision + 1,
      }
      setFileName(nextFileName)
      dirtyCellsRef.current.clear()
      structuralOpsRef.current.clear()
      const blob = new Blob([savedBytes.buffer as ArrayBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
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

  const isError = status.startsWith('Open failed') || status.startsWith('Save failed')

  return (
    <div className="excel-shell" data-testid="excel-shell">
      {/* Single application header — replaces the previous double chrome
          (OfficeFrame header + ExcelEditor header). Keeps auth + nav. */}
      <header className="excel-titlebar">
        <span className="brand">GenOffice Excel</span>
        <span className="filename" title={fileName}>
          {fileName}
        </span>
        <span className={`save-state${dirty ? ' dirty' : ''}`}>
          {dirty ? '● Unsaved changes' : '✓ Saved'}
        </span>
        <div className="quick-actions">
          <button
            className="tlb-btn"
            onClick={() => fileInputRef.current?.click()}
            title="Open a .xlsx/.csv/.xls file"
          >
            Open
          </button>
          <button
            className="tlb-btn"
            onClick={() => handleSave(false)}
            disabled={!dirty}
            title="Save the workbook"
          >
            Save
          </button>
          <button
            className="tlb-btn"
            onClick={() => handleSave(true)}
            title="Save the workbook under a new name"
          >
            Save As
          </button>
          <ThemeToggle mode={theme.mode} setMode={theme.setMode} />
          <button className="tlb-btn" onClick={() => onRoute('/office')} title="Back to Office home">
            Office
          </button>
          <button
            className="tlb-btn"
            onClick={onLogout}
            title={session?.displayName ? `Sign out (${session.displayName})` : 'Sign out'}
          >
            Sign out
          </button>
          <input
            ref={fileInputRef}
            hidden
            type="file"
            accept=".xlsx,.csv,.xls"
            onChange={(e) => {
              const f = e.target.files?.[0]
              if (f) void handleOpenFile(f)
              e.target.value = ''
            }}
          />
        </div>
      </header>

      <Ribbon api={api} />

      <div className="excel-formula-row" data-testid="excel-formula-row">
        <NameBox
          activeCellA1={api?.state.activeCellA1 ?? ''}
          onGoTo={(ref) => (api ? api.goTo(ref) : 'No workbook open')}
        />
        <span className="excel-formula-fx">ƒx</span>
        <FormulaBar
          runtime={runtime}
          selectionStamp={api?.state.seq ?? 0}
          onCommit={(text) => api?.commitFormula(text)}
        />
      </div>

      <div id="genoffice-web-excel" ref={containerRef} className="excel-grid" />

      <StatusBar api={api} status={status} isError={isError} />
    </div>
  )
}

/**
 * Subscribe to Univer's cell-value mutation events. Each event carries a
 * sparse cellValue matrix; we project it into the per-cell dirty map keyed
 * by `${sheetName}:${row}:${col}` so the latest edit wins.
 *
 * Kept identical to the pre-shell behavior — the formula-priority merge in
 * cell-mutation-merge.ts (which distinguishes an explicit formula clear from
 * the recalc echo) guarantees mutation ordering can never silently convert a
 * formula back into a literal.
 */
function subscribeToCellMutations(
  runtime: BrowserUniverRuntime,
  dirtyRef: React.MutableRefObject<Map<string, CellEdit>>,
  structuralRef: React.MutableRefObject<Map<string, JournaledStructuralOp[]>>,
  onDirty: () => void,
): { dispose(): void } {
  const sub = runtime.univerAPI.addEvent(runtime.univerAPI.Event.CommandExecuted, (event) => {
    if (STRUCTURAL_MUTATION_IDS.has(event.id)) {
      const params = event.params as
        | {
            subUnitId?: string
            range?: { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number }
          }
        | undefined
      if (!params?.subUnitId || !params.range) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      const range = params.range
      const isRow = event.id.endsWith('row') || event.id.endsWith('rows')
      const isInsert = event.id.startsWith('sheet.mutation.insert')
      const start = isRow ? (range.startRow ?? 0) : (range.startColumn ?? 0)
      const end = isRow ? (range.endRow ?? 0) : (range.endColumn ?? 0)
      const count = Math.max(1, end - start + 1)
      const delta = isInsert ? count : -count

      const kind = isRow
        ? isInsert
          ? 'insert-rows'
          : 'remove-rows'
        : isInsert
          ? 'insert-cols'
          : 'remove-cols'
      const ops = structuralRef.current.get(sheetName) ?? []
      ops.push({ kind, index: start, count })
      structuralRef.current.set(sheetName, ops)

      const shifted = new Map<string, CellEdit>()
      for (const [key, edit] of dirtyRef.current.entries()) {
        const [editSheet, rowStr, colStr] = key.split(':')
        if (editSheet !== sheetName) {
          shifted.set(key, edit)
          continue
        }
        const axis = isRow ? Number(rowStr) : Number(colStr)
        const newAxis = shiftIndex(axis, start, delta)
        if (newAxis === null) continue
        if (isRow) {
          shifted.set(dirtyKey(sheetName, newAxis, Number(colStr)), { ...edit, row: newAxis })
        } else {
          shifted.set(dirtyKey(sheetName, Number(rowStr), newAxis), { ...edit, column: newAxis })
        }
      }
      dirtyRef.current.clear()
      for (const [key, edit] of shifted.entries()) dirtyRef.current.set(key, edit)
      onDirty()
      return
    }

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
        const parsed = cellEditFromMutation(sheetName, row, column, cell)
        if (!parsed) continue
        const key = dirtyKey(sheetName, row, column)
        const existing = dirtyRef.current.get(key)
        dirtyRef.current.set(key, mergeCellEdit(existing, parsed))
        touched = true
      }
    }
    if (touched) onDirty()
  })
  return { dispose: () => sub.dispose() }
}
