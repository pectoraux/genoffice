import { useCallback, useEffect, useRef, useState } from 'react'
import { CellValueType } from '@univerjs/core'
import type {
  BooleanNumber,
  ICellData,
  IObjectArrayPrimitiveType,
  IObjectMatrixPrimitiveType,
  IRange,
  IStyleData,
  IWorksheetData,
} from '@univerjs/core'
// Side-effect import — loads @univerjs/sheets-filter's facade types, which
// augment FWorksheet with the public getFilter(): FFilter | null method
// (FWorksheetFilterMixin). The runtime side is already wired by
// UniverSheetsFilterPreset in create-browser-univer.ts; this import only
// surfaces the TypeScript signatures so the filter save snapshot
// (collectFilterStates) and the load-time install typecheck with the
// PUBLIC typed facade — no `as unknown as` casts, no private internals.
import '@univerjs/sheets-filter/facade'
import type { ICustomFilter, IFilterColumn } from '@univerjs/sheets-filter'
import type {
  CellEdit,
  CellFormatState,
  CellState,
  FilterColumnState,
  SheetFilterState,
  WorkbookSnapshot,
} from '@genoffice/xlsx-gateway'
import {
  createBrowserUniver,
  installJournalSuppressionUndoFilter,
  journalSuppression as moduleJournalSuppression,
  primeSortFormulaInterceptor,
  type BrowserUniverRuntime,
} from '../office/create-browser-univer'
import {
  createWorkbookHandle,
  openWorkbook,
  saveWorkbook,
  readFileBytes,
  type OfficeWorkbookHandle,
} from '../api/office-client'
import { parseAddress, parseRange, columnIndex } from '../office/cell-address'
import {
  cellEditFromMutation,
  mergeCellEdit,
  numfmtEditsFromMutation,
} from '../office/cell-mutation-merge'
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

/** Number-format mutation ID (sheets-numfmt preset). */
const SET_NUMFMT_MUTATION_ID = 'sheet.mutation.set.numfmt'

/** Sort / reorder-range mutation ID (sheets-sort preset). */
const REORDER_RANGE_MUTATION_ID = 'sheet.mutation.reorder-range'

/** Freeze-pane mutation ID (built-in sheets). */
const SET_FROZEN_MUTATION_ID = 'sheet.mutation.set-frozen'

/**
 * Filter mutation IDs (sheets-filter preset) — the same set the desktop's
 * App.tsx listens for. Any of these marks the sheet filter-dirty; the save
 * snapshots the LIVE filter model declaratively (never replays mutations).
 */
const FILTER_MUTATION_IDS = new Set([
  'sheet.mutation.set-filter-range',
  'sheet.mutation.set-filter-criteria',
  'sheet.mutation.remove-filter',
  'sheet.mutation.re-calc-filter',
])

/** Structural mutation IDs (insert/remove row/column). */
const STRUCTURAL_MUTATION_IDS = new Set([
  'sheet.mutation.insert-row',
  'sheet.mutation.remove-rows',
  'sheet.mutation.insert-col',
  'sheet.mutation.remove-col',
])

/** Merge mutation IDs (add/remove worksheet merge). */
const ADD_MERGE_MUTATION_ID = 'sheet.mutation.add-worksheet-merge'
const REMOVE_MERGE_MUTATION_ID = 'sheet.mutation.remove-worksheet-merge'

interface JournaledStructuralOp {
  readonly kind:
    | 'insert-rows'
    | 'remove-rows'
    | 'insert-cols'
    | 'remove-cols'
    | 'merge-cells'
    | 'unmerge-cells'
    | 'reorder-rows'
  readonly index: number
  readonly count: number
  /** Range for merge/unmerge/reorder-rows ops (startRow/endRow/startColumn/endColumn). */
  readonly range?: {
    readonly startRow: number
    readonly endRow: number
    readonly startColumn: number
    readonly endColumn: number
  }
  /** Permutation map for reorder-rows ops ({ srcRow: destRow }, 0-based). */
  readonly order?: Readonly<Record<number, number>>
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
  if (fmt.wrapText) {
    // WrapStrategy.WRAP = 3 (Univer enum: UNSPECIFIED=0, OVERFLOW=1, CLIP=2, WRAP=3).
    // The previous implementation used tb=1 (OVERFLOW), which is the default
    // non-wrapping strategy — a no-op for wrap.
    out.tb = 3
  }
  const hasAny = Object.keys(out).length > 0
  return hasAny ? out : null
}

function buildCellDataMatrix(
  cells: Readonly<Record<string, CellState>>,
  styles?: Readonly<Record<string, CellFormatState>>,
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
    const value = cell.value ?? ''
    // Seed the explicit value TYPE alongside the value. Univer's general
    // type system (getCellValueType) infers the same types from typeof v,
    // so display behavior is unchanged — but the canonical SORT
    // comparator (SheetsSortController._getCommonValue) branches on the
    // EXPLICIT `t` field and falls through to `String(v)` when it is
    // absent, which would sort numbers as strings ("10" < "9") and
    // diverge from Excel. Univer's own cell-editing path
    // (FRange.setValueForCell) always sets `t`; the snapshot seeding now
    // mirrors that so canonical sort sees typed cells.
    const valueType =
      typeof value === 'number'
        ? CellValueType.NUMBER
        : typeof value === 'boolean'
          ? CellValueType.BOOLEAN
          : typeof value === 'string' && value !== ''
            ? CellValueType.STRING
            : undefined
    // Inline the snapshot's resolved format INTO the cellData at create
    // time (ICellData.s accepts an inline IStyleData — Univer registers
    // it and swaps in the style ID). The previous post-create
    // applyCellStyles pass fired setValue({s}) mutations AFTER the
    // formula engine had started processing the live formulas; the
    // mutation/recalc interplay clobbered styled formula cells (the f
    // disappeared — see the Phase 4 Inc. 3 revision). Seeding styles
    // inline eliminates that race entirely: no style mutations fire
    // during load.
    const univerStyle = styles?.[addr] ? formatToUniverStyle(styles[addr]) : null
    const cellData: ICellData = {
      v: value,
      ...(valueType !== undefined ? { t: valueType } : {}),
      // Univer's INTERNAL cell.f convention INCLUDES the leading '='
      // (isFormulaString requires `value.substring(0, 1) === "="`). The
      // snapshot's CellState.formula already carries the '=' (see
      // readBasicWorkbook's parse: `formula: \`=${...}\``), so it is
      // seeded VERBATIM — stripping it here would seed a dead formula
      // the engine never calculates (v stays the raw cached value) and
      // the canonical sort's FormulaReorderController would crash on
      // (getFormulaStringByCell returns null for an '='-less f, then
      // moveFormulaRefOffset(null) throws
      // "Cannot read properties of null (reading 'length')" — silently
      // swallowed by sequence(), aborting the sort). The JOURNAL's
      // CellEdit wire format strips the '=' (XLSX <f> elements have no
      // '='); that stripping stays in cellEditFromMutation, untouched.
      ...(formula ? { f: formula.startsWith('=') ? formula : `=${formula}` } : {}),
      ...(univerStyle ? { s: univerStyle } : {}),
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

// applyCellStyles was REMOVED in the Phase 4 Inc. 3 revision: seeding styles
// via post-create setValue({s}) mutations raced the formula engine's live
// recalculation and clobbered styled formula cells (their f disappeared).
// Styles are now inlined into the cellData at create time by
// buildCellDataMatrix — no style mutations fire during load.

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
  /// Load-time journal suppression — the SAME convention the desktop's
  /// univer-sync.ts applies via journalSuppression: createWorkbook on a
  /// snapshot with LIVE formulas (f carries the leading '=', so the
  /// formula engine calculates them) fires synchronous set-range-values
  /// recalc echoes; without suppression those load echoes would pollute
  /// the journal and mark a freshly opened workbook "dirty". A load is
  /// not an edit.
  const journalSuppressionRef = useRef(false)
  // Per-sheet journaled freeze state. Seeded from the snapshot on open,
  // updated when the user changes View → Freeze Panes (the journal
  // subscription captures sheet.mutation.set-frozen), and emitted on save
  // as a BrowserSheetPageSetupState. Cleared on snapshot reload so stale
  // freeze state never persists across an open.
  const freezeStateRef = useRef<Map<string, { frozenRows: number; frozenColumns: number }>>(
    new Map(),
  )
  // ── AutoFilter journal (Data → Filter). Two structures, desktop parity
  //    (App.tsx filterDirty / univer-sync.ts filterOrigins):
  //    - filterDirtyRef: sheet NAMES whose filter changed in-session. A
  //      sheet absent from this set is NOT filter-dirty: its save plan
  //      carries NO filter state, so a no-op save preserves the file's own
  //      <autoFilter> XML byte-for-byte.
  //    - filterOriginsRef: per-sheet origin range — the union of the file's
  //      filter range (seeded on open) and every in-session filter range.
  //      The save's visibilityRange is this union, so removing or moving a
  //      filter still unhides the OLD span's rows (the canonical
  //      applyFilterState unhides every in-span row not listed in
  //      hiddenRows).
  const filterDirtyRef = useRef<Set<string>>(new Set())
  const filterOriginsRef = useRef<
    Map<string, { startRow: number; endRow: number; startColumn: number; endColumn: number }>
  >(new Map())
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
    // The blank workbook above was the FIRST unit creation — Univer loaded
    // all sheet plugins for it (lazily, per-type). NOW the sort's formula
    // interceptor can be force-instantiated: without this, its
    // registration waits for the async onSteady lifecycle stage and the
    // FIRST sort of a session races it (rows reorder with verbatim,
    // un-rewritten formula references — see primeSortFormulaInterceptor).
    primeSortFormulaInterceptor(rt.univer)
    // Install the load-time undo filter (desktop parity): drops undo
    // entries pushed while journalSuppression is active, so opening a
    // filtered workbook does not leave "undo the file's filter" on the
    // stack. Idempotent; patches LocalUndoRedoService.prototype once.
    installJournalSuppressionUndoFilter()
    const sub = subscribeToCellMutations(
      rt,
      dirtyCellsRef,
      structuralOpsRef,
      freezeStateRef,
      journalSuppressionRef,
      filterDirtyRef,
      filterOriginsRef,
      () => setDirty(true),
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
    // Clear the freeze-state journal before seeding from the snapshot —
    // stale freeze state from a previous workbook must never leak across
    // an open.
    freezeStateRef.current.clear()
    const sheetsConfig: Record<string, IWorksheetData> = {}
    for (const sheet of snapshot.sheets) {
      const fr = sheet.freeze
      const frozenRows = fr && fr.frozenRows > 0 ? fr.frozenRows : 0
      const frozenColumns = fr && fr.frozenColumns > 0 ? fr.frozenColumns : 0
      sheetsConfig[sheet.id] = {
        id: sheet.id,
        name: sheet.name,
        tabColor: '',
        hidden: (sheet as { hidden?: boolean }).hidden ? 1 : (0 as BooleanNumber),
        // Univer freeze config: startRow/startColumn are the first
        // scrollable (non-frozen) row/column index — so startRow = frozenRows,
        // startColumn = frozenColumns. -1 means "no freeze" on that axis.
        freeze: {
          startRow: frozenRows > 0 ? frozenRows : -1,
          startColumn: frozenColumns > 0 ? frozenColumns : -1,
          xSplit: frozenColumns,
          ySplit: frozenRows,
        },
        rowCount: 1000,
        columnCount: 26,
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 100,
        defaultRowHeight: 20,
        cellData: buildCellDataMatrix(sheet.cells, sheet.styles),
        mergeData: buildMergeData(sheet.merges),
        rowData: buildRowData(sheet.rowHeights),
        columnData: buildColumnData(sheet.colWidths),
        rowHeader: { width: 46, hidden: 0 as BooleanNumber },
        columnHeader: { height: 20, hidden: 0 as BooleanNumber },
        showGridlines: 1 as BooleanNumber,
        rightToLeft: 0 as BooleanNumber,
      }
    }
    // Suppress the journal for the whole load (desktop
    // journalSuppression parity): createWorkbook on a live-formula
    // snapshot fires set-range-values mutations (the engine registering
    // and calculating formulas) — none of them are user edits, so none
    // of them may dirty the workbook. Cell styles are inlined INTO the
    // cellData (see buildCellDataMatrix) — no post-create style
    // mutations fire during load.
    journalSuppressionRef.current = true
    moduleJournalSuppression.active = true
    // Clear the filter journal before seeding from the snapshot — stale
    // filter-dirty marks from a previous workbook must never leak across
    // an open (a reopened workbook starts NOT filter-dirty, so a no-op
    // save preserves the file's own <autoFilter> XML).
    filterDirtyRef.current.clear()
    filterOriginsRef.current.clear()
    try {
      rt.univerAPI.createWorkbook({
        id: WORKBOOK_UNIT_ID,
        name: fileNameRef.current.replace(/\.[^.]+$/, ''),
        sheets: sheetsConfig,
      })
      for (const sheet of snapshot.sheets) {
        // Seed the freeze-state journal from the snapshot — so a save
        // without any freeze change still round-trips the file's existing
        // freeze. The journal subscription updates this map when the user
        // toggles freeze interactively.
        if (sheet.freeze && (sheet.freeze.frozenRows > 0 || sheet.freeze.frozenColumns > 0)) {
          freezeStateRef.current.set(sheet.name, {
            frozenRows: sheet.freeze.frozenRows,
            frozenColumns: sheet.freeze.frozenColumns,
          })
        }
        // Render the file's existing AutoFilter in the REAL Univer UI:
        // install the filter range, then re-apply the file's criteria so
        // the live model is complete (dropdowns reflect the file state and
        // a later user edit snapshots the FULL filter, not just the delta).
        // The criteria re-application recalculates filteredOutRows, which
        // drives Univer's render-time row-hidden interceptor — the grid
        // shows exactly the rows the file's criteria keep visible. All of
        // it runs under journal suppression: installing the FILE's filter
        // is a load, not an edit.
        const fs = sheet.filterState
        if (fs && fs.filter) {
          try {
            const wb = rt.univerAPI.getActiveWorkbook()
            const ws = wb?.getSheetByName(sheet.name)
            const range = fs.filter.range
            const fFilter = ws
              ?.getRange(
                range.startRow,
                range.startColumn,
                range.endRow - range.startRow + 1,
                range.endColumn - range.startColumn + 1,
              )
              ?.createFilter()
            if (fFilter) {
              for (const column of fs.filter.columns) {
                const absoluteColumn = range.startColumn + column.colId
                // IFilterColumn criteria shape (Univer's filter model):
                // { colId, filters?: { blank?: true, filters?: string[] },
                //   customFilters?: { and?: TRUE, customFilters: [c1] | [c1, c2] } }.
                const criteria: IFilterColumn = { colId: column.colId }
                if (column.values !== undefined || column.blank) {
                  criteria.filters = {
                    ...(column.values !== undefined ? { filters: [...column.values] } : {}),
                    ...(column.blank ? { blank: true } : {}),
                  }
                }
                if (column.customs) {
                  const customFilters = column.customs.filters.map((custom) => ({
                    val: custom.val,
                    ...(custom.operator !== undefined ? { operator: custom.operator } : {}),
                  }))
                  // ICustomFilters.customFilters is a 1-or-2 tuple; the
                  // gateway's parser rejects criteria-less and >2-entry
                  // filterColumns, so the runtime shape is already valid.
                  criteria.customFilters = {
                    ...(column.customs.and === true ? { and: 1 } : {}),
                    customFilters: customFilters as
                      [ICustomFilter] | [ICustomFilter, ICustomFilter],
                  }
                }
                fFilter.setColumnFilterCriteria(absoluteColumn, criteria)
              }
            }
            // Record the origin: the file's own filter range. A later
            // move/remove unhides this span (visibilityRange = union).
            filterOriginsRef.current.set(sheet.name, { ...range })
          } catch {
            // Installing the file's filter must never fail the open —
            // the workbook still renders; only the filter dropdowns are
            // absent (and a no-op save preserves the file's XML).
          }
        }
      }
    } finally {
      journalSuppressionRef.current = false
      moduleJournalSuppression.active = false
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
      // Emit per-sheet page-setup states for every sheet with journaled
      // freeze (View → Freeze Panes). The engine's applyPageSetupState
      // writes the <pane> element into the worksheet XML on save.
      const pageSetupStates = Array.from(freezeStateRef.current.entries())
        .filter(([, f]) => f.frozenRows > 0 || f.frozenColumns > 0)
        .map(([sheetName, f]) => ({
          sheetName,
          frozenRows: f.frozenRows,
          frozenColumns: f.frozenColumns,
        }))
      // Snapshot the LIVE filter model for every filter-dirty sheet
      // (Data → Filter). Declarative, desktop collectFilterStates parity:
      // never replay mutations. Sheets NOT in the set emit NO filter state,
      // so their <autoFilter> XML survives a no-op save byte-for-byte.
      const filterStates = collectFilterStates(runtimeRef.current, filterDirtyRef, filterOriginsRef)
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
          ...(pageSetupStates.length > 0 ? { pageSetupStates } : {}),
          ...(filterStates.length > 0 ? { filterStates } : {}),
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
      // A saved filter state is now IN the source bytes — the sheet is no
      // longer filter-dirty (another no-op save must not re-emit it).
      filterDirtyRef.current.clear()
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
          <button
            className="tlb-btn"
            onClick={() => onRoute('/office')}
            title="Back to Office home"
          >
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
 * The formula-priority merge in cell-mutation-merge.ts (which distinguishes
 * an explicit formula clear from the recalc echo) guarantees mutation
 * ordering can never silently convert a formula back into a literal.
 *
 * Phase 4 Increment 3 expanded the journal to also capture:
 *   - `sheet.mutation.set.numfmt`   → style.numberFormat on CellEdit
 *   - `sheet.mutation.reorder-range` → re-read the sorted range, journal
 *      every cell's post-sort value (Univer's sort writes directly into the
 *      worksheet cellDataMatrix, NOT through set-range-values — without
 *      this handler, sort would be in-session only)
 *   - `sheet.mutation.set-frozen`    → per-sheet freeze state, emitted on
 *      save as a BrowserSheetPageSetupState (the canonical SheetPageSetupState
 *      family in applyCellEditsToXlsx writes the <pane> element)
 *
 * Phase 4 Increment 4 adds the filter family (desktop parity — the save
 * snapshots the LIVE filter model, never replays mutations):
 *   - `sheet.mutation.set-filter-range | set-filter-criteria | remove-filter
 *      | re-calc-filter` → mark the sheet filter-dirty. On save, the live
 *      FFilter model is snapshotted as a canonical SheetFilterState
 *      (range + criteria + getFilteredOutRows()), composed with the
 *      sheet's origin range so a moved/removed filter still unhides its
 *      old span.
 *
 * `suppressionRef` mirrors the desktop's journalSuppression: while a
 * snapshot load is in progress, every mutation is ignored — a load is not
 * an edit. createWorkbook on a live-formula snapshot fires synchronous
 * recalc echoes that must never reach the journal.
 */

/**
 * Snapshot the LIVE Univer filter model for every filter-dirty sheet into
 * the canonical SheetFilterState wire shape (desktop collectFilterStates
 * parity — a declarative state snapshot, never a mutation replay).
 *
 * For each dirty sheet:
 *   - No filter in the live model → the user removed it → emit
 *     `{ filter: null, hiddenRows: [] }` with the ORIGIN's visibilityRange:
 *     the gateway removes the <autoFilter> and unhides the old span.
 *   - A filter exists → emit its range + per-column criteria
 *     (getColumnFilterCriteria → values / blank / customs) +
 *     getFilteredOutRows() as hiddenRows, with the visibilityRange as the
 *     UNION of the origin and the live range (a moved filter still unhides
 *     its old span).
 *
 * Color filters cannot be serialized by the canonical gateway — they fail
 * closed (throw), matching the desktop's appColorFiltersUnsaveable. The
 * save surfaces the error instead of silently dropping criteria.
 */
function collectFilterStates(
  runtime: BrowserUniverRuntime | null,
  filterDirtyRef: React.MutableRefObject<Set<string>>,
  filterOriginsRef: React.MutableRefObject<
    Map<string, { startRow: number; endRow: number; startColumn: number; endColumn: number }>
  >,
): SheetFilterState[] {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const states: SheetFilterState[] = []
  for (const sheetName of filterDirtyRef.current) {
    const worksheet = workbook.getSheetByName(sheetName)
    if (!worksheet) continue
    const origin = filterOriginsRef.current.get(sheetName)
    // The typed filter facade (sheets-filter preset's FWorksheet mixin):
    // getFilter() returns the FFilter handle, or null when the sheet has no
    // filter (the user removed it).
    const filter = worksheet.getFilter()
    if (!filter) {
      // The user removed the filter; unhide what it was hiding. Without an
      // origin there is nothing that could have been filter-hidden.
      if (!origin) continue
      states.push({
        sheetName,
        filter: null,
        hiddenRows: [],
        visibilityRange: { ...origin },
      })
      continue
    }
    const filterRange = filter.getRange()
    const range = {
      startRow: filterRange.getRow(),
      startColumn: filterRange.getColumn(),
      endRow: filterRange.getRow() + filterRange.getHeight() - 1,
      endColumn: filterRange.getColumn() + filterRange.getWidth() - 1,
    }
    const columns: FilterColumnState[] = []
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const criteria = filter.getColumnFilterCriteria(column)
      if (!criteria) continue
      if (criteria.colorFilters) {
        throw new Error(
          'Color filters cannot be saved as XLSX yet — clear the color filter before saving.',
        )
      }
      if (!criteria.filters && !criteria.customFilters) continue
      columns.push({
        colId: column - range.startColumn,
        ...(criteria.filters?.filters ? { values: [...criteria.filters.filters] } : {}),
        ...(criteria.filters?.blank ? { blank: true } : {}),
        ...(criteria.customFilters
          ? {
              customs: {
                ...(criteria.customFilters.and === 1 ? { and: true } : {}),
                filters: criteria.customFilters.customFilters.map((custom) => ({
                  val: custom.val,
                  ...(custom.operator !== undefined ? { operator: custom.operator } : {}),
                })),
              },
            }
          : {}),
      })
    }
    const visibilityRange = origin
      ? {
          startRow: Math.min(range.startRow, origin.startRow),
          startColumn: Math.min(range.startColumn, origin.startColumn),
          endRow: Math.max(range.endRow, origin.endRow),
          endColumn: Math.max(range.endColumn, origin.endColumn),
        }
      : range
    states.push({
      sheetName,
      filter: { range, columns },
      hiddenRows: filter.getFilteredOutRows(),
      visibilityRange,
    })
  }
  return states
}

function subscribeToCellMutations(
  runtime: BrowserUniverRuntime,
  dirtyRef: React.MutableRefObject<Map<string, CellEdit>>,
  structuralRef: React.MutableRefObject<Map<string, JournaledStructuralOp[]>>,
  freezeRef: React.MutableRefObject<Map<string, { frozenRows: number; frozenColumns: number }>>,
  suppressionRef: React.MutableRefObject<boolean>,
  filterDirtyRef: React.MutableRefObject<Set<string>>,
  filterOriginsRef: React.MutableRefObject<
    Map<string, { startRow: number; endRow: number; startColumn: number; endColumn: number }>
  >,
  onDirty: () => void,
): { dispose(): void } {
  const sub = runtime.univerAPI.addEvent(runtime.univerAPI.Event.CommandExecuted, (event) => {
    // Load-time suppression (desktop journalSuppression parity): a load
    // is not an edit — skip journaling entirely while a snapshot load is
    // in progress.
    if (suppressionRef.current) return
    // The formula engine re-applies calculation results with these
    // execution options; they are derived state, never user edits
    // (desktop App.tsx parity — the same filter keeps load-time and
    // edit-time recalc echoes out of the journal). The sort's
    // FormulaReorderController mutations carry NO options, so the
    // Excel-style formula rewrites still journal.
    const options = (event as { options?: { fromFormula?: boolean } }).options
    if (options?.fromFormula) return
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

    // ── Number-format mutation (sheets-numfmt preset). The numfmt facade
    //    mixin (.n(pattern) on FRange) fires this mutation with params
    //    carrying { values: { [id]: { ranges } }, refMap: { [id]: { pattern } } }.
    //    We expand it into per-cell style-only CellEdits with
    //    style.numberFormat = pattern, which the canonical WorkbookStyleEdit
    //    persists through applyCellEditsToXlsx (xlsx-styles.ts). Without
    //    this handler, a number-format change is in-session only.
    if (event.id === SET_NUMFMT_MUTATION_ID) {
      const params = event.params as
        { subUnitId?: string; values?: unknown; refMap?: unknown; unitId?: string } | undefined
      if (!params?.subUnitId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      const edits = numfmtEditsFromMutation(sheetName, params)
      if (edits.length === 0) return
      for (const { row, column, edit } of edits) {
        const key = dirtyKey(sheetName, row, column)
        const existing = dirtyRef.current.get(key)
        // Wrap as a ParsedMutation so mergeCellEdit applies the formula-
        // priority merge rule (style-only edits never overwrite a
        // journaled formula; they only merge their style fields).
        dirtyRef.current.set(key, mergeCellEdit(existing, { edit, replacesFormula: false }))
      }
      onDirty()
      return
    }

    // ── Sort / reorder-range mutation (sheets-sort preset). Univer's sort
    //    command (FRange.sort, the public facade mixin from
    //    @univerjs/sheets-sort) fires sheet.command.sort-range →
    //    ReorderRangeMutation (sheet.mutation.reorder-range), which
    //    deepClones the entire cell record (v/f/s/p/si/t) via getCellRaw
    //    and writes it into the worksheet cellDataMatrix in-memory. It
    //    does NOT dispatch a separate set-range-values mutation, so the
    //    set-range-values journal above would miss sort.
    //
    //    The journal captures the row permutation directly as a
    //    `reorder-rows` structural op — `{ range, order }`. NOTE: Univer's
    //    order map is DEST→SRC (NEW[destRow] = OLD[order[destRow]]; the
    //    mutation source reads `getCellRaw(order[row])` and writes it to
    //    `row`) — journaled verbatim, NOT inverted here. The gateway
    //    (transformSheetRowsByPermutation) inverts it internally before
    //    renumbering <row> blocks. On save, the canonical
    //    applyStructuralOps path permutes <row> blocks atomically: the r=
    //    attributes on <row> and inner <c> renumber, but the cell
    //    contents (value, formula text, style ref, hyperlink rich-text,
    //    shared-formula si=, comment pointer) travel UNTOUCHED inside
    //    their <c> elements. This mirrors Univer's deepClone exactly —
    //    styles, numfmt, fills, borders, hyperlinks, and any other cell
    //    metadata survive save/reopen.
    //
    //    Excel does not rewrite external formula references for sort
    //    (formulas recalculated against current cell positions on
    //    reopen), so the gateway skips transformFormulas and
    //    transformRangedFeatures for this op. Formula text inside the
    //    sorted range travels verbatim — Univer's live state has the
    //    same verbatim formula text, so the saved XLSX matches Univer's
    //    pre-save state and Excel recalculates the same displayed values.
    if (event.id === REORDER_RANGE_MUTATION_ID) {
      const params = event.params as
        | {
            subUnitId?: string
            range?: { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number }
            order?: Record<string, number>
            unitId?: string
          }
        | undefined
      if (!params?.subUnitId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      const range = params.range
      if (
        !range ||
        !Number.isInteger(range.startRow) ||
        !Number.isInteger(range.endRow) ||
        !Number.isInteger(range.startColumn) ||
        !Number.isInteger(range.endColumn)
      ) {
        return
      }
      if (
        range.startRow! < 0 ||
        range.endRow! < range.startRow! ||
        range.startColumn! < 0 ||
        range.endColumn! < range.startColumn!
      ) {
        return
      }
      const orderMap = params.order
      if (!orderMap || typeof orderMap !== 'object') return
      // Normalize the order map. Univer's order is DEST→SRC
      // (NEW[destRow] = OLD[order[destRow]] — the mutation source reads
      // getCellRaw(order[row]) and writes it to row). JSON object keys
      // are strings, so normalize to number-keyed records; the gateway
      // inverts the map internally before permuting <row> blocks.
      const order: Record<number, number> = {}
      for (const [k, v] of Object.entries(orderMap)) {
        const dest = Number(k)
        const src = Number(v)
        if (!Number.isInteger(dest) || !Number.isInteger(src)) continue
        order[dest] = src
      }
      if (Object.keys(order).length === 0) return
      const startRow = range.startRow!
      const endRow = range.endRow!
      const startColumn = range.startColumn!
      const endColumn = range.endColumn!
      // ── Rebase previously-journaled cell edits into post-sort
      //    coordinates. The sort moved cells inside the range: an edit
      //    journaled at pre-sort position (srcRow, col) now describes the
      //    cell at (destRow, col), where order[destRow] === srcRow. The
      //    gateway replays structural ops BEFORE cell edits, so without
      //    rebasing the pre-sort edits would land on the wrong rows (the
      //    e2e "type then sort" scenario). Rebasing is a permutation of
      //    the dirty keys (the order map is a bijection over the
      //    participating rows), so no collisions are possible.
      const srcToDest = new Map<number, number>()
      for (const [destKey, srcRow] of Object.entries(order)) {
        srcToDest.set(srcRow, Number(destKey))
      }
      const rebasedDirty = new Map<string, CellEdit>()
      for (const [key, edit] of dirtyRef.current) {
        // key = `${sheetName}:${row}:${column}` — parse from the right so
        // sheet names containing colons could never confuse the split
        // (OOXML forbids colons in sheet names, but be robust anyway).
        const lastColon = key.lastIndexOf(':')
        const secondLastColon = key.lastIndexOf(':', lastColon - 1)
        const editSheet = key.slice(0, secondLastColon)
        const editRow = Number(key.slice(secondLastColon + 1, lastColon))
        const editColumn = Number(key.slice(lastColon + 1))
        if (
          editSheet === sheetName &&
          editRow >= startRow &&
          editRow <= endRow &&
          editColumn >= startColumn &&
          editColumn <= endColumn &&
          srcToDest.has(editRow)
        ) {
          const destRow = srcToDest.get(editRow)!
          rebasedDirty.set(dirtyKey(sheetName, destRow, editColumn), {
            ...edit,
            row: destRow,
          })
        } else {
          rebasedDirty.set(key, edit)
        }
      }
      dirtyRef.current = rebasedDirty
      const ops = structuralRef.current.get(sheetName) ?? []
      ops.push({
        kind: 'reorder-rows',
        index: 0,
        count: 1,
        range: { startRow, endRow, startColumn, endColumn },
        order,
      })
      structuralRef.current.set(sheetName, ops)
      onDirty()
      return
    }

    // ── Merge mutations (built-in sheets). The AddWorksheetMergeMutation
    //    fires when range.merge() is called; RemoveWorksheetMergeMutation
    //    fires when range.breakApart() is called. Both carry params.ranges
    //    (an array of IRange). We journal each range as a merge-cells or
    //    unmerge-cells structural op — the canonical SheetStructuralOps
    //    family in applyCellEditsToXlsx writes the <mergeCells> entries.
    if (event.id === ADD_MERGE_MUTATION_ID || event.id === REMOVE_MERGE_MUTATION_ID) {
      const params = event.params as
        { subUnitId?: string; ranges?: unknown; unitId?: string } | undefined
      if (!params?.subUnitId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      const ranges = Array.isArray(params.ranges) ? params.ranges : []
      const kind = event.id === ADD_MERGE_MUTATION_ID ? 'merge-cells' : 'unmerge-cells'
      const ops = structuralRef.current.get(sheetName) ?? []
      for (const r of ranges) {
        if (typeof r !== 'object' || r === null) continue
        const range = r as {
          startRow?: number
          endRow?: number
          startColumn?: number
          endColumn?: number
        }
        const startRow = Number.isInteger(range.startRow) ? (range.startRow as number) : -1
        const endRow = Number.isInteger(range.endRow) ? (range.endRow as number) : -1
        const startColumn = Number.isInteger(range.startColumn) ? (range.startColumn as number) : -1
        const endColumn = Number.isInteger(range.endColumn) ? (range.endColumn as number) : -1
        if (startRow < 0 || endRow < 0 || startColumn < 0 || endColumn < 0) continue
        ops.push({
          kind,
          index: 0,
          count: 1,
          range: { startRow, endRow, startColumn, endColumn },
        })
      }
      structuralRef.current.set(sheetName, ops)
      onDirty()
      return
    }

    // ── Filter mutations (sheets-filter preset). set-filter-range carries
    //    the new filter's range; the other three (set-filter-criteria,
    //    remove-filter, re-calc-filter) change an existing filter. All mark
    //    the sheet filter-dirty — the SAVE snapshots the live filter model
    //    declaratively (collectFilterStates below), exactly like the
    //    desktop. For set-filter-range, the origin union extends to cover
    //    the new range so a later move/remove still unhides this span.
    if (FILTER_MUTATION_IDS.has(event.id)) {
      const params = event.params as
        | {
            subUnitId?: string
            unitId?: string
            range?: { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number }
          }
        | undefined
      if (!params?.subUnitId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      filterDirtyRef.current.add(sheetName)
      const range = params.range
      if (
        range &&
        Number.isInteger(range.startRow) &&
        Number.isInteger(range.endRow) &&
        Number.isInteger(range.startColumn) &&
        Number.isInteger(range.endColumn) &&
        range.startRow! >= 0 &&
        range.endRow! >= range.startRow! &&
        range.startColumn! >= 0 &&
        range.endColumn! >= range.startColumn!
      ) {
        const origin = filterOriginsRef.current.get(sheetName)
        filterOriginsRef.current.set(sheetName, {
          startRow: Math.min(origin?.startRow ?? range.startRow!, range.startRow!),
          endRow: Math.max(origin?.endRow ?? range.endRow!, range.endRow!),
          startColumn: Math.min(origin?.startColumn ?? range.startColumn!, range.startColumn!),
          endColumn: Math.max(origin?.endColumn ?? range.endColumn!, range.endColumn!),
        })
      }
      onDirty()
      return
    }

    // ── Freeze-pane mutation (built-in sheets). sheet.mutation.set-frozen
    //    carries { unitId, subUnitId, startRow, startColumn, xSplit, ySplit }.
    //    startRow/startColumn are the first scrollable (non-frozen) row/column
    //    index — so frozenRows = startRow, frozenColumns = startColumn. A
    //    value of -1 on either axis means "no freeze on that axis". The
    //    journal stores per-sheet { frozenRows, frozenColumns }; on save,
    //    ExcelEditor emits them as BrowserSheetPageSetupState, which the
    //    canonical applyPageSetupState persists as the <pane> element.
    if (event.id === SET_FROZEN_MUTATION_ID) {
      const params = event.params as
        | {
            subUnitId?: string
            startRow?: number
            startColumn?: number
            xSplit?: number
            ySplit?: number
            unitId?: string
          }
        | undefined
      if (!params?.subUnitId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      const sheetName = ws.getSheetName()
      const startRow = Number.isInteger(params.startRow) ? (params.startRow as number) : -1
      const startColumn = Number.isInteger(params.startColumn) ? (params.startColumn as number) : -1
      const frozenRows = startRow > 0 ? startRow : 0
      const frozenColumns = startColumn > 0 ? startColumn : 0
      if (frozenRows === 0 && frozenColumns === 0) {
        // Freeze cleared — drop the journal entry so the save plan doesn't
        // re-emit a zero freeze (which applyPageSetupState treats as a no-op
        // anyway, but a clean journal keeps the save plan minimal).
        freezeRef.current.delete(sheetName)
      } else {
        freezeRef.current.set(sheetName, { frozenRows, frozenColumns })
      }
      onDirty()
      return
    }

    if (event.id !== SET_RANGE_VALUES_MUTATION_ID) return
    const params = event.params as
      { subUnitId?: string; cellValue?: unknown; unitId?: string } | undefined
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
