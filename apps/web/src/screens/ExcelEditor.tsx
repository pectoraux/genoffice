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
// Same pattern for the table facade (EXCEL-021): surfaces addTable /
// removeTable / addTableTheme on FWorksheet. The runtime side is wired by
// UniverSheetsTablePreset in create-browser-univer.ts (already shipped);
// this import only brings the TypeScript signatures into scope. The
// registered table is VISUAL (filter dropdowns + theme); the canonical
// persistence is the tableAdditions save family, never this model.
import '@univerjs/sheets-table/facade'
import type { ICustomFilter, IFilterColumn } from '@univerjs/sheets-filter'
import type {
  CellEdit,
  CellFormatState,
  CellState,
  DvWireRule,
  FilterColumnState,
  SheetFilterState,
  SheetNote,
  SheetTableAddition,
  SheetTableInfo,
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
import { applyTableBandingToMatrix, type TableBandingMatrix } from '../office/table-banding'
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

/**
 * Data-validation mutation IDs (data-validation preset) — the same set the
 * desktop's App.tsx listens for. Any of these marks the sheet DV-dirty; the
 * save snapshots the LIVE validation model declaratively (collectDvStates),
 * never replaying individual UI commands.
 */
const DV_MUTATION_IDS = new Set([
  'data-validation.mutation.addRule',
  'data-validation.mutation.updateRule',
  'data-validation.mutation.removeRule',
])

/**
 * Note mutation IDs (sheets-note preset) — the same set the desktop's App.tsx
 * listens for. update-note fires for create AND edit; remove-note for delete.
 * (toggle-note-popup / update-note-position are popup/geometry chrome that
 * never change the persisted comment set.) The save snapshots the LIVE note
 * model declaratively (collectNoteStates), never replaying mutations.
 */
const NOTE_MUTATION_IDS = new Set(['sheet.mutation.update-note', 'sheet.mutation.remove-note'])

/**
 * Filter PANEL commands (EXCEL-021 — desktop App.tsx FILTER_COMMAND_PATTERN
 * parity). These are the command IDs the filter toolbar emits BEFORE the
 * mutations land; a sheet whose filter origin is an Excel TABLE must refuse
 * them up front — the table part owns its filter, and editing it through the
 * worksheet filter UI cannot be persisted yet (desktop message verbatim).
 */
const FILTER_COMMAND_PATTERN =
  /^sheet\.command\.(set-filter-criteria|set-filter-range|smart-toggle-filter|clear-filter-criteria|remove-sheet-filter|re-calc-filter)$/

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
  // ── Data-validation journal (Data → Data Validation). Desktop parity
  //    (App.tsx dvDirty): sheet NAMES whose validation changed in-session.
  //    A sheet absent from this set is NOT DV-dirty: its save plan carries
  //    NO dv state, so a no-op save preserves the file's own
  //    <dataValidations> XML byte-for-byte. At save, the sheet's FULL live
  //    rule set is snapshotted (created + modified + untouched rules alike)
  //    — editing one rule never drops its neighbors.
  const dvDirtyRef = useRef<Set<string>>(new Set())
  // ── Note journal (Review → New Comment). Desktop parity (App.tsx
  //    noteDirty): sheet NAMES whose notes changed in-session. A sheet absent
  //    from this set is NOT note-dirty: its save plan carries NO note state,
  //    so a no-op save preserves the file's own comments part byte-for-byte.
  //    At save, the sheet's FULL live note set is snapshotted (created +
  //    edited + untouched notes alike) — editing one note never drops its
  //    neighbors; an empty snapshot removes the whole comment set.
  const noteDirtyRef = useRef<Set<string>>(new Set())
  // ── Protection journal (Review → Protect Sheet / Protect Workbook,
  //    EXCEL-020). Desktop parity (App.tsx sheetProtections +
  //    edit-journal.ts recordSheetProtection): protection is a FILE-level
  //    journal concern — the editor itself does not enforce it (the
  //    desktop's own status string). Two structures per level:
  //    - file refs: the state as opened from the snapshot (seeded on open,
  //      merged after save). hasPassword drives the fail-closed toggle
  //      guard — the gateway refuses to unprotect password-bearing
  //      elements, so the browser refuses up front and says why.
  //    - journal refs: desired states for sheets the user toggled
  //      in-session. An entry is DROPPED when toggled back to the file's
  //      original state (recordSheetProtection semantics), so a no-op save
  //      emits NO protection family and preserves the file's XML
  //      byte-for-byte.
  const sheetProtectionFileRef = useRef<Map<string, { protected: boolean; hasPassword: boolean }>>(
    new Map(),
  )
  const sheetProtectionJournalRef = useRef<Map<string, boolean>>(new Map())
  const workbookProtectionFileRef = useRef<{
    lockStructure: boolean
    hasPassword: boolean
  } | null>(null)
  const workbookProtectionJournalRef = useRef<boolean | null>(null)
  // ── Table state (EXCEL-021, Insert → Table / Delete Table). Desktop
  //    parity (App.tsx file.tables + edit-journal.ts tableAdds). Two
  //    structures:
  //    - tablesFileRef: the FILE's own tables per sheet NAME (SheetTableInfo
  //      from the gateway reader — metadata + pre-resolved banding colors).
  //      Seeded on open, merged after save; drives the banding paint, the
  //      Univer registration, the file-native delete refusal, and the
  //      table-owned filter origin.
  //    - tableAddsRef: the session journal (SheetTableAddition — the exact
  //      wire shape). Deleting a session table SPLICES its entry
  //      (convert-to-range: the baked cells stay, nothing reaches the
  //      file), so an unsaved table never persists.
  //    - tableUniverIdsRef: bookkeeping for the visual Univer registration
  //      (name → tableId) so Delete Table can removeTable() the right unit.
  //    - tableFilterOriginRef: sheet NAMES whose filter belongs to a table
  //      (the worksheet carries no <autoFilter> — the filter lives in the
  //      table part). Filter commands on such sheets are refused up front
  //      (BeforeCommandExecute gate, desktop appTableFilterNoEdit parity).
  const tablesFileRef = useRef<Map<string, readonly SheetTableInfo[]>>(new Map())
  const tableAddsRef = useRef<SheetTableAddition[]>([])
  const tableUniverIdsRef = useRef<Map<string, string>>(new Map())
  const tableFilterOriginRef = useRef<Set<string>>(new Set())
  // Echo for the ribbon's Protect Sheet / Protect Workbook buttons —
  // recomputed whenever the journal, the file state, or the ACTIVE SHEET
  // changes (the runtime's ActiveSheetChanged subscription re-renders the
  // editor, and the echo is derived at render time below).
  const [protectionSeq, setProtectionSeq] = useState(0)
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
      dvDirtyRef,
      noteDirtyRef,
      () => setDirty(true),
    )
    // ── EXCEL-021: table-owned filter gate (desktop App.tsx
    //    FILTER_COMMAND_PATTERN parity). A sheet whose filter belongs to an
    //    Excel TABLE (the worksheet has no <autoFilter> — the filter lives
    //    in the table part) refuses every filter command BEFORE it runs:
    //    editing it through the worksheet filter UI cannot be saved yet,
    //    and letting it through would journal an autoFilter the save would
    //    then write OUTSIDE the table — corrupting the file's semantics.
    //    The gate mirrors the desktop verbatim, message included.
    const filterGate = rt.univerAPI.addEvent(rt.univerAPI.Event.BeforeCommandExecute, (event) => {
      if (!FILTER_COMMAND_PATTERN.test(event.id)) return
      const params = event.params as { subUnitId?: string } | undefined
      const subUnitId =
        params?.subUnitId ?? rt.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
      if (subUnitId === undefined) return
      const wb = rt.univerAPI.getActiveWorkbook()
      const sheetName = wb?.getSheetBySheetId(subUnitId)?.getSheetName()
      if (sheetName === undefined) return
      if (!tableFilterOriginRef.current.has(sheetName)) return
      ;(event as { cancel?: boolean }).cancel = true
      setStatus("This sheet's filter belongs to an Excel table — editing it cannot be saved yet.")
    })
    const w = window as { __genofficeExcelRuntime?: unknown }
    w.__genofficeExcelRuntime = rt
    return () => {
      filterGate.dispose()
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

  // ── EXCEL-020: Protection echo (ribbon button state). Derived at render
  //    time from the file refs + journal + the LIVE active sheet name, so it
  //    tracks sheet switches (the runtime's ActiveSheetChanged subscription
  //    re-renders) as well as toggles and saves (protectionSeq bump).
  //    protectionSeq is read so ESLint's exhaustive-deps-style consumers see
  //    the dependency; the value itself is derived from refs.
  void protectionSeq
  const activeSheetName =
    runtimeRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetName() ?? null
  const hasFile = handleRef.current !== null
  const sheetEcho = (() => {
    if (!activeSheetName) return null
    const journaled = sheetProtectionJournalRef.current.get(activeSheetName)
    if (journaled !== undefined) return journaled
    const file = sheetProtectionFileRef.current.get(activeSheetName)
    if (file) return file.protected
    // A sheet that exists in the live workbook but not in the file map
    // (e.g. added in-session) starts unprotected — desktop parity for
    // added sheets (sheetProtectionEcho's editJournal.sheets.added branch).
    return false
  })()
  const sheetEchoHasPassword =
    activeSheetName !== null
      ? (sheetProtectionFileRef.current.get(activeSheetName)?.hasPassword ?? false)
      : false
  const workbookEcho = (() => {
    if (!hasFile) return null
    if (workbookProtectionJournalRef.current !== null) return workbookProtectionJournalRef.current
    return workbookProtectionFileRef.current?.lockStructure ?? false
  })()
  const workbookEchoHasPassword = workbookProtectionFileRef.current?.hasPassword ?? false

  const bumpProtectionEcho = useCallback(() => setProtectionSeq((n) => n + 1), [])

  /**
   * Review → Protect Sheet / Unprotect Sheet (EXCEL-020). Desktop parity
   * with ribbon-actions.ts 'sheet-protect': journal-only semantics —
   * recordSheetProtection DROPS the entry when the desired state matches
   * the file's original, refuses password-protected sheets up front
   * (fail-closed — the gateway would reject the unprotect anyway), and
   * surfaces the desktop's own status strings. The editor itself does not
   * enforce protection; the toggle writes the canonical sheetProtections
   * family on save.
   */
  const toggleSheetProtection = useCallback(() => {
    const rt = runtimeRef.current
    const wb = rt?.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    const sheetName = ws?.getSheetName()
    if (!handleRef.current || !sheetName) {
      setStatus('Open an XLSX file first — protection saves into the file.')
      return
    }
    const file = sheetProtectionFileRef.current.get(sheetName)
    const original = file?.protected ?? false
    const current = sheetProtectionJournalRef.current.get(sheetName) ?? original
    if (current && (file?.hasPassword ?? false)) {
      setStatus(
        'This sheet is protected with a password — removing its protection is not supported.',
      )
      return
    }
    const desired = !current
    if (desired === original) sheetProtectionJournalRef.current.delete(sheetName)
    else sheetProtectionJournalRef.current.set(sheetName, desired)
    setDirty(true)
    bumpProtectionEcho()
    setStatus(
      desired
        ? 'Sheet protection will be written on save (no password). The editor itself does not enforce it.'
        : 'Sheet protection will be removed on save.',
    )
  }, [bumpProtectionEcho])

  /**
   * Review → Protect Workbook / Unprotect Workbook (EXCEL-020). Desktop
   * parity with ribbon-actions.ts 'workbook-protect': journal-only
   * semantics for the workbook STRUCTURE lock; refuses a password-
   * protected structure up front (fail-closed).
   */
  const toggleWorkbookProtection = useCallback(() => {
    if (!handleRef.current) {
      setStatus('Open an XLSX file first — protection saves into the file.')
      return
    }
    const file = workbookProtectionFileRef.current
    if (file?.hasPassword) {
      setStatus('The workbook structure is password-protected — it cannot be changed here.')
      return
    }
    const original = file?.lockStructure ?? false
    const current = workbookProtectionJournalRef.current ?? original
    const desired = !current
    workbookProtectionJournalRef.current = desired === original ? null : desired
    setDirty(true)
    bumpProtectionEcho()
    setStatus(
      desired
        ? 'Workbook structure protection will be written on save.'
        : 'Workbook structure protection will be removed on save.',
    )
  }, [bumpProtectionEcho])

  /**
   * Review → Lock Cell / Unlock Cell (EXCEL-020 — "editable vs locked cell
   * behavior"). Journals canonical WorkbookStyleEdit.protectionLocked
   * deltas for every cell of the active selection — the SAME neutral-delta
   * path the desktop's Format Cells → Protection tab uses ("No Univer
   * model for cell protection — journal the neutral delta"). Univer has no
   * cell-protection model in the OSS presets, so nothing fires through the
   * mutation subscription: the edits are journaled directly and merged
   * with any existing per-cell entries via mergeCellEdit (value edits keep
   * their cell content; the protection flag rides along in `style`). The
   * gateway's buildProtection writes `<protection locked="0|1"/>` into the
   * cell's xf — together with a protected sheet that IS Excel's
   * editable-vs-locked semantics in the saved file.
   */
  const setCellsLocked = useCallback((locked: boolean) => {
    const rt = runtimeRef.current
    const wb = rt?.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    if (!wb || !ws) return
    if (!handleRef.current) {
      setStatus('Open an XLSX file first — protection saves into the file.')
      return
    }
    const sheetName = ws.getSheetName()
    const range = ws.getActiveRange()
    if (!range) {
      setStatus('Select the cells to lock or unlock first.')
      return
    }
    const startRow = Math.max(0, range.getRow())
    const startColumn = Math.max(0, range.getColumn())
    // FRange dimensions: getHeight() = row count, getWidth() = column
    // count (f-range.d.ts) — the facade's own selection-extent API.
    const numRows = range.getHeight()
    const numColumns = range.getWidth()
    let touched = false
    for (let row = startRow; row < startRow + numRows; row += 1) {
      for (let column = startColumn; column < startColumn + numColumns; column += 1) {
        const key = dirtyKey(sheetName, row, column)
        const edit: CellEdit = {
          sheetName,
          row,
          column,
          writeValue: false,
          cell: { value: null },
          style: { protectionLocked: locked },
        }
        const existing = dirtyCellsRef.current.get(key)
        dirtyCellsRef.current.set(key, mergeCellEdit(existing, { edit, replacesFormula: false }))
        touched = true
      }
    }
    if (touched) {
      setDirty(true)
      setStatus(
        locked
          ? 'Cells locked — they will be read-only when the sheet is protected.'
          : 'Cells unlocked — they stay editable when the sheet is protected.',
      )
    }
  }, [])

  /**
   * EXCEL-021: does the rectangle overlap any table area on the sheet?
   * Used by Insert → Table to fail closed up front — the gateway re-checks
   * at save time (against the file's own parts), but refusing before the
   * journal entry exists gives the desktop's clear status message instead
   * of a save-time error.
   */
  const tableOverlaps = useCallback(
    (sheetName: string, area: SheetTableAddition['area']): string | null => {
      const apart = (other: {
        startRow: number
        endRow: number
        startColumn: number
        endColumn: number
      }) =>
        area.endRow < other.startRow ||
        other.endRow < area.startRow ||
        area.endColumn < other.startColumn ||
        other.endColumn < area.startColumn
      for (const table of tableAddsRef.current) {
        if (table.sheetName !== sheetName) continue
        if (!apart(table.area)) return table.name
      }
      const fileTables = tablesFileRef.current.get(sheetName) ?? []
      for (const table of fileTables) {
        if (!apart(table.area)) return table.name ?? ''
      }
      return null
    },
    [],
  )

  /**
   * Insert → Table (EXCEL-021). Browser port of the desktop's
   * handleFormatAsTable → applyAiTableAdd: validate the active range (a
   * header row plus at least one data row, at most 1,000 columns), name it
   * TableN skipping names the session used, sanitize the column names from
   * the header row (trim, 255 cap, blank → ColumnN, dedupe with a numeric
   * suffix — corrected values are written back through the REAL facade so
   * they journal as cell edits), register the visual Univer table with its
   * DEFAULT theme (the desktop does NOT mute created tables — Univer's own
   * table styling is the immediate feedback), and journal the canonical
   * SheetTableAddition. Status strings are the desktop's own English.
   */
  const handleInsertTable = useCallback(() => {
    const rt = runtimeRef.current
    const wb = rt?.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    if (!handleRef.current || !wb || !ws) {
      setStatus('Open an XLSX file first — tables are written into the file.')
      return
    }
    const range = ws.getActiveRange()
    if (!range) {
      setStatus('Select the data range first — headers in its first row.')
      return
    }
    const sheetName = ws.getSheetName()
    const startRow = range.getRow()
    const startColumn = range.getColumn()
    const endRow = startRow + range.getHeight() - 1
    const endColumn = startColumn + range.getWidth() - 1
    if (endRow <= startRow) {
      setStatus('A table needs a header row plus at least one data row.')
      return
    }
    const width = endColumn - startColumn + 1
    if (width > 1_000) {
      setStatus('A table can span at most 1,000 columns.')
      return
    }
    // Default session names: Table1, Table2, … skipping names the session
    // already used (desktop nextSessionTableName). Collisions with the
    // file's own tables are refused up front here (the gateway re-checks
    // at save time) — clear message instead of a failed save.
    const taken = new Set(tableAddsRef.current.map((table) => table.name.toLowerCase()))
    for (const fileTables of tablesFileRef.current.values()) {
      for (const table of fileTables) {
        if (table.name !== undefined) taken.add(table.name.toLowerCase())
      }
    }
    let index = tableAddsRef.current.length + 1
    while (taken.has(`table${index}`)) index += 1
    const name = `Table${index}`
    // Overlap check: session tables AND the file's own tables on this
    // sheet (the gateway re-checks at save time against the file parts).
    const overlaps = tableOverlaps(sheetName, { startRow, startColumn, endRow, endColumn })
    if (overlaps !== null) {
      setStatus(`The range overlaps table "${overlaps}" created this session.`)
      return
    }
    // Sanitize the column names from the header row (desktop parity):
    // trim, 255 cap, blank → ColumnN, dedupe with a numeric suffix; the
    // corrected values are written back through the REAL facade so they
    // journal as canonical cell edits and save with the table.
    const headerValues = range.getValues()[0] ?? []
    const columnNames: string[] = []
    const used = new Set<string>()
    for (let offset = 0; offset < width; offset += 1) {
      const raw = String(headerValues[offset] ?? '')
        .trim()
        .slice(0, 255)
      const base = raw.length === 0 ? `Column${offset + 1}` : raw
      let candidate = base
      for (let suffix = 2; used.has(candidate.toLowerCase()); suffix += 1) {
        candidate = `${base}${suffix}`
      }
      used.add(candidate.toLowerCase())
      columnNames.push(candidate)
      if (candidate !== raw) {
        ws.getRange(startRow, startColumn + offset).setValue(candidate)
      }
    }
    // Visual registration — best-effort (the journal entry below is what
    // the save writes, and the gateway re-checks conflicts). The DEFAULT
    // theme stays: created tables take Univer's own styling (the desktop
    // deliberately does not mute them — instant visual feedback).
    const tableId = `ai-table-${tableAddsRef.current.length + 1}-${Date.now().toString(36)}`
    try {
      void ws.addTable(name, { startRow, startColumn, endRow, endColumn }, tableId)
      tableUniverIdsRef.current.set(`${sheetName}::${name}`, tableId)
    } catch {
      // Rendering is best-effort; the journal entry is the source of truth.
    }
    tableAddsRef.current.push({
      sheetName,
      area: { startRow, startColumn, endRow, endColumn },
      name,
      columnNames,
      style: 'TableStyleMedium2',
      bandedRows: true,
    })
    setDirty(true)
    setStatus('Table created — save with ⌘S.')
  }, [tableOverlaps])

  /**
   * Insert → Delete Table (EXCEL-021 — the desktop has no table-delete
   * ribbon button; the web adds it for the required delete verification).
   * Convert-to-range semantics, desktop removeTableAdd parity: ONLY a
   * table created THIS SESSION can be deleted — the journal entry is
   * spliced (nothing reaches the file; the baked cells stay) and the
   * visual Univer registration is removed. A file-native table under the
   * active cell is REFUSED with the desktop's exact message; so is an
   * empty selection.
   */
  const handleDeleteTable = useCallback(() => {
    const rt = runtimeRef.current
    const wb = rt?.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    if (!handleRef.current || !wb || !ws) {
      setStatus('Open an XLSX file first — tables are written into the file.')
      return
    }
    const sheetName = ws.getSheetName()
    const cell = ws.getActiveRange()
    const row = cell?.getRow() ?? -1
    const column = cell?.getColumn() ?? -1
    const contains = (area: {
      startRow: number
      endRow: number
      startColumn: number
      endColumn: number
    }) =>
      row >= area.startRow &&
      row <= area.endRow &&
      column >= area.startColumn &&
      column <= area.endColumn
    // A session table under the active cell → convert-to-range.
    const sessionIndex = tableAddsRef.current.findIndex(
      (table) => table.sheetName === sheetName && contains(table.area),
    )
    if (sessionIndex >= 0) {
      const entry = tableAddsRef.current[sessionIndex]!
      tableAddsRef.current.splice(sessionIndex, 1)
      const tableId = tableUniverIdsRef.current.get(`${sheetName}::${entry.name}`)
      if (tableId !== undefined) {
        tableUniverIdsRef.current.delete(`${sheetName}::${entry.name}`)
        try {
          void ws.removeTable(tableId)
        } catch {
          // The registration is visual; the journal splice already won.
        }
      }
      setDirty(true)
      setStatus(`Table "${entry.name}" removed — the cells stay as they are.`)
      return
    }
    // A file-native table under the active cell (or no table at all) → the
    // desktop's refusal, verbatim (removeTableAdd returns false).
    const fileTables = tablesFileRef.current.get(sheetName) ?? []
    const fileHit = fileTables.find((table) => contains(table.area))
    const refusedName = fileHit?.name ?? ''
    setStatus(
      `Table "${refusedName}" does not exist or was not created this session — ` +
        'tables already in the file cannot be deleted yet.',
    )
  }, [])

  const loadSnapshot = useCallback(
    (snapshot: WorkbookSnapshot) => {
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
      // Same leak guard for the EXCEL-020 protection state: the file refs are
      // re-seeded from THIS snapshot (per-sheet sheetProtection + the
      // workbook-level workbookProtection the gateway reader surfaces), and
      // the toggle journals start empty — a freshly opened workbook carries
      // no protection decisions, so a no-op save emits no protection family
      // and preserves the file's XML byte-for-byte.
      sheetProtectionFileRef.current.clear()
      sheetProtectionJournalRef.current.clear()
      workbookProtectionFileRef.current = snapshot.workbookProtection ?? null
      workbookProtectionJournalRef.current = null
      for (const sheet of snapshot.sheets) {
        if (sheet.sheetProtection) {
          sheetProtectionFileRef.current.set(sheet.name, sheet.sheetProtection)
        }
      }
      bumpProtectionEcho()
      // EXCEL-021 table state: same leak guard. The file refs re-seed from
      // THIS snapshot (per-sheet tables with pre-resolved banding colors),
      // the session journal starts empty (a freshly opened workbook saves
      // NO tableAdditions — a no-op save preserves the table parts
      // byte-for-byte), and the visual-registration bookkeeping resets.
      tablesFileRef.current.clear()
      tableAddsRef.current = []
      tableUniverIdsRef.current.clear()
      tableFilterOriginRef.current.clear()
      for (const sheet of snapshot.sheets) {
        if (sheet.tables && sheet.tables.length > 0) {
          tablesFileRef.current.set(sheet.name, sheet.tables)
        }
      }
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
          // EXCEL-021: table banding is painted INTO the cell matrix before
          // the workbook is created (desktop applyTableBanding parity — the
          // gateway already resolved every band color from the file's theme
          // accents / custom tableStyle dxfs, so this stays a pure value
          // transform; explicit cell fills always WIN over banding).
          cellData: (() => {
            const matrix = buildCellDataMatrix(sheet.cells, sheet.styles) as TableBandingMatrix
            const tables = tablesFileRef.current.get(sheet.name)
            if (tables && tables.length > 0) {
              applyTableBandingToMatrix(matrix, tables)
            }
            return matrix as IObjectMatrixPrimitiveType<ICellData>
          })(),
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
      // save preserves the file's own <autoFilter> XML). Same for DV.
      filterDirtyRef.current.clear()
      filterOriginsRef.current.clear()
      dvDirtyRef.current.clear()
      noteDirtyRef.current.clear()
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
          // Install the file's data-validation rules into the REAL Univer
          // model (desktop applyDataValidations parity): each DvWireRule
          // becomes a data-validation.mutation.addRule execution under
          // journal suppression — loading a workbook is not an edit, and the
          // installed rules give real in-cell validation behavior (reject
          // dialogs on invalid input, list dropdowns, prompts). The rule
          // shape is the Univer IDataValidationRule wire form with the
          // desktop's toUniverDvRule transforms: list literals unquote
          // ("a,b" → a,b) while references/custom gain a leading '='.
          if (sheet.dvRules && sheet.dvRules.length > 0) {
            try {
              const wb = rt.univerAPI.getActiveWorkbook()
              const ws = wb?.getSheetByName(sheet.name)
              if (ws) {
                for (const [index, wire] of sheet.dvRules.entries()) {
                  const mapped = toUniverDvRule(wire, `file-dv-${sheet.id}-${index}`)
                  if (!mapped) continue
                  try {
                    rt.univerAPI.syncExecuteCommand('data-validation.mutation.addRule', {
                      unitId: wb!.getId(),
                      subUnitId: ws.getSheetId(),
                      rule: mapped,
                    })
                  } catch {
                    // An unsupported rule shape must not break the open.
                  }
                }
              }
            } catch {
              // Installing the file's validations must never fail the open.
            }
          }
          // Install the file's legacy notes into the REAL Univer note model
          // (desktop applyWorkbookNotes parity): each SheetNote becomes a
          // createOrUpdateNote under journal suppression — the canonical
          // "Author:\nText" blob convention (re-split at save), the same id
          // scheme, and the Inc-4 undo filter keeps ⌘Z clean. The notes then
          // render through the real Univer note UI (cell markers + popup).
          if (sheet.notes && sheet.notes.length > 0) {
            try {
              const wb = rt.univerAPI.getActiveWorkbook()
              const ws = wb?.getSheetByName(sheet.name)
              if (ws) {
                for (const note of sheet.notes) {
                  try {
                    ws.getRange(note.row, note.column).createOrUpdateNote({
                      id: `note-${sheet.id}-${note.row}-${note.column}`,
                      row: note.row,
                      col: note.column,
                      width: 220,
                      height: 90,
                      note: note.author ? `${note.author}:\n${note.text}` : note.text,
                    })
                  } catch {
                    // Notes are best-effort decoration (desktop parity).
                  }
                }
              }
            } catch {
              // Installing the file's notes must never fail the open.
            }
          }
          // ── EXCEL-021: register the file's tables into the REAL Univer
          //    table model (desktop App.tsx parity) so the grid renders
          //    filter dropdowns over the table range. The registration is
          //    VISUAL-ONLY (the journal stays empty for file tables — the
          //    banding lives in the painted cell fills), so it runs under
          //    journal suppression and every failure is swallowed. Univer's
          //    table header is not optional yet: registering a headerless
          //    table injects synthesized "Column N" labels over the first
          //    data row, so those skip registration (banding still paints).
          //    Univer also paints its own lavender default theme over the
          //    cells — mute it to a plain theme so only the banding shows.
          const fileTables = tablesFileRef.current.get(sheet.name) ?? []
          if (fileTables.length > 0) {
            const wb = rt.univerAPI.getActiveWorkbook()
            const ws = wb?.getSheetByName(sheet.name)
            if (ws) {
              for (let index = 0; index < fileTables.length; index += 1) {
                const table = fileTables[index]!
                if (table.headerRowCount === 0) continue
                const tableId = `file-table-${sheet.id}-${index}`
                const tableName = `Table${index + 1}_${sheet.id.slice(0, 6)}`
                try {
                  const added = ws.addTable(tableName, { ...table.area }, tableId) as unknown
                  tableUniverIdsRef.current.set(
                    `${sheet.name}::${table.name ?? tableName}`,
                    tableId,
                  )
                  // Univer paints its own default table theme over the
                  // cells; the file's real banding is already in the cell
                  // fills, so mute the theme to plain. Best-effort.
                  void (added as Promise<unknown>)?.then?.(() => {
                    try {
                      ;(
                        ws as unknown as {
                          addTableTheme(id: string, theme: { name: string }): unknown
                        }
                      ).addTableTheme(tableId, { name: `plain-${tableId}` })
                    } catch {
                      // Theme muting is cosmetic; the table itself is
                      // registered.
                    }
                  })
                } catch {
                  // Best-effort: skip if Univer rejects (e.g. overlapping
                  // ranges) — the data itself is still usable.
                }
              }
            }
            // ── EXCEL-021: table-owned filter origin (desktop
            //    applySheetFilter parity). Excel allows one filter per
            //    sheet: the worksheet's own <autoFilter> (installed above
            //    from filterState) wins; otherwise the FIRST table's range
            //    IS the sheet's filter. Install it so the dropdowns render,
            //    record the origin (for later unhide semantics), and mark
            //    the sheet's filter origin as table-owned — the
            //    BeforeCommandExecute gate then refuses filter edits with
            //    the desktop's exact message.
            if (!sheet.filterState?.filter && fileTables[0]) {
              const area = fileTables[0].area
              try {
                const wb = rt.univerAPI.getActiveWorkbook()
                const ws = wb?.getSheetByName(sheet.name)
                ws?.getRange(
                  area.startRow,
                  area.startColumn,
                  area.endRow - area.startRow + 1,
                  area.endColumn - area.startColumn + 1,
                ).createFilter()
                filterOriginsRef.current.set(sheet.name, { ...area })
                tableFilterOriginRef.current.add(sheet.name)
              } catch {
                // Installing the table's filter must never fail the open.
              }
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
    },
    [bumpProtectionEcho],
  )

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

  const handleSave = useCallback(
    async (saveAs: boolean) => {
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
        const filterStates = collectFilterStates(
          runtimeRef.current,
          filterDirtyRef,
          filterOriginsRef,
        )
        // Snapshot the LIVE validation model for every DV-dirty sheet
        // (Data → Data Validation). Declarative, desktop collectDvStates
        // parity: the FULL rule set of each dirty sheet — untouched rules
        // ride along, so editing one rule never drops its neighbors. An
        // empty list means all validation on the sheet was cleared.
        const dvStates = collectDvStates(runtimeRef.current, dvDirtyRef)
        // Snapshot the LIVE note model for every note-dirty sheet
        // (Review → New Comment). Declarative, desktop collectNoteStates
        // parity: the FULL note set of each dirty sheet — untouched notes
        // ride along, so editing one note never drops its neighbors. An
        // empty list means all notes on the sheet were cleared.
        const noteStates = collectNoteStates(runtimeRef.current, noteDirtyRef)
        // EXCEL-020: emit the journaled protection decisions (Review →
        // Protect Sheet / Protect Workbook). Desktop save-actions parity —
        // only sheets the user toggled carry a state; an untouched workbook
        // emits NOTHING, so a no-op save preserves the file's protection
        // XML byte-for-byte. The wire validates the typed family and the
        // gateway writes/removes the OOXML elements.
        const sheetProtections = Array.from(sheetProtectionJournalRef.current.entries()).map(
          ([sheetName, protect]) => ({ sheetName, protected: protect }),
        )
        const workbookProtectionState =
          workbookProtectionJournalRef.current !== null
            ? { lockStructure: workbookProtectionJournalRef.current }
            : null
        // EXCEL-021: the session's table creations (Insert → Table). Desktop
        //   toSaveTableAdds parity — the journal is the wire payload. A
        //   workbook without table creations emits NO family, so a no-op
        //   save preserves the file's table parts byte-for-byte.
        const tableAdditions = [...tableAddsRef.current]
        // SPLIT-SAVE (desktop save-actions heldTables parity): the gateway
        // fails closed when a new table rides with row/column changes on
        // its sheet ("A new table cannot be saved together with row/column
        // changes on its sheet — save the table first."). Instead of
        // bouncing the user, hold the tables back: phase 1 saves the
        // structure (and every other family) without them, phase 2 saves
        // the held tables ALONE against the phase-1 bytes.
        const heldTables = structuralOps.length > 0 ? tableAdditions : []
        const splitSave = heldTables.length > 0
        let nextFileName = handle.fileName
        if (saveAs) {
          const newName = window.prompt('Save as:', nextFileName)
          if (!newName) {
            setStatus('Save cancelled')
            return
          }
          nextFileName = newName.endsWith('.xlsx') ? newName : `${newName}.xlsx`
        }
        let savedBytes = await saveWorkbook({
          fileName: nextFileName,
          fileBytes: handle.sourceBytes,
          savePlan: {
            edits,
            ...(structuralOps.length > 0 ? { structuralOps } : {}),
            ...(pageSetupStates.length > 0 ? { pageSetupStates } : {}),
            ...(filterStates.length > 0 ? { filterStates } : {}),
            ...(dvStates.length > 0 ? { dvStates } : {}),
            ...(noteStates.length > 0 ? { noteStates } : {}),
            ...(sheetProtections.length > 0 ? { sheetProtections } : {}),
            ...(workbookProtectionState !== null ? { workbookProtectionState } : {}),
            ...(!splitSave && tableAdditions.length > 0 ? { tableAdditions } : {}),
          },
        })
        if (splitSave) {
          // Phase 2: the held tables alone, against the phase-1 bytes —
          // the same two-phase save the desktop runs when structural ops
          // and new tables collide.
          savedBytes = await saveWorkbook({
            fileName: nextFileName,
            fileBytes: savedBytes,
            savePlan: { edits: [], tableAdditions: heldTables },
          })
        }
        handleRef.current = {
          fileName: nextFileName,
          sourceBytes: savedBytes,
          revision: handle.revision + 1,
        }
        setFileName(nextFileName)
        dirtyCellsRef.current.clear()
        structuralOpsRef.current.clear()
        // A saved filter state is now IN the source bytes — the sheet is no
        // longer filter-dirty (another no-op save must not re-emit it). Same
        // for DV: the snapshot is in the file; a no-op save must preserve it.
        filterDirtyRef.current.clear()
        dvDirtyRef.current.clear()
        noteDirtyRef.current.clear()
        // EXCEL-020: the saved protection state is now IN the source bytes —
        // merge the journal into the file refs (the ribbon echo must reflect
        // the saved state) and clear the journals, so a subsequent no-op
        // save emits no protection family and preserves the file's XML.
        for (const [sheetName, protect] of sheetProtectionJournalRef.current) {
          const prior = sheetProtectionFileRef.current.get(sheetName)
          sheetProtectionFileRef.current.set(sheetName, {
            protected: protect,
            hasPassword: prior?.hasPassword ?? false,
          })
        }
        sheetProtectionJournalRef.current.clear()
        if (workbookProtectionJournalRef.current !== null) {
          workbookProtectionFileRef.current = {
            lockStructure: workbookProtectionJournalRef.current,
            hasPassword: workbookProtectionFileRef.current?.hasPassword ?? false,
          }
        }
        workbookProtectionJournalRef.current = null
        bumpProtectionEcho()
        // EXCEL-021: the saved tables are now IN the source bytes — they are
        // file-native. Merge the journal into the file refs (delete-refusal
        // and overlap checks must see them; the banding/registration for a
        // created table was already painted at create time) and clear the
        // journal, so a subsequent no-op save emits no tableAdditions and
        // preserves the table XML byte-for-byte.
        if (tableAdditions.length > 0) {
          for (const addition of tableAdditions) {
            const existing = tablesFileRef.current.get(addition.sheetName) ?? []
            tablesFileRef.current.set(addition.sheetName, [
              ...existing,
              {
                area: { ...addition.area },
                headerRowCount: 1,
                showRowStripes: addition.bandedRows,
                showColumnStripes: false,
                ...(addition.name !== '' ? { name: addition.name } : {}),
                columns: [...addition.columnNames],
                ...(addition.style !== undefined ? { styleName: addition.style } : {}),
              },
            ])
          }
          tableAddsRef.current = []
        }
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
    },
    [bumpProtectionEcho],
  )

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

      <Ribbon
        api={api}
        protection={{
          sheetProtected: sheetEcho,
          sheetHasPassword: sheetEchoHasPassword,
          workbookLocked: workbookEcho,
          workbookHasPassword: workbookEchoHasPassword,
          onToggleSheetProtection: toggleSheetProtection,
          onToggleWorkbookProtection: toggleWorkbookProtection,
          onSetCellsLocked: setCellsLocked,
        }}
        tables={{
          onInsertTable: handleInsertTable,
          onDeleteTable: handleDeleteTable,
        }}
      />

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
 * File DvWireRule → Univer IDataValidationRule (browser port of the
 * desktop's toUniverDvRule; bijective with the gateway's serializeRule):
 *   - type none → any (messages-only), everything else verbatim
 *   - list literals unquote ("a,b" → a,b); references gain a leading '='
 *   - custom formulas gain a leading '='
 *   - the checkbox degrade list "1,0" restores to a checkbox rule
 *   - errorStyle names arrive as Univer numbers already (parse side)
 *   - list rules carry showDropDown + renderMode TEXT (the file's normal
 *     cell appearance; the preset overlays only the small dropdown arrow)
 * Returns null for types the model cannot install — the caller skips them.
 */
function toUniverDvRule(wire: DvWireRule, uid: string): Record<string, unknown> | null {
  const raw = wire.rule
  const type = raw.type === 'none' ? 'any' : String(raw.type ?? 'any')
  if (!['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'].includes(type)) {
    return null
  }
  let formula1 = raw.formula1 === undefined ? undefined : String(raw.formula1)
  const formula2 = raw.formula2 === undefined ? undefined : String(raw.formula2)
  if (type === 'list' && formula1 !== undefined) {
    const literal = formula1.trim()
    // The insert-checkbox degrade writes list "1,0" (gateway); restore it.
    if (literal === '"1,0"') {
      return {
        uid,
        type: 'checkbox',
        ranges: wire.ranges.map((area) => ({ ...area })),
        ...(raw.allowBlank === true ? { allowBlank: true } : {}),
      }
    }
    formula1 =
      literal.startsWith('"') && literal.endsWith('"')
        ? literal.slice(1, -1)
        : `=${literal.replace(/^=/, '')}`
  } else if (type === 'custom' && formula1 !== undefined) {
    formula1 = `=${formula1.replace(/^=/, '')}`
  }
  return {
    uid,
    type,
    ranges: wire.ranges.map((area) => ({ ...area })),
    ...(raw.allowBlank === true ? { allowBlank: true } : {}),
    ...(raw.operator === undefined ? {} : { operator: raw.operator }),
    ...(formula1 === undefined ? {} : { formula1 }),
    ...(formula2 === undefined ? {} : { formula2 }),
    ...(type === 'list'
      ? {
          // OOXML showDropDown="1" suppresses the dropdown — the gateway's
          // read already inverted it into the Univer sense.
          showDropDown: raw.showDropDown !== false,
          renderMode: 1,
        }
      : {}),
    ...(raw.showInputMessage === true ? { showInputMessage: true } : {}),
    ...(raw.showErrorMessage === true ? { showErrorMessage: true } : {}),
    ...(raw.errorStyle === undefined ? {} : { errorStyle: raw.errorStyle }),
    ...(raw.errorTitle === undefined ? {} : { errorTitle: raw.errorTitle }),
    ...(raw.error === undefined ? {} : { error: raw.error }),
    ...(raw.promptTitle === undefined ? {} : { promptTitle: raw.promptTitle }),
    ...(raw.prompt === undefined ? {} : { prompt: raw.prompt }),
  }
}

/**
 * Snapshot the LIVE Univer validation model for every DV-dirty sheet into
 * the canonical SheetDvState wire shape (desktop collectDvStates parity —
 * a declarative full-rule-set snapshot, never a UI-command replay). The
 * FWorksheet.getDataValidations() facade returns FDataValidation handles;
 * each .rule is the IDataValidationRule. The wire rule strips ranges out
 * of the rule object and emits them as the sibling `ranges` array, exactly
 * the shape the gateway's applyDvRules consumes. An empty rule list means
 * "all validation on the sheet was cleared" (the engine removes
 * <dataValidations>).
 */
function collectDvStates(
  runtime: BrowserUniverRuntime | null,
  dvDirtyRef: React.MutableRefObject<Set<string>>,
): Array<{ sheetName: string; rules: DvWireRule[] }> {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const states: Array<{ sheetName: string; rules: DvWireRule[] }> = []
  for (const sheetName of dvDirtyRef.current) {
    const worksheet = workbook.getSheetByName(sheetName)
    if (!worksheet) continue
    const rules: DvWireRule[] = []
    for (const handle of worksheet.getDataValidations()) {
      const { ranges, ...rest } = (handle as unknown as { rule: Record<string, unknown> }).rule as {
        ranges?: Array<{
          startRow: number
          endRow: number
          startColumn: number
          endColumn: number
        }>
      } & Record<string, unknown>
      rules.push({
        ranges: (ranges ?? []).map((range) => ({
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
        })),
        rule: rest,
      })
    }
    states.push({ sheetName, rules })
  }
  return states
}

/**
 * Snapshot the LIVE Univer note model for every note-dirty sheet into the
 * canonical SheetNoteState wire shape (desktop collectNoteStates parity —
 * a declarative full-set snapshot, never a mutation replay). Univer's note
 * model stores one flat text blob per cell; the desktop convention
 * serializes authorship into it as "Author:\nText", re-split here via the
 * same regex. An empty notes array means "all notes on the sheet were
 * cleared" (the engine removes the comments part).
 */
function collectNoteStates(
  runtime: BrowserUniverRuntime | null,
  noteDirtyRef: React.MutableRefObject<Set<string>>,
): Array<{ sheetName: string; notes: SheetNote[] }> {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const states: Array<{ sheetName: string; notes: SheetNote[] }> = []
  for (const sheetName of noteDirtyRef.current) {
    const worksheet = workbook.getSheetByName(sheetName)
    if (!worksheet) continue
    const notes: SheetNote[] = []
    for (const note of worksheet.getNotes()) {
      const blob = note.note
      const split = /^([^\n]{1,60}):\n([\s\S]*)$/.exec(blob)
      notes.push({
        row: note.row,
        column: note.col,
        author: split?.[1] ?? '',
        text: split?.[2] ?? blob,
      })
    }
    states.push({ sheetName, notes })
  }
  return states
}

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
  dvDirtyRef: React.MutableRefObject<Set<string>>,
  noteDirtyRef: React.MutableRefObject<Set<string>>,
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

    // ── Data-validation mutations (data-validation preset). Any add/update/
    //    remove marks the sheet DV-dirty — the SAVE snapshots the live
    //    validation model declaratively (collectDvStates), exactly like the
    //    desktop's DV_MUTATIONS handler. Individual commands are never
    //    replayed; created/modified/deleted all collapse into "the sheet's
    //    full rule set at save time".
    if (DV_MUTATION_IDS.has(event.id)) {
      const params = event.params as { subUnitId?: string; unitId?: string } | undefined
      if (!params?.subUnitId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(params.subUnitId)
      if (!ws) return
      dvDirtyRef.current.add(ws.getSheetName())
      onDirty()
      return
    }

    // ── Note mutations (sheets-note preset). update-note covers create and
    //    edit; remove-note covers delete. All mark the sheet note-dirty —
    //    the SAVE snapshots the live note model declaratively
    //    (collectNoteStates), exactly like the desktop's NOTE_MUTATIONS
    //    handler. Individual mutations are never replayed.
    if (NOTE_MUTATION_IDS.has(event.id)) {
      // The note mutations carry `sheetId` (not `subUnitId` like the other
      // sheet mutation families) — accept both spellings.
      const params = event.params as
        { subUnitId?: string; sheetId?: string; unitId?: string } | undefined
      const sheetId = params?.subUnitId ?? params?.sheetId
      if (!sheetId) return
      const wb = runtime.univerAPI.getActiveWorkbook()
      if (!wb) return
      const ws = wb.getSheetBySheetId(sheetId)
      if (!ws) return
      noteDirtyRef.current.add(ws.getSheetName())
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
