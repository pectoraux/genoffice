/**
 * GenOffice web Sheets — bridge between the React shell and Univer's facade.
 *
 * This hook owns ALL Univer interaction for the shell (ribbon, name box,
 * formula bar, status bar). It exposes reactive state (active cell, selection
 * format, zoom, undo/redo occupancy, gridlines visibility) and stable action
 * callbacks (toggle bold, set font, merge, zoom, gridlines, undo/redo, name-box
 * navigation, formula-bar commit).
 *
 * Fidelity invariant: every value/style edit is committed through Univer's
 * FRange.setValue / setValueForCell, which fires `sheet.mutation.set-range-values`.
 * ExcelEditor's existing CommandExecuted subscription journals that mutation
 * via cell-mutation-merge.ts — the browser never serializes the workbook and
 * never runs a second formula engine. This hook does NOT touch the journal or
 * the save plan directly.
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import { BooleanNumber, WrapStrategy } from '@univerjs/core'
import type { IStyleData } from '@univerjs/core'
import { ILayoutService } from '@univerjs/ui'
// Side-effect import — loads the @univerjs/sheets-sort facade types, which
// augment `FRange` (from @univerjs/sheets/facade) with the public
// `sort(column: SortColumnSpec | SortColumnSpec[]): FRange` method declared
// by `FRangeSheetsSortMixin`. The runtime side is already wired by the
// `UniverSheetsSortPreset` in create-browser-univer.ts; this import only
// surfaces the TypeScript signature so `range.sort(...)` typechecks with
// NO `as unknown as` / `as never` cast and NO private-field reach. The
// method on FRange is the canonical Univer sort facade — the same path
// the desktop's renderer uses.
import '@univerjs/sheets-sort/facade'
import type { BrowserUniverRuntime } from '../../office/create-browser-univer'
import { parseAddress, parseRange } from '../../office/cell-address'

/** Selection format the ribbon echoes — a trimmed view of IStyleData. */
export interface SelectionFormat {
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly fontFamily: string | null
  readonly fontSize: number | null
  readonly fontColor: string | null
  readonly fillColor: string | null
  readonly hAlign: 'left' | 'center' | 'right' | null
  readonly vAlign: 'top' | 'middle' | 'bottom' | null
  readonly wrap: boolean
}

const EMPTY_FORMAT: SelectionFormat = {
  bold: false,
  italic: false,
  underline: false,
  fontFamily: null,
  fontSize: null,
  fontColor: null,
  fillColor: null,
  hAlign: null,
  vAlign: null,
  wrap: false,
}

export interface ExcelRuntimeState {
  ready: boolean
  activeCellA1: string
  selectionFormat: SelectionFormat
  zoomPercent: number
  gridlinesVisible: boolean
  canUndo: boolean
  canRedo: boolean
  /** Bumps on every refresh (selection/zoom/sheet/command) — lets the formula
   * bar re-read the active cell even when the cell address didn't change. */
  seq: number
}

const H_ALIGN_NAMES: Record<number, 'left' | 'center' | 'right'> = {
  1: 'left',
  2: 'center',
  3: 'right',
}
const V_ALIGN_NAMES: Record<number, 'top' | 'middle' | 'bottom'> = {
  1: 'top',
  2: 'middle',
  3: 'bottom',
}

/** Clamp a row/column index to a non-negative integer (Univer uses 0-based). */
function range0(n: number): number {
  if (!Number.isInteger(n) || n < 0) return 0
  return n
}

function toSelectionFormat(style: IStyleData | null | undefined): SelectionFormat {
  if (!style) return EMPTY_FORMAT
  const colorHex = (rgb: string | undefined): string | null => {
    if (!rgb) return null
    const s = rgb.startsWith('#') ? rgb : `#${rgb}`
    return /^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(s) ? s.slice(0, 7).toUpperCase() : null
  }
  return {
    bold: style.bl === BooleanNumber.TRUE,
    italic: style.it === BooleanNumber.TRUE,
    underline: style.ul?.s === BooleanNumber.TRUE,
    fontFamily: style.ff ?? null,
    fontSize: typeof style.fs === 'number' ? style.fs : null,
    fontColor: colorHex(style.cl?.rgb ?? undefined),
    fillColor: colorHex(style.bg?.rgb ?? undefined),
    hAlign: style.ht != null && H_ALIGN_NAMES[style.ht] ? H_ALIGN_NAMES[style.ht] : null,
    vAlign: style.vt != null && V_ALIGN_NAMES[style.vt] ? V_ALIGN_NAMES[style.vt] : null,
    wrap: style.tb === WrapStrategy.WRAP,
  }
}

/** Read the live selection state from the active workbook (or empty). */
function readState(rt: BrowserUniverRuntime): ExcelRuntimeState {
  const wb = rt.univerAPI.getActiveWorkbook()
  if (!wb) {
    return {
      ready: false,
      activeCellA1: '',
      selectionFormat: EMPTY_FORMAT,
      zoomPercent: 100,
      gridlinesVisible: true,
      canUndo: false,
      canRedo: false,
      seq: 0,
    }
  }
  const ws = wb.getActiveSheet()
  const range = ws?.getActiveRange() ?? null
  // FRange.getCellStyleData() returns the resolved IStyleData for the
  // anchor cell — same source the desktop's selection-format.ts reads.
  const style = range ? range.getCellStyleData() : null
  const zoom = ws?.getZoom() ?? 1
  return {
    ready: true,
    activeCellA1: range ? range.getA1Notation() : '',
    selectionFormat: toSelectionFormat(style),
    zoomPercent: Math.round(zoom * 100),
    gridlinesVisible: ws ? !ws.hasHiddenGridLines() : true,
    canUndo: false,
    canRedo: false,
    seq: 0,
  }
}

export interface ExcelRuntimeApi {
  readonly state: ExcelRuntimeState
  /** Apply a partial IStyleData to the active range (fires set-range-values). */
  applyStyle(patch: Partial<IStyleData>): void
  toggleBold(): void
  toggleItalic(): void
  toggleUnderline(): void
  setFontFamily(ff: string): void
  setFontSize(fs: number): void
  setFontColor(hex: string): void
  setFillColor(hex: string): void
  setHAlign(v: 1 | 2 | 3): void
  setVAlign(v: 1 | 2 | 3): void
  toggleWrap(): void
  toggleMerge(): void
  undo(): void
  redo(): void
  setZoom(ratio: number): void
  zoomIn(): void
  zoomOut(): void
  toggleGridlines(): void
  /** Jump the selection to an A1/range ref. Returns null on success, an error string on failure. */
  goTo(ref: string): string | null
  /** Commit the formula bar's text to the active cell (fires set-range-values). */
  commitFormula(text: string): void
  /**
   * Apply a number-format pattern to the active range via the numfmt facade
   * mixin (.n(pattern)). The mutation `sheet.mutation.set.numfmt` is
   * journaled by ExcelEditor's expanded subscription as a per-cell
   * style.numberFormat CellEdit, which the canonical WorkbookStyleEdit
   * persists through applyCellEditsToXlsx — the pattern survives
   * save/reopen.
   */
  setNumberFormat(pattern: string): void
  /**
   * Sort the active range by its first column. asc=true → ascending,
   * asc=false → descending. Uses the sheets-sort preset's FRange.sort()
   * facade, which fires sheet.command.sort-range → ReorderRangeMutation.
   * ExcelEditor's expanded journal subscription captures the post-sort cell
   * values and journals them as writeValue CellEdits — on save, the
   * canonical applyCellEditsToXlsx writes the new row order into the XLSX.
   */
  sortRange(asc: boolean): void
  /**
   * Freeze panes at the active cell — freezes all rows above and all
   * columns to the left of the active cell. Fires sheet.command.set-frozen,
   * journaled by ExcelEditor as a per-sheet BrowserSheetPageSetupState, and
   * persisted by the canonical applyPageSetupState (the <pane> element).
   * Calling freezePanes() again on a frozen sheet first clears the freeze
   * (the command toggles). Returns null on success, an error string on
   * failure (e.g. no active selection).
   */
  toggleFreezePanes(): string | null
  /**
   * Insert a function (=SUM by default, or any formula body) into the
   * active cell through the existing commitFormula path — fires
   * set-range-values, journaled by the existing subscription. The leading
   * '=' is optional (added if absent). No second formula engine runs.
   */
  insertFunction(formulaBody: string): void
}

export function useExcelRuntime(rt: BrowserUniverRuntime | null): ExcelRuntimeApi | null {
  const [state, setState] = useState<ExcelRuntimeState>(() =>
    rt
      ? readState(rt)
      : {
          ready: false,
          activeCellA1: '',
          selectionFormat: EMPTY_FORMAT,
          zoomPercent: 100,
          gridlinesVisible: true,
          canUndo: false,
          canRedo: false,
          seq: 0,
        },
  )
  const rtRef = useRef(rt)
  rtRef.current = rt

  const refresh = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    setState((prev) => {
      const next = readState(r)
      // Preserve undo/redo occupancy (driven by its own subscription) and
      // bump seq so subscribers (formula bar) re-read on same-cell edits.
      return { ...next, canUndo: prev.canUndo, canRedo: prev.canRedo, seq: prev.seq + 1 }
    })
  }, [])

  // Subscribe to Univer selection/zoom/sheet events + undo/redo occupancy.
  useEffect(() => {
    if (!rt) return
    const { univerAPI, undoRedoService } = rt
    const subs: Array<() => void> = []
    const E = univerAPI.Event
    // addEvent returns IDisposable; wrap each so the cleanup array has a
    // uniform () => void signature.
    const track = (d: { dispose(): void }) => subs.push(() => d.dispose())
    track(univerAPI.addEvent(E.SelectionChanged, () => refresh()))
    track(univerAPI.addEvent(E.SheetZoomChanged, () => refresh()))
    track(univerAPI.addEvent(E.ActiveSheetChanged, () => refresh()))
    // Style/numfmt edits surface as commands — refresh so the ribbon toggles
    // re-toggle after the mutation lands.
    track(univerAPI.addEvent(E.CommandExecuted, () => refresh()))
    const undoSub = undoRedoService.undoRedoStatus$.subscribe(
      (s: { undos: number; redos: number }) => {
        setState((prev) => ({ ...prev, canUndo: s.undos > 0, canRedo: s.redos > 0 }))
      },
    )
    subs.push(() => undoSub.unsubscribe())
    refresh()
    return () => subs.forEach((fn) => fn())
  }, [rt, refresh])

  const applyStyle = useCallback((patch: Partial<IStyleData>) => {
    const r = rtRef.current
    if (!r) return
    const wb = r.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    const range = ws?.getActiveRange()
    if (!range) return
    try {
      range.setValue({ s: patch } as never)
    } catch {
      /* some style combinations are rejected; the grid stays canonical */
    }
  }, [])

  const toggleBold = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style =
      r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ??
      null
    applyStyle({ bl: style?.bl === BooleanNumber.TRUE ? BooleanNumber.FALSE : BooleanNumber.TRUE })
  }, [applyStyle])

  const toggleItalic = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style =
      r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ??
      null
    applyStyle({ it: style?.it === BooleanNumber.TRUE ? BooleanNumber.FALSE : BooleanNumber.TRUE })
  }, [applyStyle])

  const toggleUnderline = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style =
      r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ??
      null
    const on = style?.ul?.s === BooleanNumber.TRUE
    applyStyle({ ul: { s: on ? BooleanNumber.FALSE : BooleanNumber.TRUE } })
  }, [applyStyle])

  const setFontFamily = useCallback((ff: string) => applyStyle({ ff }), [applyStyle])
  const setFontSize = useCallback((fs: number) => applyStyle({ fs }), [applyStyle])
  const setFontColor = useCallback(
    (hex: string) => {
      const rgb = hex.startsWith('#') ? hex : `#${hex}`
      applyStyle({ cl: { rgb } })
    },
    [applyStyle],
  )
  const setFillColor = useCallback(
    (hex: string) => {
      const rgb = hex.startsWith('#') ? hex : `#${hex}`
      applyStyle({ bg: { rgb } })
    },
    [applyStyle],
  )

  const setHAlign = useCallback((v: 1 | 2 | 3) => applyStyle({ ht: v }), [applyStyle])
  const setVAlign = useCallback((v: 1 | 2 | 3) => applyStyle({ vt: v }), [applyStyle])

  const toggleWrap = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style =
      r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ??
      null
    // WrapStrategy: 1=WRAP, 2=OVERFLOW (default). Toggle between them.
    const next = style?.tb === WrapStrategy.WRAP ? WrapStrategy.OVERFLOW : WrapStrategy.WRAP
    applyStyle({ tb: next })
  }, [applyStyle])

  const toggleMerge = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const range = r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()
    if (!range) return
    try {
      if (range.isMerged()) range.breakApart()
      else range.merge()
    } catch {
      /* merge not applicable to the current selection */
    }
  }, [])

  const undo = useCallback(() => {
    void rtRef.current?.univerAPI.undo()
  }, [])
  const redo = useCallback(() => {
    void rtRef.current?.univerAPI.redo()
  }, [])

  const setZoom = useCallback((ratio: number) => {
    const r = rtRef.current
    if (!r) return
    const clamped = Math.max(0.1, Math.min(4, ratio))
    r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.zoom(clamped)
  }, [])
  const zoomIn = useCallback(() => {
    const r = rtRef.current
    const cur = r?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getZoom() ?? 1
    setZoom(cur + 0.1)
  }, [setZoom])
  const zoomOut = useCallback(() => {
    const r = rtRef.current
    const cur = r?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getZoom() ?? 1
    setZoom(cur - 0.1)
  }, [setZoom])

  const toggleGridlines = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const ws = r.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!ws) return
    // Toggle: hidden → visible, visible → hidden. The previous
    // implementation passed `ws.hasHiddenGridLines()` (the CURRENT state),
    // which was a no-op — the gridlines never actually toggled.
    ws.setHiddenGridlines(!ws.hasHiddenGridLines())
  }, [])

  const goTo = useCallback(
    (ref: string): string | null => {
      const r = rtRef.current
      if (!r) return 'No workbook open'
      const wb = r.univerAPI.getActiveWorkbook()
      const ws = wb?.getActiveSheet()
      if (!wb || !ws) return 'No active sheet'
      const trimmed = ref.trim()
      if (!trimmed) return null
      // Validate A1 / range syntax first (the web has no defined-names layer,
      // so plain syntax validation is the resolution step the desktop's
      // resolveGoToRef performs for named refs).
      try {
        try {
          parseRange(trimmed)
        } catch {
          parseAddress(trimmed)
        }
      } catch {
        return 'Invalid reference'
      }
      // The FWorkbook facade carries the authoritative selection API
      // (setActiveRange, declared at @univerjs/sheets f-workbook.d.ts:384) —
      // the SAME path the desktop's goToReference (apps/sheets/src/renderer/
      // data-tools-actions.ts:134-180) uses. The previous implementation
      // called ws.setActiveSelection(range) + range.activate()/activateAs-
      // CurrentCell() (FWorksheet/FRange facades), which throw silently under
      // the toolbar:false config and never move the active cell. workbook.
      // setActiveRange is the correct, proven entry point.
      try {
        // A jump must not leave an editor open on the previous cell — later
        // keystrokes would land there. Commit it before moving (fire-and-
        // forget; matches the desktop's `void endEditingAsync(true)`).
        if (wb.isCellEditing()) void wb.endEditingAsync(true)
        const range = ws.getRange(trimmed)
        // If the ref named a different sheet (e.g. "Sheet2!B5"), switch to it;
        // otherwise stay on the active sheet.
        const target = wb.getSheetBySheetId(range.getSheetId()) ?? ws
        wb.setActiveRange(range)
        target.scrollToCell(range.getRow(), range.getColumn())
        // Hand keyboard focus back to the grid (Univer's hidden editor host,
        // the same handoff the desktop performs via ILayoutService.focus()).
        // Without this, the browser keeps focus on the Name Box <input> and
        // typing after a jump lands in the input, not the target cell.
        try {
          r.univer.__getInjector().get(ILayoutService).focus()
        } catch {
          /* ILayoutService not registered in this build — grid click + name
           box echo still work; only direct keyboard entry after a jump
           would land on <body> instead of the target cell. Best-effort. */
        }
        // Trigger a refresh so the Name Box echoes the new active cell
        // immediately (the SelectionChanged event fires asynchronously).
        refresh()
      } catch (error: unknown) {
        const detail = error instanceof Error ? error.message : ''
        if (detail.includes('out of bounds') || detail.includes('out of range')) {
          return 'Reference out of range'
        }
        if (detail.includes('Range not found')) return 'Sheet not found'
        return detail === '' ? 'Invalid reference' : `Navigation failed: ${detail}`
      }
      return null
    },
    [refresh],
  )

  const commitFormula = useCallback((text: string) => {
    const r = rtRef.current
    if (!r) return
    const wb = r.univerAPI.getActiveWorkbook()
    const cell = wb?.getActiveSheet()?.getActiveCell()
    if (!cell) return
    const trimmed = text
    try {
      if (trimmed.startsWith('=')) {
        // Formula commit — KEEP the leading '='. Univer's internal cell.f
        // convention INCLUDES it (isFormulaString requires
        // `value.substring(0, 1) === "="`): setValueForCell({f:'=SUM(..)'})
        // stores a LIVE formula the engine calculates (v gets the result);
        // stripping the '=' would seed a dead formula that never
        // calculates. The JOURNAL's CellEdit wire format strips the '=' in
        // cellEditFromMutation (XLSX <f> elements have no '=') — untouched.
        cell.setValueForCell({ f: trimmed } as never)
      } else {
        // Literal value commit. An empty string clears the cell value.
        cell.setValueForCell((trimmed === '' ? { v: '' } : { v: trimmed }) as never)
      }
    } catch {
      /* rejected commit — the cell stays canonical */
    }
  }, [])

  const setNumberFormat = useCallback((pattern: string) => {
    const r = rtRef.current
    if (!r) return
    const range = r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()
    if (!range) return
    // The numfmt facade mixin (from @univerjs/sheets-numfmt, bundled by the
    // core preset) adds the setNumberFormat(pattern) method to FRange
    // (FRangeSheetsNumfmtMixin). It fires SetNumfmtCommand, which internally
    // dispatches sheet.mutation.set.numfmt — journaled by ExcelEditor's
    // expanded subscription as a per-cell style.numberFormat CellEdit. The
    // type isn't surfaced through the @univerjs/core re-export, so the cast
    // is required for typecheck; the method exists at runtime.
    try {
      ;(range as unknown as { setNumberFormat(p: string): unknown }).setNumberFormat(pattern)
    } catch {
      /* pattern not applicable — the cell stays canonical */
    }
  }, [])

  const sortRange = useCallback(
    (asc: boolean) => {
      const r = rtRef.current
      if (!r) return
      const wb = r.univerAPI.getActiveWorkbook()
      const ws = wb?.getActiveSheet()
      const range = ws?.getActiveRange()
      if (!wb || !ws || !range) return
      // Canonical Univer sort path: FRange.sort (the public facade mixin from
      // @univerjs/sheets-sort, surfaced via the side-effect import at the top
      // of this module). The mixin delegates to SortRangeCommand →
      // ReorderRangeCommand → ReorderRangeMutation (sheet.mutation.reorder-
      // range), which deepClones the entire cell record (v/f/s/p/si/t) via
      // getCellRaw and writes it into the worksheet cellDataMatrix. Styles,
      // numfmt, fills, borders, hyperlinks, comments, validation, and any
      // other cell metadata travel atomically with the row — the cardinal
      // requirement the previous JS-sort implementation violated.
      //
      // ExcelEditor's `sheet.mutation.reorder-range` subscription journals
      // the row permutation as a `reorder-rows` structural op (range + order
      // map). On save, the canonical applyStructuralOps path in xlsx-gateway
      // permutes <row> blocks atomically — only the r= attributes on <row>
      // and inner <c> renumber, the cell contents travel UNTOUCHED inside
      // their <c> elements. Save/reopen is faithful to Univer's live state.
      //
      // No `_range` private-field access, no `as never` cast, no
      // `as unknown as` cast — `range.sort(...)` typechecks directly via
      // the FRangeSheetsSortMixin augmentation.
      try {
        range.sort({ column: 0, ascending: asc })
        refresh()
      } catch {
        /* sort not applicable — selection may be a single cell */
      }
    },
    [refresh],
  )

  const toggleFreezePanes = useCallback((): string | null => {
    const r = rtRef.current
    if (!r) return 'No workbook open'
    const wb = r.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    const cell = ws?.getActiveCell()
    if (!wb || !ws || !cell) return 'No active selection'
    // FWorksheet exposes setFreeze({ startRow, startColumn, xSplit, ySplit })
    // and getFreeze() — the facade's documented freeze API
    // (f-worksheet.d.ts:1019/1033). startRow/startColumn are the first
    // scrollable (non-frozen) row/column index — so freezing at the active
    // cell means startRow = activeRow, startColumn = activeColumn, xSplit =
    // activeColumn (count of frozen cols), ySplit = activeRow (count of
    // frozen rows). setFreeze fires sheet.mutation.set-frozen, journaled
    // by ExcelEditor's expanded subscription as a per-sheet
    // BrowserSheetPageSetupState.
    try {
      const activeRow = range0(cell.getRow())
      const activeCol = range0(cell.getColumn())
      // Toggle: if already frozen at exactly this cell, clear; otherwise
      // (re)freeze. A "no freeze" config uses startRow=-1, startColumn=-1,
      // xSplit=0, ySplit=0.
      const current = (
        ws as unknown as {
          getFreeze(): { startRow: number; startColumn: number; xSplit: number; ySplit: number }
        }
      ).getFreeze()
      const isFrozenHere = current.startRow === activeRow && current.startColumn === activeCol
      const next = isFrozenHere
        ? { startRow: -1, startColumn: -1, xSplit: 0, ySplit: 0 }
        : { startRow: activeRow, startColumn: activeCol, xSplit: activeCol, ySplit: activeRow }
      ;(
        ws as unknown as {
          setFreeze(f: {
            startRow: number
            startColumn: number
            xSplit: number
            ySplit: number
          }): unknown
        }
      ).setFreeze(next)
      // Trigger a refresh so the ribbon state re-reads immediately.
      refresh()
    } catch (error: unknown) {
      const detail = error instanceof Error ? error.message : ''
      return detail === '' ? 'Freeze failed' : `Freeze failed: ${detail}`
    }
    return null
  }, [refresh])

  const insertFunction = useCallback(
    (formulaBody: string) => {
      // Strip an optional leading '=' so callers can pass either "SUM(A1:A2)"
      // or "=SUM(A1:A2)". The commit path adds the leading '=' back for
      // formula echo display, and cellEditFromMutation stores the body
      // without the '=' (the canonical CellEdit.formula convention).
      const body = formulaBody.startsWith('=') ? formulaBody.slice(1) : formulaBody
      commitFormula(`=${body}`)
    },
    [commitFormula],
  )

  if (!rt) return null
  return {
    state,
    applyStyle,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    setFontFamily,
    setFontSize,
    setFontColor,
    setFillColor,
    setHAlign,
    setVAlign,
    toggleWrap,
    toggleMerge,
    undo,
    redo,
    setZoom,
    zoomIn,
    zoomOut,
    toggleGridlines,
    goTo,
    commitFormula,
    setNumberFormat,
    sortRange,
    toggleFreezePanes,
    insertFunction,
  }
}
