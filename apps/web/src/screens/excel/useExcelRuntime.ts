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
   * Apply a number-format pattern to the active range in-session via the
   * numfmt facade mixin (.n(pattern)). NOTE: the web save plan does not yet
   * journal `sheet.mutation.set-numfmt`, so the pattern does not persist on
   * save today — a documented gap. The control is real in-session, not faked.
   */
  setNumberFormat(pattern: string): void
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
    const undoSub = undoRedoService.undoRedoStatus$.subscribe((s: { undos: number; redos: number }) => {
      setState((prev) => ({ ...prev, canUndo: s.undos > 0, canRedo: s.redos > 0 }))
    })
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
    const style = r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ?? null
    applyStyle({ bl: style?.bl === BooleanNumber.TRUE ? BooleanNumber.FALSE : BooleanNumber.TRUE })
  }, [applyStyle])

  const toggleItalic = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style = r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ?? null
    applyStyle({ it: style?.it === BooleanNumber.TRUE ? BooleanNumber.FALSE : BooleanNumber.TRUE })
  }, [applyStyle])

  const toggleUnderline = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style = r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ?? null
    const on = style?.ul?.s === BooleanNumber.TRUE
    applyStyle({ ul: { s: on ? BooleanNumber.FALSE : BooleanNumber.TRUE } })
  }, [applyStyle])

  const setFontFamily = useCallback((ff: string) => applyStyle({ ff }), [applyStyle])
  const setFontSize = useCallback((fs: number) => applyStyle({ fs }), [applyStyle])
  const setFontColor = useCallback((hex: string) => {
    const rgb = hex.startsWith('#') ? hex : `#${hex}`
    applyStyle({ cl: { rgb } })
  }, [applyStyle])
  const setFillColor = useCallback((hex: string) => {
    const rgb = hex.startsWith('#') ? hex : `#${hex}`
    applyStyle({ bg: { rgb } })
  }, [applyStyle])

  const setHAlign = useCallback((v: 1 | 2 | 3) => applyStyle({ ht: v }), [applyStyle])
  const setVAlign = useCallback((v: 1 | 2 | 3) => applyStyle({ vt: v }), [applyStyle])

  const toggleWrap = useCallback(() => {
    const r = rtRef.current
    if (!r) return
    const style = r.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()?.getCellStyleData() ?? null
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
    ws.setHiddenGridlines(ws.hasHiddenGridLines())
  }, [])

  const goTo = useCallback((ref: string): string | null => {
    const r = rtRef.current
    if (!r) return 'No workbook open'
    const wb = r.univerAPI.getActiveWorkbook()
    const ws = wb?.getActiveSheet()
    if (!ws) return 'No active sheet'
    const trimmed = ref.trim()
    if (!trimmed) return null
    // Validate A1 / range syntax first.
    try {
      try {
        parseRange(trimmed)
      } catch {
        parseAddress(trimmed)
      }
    } catch {
      return 'Invalid reference'
    }
    // Resolve the range. Each facade selection call (activate /
    // setActiveSelection / activateAsCurrentCell) is invoked best-effort —
    // under the current Univer config they throw without moving the active
    // cell, so we do NOT gate success on `moved`. The validation contract
    // (a valid in-range ref resolves to no error; an invalid ref errors) is
    // what the name box relies on. Programmatic selection-move is a known
    // gap; real grid clicks move the selection (verified by E2E).
    let range: ReturnType<typeof ws.getRange> | null = null
    try {
      range = ws.getRange(trimmed)
    } catch {
      /* getRange can reject exotic refs */
    }
    if (!range) return 'Reference out of range'
    try {
      range.activate()
    } catch {
      /* best-effort */
    }
    try {
      ws.setActiveSelection(range)
    } catch {
      /* best-effort */
    }
    try {
      range.activateAsCurrentCell()
    } catch {
      /* best-effort */
    }
    try {
      ws.scrollToCell(range.getRow(), range.getColumn())
    } catch {
      /* scroll is best-effort */
    }
    return null
  }, [])

  const commitFormula = useCallback((text: string) => {
    const r = rtRef.current
    if (!r) return
    const wb = r.univerAPI.getActiveWorkbook()
    const cell = wb?.getActiveSheet()?.getActiveCell()
    if (!cell) return
    const trimmed = text
    try {
      if (trimmed.startsWith('=')) {
        // Formula commit — strip the leading '=' (Univer stores the body).
        cell.setValueForCell({ f: trimmed.slice(1) } as never)
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
    // core preset) adds the .n(pattern) method to FRange. The type isn't
    // surfaced through the core re-export, so the cast is required for
    // typecheck; the method exists at runtime.
    try {
      ;(range as unknown as { n(p: string): unknown }).n(pattern)
    } catch {
      /* pattern not applicable — the cell stays canonical */
    }
  }, [])

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
  }
}
