/**
 * Formula view — Excel's Show Formulas, persisted per sheet as
 * sheetView/@showFormulas (EXCEL-026). Desktop parity: the per-sheet flag
 * lives in the shell's state (NOT the Univer model — the engine has no
 * public formula-view API in 0.25.1), the RENDER_RAW_FORMULA_KEY context
 * flips per ACTIVE sheet so every skeleton drops its text cache and
 * repaints, and a CELL_CONTENT interceptor performs the actual value →
 * formula-text swap (the engine's raw-formula render path only reaches
 * rich-text/rotated cells in this version).
 *
 * The shell journals toggles through the canonical pageSetupStates save
 * family ({ showFormulas: boolean } — the gateway writes
 * sheetView@showFormulas); this module holds only the RENDER state.
 */
import { CellValueType, IContextService } from '@univerjs/core'
import {
  INTERCEPTOR_POINT,
  RENDER_RAW_FORMULA_KEY,
  SheetInterceptorService,
} from '@univerjs/preset-sheets-core'

import type { BrowserUniverRuntime } from './create-browser-univer'

/** Per-sheet formula-view render state (sheetIds in formula view). */
export interface FormulaViewState {
  readonly sheets: Set<string>
}

export function createFormulaViewState(): FormulaViewState {
  return { sheets: new Set() }
}

/**
 * Sync the (global) raw-formula render key to the ACTIVE sheet's flag.
 * Flipping the context value makes every sheet skeleton drop its text
 * cache and repaint; the interceptor below decides per-cell what to draw.
 */
export function applyShowFormulasView(
  runtime: BrowserUniverRuntime,
  state: FormulaViewState,
  sheetId: string,
): void {
  const contextService = runtime.univer.__getInjector().get(IContextService)
  const next = state.sheets.has(sheetId)
  if (Boolean(contextService.getContextValue(RENDER_RAW_FORMULA_KEY)) !== next) {
    contextService.setContextValue(RENDER_RAW_FORMULA_KEY, next)
  }
}

/**
 * Swap the displayed value of every formula cell on a formula-view sheet
 * with its formula text. Above NUMFMT (10): a formula cell in formula view
 * shows its formula, not the formatted value, so the chain stops here.
 * The web's model seeds `f` WITH the leading '=' (buildCellDataMatrix), so
 * the text drawn is Excel's own "=A1*2" display form — no re-synthesis.
 */
export function installFormulaViewInterceptor(
  runtime: BrowserUniverRuntime,
  stateRef: { readonly current: FormulaViewState },
): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    priority: 9999,
    handler: (cell, location, next) => {
      if (!stateRef.current.sheets.has(location.subUnitId)) {
        return next(cell)
      }
      const formula = location.rawData?.f
      if (typeof formula !== 'string' || formula === '') return next(cell)
      return { ...(cell ?? {}), v: formula, t: CellValueType.STRING, p: null }
    },
  })
}
