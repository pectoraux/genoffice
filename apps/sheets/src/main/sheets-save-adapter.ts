/**
 * Sheets save adapter — typed conversion between the renderer's frozen
 * WorkbookSaveRequest/WorkbookFile schemas and the domain SavePlan/SaveRequest
 * types.
 *
 * This module is the shell/application conversion boundary. It lives in
 * apps/sheets (shell-owned) — NOT in runtime-contracts (which must remain
 * runtime-independent and free of renderer/IPC schema knowledge).
 *
 * ARCHITECTURE GUARDS (verified by tests):
 *   - ZERO type assertions (`as unknown as`, `as any`, `as never`)
 *   - ZERO Electron APIs (this is a pure data conversion module)
 *   - ZERO filesystem APIs
 *   - ZERO sidecar client imports
 *   - ZERO xlsx-gateway imports
 *   - ZERO IPC handler logic
 *
 * CONVERSION DIRECTION:
 *   1. WorkbookSaveRequest → SaveRequest (renderer IPC → domain)
 *      Each mutation family is mapped via an explicit typed conversion
 *      function. The domain types use `readonly` fields; the Zod-inferred
 *      types use mutable fields. We bridge this by constructing fresh
 *      object literals for each element (TypeScript treats fresh literals
 *      as assignable to readonly interfaces).
 *
 *   2. ShellWorkbookSession → WorkbookFile (domain → renderer IPC)
 *      The session's metadata is mapped to the renderer's WorkbookFile
 *      shape, then validated via workbookFileSchema.parse() to guarantee
 *      the frozen contract is satisfied. If validation fails, an Error is
 *      thrown — no unvalidated data reaches the renderer.
 */

import {
  workbookFileSchema,
  type WorkbookFile,
  type WorkbookSaveRequest,
} from '../shared/desktop-api'
import type {
  SaveRequest,
  SavePlan,
  SheetCellEdit,
  SheetStructuralOp,
  SheetOp,
  SheetHyperlinkEdit,
  SheetFilterState,
  SheetCfState,
  SheetDvState,
  SheetPageSetupState,
  SheetNoteState,
  SheetProtectionState,
  SheetProtectedRangesState,
  SheetVisualAddition,
  SheetTableAddition,
  SheetPivotAddition,
  SheetSparklineAddition,
  SheetFormulaValue,
  PivotRefreshUpdate,
  WorkbookChartEdit,
  WorkbookVisualEdit,
  DefinedNamesState,
  WorkbookThemeState,
  WorkbookProtectionState,
} from '@genoffice/runtime-contracts'
import type { ShellWorkbookSession } from './sheets-shell-coordinator'

// ── WorkbookSaveRequest → SaveRequest ───────────────────────────────

/**
 * Convert the renderer's frozen WorkbookSaveRequest to the domain SaveRequest.
 *
 * Every mutation family is mapped via an explicit typed conversion function.
 * The service resolves sheetIds → file sheet names internally (fail-closed).
 *
 * This function does NOT:
 *   - resolve sheetIds to sheetNames (the service does this)
 *   - call xlsx-gateway planning functions (the engine does this)
 *   - perform filesystem operations
 *   - use any type assertions
 */
export function translateSaveRequest(request: WorkbookSaveRequest): SaveRequest {
  const plan: SavePlan = {
    edits: mapEdits(request.edits),
    structuralOps: mapStructuralOps(request.structuralOps),
    formulaValues: mapFormulaValues(request.formulaValues),
    sheetOps: mapSheetOps(request.sheetOps),
    sheetOrder: request.sheetOrder,
    filterStates: mapFilterStates(request.filterStates),
    hyperlinkEdits: mapHyperlinkEdits(request.hyperlinkEdits),
    cfStates: mapCfStates(request.cfStates),
    dvStates: mapDvStates(request.dvStates),
    pageSetupStates: mapPageSetupStates(request.pageSetupStates),
    noteStates: mapNoteStates(request.noteStates),
    sheetProtections: mapSheetProtections(request.sheetProtections),
    protectedRangeStates: mapProtectedRangeStates(request.protectedRangeStates),
    visualAdditions: mapVisualAdditions(request.visualAdditions),
    tableAdditions: mapTableAdditions(request.tableAdditions),
    pivotAdditions: mapPivotAdditions(request.pivotAdditions),
    sparklineAdditions: mapSparklineAdditions(request.sparklineAdditions),
    chartEdits: mapChartEdits(request.chartEdits),
    visualEdits: mapVisualEdits(request.visualEdits),
    pivotCacheRefreshPaths: request.pivotCacheRefreshPaths,
    pivotRefreshUpdates: mapPivotRefreshUpdates(request.pivotRefreshUpdates),
    definedNamesState: mapDefinedNamesState(request.definedNamesState),
    themeState: mapThemeState(request.themeState),
    workbookProtectionState: mapWorkbookProtectionState(request.workbookProtectionState),
  }
  return { plan }
}

// ── Per-family mappers (WorkbookSaveRequest → SavePlan) ──────────────
//
// Each mapper constructs fresh object literals for each element. TypeScript
// treats fresh literals as assignable to readonly interfaces, even under
// `exactOptionalPropertyTypes: true`. Conditional spreads (`...`)
// preserve optionality without introducing `undefined`.

function mapEdits(edits: WorkbookSaveRequest['edits']): SheetCellEdit[] {
  return edits.map((e): SheetCellEdit => ({
    sheetId: e.sheetId,
    row: e.row,
    column: e.column,
    writeValue: e.writeValue,
    value: e.value,
    ...(e.formula !== undefined ? { formula: e.formula } : {}),
    ...(e.style !== undefined ? { style: e.style } : {}),
    ...(e.rich !== undefined ? { rich: e.rich } : {}),
    ...(e.styleReset !== undefined ? { styleReset: e.styleReset } : {}),
  }))
}

function mapStructuralOps(ops: WorkbookSaveRequest['structuralOps']): SheetStructuralOp[] {
  return ops.map((op): SheetStructuralOp => {
    if (op.kind === 'merge-cells' || op.kind === 'unmerge-cells') {
      return { sheetId: op.sheetId, kind: op.kind, range: op.range }
    }
    if (op.kind === 'set-row-size' || op.kind === 'set-col-size') {
      return { sheetId: op.sheetId, kind: op.kind, start: op.start, end: op.end, size: op.size }
    }
    if (op.kind === 'set-rows-hidden' || op.kind === 'set-cols-hidden') {
      return { sheetId: op.sheetId, kind: op.kind, start: op.start, end: op.end, hidden: op.hidden }
    }
    if (op.kind === 'set-rows-outline' || op.kind === 'set-cols-outline') {
      if (op.collapsed !== undefined) {
        return { sheetId: op.sheetId, kind: op.kind, start: op.start, end: op.end, level: op.level, collapsed: op.collapsed }
      }
      return { sheetId: op.sheetId, kind: op.kind, start: op.start, end: op.end, level: op.level }
    }
    if (op.kind === 'move-rows') {
      return { sheetId: op.sheetId, kind: op.kind, index: op.index, count: op.count, before: op.before }
    }
    // Remaining kinds: insert-rows, remove-rows, insert-cols, remove-cols.
    // These all carry `index` + `count`. The Zod schema guarantees their
    // presence for these kinds. TypeScript narrows the union to this
    // remaining set, which DOES have `index` and `count` — but the type
    // checker sees them as part of a complex union. We access them via
    // `'index' in op` narrowing for full type safety (no assertion).
    if ('index' in op && 'count' in op) {
      return { sheetId: op.sheetId, kind: op.kind, index: op.index, count: op.count }
    }
    // Unreachable: the Zod schema guarantees all kinds are covered above.
    throw new Error(`Unreachable: unexpected structural op kind ${op.kind}`)
  })
}

function mapSheetOps(ops: WorkbookSaveRequest['sheetOps']): SheetOp[] {
  return ops.map((op): SheetOp => {
    if (op.kind === 'add-sheet') {
      return { kind: op.kind, sheetId: op.sheetId, name: op.name }
    }
    if (op.kind === 'duplicate-sheet') {
      return { kind: op.kind, sheetId: op.sheetId, name: op.name, sourceSheetId: op.sourceSheetId }
    }
    if (op.kind === 'rename-sheet') {
      return { kind: op.kind, sheetId: op.sheetId, newName: op.newName }
    }
    if (op.kind === 'set-sheet-hidden') {
      return { kind: op.kind, sheetId: op.sheetId, hidden: op.hidden }
    }
    // remove-sheet: has sheetId. reorder-sheets: has NO sheetId.
    // The domain SheetOp type has `sheetId: string` (required for all kinds),
    // but the renderer's reorder-sheets doesn't send one. We default to ''
    // — the service doesn't resolve sheetId for reorder-sheets (it only
    // reads sheetOrder).
    if (op.kind === 'reorder-sheets') {
      return { kind: 'reorder-sheets', sheetId: '' }
    }
    return { kind: op.kind, sheetId: op.sheetId }
  })
}

function mapHyperlinkEdits(edits: WorkbookSaveRequest['hyperlinkEdits']): SheetHyperlinkEdit[] {
  return edits.map((e) => ({
    sheetId: e.sheetId,
    row: e.row,
    column: e.column,
    target: e.target,
  }))
}

function mapFilterStates(states: WorkbookSaveRequest['filterStates']): SheetFilterState[] {
  return states.map((s): SheetFilterState => ({
    sheetId: s.sheetId,
    filter: s.filter,
    hiddenRows: s.hiddenRows,
    ...(s.visibilityRange !== undefined ? { visibilityRange: s.visibilityRange } : {}),
  }))
}

function mapCfStates(states: WorkbookSaveRequest['cfStates']): SheetCfState[] {
  return states.map((s) => ({
    sheetId: s.sheetId,
    rules: s.rules,
  }))
}

function mapDvStates(states: WorkbookSaveRequest['dvStates']): SheetDvState[] {
  return states.map((s) => ({
    sheetId: s.sheetId,
    rules: s.rules,
  }))
}

function mapPageSetupStates(states: WorkbookSaveRequest['pageSetupStates']): SheetPageSetupState[] {
  // SheetPageSetupState is { readonly sheetId: string; readonly [key: string]: unknown }.
  // The Zod-inferred type has `sheetId` plus optional named fields. Spreading
  // all properties into a fresh object literal is assignable to the index
  // signature — no cast needed.
  return states.map((s): SheetPageSetupState => ({ ...s }))
}

function mapNoteStates(states: WorkbookSaveRequest['noteStates']): SheetNoteState[] {
  return states.map((s) => ({
    sheetId: s.sheetId,
    notes: s.notes,
  }))
}

function mapSheetProtections(states: WorkbookSaveRequest['sheetProtections']): SheetProtectionState[] {
  return states.map((s) => ({
    sheetId: s.sheetId,
    protected: s.protected,
  }))
}

function mapProtectedRangeStates(states: WorkbookSaveRequest['protectedRangeStates']): SheetProtectedRangesState[] {
  return states.map((s) => ({
    sheetId: s.sheetId,
    ranges: s.ranges,
  }))
}

function mapVisualAdditions(adds: WorkbookSaveRequest['visualAdditions']): SheetVisualAddition[] {
  return adds.map((a): SheetVisualAddition => ({
    sheetId: a.sheetId,
    anchor: a.anchor,
    ...(a.chart !== undefined ? { chart: a.chart } : {}),
    ...(a.shape !== undefined ? { shape: a.shape } : {}),
    ...(a.image !== undefined ? { image: a.image } : {}),
  }))
}

function mapTableAdditions(adds: WorkbookSaveRequest['tableAdditions']): SheetTableAddition[] {
  return adds.map((t): SheetTableAddition => ({
    sheetId: t.sheetId,
    area: t.area,
    name: t.name,
    columnNames: t.columnNames,
    ...(t.style !== undefined ? { style: t.style } : {}),
    ...(t.bandedRows !== undefined ? { bandedRows: t.bandedRows } : {}),
  }))
}

function mapPivotAdditions(adds: WorkbookSaveRequest['pivotAdditions']): SheetPivotAddition[] {
  // SheetPivotAddition is { readonly sheetId, sourceSheetId, sourceArea,
  // location, name, [key: string]: unknown }. The Zod-inferred type has
  // these plus extra fields. Spreading all into a fresh object is assignable
  // to the index signature — no cast needed.
  return adds.map((p): SheetPivotAddition => ({ ...p }))
}

function mapSparklineAdditions(adds: WorkbookSaveRequest['sparklineAdditions']): SheetSparklineAddition[] {
  return adds.map((a): SheetSparklineAddition => ({
    sheetId: a.sheetId,
    type: a.type,
    cells: a.cells.map((c) => ({ cell: c.cell, sourceRef: c.sourceRef })),
    ...(a.color !== undefined ? { color: a.color } : {}),
  }))
}

function mapChartEdits(edits: WorkbookSaveRequest['chartEdits']): WorkbookChartEdit[] {
  // WorkbookChartEdit (Increment 6A) is `{ readonly [key: string]: unknown }`
  // — no named fields. The renderer's schema uses `chartPath`; the gateway
  // reads `chartPath` directly (see xlsx-gateway.ts:928). We spread all
  // properties into a fresh object. No type assertion needed because the
  // index signature accepts any string-keyed properties.
  return edits.map((e): WorkbookChartEdit => ({ ...e }))
}

function mapVisualEdits(edits: WorkbookSaveRequest['visualEdits']): WorkbookVisualEdit[] {
  // WorkbookVisualEdit (Increment 6A) is `{ readonly [key: string]: unknown }`
  // — no named fields. The renderer's schema uses `drawingPath`. Spread all
  // properties into a fresh object.
  return edits.map((e): WorkbookVisualEdit => ({ ...e }))
}

function mapPivotRefreshUpdates(updates: WorkbookSaveRequest['pivotRefreshUpdates']): PivotRefreshUpdate[] {
  return updates.map((u): PivotRefreshUpdate => {
    if (u.relayout !== undefined) {
      // Strip sheetId from the relayout (the service resolves it from the
      // outer update). SheetPivotAddition has an index signature so the
      // remaining properties flow through.
      const { sheetId: _sheetId, ...rest } = u.relayout
      void _sheetId
      const relayout: SheetPivotAddition = {
        sheetId: u.sheetId,
        ...rest,
      }
      return {
        cachePath: u.cachePath,
        sheetId: u.sheetId,
        newOutputRef: u.newOutputRef,
        relayout,
      }
    }
    return {
      cachePath: u.cachePath,
      sheetId: u.sheetId,
      newOutputRef: u.newOutputRef,
    }
  })
}

function mapDefinedNamesState(state: WorkbookSaveRequest['definedNamesState']): DefinedNamesState | null {
  if (state === null) return null
  return {
    names: state.names.map((n) => ({
      name: n.name,
      formula: n.formula,
      ...(n.sheetIndex !== undefined ? { sheetIndex: n.sheetIndex } : {}),
    })),
    preserveNames: state.preserveNames,
  }
}

function mapThemeState(state: WorkbookSaveRequest['themeState']): WorkbookThemeState | null {
  if (state === null) return null
  return {
    ...(state.colors !== undefined ? { colors: state.colors } : {}),
    ...(state.fonts !== undefined ? { fonts: state.fonts } : {}),
  }
}

function mapWorkbookProtectionState(state: WorkbookSaveRequest['workbookProtectionState']): WorkbookProtectionState | null {
  if (state === null) return null
  return { lockStructure: state.lockStructure }
}

function mapFormulaValues(values: WorkbookSaveRequest['formulaValues']): SheetFormulaValue[] {
  return values.map((v) => ({
    sheetId: v.sheetId,
    row: v.row,
    column: v.column,
    value: v.value,
  }))
}

// ── ShellWorkbookSession → WorkbookFile ──────────────────────────────

/**
 * Build the renderer's frozen WorkbookFile from the coordinator's
 * ShellWorkbookSession, then validate it via workbookFileSchema.parse().
 *
 * After a successful save, the coordinator has replaced the old session with
 * a new one (same sessionId, new engine handle, new snapshot, new fingerprint).
 * The replacement session carries the full WorkbookMetadata from engine.open().
 *
 * This function maps the contract metadata to the renderer's WorkbookFile
 * shape and validates the result via the frozen Zod schema. If validation
 * fails, an Error is thrown — no unvalidated data reaches the renderer.
 *
 * INCREMENT 6A: the return type is the frozen `WorkbookFile` (Zod-inferred),
 * NOT `unknown` or `Record<string, unknown>`. The schema parse guarantees
 * the shape.
 */
export function buildWorkbookFile(session: ShellWorkbookSession): WorkbookFile {
  const m = session.metadata
  const candidate = {
    sessionId: session.sessionId,
    name: m.name,
    path: session.originalPath,
    // Use the coordinator's diskFingerprint (computed via sha256File(snapshot))
    // — NOT metadata.sha256, which comes from the sidecar's open response
    // and is often empty. The renderer's preload validates sha256 as
    // /^[a-f0-9]{64}$/.
    sha256: session.diskFingerprint,
    entryCount: m.entryCount,
    sheets: m.sheets.map((s) => ({
      id: s.id,
      name: s.name,
      rowCount: s.rowCount,
      columnCount: s.columnCount,
      columnWidths: s.columnWidths ?? [],
      defaultRowHeight: s.defaultRowHeight,
      defaultColumnWidth: s.defaultColumnWidth,
      freeze: null,
      hidden: s.hidden,
      tabColor: s.tabColor ?? null,
      showGridLines: s.showGridlines,
      tables: s.tables ?? [],
      comments: s.comments ?? [],
      pivotRanges: s.pivotRanges ?? [],
    })),
    activeTab: m.activeTab,
    styles: m.styles ?? [],
    dxfStyles: m.dxfStyles ?? [],
    visuals: m.visuals ?? [],
    definedNames: m.definedNames.map((d) => {
      const result: { name: string; formula: string; sheetIndex?: number } = {
        name: d.name,
        formula: d.formula,
      }
      if (d.sheetIndex !== undefined) result.sheetIndex = d.sheetIndex
      return result
    }),
    readOnly: false,
    needsSaveAs: session.needsSaveAs ?? (session.suggestSaveAs !== undefined),
    restoredFromRecovery: session.restoredFromRecovery ?? (session.restoreTarget !== undefined),
    ...(m.themeColors.length > 0 ? { themeColors: m.themeColors } : {}),
    ...(m.themeFonts.major !== '' || m.themeFonts.minor !== ''
      ? { themeFonts: m.themeFonts }
      : {}),
  }
  return workbookFileSchema.parse(candidate)
}
