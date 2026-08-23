/**
 * SpreadsheetEngine — runtime-independent interface for the spreadsheet
 * execution engine (ADR-004).
 *
 * This is the Sheets-specific execution-engine port. It is NOT a generic
 * platform capability and NOT a shell responsibility. The interface lives
 * in runtime-contracts (Layer 1); the implementation lives in
 * platform-electron (Layer 4a) as ElectronXlsxSidecarEngine.
 *
 * RUNTIME INDEPENDENCE:
 *   The engine contract is data-oriented. It accepts and returns Uint8Array
 *   for workbook content — NOT filesystem paths. This allows:
 *     - ElectronXlsxSidecarEngine: internally writes Uint8Array to a temp
 *       file, passes the path to the Rust sidecar, reads the result back
 *       as Uint8Array. The temp-file translation is private to the adapter.
 *     - WasmSpreadsheetEngine: passes Uint8Array directly to in-process
 *       IronCalc — no filesystem needed.
 *     - CloudSpreadsheetEngine: uploads Uint8Array to a server — no local
 *       filesystem needed.
 *
 * The engine operates on opaque EngineSessionHandle tokens — the domain
 * service and runtime contracts NEVER inspect the token's internal
 * representation.
 *
 * FORBIDDEN in this file (and all runtime-contracts):
 *   sidecarSessionId, sidecar, Rust, stdio, child_process, snapshotPath,
 *   BrowserWindow, WebContents, wcId, Electron, node:fs, node:path,
 *   engineSessionId, filesystem path parameters
 *
 * IMPORTANT (ADR-001 Correction A): constructor injection. No getRuntime().
 */

import type { SavePlan } from './save-plan.js'

// ── Opaque engine session handle ───────────────────────────────────────

/**
 * An opaque token representing an engine session. Created by
 * `SpreadsheetEngine.open()`, passed to all subsequent engine operations.
 *
 * The token has NO inspectable fields. Consumers must not attempt to
 * read, compare, or construct one. The only way to obtain an
 * EngineSessionHandle is as the return value of `engine.open()`.
 *
 * The Electron adapter maps this token to a Rust sidecar UUID internally.
 * A WASM adapter would map it to an in-memory table key. A Cloud adapter
 * would map it to a server session token. The mapping is private to the
 * adapter implementation.
 *
 * INCREMENT 5B (build-fix): the brand is now a real runtime `Symbol()`
 * (instead of `declare const ... unique symbol` which emitted no runtime
 * binding and could not be re-exported by rollup's `export *`). The
 * unique-symbol type guarantee is preserved via the explicit type
 * annotation on the const. The brand remains opaque and unforgeable —
 * symbol keys are invisible to `Object.keys()` / `JSON.stringify`, and
 * external code cannot construct a handle without the brand reference.
 */
export const ENGINE_SESSION_HANDLE_BRAND: unique symbol = Symbol('ENGINE_SESSION_HANDLE_BRAND')

export interface EngineSessionHandle {
  /** @internal Brand marker — do not access. */
  readonly [ENGINE_SESSION_HANDLE_BRAND]: typeof ENGINE_SESSION_HANDLE_BRAND
}

// ── Engine errors ───────────────────────────────────────────────────────

/**
 * Base error for all engine failures. Contains domain-safe information only —
 * no Rust/stdio/child-process implementation details.
 */
export class EngineError extends Error {
  constructor(
    message: string,
    readonly code: string,
  ) {
    super(message)
    this.name = 'EngineError'
  }
}

/**
 * The engine session handle is unknown — the session was closed, expired,
 * or never opened. The coordinator should clean up the session.
 */
export class InvalidSessionError extends EngineError {
  constructor(message: string) {
    super(message, 'INVALID_SESSION')
    this.name = 'InvalidSessionError'
  }
}

/**
 * The request payload failed validation. The domain service should reject
 * the caller.
 */
export class InvalidInputError extends EngineError {
  constructor(message: string) {
    super(message, 'INVALID_INPUT')
    this.name = 'InvalidInputError'
  }
}

// ── External change status ─────────────────────────────────────────────

/**
 * Runtime-neutral fact about whether the file on disk has changed since
 * the workbook was opened.
 *
 * Supplied by the shell (which observes filesystem state via the Files
 * capability). The domain service applies the frozen policy:
 *   'unchanged' → save permitted
 *   'changed'   → in-place save refused
 *   'unknown'   → in-place save refused (safe default)
 *
 * Save-As is always available (targets a user-selected path).
 */
export type ExternalChangeStatus = 'unchanged' | 'changed' | 'unknown'

// ── Engine domain types ────────────────────────────────────────────────

/** Worksheet metadata returned by engine.open(). */
export interface WorksheetMetadata {
  /**
   * The stable XLSX sheet identifier — the `<sheet sheetId="...">`
   * attribute from the workbook.xml. This id is IMMUTABLE across renames
   * (a rename changes `name`, not `id`), so it is the correct key for the
   * domain sheetId → file sheet name mapping maintained by
   * SpreadsheetService.
   *
   * The renderer's Univer sheet id is mapped to this stable id at the
   * shell/compatibility boundary; the engine contract exposes it directly
   * so the service can build `sheetNames: Map<sheetId, sheetName>` from
   * `[sheet.id, sheet.name]` (mirroring the legacy runtime at
   * apps/sheets/src/main/sheets-main.ts:2805).
   */
  id: string
  /** The sheet name as it appears in the xlsx file (mutable via rename). */
  name: string
  /** Zero-based sheet index in the workbook. */
  index: number
  /** Whether the sheet is hidden. */
  hidden: boolean
  /** RTL layout. */
  rtl: boolean
  /** Gridline color (ARGB hex), if set. */
  gridlineColor?: string
  /** Whether gridlines are visible. */
  showGridlines: boolean
  /** Row count. */
  rowCount: number
  /** Column count. */
  columnCount: number
  /** Default row height in points. */
  defaultRowHeight: number
  /** Default column width in character units. */
  defaultColumnWidth: number
  /** Tab color (ARGB hex), if set. */
  tabColor?: string
  /**
   * Column width overrides (opaque array — the renderer interprets these).
   * Captured from the sidecar's open response so the save response can
   * carry them back to the renderer without loss (Increment 6).
   */
  columnWidths?: unknown[]
  /**
   * Tables in this sheet (opaque array — the renderer interprets these).
   * Captured from the sidecar's open response so the save response can
   * carry them back to the renderer without loss (Increment 6).
   */
  tables?: unknown[]
  /**
   * Cell comments/notes (opaque array — the renderer interprets these).
   * Captured from the sidecar's open response so the save response can
   * carry them back to the renderer without loss (Increment 6).
   */
  comments?: unknown[]
  /**
   * Pivot table range info (opaque array — the renderer interprets these).
   * Captured from the sidecar's open response so the save response can
   * carry them back to the renderer without loss (Increment 6).
   */
  pivotRanges?: unknown[]
}

/**
 * Workbook metadata returned by engine.open().
 *
 * Contains workbook/domain metadata ONLY — no filesystem paths, no
 * snapshot paths, no absolute paths. Those belong to ShellWorkbookSession.
 */
export interface WorkbookMetadata {
  /** The workbook name (basename, e.g. 'budget.xlsx'). */
  name: string
  /** SHA-256 hash of the workbook content. */
  sha256: string
  /** Number of ZIP entries in the archive. */
  entryCount: number
  /** Worksheet metadata for each sheet. */
  sheets: WorksheetMetadata[]
  /** Active sheet index (workbookView/@activeTab). */
  activeTab: number
  /**
   * Defined names (workbook-level named ranges).
   *
   * INCREMENT 6: Changed from `{ name: string; value: string }` to
   * `{ name: string; formula: string; sheetIndex?: number }` to match
   * the sidecar's native response shape. This eliminates the lossy
   * `formula → value` translation that discarded `sheetIndex` (the
   * localSheetId attribute for sheet-scoped names). The renderer's
   * `WorkbookFile.definedNames` expects `{ name, formula, sheetIndex? }`.
   */
  definedNames: Array<{ name: string; formula: string; sheetIndex?: number }>
  /** Theme color scheme (ARGB hex values). */
  themeColors: string[]
  /** Theme font scheme (major/minor font names). */
  themeFonts: { major: string; minor: string }
  /**
   * Cell styles (opaque array — the renderer interprets these).
   * Captured from the sidecar's open response so the save response can
   * carry them back to the renderer without loss (Increment 6).
   */
  styles?: unknown[]
  /**
   * Differential styles for conditional formatting (opaque array).
   * Captured from the sidecar's open response (Increment 6).
   */
  dxfStyles?: unknown[]
  /**
   * Visual objects (charts, images, shapes) in the workbook (opaque array).
   * Captured from the sidecar's open response (Increment 6).
   */
  visuals?: unknown[]
}

/** A cell record within a range result. */
export interface EngineCellRecord {
  row: number
  column: number
  /** The cell value as a string (formatted). */
  value: string
  /** The raw typed value, if numeric. */
  number?: number
  /** Whether the cell contains a formula. */
  isFormula: boolean
  /** Style index (0 = default). */
  styleIndex: number
  /** Hyperlink target, if set. */
  hyperlink?: string
}

/** A merged cell range. */
export interface EngineCellArea {
  firstRow: number
  firstColumn: number
  lastRow: number
  lastColumn: number
}

/** Row metadata within a range result. */
export interface EngineRowMetadata {
  row: number
  height?: number
  customHeight?: boolean
  hidden: boolean
  outlineLevel?: number
  collapsed?: boolean
  styleIndex?: number
}

/** Column metadata within a range result. */
export interface EngineColumnMetadata {
  column: number
  width?: number
  customWidth?: boolean
  hidden: boolean
  outlineLevel?: number
  collapsed?: boolean
  styleIndex?: number
}

/** Result of reading a range from the engine. */
export interface EngineRangeResult {
  cells: EngineCellRecord[]
  rows: EngineRowMetadata[]
  merges: EngineCellArea[]
  columns: EngineColumnMetadata[]
  /** Hyperlinks in the range. */
  hyperlinks: Array<{ cell: string; target: string }>
  /** Conditional formatting rules. */
  conditionalFormatting: unknown[]
  /** Data validation rules. */
  dataValidation: unknown[]
  /** Auto-filter state, if set. */
  autoFilter?: { startRow: number; startColumn: number; endRow: number; endColumn: number }
  /** Page break rows. */
  rowBreaks: number[]
  /** Page break columns. */
  columnBreaks: number[]
  /** Sheet protection state. */
  sheetProtection: boolean
}

/** A formula cell within a formula-cells result. */
export interface EngineFormulaCell {
  row: number
  column: number
  /** The formula string (without leading =). */
  formula: string
  /** The cached value, if any. */
  cachedValue?: string
}

/** Result of reading formula cells from the engine. */
export interface EngineFormulaCellsResult {
  cells: EngineFormulaCell[]
}

/** A recalculation edit (user input to apply before evaluation). */
export interface EngineRecalcEdit {
  sheetName: string
  row: number
  column: number
  /** The value to set (string for formulas, number for numeric). */
  value: string
}

/** A recalculation read request (which cells to return computed values for). */
export interface EngineRecalcRead {
  sheetName: string
  row: number
  column: number
}

/** A computed cell after recalculation. */
export interface EngineRecalcCell {
  sheetName: string
  row: number
  column: number
  /** The formatted display value. */
  formatted: string
  /** The numeric value, if numeric. */
  number?: number
  /** Whether the cell contains a formula. */
  isFormula: boolean
}

/** Result of recalculation. */
export interface EngineRecalcResult {
  cells: EngineRecalcCell[]
}

/** Result of reading media (image bytes) from the engine. */
export interface EngineMediaResult {
  /** MIME type (e.g., 'image/png'). */
  mediaType: string
  /** Base64-encoded image bytes. */
  base64: string
}

/**
 * Result of applying a SavePlan — the saved archive bytes plus the list of
 * archive entry paths that were touched (for shell-layer recovery/recent-
 * files tracking). Contains NO engine-specific archive type.
 */
export interface EngineSaveResult {
  /** The complete saved archive bytes. */
  readonly data: Uint8Array
  /** Archive entry paths that were touched (e.g., 'xl/worksheets/sheet1.xml'). */
  readonly touchedEntries: string[]
}

// NOTE (Increment 3C):
//   EngineArchivePatch has been REMOVED from runtime-contracts. It is an
//   engine-specific archive representation that must NOT leak above the
//   engine boundary. The ElectronXlsxSidecarEngine internally defines its
//   own archive-patch type (in packages/platform-electron/) and uses it
//   to translate the domain SavePlan → archive patches.
//
//   The engine contract now exposes `applySavePlan(handle, plan)`, which
//   accepts the domain SavePlan (from save-plan.ts) and returns the saved
//   archive bytes + touched entry paths. The translation is private to
//   the engine implementation.

// ── SpreadsheetEngine interface ────────────────────────────────────────

/**
 * The spreadsheet execution engine interface.
 *
 * This interface is implemented by:
 *   - ElectronXlsxSidecarEngine (current: Rust sidecar via child_process)
 *   - WasmSpreadsheetEngine (future: IronCalc compiled to WASM)
 *   - CloudSpreadsheetEngine (future: server-side computation)
 *
 * RUNTIME INDEPENDENCE:
 *   The engine accepts workbook content as Uint8Array — NOT filesystem
 *   paths. The Electron adapter may internally write the bytes to a temp
 *   file and pass the path to the Rust sidecar, but that translation is
 *   private to the adapter. A WASM engine passes the bytes directly to
 *   in-process IronCalc. A Cloud engine uploads the bytes to a server.
 *
 *   The engine accepts the domain `SavePlan` for saves — NOT engine-
 *   specific archive patches. The implementation translates the plan to
 *   its own archive format internally (Increment 3C: no EngineArchivePatch
 *   leakage above the engine boundary).
 *
 * The interface uses `Promise<T>` for all operations. It must not assume
 * a process boundary — a WASM engine returns results via async in-process
 * calls, not IPC.
 */
export interface SpreadsheetEngine {
  /**
   * Open a workbook from raw bytes. Creates a new engine session and
   * returns an opaque handle + workbook metadata.
   *
   * The handle does NOT exist before this call. All subsequent operations
   * receive the handle returned here.
   *
   * @param workbook — the raw xlsx file content (Uint8Array)
   * @param locale — the UI locale for formula/function name resolution
   * @param fileName — the workbook file name (basename, e.g. 'budget.xlsx')
   */
  open(
    workbook: Uint8Array,
    locale: string,
    fileName: string,
  ): Promise<{
    handle: EngineSessionHandle
    metadata: WorkbookMetadata
  }>

  /**
   * Read a range of cells from a worksheet.
   * @param handle — opaque engine session handle (from open())
   * @param sheetName — the file sheet name (NOT a renderer sheet id)
   * @param range — the cell range (e.g., 'A1:Z100')
   */
  readRange(
    handle: EngineSessionHandle,
    sheetName: string,
    range: string,
  ): Promise<EngineRangeResult>

  /**
   * Read all formula cells from a worksheet.
   * @param handle — opaque engine session handle
   * @param sheetName — the file sheet name
   */
  readFormulaCells(
    handle: EngineSessionHandle,
    sheetName: string,
  ): Promise<EngineFormulaCellsResult>

  /**
   * Recalculate formulas. Applies pending edits as user input, evaluates,
   * and returns computed values for the requested cells.
   * @param handle — opaque engine session handle
   * @param edits — the edits to apply before evaluation
   * @param reads — the cells to return computed values for
   */
  recalculate(
    handle: EngineSessionHandle,
    edits: EngineRecalcEdit[],
    reads: EngineRecalcRead[],
  ): Promise<EngineRecalcResult>

  /**
   * Read media (image bytes) from the workbook.
   * @param handle — opaque engine session handle
   * @param visualId — the visual object identifier
   */
  readMedia(
    handle: EngineSessionHandle,
    visualId: string,
  ): Promise<EngineMediaResult>

  /**
   * Apply a domain SavePlan and return the saved archive bytes.
   *
   * The engine implementation translates the SavePlan to its own internal
   * archive representation (e.g., EngineArchivePatch[] for the Rust sidecar,
   * an in-memory entry map for WASM) and produces the complete archive bytes.
   * The translation is PRIVATE to the implementation — the runtime-
   * independent contract does NOT expose any engine-specific archive type.
   *
   * The caller (SpreadsheetService) is responsible for validating all
   * sheetIds in the SavePlan against `session.sheetNames` (fail-closed)
   * BEFORE calling this method. The engine receives a plan whose sheetIds
   * have already been resolved to file sheet names where applicable —
   * HOWEVER, the engine MAY also perform its own validation.
   *
   * @param handle — opaque engine session handle
   * @param plan — the domain save plan (sheetOps, edits, structuralOps, etc.)
   * @returns the saved archive bytes + touched entry paths
   */
  applySavePlan(
    handle: EngineSessionHandle,
    plan: SavePlan,
  ): Promise<EngineSaveResult>

  /**
   * Convert a legacy workbook (.xls) to .xlsx format.
   *
   * @param workbook — the raw legacy file content (Uint8Array)
   * @param fileName — the legacy file name (e.g. 'old.xls')
   * @returns the converted .xlsx content as bytes + the new file name
   */
  convertWorkbook(
    workbook: Uint8Array,
    fileName: string,
  ): Promise<{
    data: Uint8Array
    fileName: string
  }>

  /**
   * Close an engine session. The handle becomes invalid after this call.
   * @param handle — opaque engine session handle
   */
  close(handle: EngineSessionHandle): Promise<void>

  /**
   * Read a single archive entry (XML text) from the workbook's on-disk
   * temp file. Used by pivot-definition reads: the handler needs the raw
   * XML from specific parts (e.g. xl/pivotTables/pivotTable1.xml) that
   * are not exposed by the structured read methods.
   *
   * @param handle — opaque engine session handle (determines which temp file)
   * @param entryName — the archive entry path (e.g. 'xl/pivotTables/pivotTable1.xml')
   * @returns the entry's text content (UTF-8)
   *
   * THROWS on failure:
   *   - InvalidSessionError — handle was closed or never opened
   *   - InvalidInputError    — entry not found in the archive
   *   - EngineError          — engine/protocol failure
   */
  readArchiveEntry(
    handle: EngineSessionHandle,
    entryName: string,
  ): Promise<string>

  /**
   * Stop the engine entirely. Kills any background processes, releases
   * all resources. Called on app shutdown.
   */
  stop(): Promise<void>
}
