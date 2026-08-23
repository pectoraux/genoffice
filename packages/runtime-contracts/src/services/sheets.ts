/**
 * SpreadsheetService — domain runtime service for the sheets (`.xlsx`) editor.
 *
 * ADR-004 / Phase 2 Architecture (Increment 3C correction):
 *   The service owns DOMAIN semantics only — workbook open/read/recalc/save.
 *   It delegates engine operations to SpreadsheetEngine (opaque handle).
 *   It receives dependencies via constructor injection — no getRuntime().
 *
 *   The service does NOT own:
 *     - File dialogs (shell owns Files.pickOpen / Files.pickSave)
 *     - BrowserWindow / WebContents (shell owns window management)
 *     - wcId lookup (shell coordinator owns the session registry)
 *     - Snapshot management (shell owns snapshot paths)
 *     - Disk fingerprint computation (shell observes filesystem state)
 *     - Recovery UI dialogs (shell owns Restore/Discard prompts)
 *     - child_process spawning (engine adapter owns sidecar lifecycle)
 *     - Renderer event routing / lifecycle notifications (shell coordinator
 *       owns `docs/workbook opened`, `renamed`, `teardown` notifications)
 *     - SavePlan → archive-patch translation (engine implementation owns
 *       this internally — Increment 3C removed the runtime-contract
 *       SavePlanTranslator and EngineArchivePatch)
 *
 *   The service DOES own:
 *     - Workbook open/save semantics
 *     - Sheet-id translation (domain sheetId ↔ engine sheet name) — FAIL-CLOSED
 *       on unknown sheetIds (throws InvalidInputError)
 *     - ExternalChangeStatus policy (unchanged → save; changed/unknown → refuse)
 *     - Save-plan validation (sheetId resolution before delegation)
 *     - Engine coordination (delegates to SpreadsheetEngine.applySavePlan)
 *     - Recovery path derivation (pure computation, no filesystem)
 *
 * SAVE DOMAIN MODEL (Increment 3B + 3C):
 *   The service accepts a domain `SavePlan` (defined in save-plan.ts). The
 *   SavePlan preserves ALL renderer-independent Sheets mutation families
 *   (sheetOps, edits, structuralOps, filterStates, hyperlinkEdits, cfStates,
 *   dvStates, pageSetupStates, noteStates, formulaValues, visualAdditions,
 *   tableAdditions, pivotAdditions, sparklineAdditions, pivotRefreshUpdates,
 *   sheetProtections, protectedRangeStates, definedNamesState, themeState,
 *   workbookProtectionState, chartEdits, visualEdits). This mirrors the
 *   legacy `WorkbookSaveRequest` (apps/sheets/src/shared/desktop-api.ts:1476)
 *   but as domain types, not Zod schemas.
 *
 *   The service validates all sheetIds in the SavePlan (fail-closed), then
 *   delegates to `engine.applySavePlan(handle, plan)`. The engine
 *   implementation translates the SavePlan to its own internal archive
 *   representation PRIVATELY. The runtime-independent contract does NOT
 *   expose `EngineArchivePatch` or any engine-specific archive type.
 *
 * SHEET-ID MAPPING (Increment 3B):
 *   The service builds `sheetNames: Map<sheetId, sheetName>` from
 *   `[sheet.id, sheet.name]` (NOT `[sheet.name, sheet.name]`). The `id`
 *   is the stable XLSX sheetId attribute (immutable across renames);
 *   `name` is the visible tab name (mutable). Unknown sheetIds in any
 *   operation → `InvalidInputError` (fail-closed, mirroring the legacy
 *   runtime at sheets-main.ts:1785-1789 and 2544-2545).
 *
 * IMPORTANT (ADR-001 Correction A): constructor injection. No getRuntime().
 */

import type {
  SpreadsheetEngine,
  EngineSessionHandle,
  ExternalChangeStatus,
  WorkbookMetadata,
  EngineRangeResult,
  EngineFormulaCellsResult,
  EngineRecalcEdit,
  EngineRecalcRead,
  EngineRecalcResult,
  EngineMediaResult,
} from './spreadsheet-engine.js'

// Re-export all SavePlan domain types so callers can construct save requests
// without importing from two files. These types are defined in save-plan.ts
// (separate file to avoid a circular import: spreadsheet-engine.ts needs
// SavePlan, and sheets.ts needs SpreadsheetEngine from spreadsheet-engine.ts).
export type {
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
} from './save-plan.js'
import type { SavePlan } from './save-plan.js'

// ── Domain session ───────────────────────────────────────────────────

/**
 * Domain-level workbook session.
 *
 * Contains ONLY domain data — no filesystem paths, no engine handles,
 * no snapshot paths, no disk fingerprints. Those belong to
 * ShellWorkbookSession (shell coordinator).
 *
 * Note: the field `workbookName` is a basename (e.g. 'budget.xlsx'), NOT an
 * absolute filesystem path. Filesystem paths remain exclusively in the
 * shell layer (ShellWorkbookSession).
 */
export interface WorkbookSession {
  /**
   * The workbook name (basename, e.g. 'budget.xlsx').
   * This is a domain identifier — NOT a filesystem path. The shell
   * layer (ShellWorkbookSession) owns the absolute path separately.
   */
  readonly workbookName: string
  /** SHA-256 hash of the workbook content at open time. */
  readonly workbookHash: string
  /**
   * Domain sheetId → file sheet name mapping.
   * Built from `[sheet.id, sheet.name]` (NOT `[sheet.name, sheet.name]`).
   * The `id` is the stable XLSX sheetId attribute (immutable across renames).
   */
  readonly sheetNames: ReadonlyMap<string, string>
}

/**
 * Result of opening a workbook.
 *
 * `engineHandle` is an opaque engine context token. Callers MUST NOT
 * inspect, compare, serialize, or construct one. The only way to obtain
 * an `EngineSessionHandle` is as the return value of `service.open()`.
 *
 * The handle is included in this result so that the shell coordinator
 * (which owns the WorkbookSession ↔ engineHandle mapping) can pass it
 * to subsequent service operations. The shell stores it inside
 * `ShellWorkbookSession.engineHandle`; the domain WorkbookSession above
 * does NOT contain it.
 *
 * The handle exposes NO sidecar UUID, NO engineSessionId, NO
 * implementation details — only an opaque brand symbol. Any attempt to
 * read its fields via `Object.keys()`, `Reflect.ownKeys()`, or similar
 * reflection returns nothing useful.
 */
export interface WorkbookOpenResult {
  /** Domain-level session (workbookName, hash, sheetNames). */
  session: WorkbookSession
  /**
   * Opaque engine context token — pass to subsequent service operations.
   * Not inspectable, not serializable, not constructable by callers.
   */
  engineHandle: EngineSessionHandle
  /** Workbook metadata from the engine. */
  metadata: WorkbookMetadata
}

/**
 * The domain save request — a rich SavePlan preserving all mutation families.
 *
 * This REPLACES the Increment 3A `SaveRequest = EngineArchivePatch[]`,
 * which discarded renderer-independent Sheets mutation semantics. The
 * service now receives a domain SavePlan, validates sheetIds (fail-closed),
 * resolves them to file sheet names, and delegates to
 * `engine.applySavePlan(handle, plan)`.
 */
export interface SaveRequest {
  /** The domain save plan (sheetOps, edits, structuralOps, etc.). */
  readonly plan: SavePlan
}

/**
 * Result of a save operation.
 *
 * `ok: false` is a LEGITIMATE business outcome, NOT an error — it
 * indicates that in-place save was refused because the external file
 * changed (or its state is unknown). The shell prompts the user to
 * Save-As instead.
 *
 * Engine failures (InvalidSessionError, EngineError) do NOT produce
 * `ok: false` — they propagate as typed errors. The caller can
 * distinguish:
 *   - externalChange policy refusal → { ok: false, reason: 'external-modified' }
 *   - engine failure              → throws EngineError | InvalidSessionError
 *
 * `touchedEntries` is a list of archive-entry path strings (e.g.
 * 'xl/worksheets/sheet1.xml') — NOT an engine-specific structure.
 */
export interface SaveResult {
  /** true when the save succeeded; false when refused by external-change policy. */
  ok: boolean
  /**
   * Present when ok === false. Currently always 'external-modified'
   * (the only legitimate soft-failure reason).
   */
  reason?: 'external-modified'
  /** The saved workbook bytes — present when ok === true. */
  data?: Uint8Array
  /** Archive entry paths that were touched — present when ok === true. */
  touchedEntries?: string[]
}

// ── Service dependencies ─────────────────────────────────────────────

/**
 * Dependencies for SpreadsheetServiceImpl.
 *
 * The service receives only the SpreadsheetEngine — it does NOT need a
 * SavePlanTranslator (Increment 3C removed that abstraction). The engine
 * accepts the domain SavePlan directly via `applySavePlan(handle, plan)`
 * and translates it to its own internal archive format privately.
 */
export interface SpreadsheetServiceDeps {
  /** The spreadsheet execution engine (injected — runtime chooses impl). */
  readonly engine: SpreadsheetEngine
}

// ── Service interface ───────────────────────────────────────────────

/**
 * The runtime-independent spreadsheet domain service.
 *
 * Uses SpreadsheetEngine for workbook I/O and computation.
 * Uses Storage / Files capabilities for persistence (via the shell).
 *
 * The service NEVER touches Electron, Node builtins, or filesystem paths
 * directly. It receives bytes and returns bytes. The shell persists them.
 *
 * ERROR MODEL:
 *   - `open()`, `close()`, `writeRecovery()` throw typed errors on failure
 *     (EngineError, InvalidSessionError, InvalidInputError). They do NOT
 *     return null or `{ ok: false }` — the caller must catch typed errors.
 *   - `save()` returns `SaveResult` because external-change policy refusal
 *     is a legitimate business outcome. Engine failures still throw.
 *   - `readRange`, `readFormulaCells`, `recalculate`, `readMedia` throw
 *     typed errors on engine failure (no swallowing).
 *
 * SHEET-ID FAIL-CLOSED:
 *   All operations that accept a `sheetId` (readRange, readFormulaCells,
 *   recalculate, readMedia, save, writeRecovery) validate the sheetId
 *   against `session.sheetNames` BEFORE delegation. Unknown sheetIds →
 *   `InvalidInputError` (mirrors the legacy runtime at sheets-main.ts:1787,
 *   2545, 2554).
 *
 * MEDIA SESSION SAFETY:
 *   `readMedia` accepts `session` + `engineHandle` for API consistency.
 *   The engineHandle is the complete session scope — the sidecar maps
 *   engineHandle → sidecar sessionId internally, and visualId is scoped
 *   to that sessionId. Cross-session misuse (passing session A's
 *   visualId with session B's engineHandle) fails at the engine level
 *   (visualId not found in session B's sidecar session). The service
 *   does NOT need to validate session ↔ engineHandle binding — the
 *   engine's own session isolation enforces it.
 */
export interface SpreadsheetService {
  // ── Workbook lifecycle ──

  /**
   * Open a workbook from raw bytes. Internally calls engine.open() which
   * creates the opaque handle. Returns domain session + engine handle + metadata.
   *
   * The sheetNames map is built from `[sheet.id, sheet.name]` — the stable
   * XLSX sheetId attribute, NOT the mutable sheet name. This preserves
   * the legacy mapping at sheets-main.ts:2805.
   *
   * THROWS on failure (does NOT return null):
   *   - InvalidInputError     — workbook bytes are not a valid xlsx
   *   - InvalidSessionError   — engine could not establish a session
   *   - EngineError           — engine failure (INTERNAL_ERROR) or
   *                             protocol failure (PROTOCOL_ERROR)
   */
  open(
    workbook: Uint8Array,
    locale: string,
    fileName: string,
  ): Promise<WorkbookOpenResult>

  /**
   * Close an engine session. The handle becomes invalid after this call.
   *
   * THROWS on failure (does NOT return `{ ok: false }`):
   *   - InvalidSessionError — handle was already closed or never opened
   *   - EngineError         — engine failure or protocol failure
   */
  close(engineHandle: EngineSessionHandle): Promise<void>

  // ── Workbook operations ──

  /**
   * Read a range of cells. The service resolves domain sheetId →
   * engine sheet name using the session's sheetNames map.
   * Unknown sheetId → InvalidInputError (fail-closed).
   *
   * THROWS on engine failure (InvalidSessionError, EngineError).
   */
  readRange(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    sheetId: string,
    range: string,
  ): Promise<EngineRangeResult>

  /**
   * Read all formula cells from a worksheet.
   * Unknown sheetId → InvalidInputError (fail-closed).
   *
   * THROWS on engine failure (InvalidSessionError, EngineError).
   */
  readFormulaCells(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    sheetId: string,
  ): Promise<EngineFormulaCellsResult>

  /**
   * Recalculate formulas. The service resolves domain sheetIds → engine
   * sheet names before delegating to the engine.
   * Unknown sheetId → InvalidInputError (fail-closed).
   *
   * THROWS on engine failure (InvalidSessionError, InvalidInputError, EngineError).
   */
  recalculate(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    edits: EngineRecalcEdit[],
    reads: EngineRecalcRead[],
  ): Promise<EngineRecalcResult>

  /**
   * Read media (image bytes) from the workbook.
   *
   * The `visualId` is scoped to the engine session (the sidecar maps
   * engineHandle → sidecar sessionId; visualId is unique within that
   * sessionId). The `session` parameter is accepted for API consistency
   * with readRange/readFormulaCells/recalculate. Cross-session misuse
   * (session A's visualId with session B's engineHandle) fails at the
   * engine level — the service does not need to validate the binding.
   *
   * THROWS on engine failure (InvalidSessionError, EngineError).
   */
  readMedia(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    visualId: string,
  ): Promise<EngineMediaResult>

  // ── Save ──

  /**
   * Save the workbook. The service applies the frozen external-change policy:
   *   'unchanged' → save permitted, returns archive bytes
   *   'changed'   → in-place save refused ({ ok: false, reason: 'external-modified' })
   *   'unknown'   → in-place save refused (safe default)
   *
   * The service validates all sheetIds in the SavePlan against
   * `session.sheetNames` (fail-closed → InvalidInputError) before
   * delegation. Unknown sheetIds in ANY mutation family (edits,
   * structuralOps, sheetOps, filterStates, etc.) → InvalidInputError.
   *
   * The service delegates to `engine.applySavePlan(handle, plan)`, which
   * internally translates the SavePlan to the engine's own archive format
   * and produces the saved bytes + touched entry paths. The translation is
   * PRIVATE to the engine implementation — no EngineArchivePatch leaks
   * above the engine boundary.
   *
   * RETURNS `SaveResult` — the external-change refusal is a legitimate
   * business outcome (NOT an error). Engine failures (InvalidSessionError,
   * EngineError) PROPAGATE as typed errors — they do NOT produce
   * `{ ok: false }`.
   */
  save(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    request: SaveRequest,
    externalChange: ExternalChangeStatus,
  ): Promise<SaveResult>

  /**
   * Write a recovery copy. The service validates all sheetIds in the
   * SavePlan (fail-closed), delegates to `engine.applySavePlan(handle, plan)`,
   * and returns the saved bytes. It does NOT write to a filesystem path —
   * the shell persists the bytes.
   *
   * THROWS on failure (does NOT return `{ ok: false }`):
   *   - InvalidInputError    — unknown sheetId in the SavePlan
   *   - InvalidSessionError  — handle was closed or never opened
   *   - EngineError          — engine failure or protocol failure
   *
   * @returns the recovery archive bytes (the shell persists them)
   */
  writeRecovery(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    request: SaveRequest,
  ): Promise<Uint8Array>

  /**
   * Read a pivot table definition from the workbook's archive.
   *
   * The service reads the pivotTable XML and pivotCacheDefinition XML
   * from the engine's temp file via `engine.readArchiveEntry()`, then
   * parses them via the canonical @genoffice/xlsx-gateway
   * `parsePivotDefinition()` — runtime-independent, no duplication.
   *
   * THROWS on failure:
   *   - InvalidSessionError — handle was closed or never opened
   *   - InvalidInputError    — entry not found or malformed XML
   *   - EngineError          — engine/protocol failure
   *
   * @returns the parsed pivot definition (typed by xlsx-gateway)
   */
  readPivotDefinition(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    pivotTablePath: string,
    cacheDefinitionPath: string,
  ): Promise<unknown>
}
