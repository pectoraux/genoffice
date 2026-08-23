/**
 * SpreadsheetServiceImpl — the Sheets domain service.
 *
 * Composes SpreadsheetEngine (runtime-independent). Owns domain semantics:
 * workbook open/read/recalc/save, sheet-id translation (fail-closed),
 * external-change policy, recovery path derivation.
 *
 * ZERO node:* / Electron imports (verified by architecture test).
 * ZERO shell-hook deps (no wcId, no BrowserWindow, no dialogs).
 * Session-scoped: open returns { session, engineHandle, metadata };
 * subsequent operations receive session + engineHandle.
 *
 * ERROR MODEL (Increment 3A correction):
 *   The service preserves typed engine/domain failures. It does NOT silently
 *   convert every engine exception into `null` or `{ ok: false }`:
 *     - open()          → throws EngineError | InvalidSessionError | InvalidInputError
 *     - close()         → throws EngineError | InvalidSessionError
 *     - writeRecovery() → throws EngineError | InvalidSessionError
 *   Only `save()` returns a soft-failure result, because external-change
 *   policy refusal ({ ok: false, reason: 'external-modified' }) is a
 *   legitimate business outcome. Engine failures during save still throw.
 *
 * SHEET-ID MAPPING (Increment 3B correction):
 *   The service builds `sheetNames: Map<sheetId, sheetName>` from
 *   `[sheet.id, sheet.name]` — the stable XLSX sheetId attribute, NOT the
 *   mutable sheet name. This mirrors the legacy runtime at
 *   sheets-main.ts:2805.
 *
 * UNKNOWN SHEET FAIL-CLOSED (Increment 3B correction):
 *   All operations that accept a `sheetId` validate it against
 *   `session.sheetNames` BEFORE delegation. Unknown sheetIds →
 *   `InvalidInputError`. This mirrors the legacy runtime at
 *   sheets-main.ts:1787 (recalc), 2545 and 2554 (save).
 *   The forbidden `if (!sheetName) return sheetId` fallback is gone.
 *
 * SAVE DOMAIN MODEL (Increment 3B + 3C correction):
 *   The service accepts a domain `SavePlan` (defined in save-plan.ts).
 *   The service validates all sheetIds in the plan (fail-closed), then
 *   delegates to `engine.applySavePlan(handle, plan)`. The engine
 *   implementation translates the SavePlan to its own internal archive
 *   format PRIVATELY. The runtime-independent contract does NOT expose
 *   `EngineArchivePatch` or any engine-specific archive type (Increment 3C
 *   removed the SavePlanTranslator and SavePlanTranslation abstractions —
 *   the translation is now entirely below the engine boundary).
 *
 * DOMAIN-EVENT PURITY (Increment 3A correction):
 *   The service does NOT own renderer/event routing. It exposes NO
 *   `onOpened`, `onRenamed`, `onTeardown`, `SheetsEventBus`, or
 *   `{ oldPath, newPath }` payloads. The shell coordinator owns:
 *     - `docs/workbook opened` notification
 *     - `renamed` notification
 *     - `teardown` notification
 *     - renderer notification dispatch
 *
 * IMPORTANT (ADR-001 Correction A): constructor injection. No getRuntime().
 */

import type {
  SpreadsheetEngine,
  EngineSessionHandle,
  ExternalChangeStatus,
  EngineRangeResult,
  EngineFormulaCellsResult,
  EngineRecalcEdit,
  EngineRecalcRead,
  EngineRecalcResult,
  EngineMediaResult,
  WorkbookPivotDefinition,
} from '@genoffice/runtime-contracts'
import { InvalidInputError } from '@genoffice/runtime-contracts'
import type {
  SpreadsheetService,
  SpreadsheetServiceDeps,
  WorkbookSession,
  WorkbookOpenResult,
  SaveRequest,
  SaveResult,
  SavePlan,
} from '@genoffice/runtime-contracts'

// ── Implementation ────────────────────────────────────────────────────

export class SpreadsheetServiceImpl implements SpreadsheetService {
  constructor(private readonly deps: SpreadsheetServiceDeps) {}

  // ── Workbook lifecycle ──────────────────────────────────────────────

  async open(
    workbook: Uint8Array,
    locale: string,
    fileName: string,
  ): Promise<WorkbookOpenResult> {
    // Delegate to the engine. Typed engine failures (InvalidInputError,
    // InvalidSessionError, EngineError) propagate to the caller.
    const { handle, metadata } = await this.deps.engine.open(workbook, locale, fileName)

    // Build the domain session — no engine handle, no snapshot path,
    // no absolute filesystem path. workbookName is a basename only.
    //
    // SHEET-ID MAPPING (Increment 3B): build sheetNames from [sheet.id, sheet.name].
    // The `id` is the stable XLSX sheetId attribute (immutable across renames);
    // the `name` is the visible tab name (mutable). This mirrors the legacy
    // runtime at sheets-main.ts:2805:
    //   sheetNames: new Map(opened.sheets.map((sheet) => [sheet.id, sheet.name]))
    const sheetNames = new Map<string, string>()
    for (const sheet of metadata.sheets) {
      sheetNames.set(sheet.id, sheet.name)
    }

    const session: WorkbookSession = {
      workbookName: fileName,
      workbookHash: metadata.sha256,
      sheetNames,
    }

    return {
      session,
      engineHandle: handle,
      metadata,
    }
  }

  async close(engineHandle: EngineSessionHandle): Promise<void> {
    // Delegate to the engine. Typed engine failures (InvalidSessionError,
    // EngineError) propagate to the caller — do NOT swallow them as
    // { ok: false }. The caller must distinguish invalid-session from
    // protocol failure from engine failure.
    await this.deps.engine.close(engineHandle)
  }

  // ── Workbook operations ─────────────────────────────────────────────

  async readRange(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    sheetId: string,
    range: string,
  ): Promise<EngineRangeResult> {
    // FAIL-CLOSED: unknown sheetId → InvalidInputError (NOT return sheetId).
    // Mirrors the legacy runtime at sheets-main.ts:1787.
    const sheetName = this.resolveSheetNameOrThrow(session, sheetId)
    return this.deps.engine.readRange(engineHandle, sheetName, range)
  }

  async readFormulaCells(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    sheetId: string,
  ): Promise<EngineFormulaCellsResult> {
    const sheetName = this.resolveSheetNameOrThrow(session, sheetId)
    return this.deps.engine.readFormulaCells(engineHandle, sheetName)
  }

  async recalculate(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    edits: EngineRecalcEdit[],
    reads: EngineRecalcRead[],
  ): Promise<EngineRecalcResult> {
    // Resolve domain sheet ids → engine sheet names (fail-closed).
    const resolvedEdits = edits.map((e) => ({
      ...e,
      sheetName: this.resolveSheetNameOrThrow(session, e.sheetName),
    }))
    const resolvedReads = reads.map((r) => ({
      ...r,
      sheetName: this.resolveSheetNameOrThrow(session, r.sheetName),
    }))
    return this.deps.engine.recalculate(engineHandle, resolvedEdits, resolvedReads)
  }

  async readMedia(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    visualId: string,
  ): Promise<EngineMediaResult> {
    // The engineHandle is the complete session scope — the sidecar maps
    // engineHandle → sidecar sessionId internally, and visualId is scoped
    // to that sessionId. Cross-session misuse (session A's visualId with
    // session B's engineHandle) fails at the engine level (visualId not
    // found). The service does NOT validate session ↔ engineHandle binding.
    //
    // The `session` parameter is accepted for API consistency with
    // readRange/readFormulaCells/recalculate. It is reserved for future
    // domain-level validation but not currently used for safety.
    void session
    return this.deps.engine.readMedia(engineHandle, visualId)
  }

  // ── Save ────────────────────────────────────────────────────────────

  async save(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    request: SaveRequest,
    externalChange: ExternalChangeStatus,
  ): Promise<SaveResult> {
    // Apply the frozen external-change policy. This is the ONLY legitimate
    // soft-failure outcome — refused in-place save is a business decision,
    // NOT an error. The shell prompts the user to Save-As.
    if (externalChange === 'changed' || externalChange === 'unknown') {
      return { ok: false, reason: 'external-modified' }
    }
    // externalChange === 'unchanged' → proceed with save.

    // Validate ALL sheetIds in the SavePlan (fail-closed). Unknown sheetIds
    // in ANY mutation family → InvalidInputError. This mirrors the legacy
    // runtime at sheets-main.ts:2544-2545, 2553-2554.
    this.validateSavePlanSheetIds(request.plan, session.sheetNames)

    // Delegate to the engine. The engine implementation translates the
    // SavePlan to its own internal archive representation PRIVATELY — no
    // EngineArchivePatch leaks above the engine boundary (Increment 3C).
    // Typed engine failures (InvalidSessionError, InvalidInputError,
    // EngineError) propagate to the caller.
    const result = await this.deps.engine.applySavePlan(engineHandle, request.plan)
    return {
      ok: true,
      data: result.data,
      touchedEntries: result.touchedEntries,
    }
  }

  async writeRecovery(
    session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    request: SaveRequest,
  ): Promise<Uint8Array> {
    // Validate ALL sheetIds in the SavePlan (fail-closed).
    this.validateSavePlanSheetIds(request.plan, session.sheetNames)

    // Delegate to the engine. The engine translates the SavePlan to its
    // internal archive format PRIVATELY. Typed engine failures propagate
    // to the caller — do NOT swallow them as { ok: false }.
    const result = await this.deps.engine.applySavePlan(engineHandle, request.plan)
    return result.data
  }

  async readPivotDefinition(
    _session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    pivotTablePath: string,
    cacheDefinitionPath: string,
  ): Promise<WorkbookPivotDefinition> {
    // The engine is the SINGLE translation point between the OOXML wire
    // format and the runtime-independent `WorkbookPivotDefinition`. It
    // reads both XML parts from its on-disk temp file (private to the
    // adapter) and parses them via the canonical xlsx-gateway parser.
    //
    // The service performs NO archive I/O and NO parsing — it delegates
    // directly. This closes two prior defects:
    //   (a) The service used to call `engine.readArchiveEntry()` (a
    //       generic ZIP-entry escape-hatch on the engine contract).
    //   (b) The service used to `await import('@genoffice/xlsx-gateway/...')`
    //       at runtime — a violation of the services-sheets architecture
    //       test (which forbids xlsx-gateway imports).
    //
    // The `session` parameter is accepted for API consistency with
    // readRange / readFormulaCells / recalculate / readMedia. The
    // engineHandle is the complete session scope.
    void _session
    return this.deps.engine.readPivotDefinition(
      engineHandle,
      pivotTablePath,
      cacheDefinitionPath,
    )
  }

  // ── Internal: sheet-id translation (fail-closed) ───────────────────

  /**
   * Resolve a domain sheetId to the engine's file sheet name.
   * The service owns this translation — the engine never sees domain sheet ids.
   *
   * FAIL-CLOSED: unknown sheetId → InvalidInputError (NOT return sheetId).
   * Mirrors the legacy runtime at sheets-main.ts:1785-1789:
   *   const fileSheetName = (sheetId: string): string => {
   *     const name = session.sheetNames.get(sheetId)
   *     if (name === undefined) throw new Error(`Unknown sheet for recalculation: ${sheetId}`)
   *     return name
   *   }
   */
  private resolveSheetNameOrThrow(session: WorkbookSession, sheetId: string): string {
    const sheetName = session.sheetNames.get(sheetId)
    if (sheetName === undefined) {
      throw new InvalidInputError(`Unknown sheetId: ${sheetId}`)
    }
    return sheetName
  }

  /**
   * Validate ALL sheetIds in a SavePlan against the session's sheetNames map.
   * Unknown sheetIds in ANY mutation family → InvalidInputError (fail-closed).
   *
   * This mirrors the legacy runtime at sheets-main.ts:2544-2545, 2553-2554,
   * 2578, 2589, 2617, 2624, 2634, 2638, 2642, 2645, 2650, 2654, 2658, 2665,
   * 2674, 2696, 2706, 2749, 2756 — every mutation family resolves sheetIds
   * via `session.sheetNames.get(sheetId)` and throws on unknown.
   *
   * Note: sheetOps with kind 'add-sheet' or 'duplicate-sheet' introduce NEW
   * sheetIds that are not in the session map. The legacy runtime tracks these
   * in `addedSheetNames` (sheets-main.ts:2519, 2528, 2536). The engine
   * implementation handles added sheets internally — the service validates
   * only KNOWN sheetIds (those expected to already exist in the session map).
   */
  private validateSavePlanSheetIds(plan: SavePlan, sheetNames: ReadonlyMap<string, string>): void {
    // Cell edits
    for (const edit of plan.edits) {
      if (!sheetNames.has(edit.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in edit: ${edit.sheetId}`)
      }
    }
    // Structural ops
    for (const op of plan.structuralOps) {
      if (!sheetNames.has(op.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in structural op: ${op.sheetId}`)
      }
    }
    // Formula values
    for (const fv of plan.formulaValues) {
      if (!sheetNames.has(fv.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in formula value: ${fv.sheetId}`)
      }
    }
    // Sheet ops (validate non-add/duplicate — added sheets are new)
    for (const op of plan.sheetOps) {
      if (op.kind === 'add-sheet' || op.kind === 'duplicate-sheet') {
        // Added sheets introduce new sheetIds — skip validation.
        // The engine handles added sheets via the `addedSheetNames` map
        // (mirroring the legacy runtime at sheets-main.ts:2519, 2528, 2536).
        // For duplicate-sheet, validate the SOURCE sheetId exists.
        if (op.kind === 'duplicate-sheet' && op.sourceSheetId !== undefined) {
          if (!sheetNames.has(op.sourceSheetId)) {
            throw new InvalidInputError(`Unknown source sheetId in duplicate: ${op.sourceSheetId}`)
          }
        }
        continue
      }
      // rename/remove/set-sheet-hidden/reorder: sheetId must exist
      if (!sheetNames.has(op.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in sheet op (${op.kind}): ${op.sheetId}`)
      }
    }
    // Filter states
    for (const s of plan.filterStates) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in filter state: ${s.sheetId}`)
      }
    }
    // Hyperlink edits
    for (const link of plan.hyperlinkEdits) {
      if (!sheetNames.has(link.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in hyperlink edit: ${link.sheetId}`)
      }
    }
    // CF states
    for (const s of plan.cfStates) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in CF state: ${s.sheetId}`)
      }
    }
    // DV states
    for (const s of plan.dvStates) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in DV state: ${s.sheetId}`)
      }
    }
    // Page setup states
    for (const s of plan.pageSetupStates) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in page setup: ${s.sheetId}`)
      }
    }
    // Note states
    for (const s of plan.noteStates) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in note state: ${s.sheetId}`)
      }
    }
    // Sheet protections
    for (const s of plan.sheetProtections) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in sheet protection: ${s.sheetId}`)
      }
    }
    // Protected range states
    for (const s of plan.protectedRangeStates) {
      if (!sheetNames.has(s.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in protected range: ${s.sheetId}`)
      }
    }
    // Visual additions
    for (const add of plan.visualAdditions) {
      if (!sheetNames.has(add.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in visual addition: ${add.sheetId}`)
      }
    }
    // Table additions
    for (const add of plan.tableAdditions) {
      if (!sheetNames.has(add.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in table addition: ${add.sheetId}`)
      }
    }
    // Pivot additions (sheetId + sourceSheetId)
    for (const add of plan.pivotAdditions) {
      if (!sheetNames.has(add.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in pivot addition: ${add.sheetId}`)
      }
      if (!sheetNames.has(add.sourceSheetId)) {
        throw new InvalidInputError(`Unknown source sheetId in pivot addition: ${add.sourceSheetId}`)
      }
    }
    // Sparkline additions
    for (const add of plan.sparklineAdditions) {
      if (!sheetNames.has(add.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in sparkline addition: ${add.sheetId}`)
      }
    }
    // Pivot refresh updates
    for (const upd of plan.pivotRefreshUpdates) {
      if (!sheetNames.has(upd.sheetId)) {
        throw new InvalidInputError(`Unknown sheetId in pivot refresh update: ${upd.sheetId}`)
      }
    }
    // Note: chartEdits, visualEdits, pivotCacheRefreshPaths,
    // definedNamesState, themeState, workbookProtectionState are NOT
    // sheetId-keyed — they use package-absolute paths or workbook-level
    // state. No sheetId validation needed for these.
  }
}
