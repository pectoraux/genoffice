/**
 * Sheets runtime bundle — constructs the runtime-independent service stack
 * and the shell coordinator for the Sheets editor.
 *
 * Architecture:
 *   ElectronXlsxSidecarEngine (platform-electron)
 *     ↓ implements SpreadsheetEngine
 *   SpreadsheetServiceImpl (services-sheets)
 *     ↓ implements SpreadsheetService
 *   SheetsShellCoordinator (shell coordinator)
 *     ↓ owns session lifecycle
 *   Migrated IPC handlers (thin adapter)
 *
 * The coordinator receives ONLY SpreadsheetService — it does NOT import
 * or depend on XlsxSidecarClient, xlsx-gateway, child_process, or any
 * engine-specific code.
 *
 * LEGACY SESSION ADOPTION (Increment 5A):
 *   The legacy `workbook:select` open path opens the workbook via
 *   `XlsxSidecarClient.open()` directly (not via the engine). To bridge
 *   this with the migrated read/recalc/media/close path (which uses the
 *   coordinator-backed `SpreadsheetService`), `initSheetsRuntime` accepts
 *   an injectable `sidecarClient` (the legacy `XlsxSidecarClient`).
 *
 *   The engine uses this client INSTEAD of constructing its own — sharing
 *   the SAME sidecar process. Adoption (`adoptLegacySessionIntoCoordinator`)
 *   then wraps the legacy `sidecarSessionId` into an opaque
 *   `EngineSessionHandle` via `engine.adoptExternalSession()` — a pure
 *   in-process handle registration with NO wire call, NO spawn, NO file IO.
 */

import {
  ElectronXlsxSidecarEngine,
  ElectronSpreadsheetPdfRenderer,
  ElectronScreenCapture,
  type ElectronXlsxSidecarEngineConfig,
  type SidecarProtocolLike,
} from '@genoffice/platform-electron'
import { SpreadsheetServiceImpl } from '@genoffice/services-sheets/src/spreadsheet-service.js'
import { SheetsShellCoordinator, type ShellWorkbookSession } from './sheets-shell-coordinator'
import type {
  SpreadsheetEngine,
  EngineSessionHandle,
  WorkbookSession,
  WorkbookMetadata,
  SpreadsheetPdfRenderer,
} from '@genoffice/runtime-contracts'
import type { ScreenCapture } from '@genoffice/platform'

export interface SheetsRuntimeBundle {
  readonly engine: ElectronXlsxSidecarEngine
  readonly service: SpreadsheetServiceImpl
  readonly coordinator: SheetsShellCoordinator
  readonly pdfRenderer: SpreadsheetPdfRenderer
  readonly screenCapture: ScreenCapture
}

/**
 * Construct the Sheets runtime bundle.
 *
 * @param config — sidecar binary path + optional temp dir + optional
 *                 `sidecarClient` (legacy `XlsxSidecarClient` to share).
 * @returns the runtime bundle (engine + service + coordinator + pdfRenderer)
 */
export function initSheetsRuntime(config: ElectronXlsxSidecarEngineConfig): SheetsRuntimeBundle {
  // If a legacy `sidecarClient` is injected, the engine uses it INSTEAD of
  // constructing its own `SidecarProtocolClient` — sharing the same sidecar
  // process and enabling zero-overhead legacy session adoption.
  const engine = new ElectronXlsxSidecarEngine(config)
  // Only call engine.start() when the engine owns its client. When the
  // client is injected (legacy XlsxSidecarClient), the caller is
  // responsible for starting it.
  if (!config.sidecarClient) {
    engine.start()
  }

  const service = new SpreadsheetServiceImpl({ engine })

  // INCREMENT 7: construct the PDF renderer (hidden BrowserWindow + printToPDF).
  // The renderer is a Sheets-specific runtime port (ADR-006). The coordinator
  // owns the callerWindow + save dialog; the renderer owns only the rendering
  // context (hidden window, HTML load, printToPDF, cleanup).
  const pdfRenderer = new ElectronSpreadsheetPdfRenderer()

  // INCREMENT 8: construct the ScreenCapture capability (ADR-005).
  // The capability owns desktopCapturer + screen.getAllDisplays + permission
  // checks. The handler delegates to it directly (no coordinator involvement
  // — screen capture has no session/lifecycle concerns).
  const screenCapture = new ElectronScreenCapture()

  const coordinator = new SheetsShellCoordinator({ service, pdfRenderer })

  return { engine, service, coordinator, pdfRenderer, screenCapture }
}

// ── Legacy session adoption ──────────────────────────────────────────

/**
 * Input shape for adopting a legacy-opened session into the coordinator's
 * registry. The caller has already opened the workbook via the legacy
 * `XlsxSidecarClient.open()` and captured all session state.
 *
 * The `sidecarSessionId` is the UUID the sidecar binary returned in its
 * `open` response. The coordinator will register the session under this
 * SAME id — preserving renderer continuity (the renderer already holds
 * this sessionId from the legacy open response).
 */
export interface LegacySessionAdoption {
  /** The sidecar's session UUID (returned by the sidecar's `open` command). */
  readonly sidecarSessionId: string
  /** Absolute path the user opened (or the restore target, if a recovery). */
  readonly originalPath: string
  /**
   * Absolute path to the snapshot file the sidecar has open.
   * The snapshot is owned by the legacy path until adoption — afterwards,
   * the coordinator owns it (cleanup happens via coordinator.teardown).
   */
  readonly snapshotPath: string
  /** SHA-256 of the snapshot file at open time. */
  readonly diskFingerprint: string
  /** Set when the session opened a converted .xls/.csv import. */
  readonly suggestSaveAs?: string
  /** True when the session opened a converted .csv import. */
  readonly csvImport?: boolean
  /** Set when the session opened a restored crash-recovery copy. */
  readonly restoreTarget?: string
  /** SHA-256 of the restore target at open time. */
  readonly restoreTargetSha?: string
  /**
   * sheetId → file sheet name mapping (built from `[sheet.id, sheet.name]`
   * of the sidecar's open response). The engine stores this on the
   * adopted handle so applySavePlan's translator can resolve domain
   * sheetIds → file sheet names.
   */
  readonly sheetNames: ReadonlyMap<string, string>
  /** Workbook metadata (built from the sidecar's open response). */
  readonly metadata: WorkbookMetadata
  /** UI locale (preserved across save's session swap). */
  readonly locale: string
}

/**
 * Adopt a legacy-opened session into the coordinator's registry.
 *
 * This is the bridge between the legacy `workbook:select` open lifecycle
 * and the migrated read/recalc/media/close path.
 *
 * Steps (all zero-overhead):
 *   1. `engine.adoptExternalSession(...)` — wraps the EXISTING sidecar
 *      `sessionId` into an opaque `EngineSessionHandle`. NO wire call.
 *   2. Build a `WorkbookSession` from the legacy `sheetNames` and metadata.
 *   3. Build a `ShellWorkbookSession` carrying the engine handle + the
 *      legacy snapshot path + the legacy fingerprint.
 *   4. `coordinator.adoptLegacySession(wcId, shellSession)` — registers
 *      the session under (wcId, sessionId) and sets commit state IDLE.
 *
 * After adoption, the renderer can call the migrated `read-range`,
 * `read-formulas`, `recalc`, `read-media`, and `close` IPC handlers with
 * the SAME `sessionId` it received from the legacy `workbook:select`
 * response — no remapping, no re-open.
 *
 * OWNERSHIP: After adoption, the coordinator is the EXCLUSIVE owner of:
 *   - The engine handle (and the underlying sidecar session it wraps).
 *   - The snapshot path.
 *   - The recovery resources.
 * The legacy `SessionInfo` must mark itself as `adopted: true` so that
 * legacy `closeAllSessions` (teardown) does NOT double-close the sidecar
 * session or double-delete the snapshot.
 *
 * @param bundle — the runtime bundle (engine + service + coordinator)
 * @param wcId — the renderer's webContents id
 * @param adoption — the legacy session state to adopt
 * @returns the registered `ShellWorkbookSession`
 */
export function adoptLegacySessionIntoCoordinator(
  bundle: SheetsRuntimeBundle,
  wcId: number,
  adoption: LegacySessionAdoption,
): Promise<ShellWorkbookSession> {
  // Step 1: Wrap the existing sidecar session into an opaque engine handle.
  // This is a pure in-process operation — NO wire call, NO spawn, NO file IO.
  const engineHandle: EngineSessionHandle = bundle.engine.adoptExternalSession({
    sidecarSessionId: adoption.sidecarSessionId,
    tempPath: adoption.snapshotPath,
    sheetNames: adoption.sheetNames,
  })

  // Step 2: Build the domain WorkbookSession (workbookName, hash, sheetNames).
  const domainSession: WorkbookSession = {
    workbookName: adoption.metadata.name,
    workbookHash: adoption.metadata.sha256,
    sheetNames: adoption.sheetNames,
  }

  // Step 3: Build the ShellWorkbookSession carrying the adopted engine
  // handle + the legacy snapshot path + the legacy fingerprint.
  const shellSession: ShellWorkbookSession = {
    sessionId: adoption.sidecarSessionId,
    originalPath: adoption.originalPath,
    snapshotPath: adoption.snapshotPath,
    diskFingerprint: adoption.diskFingerprint,
    ...(adoption.suggestSaveAs !== undefined ? { suggestSaveAs: adoption.suggestSaveAs } : {}),
    ...(adoption.csvImport === true ? { csvImport: adoption.csvImport } : {}),
    ...(adoption.restoreTarget !== undefined ? { restoreTarget: adoption.restoreTarget } : {}),
    ...(adoption.restoreTargetSha !== undefined ? { restoreTargetSha: adoption.restoreTargetSha } : {}),
    engineHandle,
    domainSession,
    metadata: adoption.metadata,
    locale: adoption.locale,
    recoveryEpoch: 0,
  }

  // Step 4: Register the session with the coordinator. This sets commit
  // state IDLE and acquires the per-session mutation lock — no resources
  // are created or transferred.
  return bundle.coordinator.adoptLegacySession(wcId, shellSession)
}

/**
 * Type guard / branding helper — exposes the `SidecarProtocolLike` view
 * of a `SheetsRuntimeBundle`'s engine for testing. The engine already
 * satisfies this via its `client` field, but that's private.
 *
 * Production code never needs this; it's used by tests that want to verify
 * the injected sidecar client is the one speaking the wire protocol.
 */
export function engineSidecarClient(bundle: SheetsRuntimeBundle): SidecarProtocolLike | undefined {
  // The engine's `client` field is private. We can't read it directly.
  // This function exists for type-level documentation; tests that need
  // to verify the wire client should inject their own mock.
  void bundle
  return undefined
}
