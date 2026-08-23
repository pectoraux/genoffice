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
 *     ↓ owns session lifecycle (sole owner — Phase 2 Increment 16)
 *   Migrated IPC handlers (thin adapter)
 *
 * The coordinator receives ONLY SpreadsheetService — it does NOT import
 * or depend on XlsxSidecarClient, xlsx-gateway, child_process, or any
 * engine-specific code.
 *
 * SIDECAR PROCESS SHARING:
 *   The engine accepts an injectable `sidecarClient` (a `SidecarProtocolLike`)
 *   so a single Rust sidecar process serves both the engine's wire commands
 *   and (during the legacy cutover window) any remaining legacy callers.
 *   The engine does NOT own the injected client's lifecycle: stop() is a
 *   no-op for the client (the caller owns starting/stopping it).
 */

import {
  ElectronXlsxSidecarEngine,
  ElectronSpreadsheetPdfRenderer,
  ElectronScreenCapture,
  type ElectronXlsxSidecarEngineConfig,
} from '@genoffice/platform-electron'
import { SpreadsheetServiceImpl } from '@genoffice/services-sheets/src/spreadsheet-service.js'
import { SheetsShellCoordinator } from './sheets-shell-coordinator'
import type {
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
 * Coordinator-level configuration that the runtime construction cannot
 * derive from the engine config alone. Currently carries only the
 * `onWorkbookRenamed` callback (used to update the legacy
 * `SessionInfo.path` mirror after a successful auto-rename).
 *
 * Kept separate from `ElectronXlsxSidecarEngineConfig` so the engine
 * config stays focused on the sidecar lifecycle.
 */
export interface SheetsCoordinatorConfig {
  /**
   * Localized recovery-dialog text provider. See
   * `SheetsShellCoordinatorDeps.recoveryDialogText`.
   */
  readonly recoveryDialogText?: () => {
    readonly restoreButton: string
    readonly discardButton: string
    readonly title: string
    readonly body: string
  }
  /**
   * Invoked after a SUCCESSFUL open. See
   * `SheetsShellCoordinatorDeps.onWorkbookOpened`.
   */
  readonly onWorkbookOpened?: (wcId: number, openedPath: string) => void
  /**
   * Returns the shell-queued workbook path (if any). See
   * `SheetsShellCoordinatorDeps.consumeQueuedWorkbookPath`.
   */
  readonly consumeQueuedWorkbookPath?: () => string | undefined
}

/**
 * Construct the Sheets runtime bundle.
 *
 * @param config — sidecar binary path + optional temp dir + optional
 *                 `sidecarClient` (a `SidecarProtocolLike` to share the
 *                 sidecar process).
 * @param coordinatorConfig — coordinator-level configuration (legacy mirror
 *                             update callback, recovery dialog text, opened
 *                             hook, queued-path consumer).
 * @returns the runtime bundle (engine + service + coordinator + pdfRenderer)
 */
export function initSheetsRuntime(
  config: ElectronXlsxSidecarEngineConfig,
  coordinatorConfig: SheetsCoordinatorConfig = {},
): SheetsRuntimeBundle {
  // If a `sidecarClient` is injected, the engine uses it INSTEAD of
  // constructing its own `SidecarProtocolClient` — sharing the same sidecar
  // process. The engine does NOT own the injected client's lifecycle.
  //
  // During the Phase 2 Increment 16 cutover window: the legacy
  // `XlsxSidecarClient` (constructed in `createSheetsWindow` /
  // `createSheetsView`) is injected here so the engine speaks to the SAME
  // sidecar process. This preserves the "exactly ONE sidecar process"
  // invariant verified by the CDP smoke test. Once all legacy handlers are
  // removed (Phase E), the engine will own its own client.
  const engine = new ElectronXlsxSidecarEngine(config)
  // Only call engine.start() when the engine owns its client. When the
  // client is injected, the caller is responsible for starting it.
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

  // Phase 2 Increment 17: the coordinator is the SOLE owner of workbook
  // sessions. The shell plumbs three coordinator-level callbacks:
  //   - recoveryDialogText: localized recovery prompt text (the coordinator
  //     is language-agnostic).
  //   - onWorkbookOpened: fire `workbookOpenedHook` (tab tracking) after a
  //     successful open.
  //   - consumeQueuedWorkbookPath: return the shell-queued path (set by
  //     setForcedWorkbookPath / the dev capture server). The coordinator
  //     consumes it WITHOUT showing the dialog — matching legacy behavior.
  //
  // The manual rename path (`sheetsFileRenamed`) now calls
  // `coordinator.renameWorkbookFromShell(wcId, oldPath, newPath)` directly
  // — NO legacy SessionInfo mirror, NO onWorkbookRenamed callback.
  const coordinator = new SheetsShellCoordinator({
    service,
    pdfRenderer,
    ...(coordinatorConfig.recoveryDialogText !== undefined
      ? { recoveryDialogText: coordinatorConfig.recoveryDialogText }
      : {}),
    ...(coordinatorConfig.onWorkbookOpened !== undefined
      ? { onWorkbookOpened: coordinatorConfig.onWorkbookOpened }
      : {}),
    ...(coordinatorConfig.consumeQueuedWorkbookPath !== undefined
      ? { consumeQueuedWorkbookPath: coordinatorConfig.consumeQueuedWorkbookPath }
      : {}),
  })

  return { engine, service, coordinator, pdfRenderer, screenCapture }
}

