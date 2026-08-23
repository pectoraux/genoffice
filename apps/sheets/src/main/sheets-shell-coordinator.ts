/**
 * SheetsShellCoordinator — owns the per-renderer workbook session lifecycle.
 *
 * RESOURCE OWNERSHIP (Increment 4D/4E/4F/4G):
 *   OwnedResources is created BEFORE the first resource. The operation owns
 *   every resource from creation until transfer() or release().
 *   Conversion temp dirs are owned: cleaned up eagerly after snapshot creation,
 *   but if cleanup fails, ownership is RETAINED and release() retries.
 *
 * SAVE COMMIT PROTOCOL (Increment 4F/4G):
 *   Phase A — Prepare: temp target + snapshot + open + validate.
 *   Phase B — Commit: transition to COMMITTING → write marker → rename →
 *             install replacement → clear marker → transition to IDLE.
 *             If rename fails, save fails explicitly — NO copyFile fallback.
 *   Phase C — Cleanup: close old handle, remove old snapshot (best-effort).
 *
 * SESSION COMMIT LIFECYCLE (Increment 4G):
 *   Each session has a commit state: IDLE | COMMITTING | TEARING_DOWN | CLOSED.
 *   - Teardown checks: if session is COMMITTING, teardown waits for the
 *     commit to complete (it cannot preempt an in-progress commit).
 *   - Save checks: if session is TEARING_DOWN, save aborts before commit.
 *   This gives deterministic semantics: teardown cannot corrupt a committed save.
 *
 * CRASH RECONCILIATION (Increment 4G):
 *   Commit markers are stored in a DETERMINISTIC userData directory:
 *     userData/sheets-save-commits/<sessionId>.json
 *   Both save and reconcileSaveCommit() use the SAME location.
 *   Markers are validated at read time (not unchecked casts).
 */

import { randomUUID } from 'node:crypto'
import { mkdir, rm, readFile, writeFile, rename, readdir } from 'node:fs/promises'
import { existsSync, statSync, unlinkSync, renameSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { join, dirname, basename } from 'node:path'
import { app, BrowserWindow, dialog, type WebContents, type IpcMainInvokeEvent } from 'electron'

import type {
  SpreadsheetService, WorkbookSession, WorkbookOpenResult, EngineSessionHandle,
  ExternalChangeStatus, SaveRequest, SaveResult, EngineRangeResult, EngineFormulaCellsResult,
  EngineRecalcEdit, EngineRecalcRead, EngineRecalcResult, EngineMediaResult,
  SpreadsheetPdfRenderer, SpreadsheetPdfOptions, WorkbookPivotDefinition,
} from '@genoffice/runtime-contracts'
import { EngineError, InvalidInputError, InvalidSessionError } from '@genoffice/runtime-contracts'

// ── ShellWorkbookSession ──

export interface ShellWorkbookSession {
  readonly sessionId: string
  readonly originalPath: string
  readonly snapshotPath: string
  readonly diskFingerprint: string
  readonly suggestSaveAs?: string | undefined
  readonly csvImport?: boolean | undefined
  readonly restoreTarget?: string | undefined
  readonly restoreTargetSha?: string | undefined
  readonly engineHandle: EngineSessionHandle
  readonly domainSession: WorkbookSession
  readonly metadata: import('@genoffice/runtime-contracts').WorkbookMetadata
  readonly locale: string
  readonly recoveryEpoch: number
}

export interface SheetsShellCoordinatorDeps {
  readonly service: SpreadsheetService
  readonly onCommitGate?: (sessionId: string) => Promise<void>
  readonly onMarkerWritten?: (markerPath: string, sessionId: string) => Promise<void>
  readonly pdfRenderer?: SpreadsheetPdfRenderer
  /**
   * Optional callback invoked after a SUCCESSFUL auto-rename to update the
   * legacy `SessionInfo.path` mirror in the shell (`sheetsTabs`).
   *
   * The coordinator is the AUTHORITATIVE owner of the session path
   * (`ShellWorkbookSession.originalPath`); the legacy `SessionInfo` is a
   * NON-OWNING compatibility reference kept only so legacy consumers
   * (e.g. `resolveSheetsSessionPath` used by `project:rebindChat`) see
   * the new path. Without this callback the legacy mirror would go stale
   * after an auto-rename.
   *
   * The callback is invoked AFTER the coordinator has:
   *   - updated `ShellWorkbookSession.originalPath`
   *   - removed the old path from `untitledPaths`
   *   - pushed the `workbook:renamed` event to the initiating renderer
   * The callback MUST NOT re-push the event (the coordinator already did).
   *
   * IMPLEMENTATION (Increment 15A):
   *   The shell plumbs `updateLegacySessionPath(wcId, oldPath, newPath)`
   *   (extracted from `sheetsFileRenamed`) as this callback. That helper
   *   updates the legacy `sheetsTabs.sessions[].path` mirror ONLY — it
   *   does NOT push any IPC event, avoiding the duplicate-push hazard.
   */
  readonly onWorkbookRenamed?: (wcId: number, oldPath: string, newPath: string) => void
}

// ── Session commit lifecycle ──

const enum SessionCommitState {
  IDLE = 0,
  COMMITTING = 1,
  TEARING_DOWN = 2,
  CLOSED = 3,
}

interface RendererState {
  readonly sessions: Map<string, ShellWorkbookSession>
  epoch: number
  readonly locks: Map<string, Promise<unknown>>
  /** Per-session commit state — controls teardown/commit race semantics. */
  readonly commitStates: Map<string, SessionCommitState>
}

// ── OwnedResources ──

class OwnedResources {
  private snapshotPath: string | undefined
  private engineHandle: EngineSessionHandle | undefined
  private tempTargetPath: string | undefined
  private conversionDir: string | undefined
  private _transferred = false

  setSnapshot(path: string): void { this.snapshotPath = path }
  setEngineHandle(handle: EngineSessionHandle): void { this.engineHandle = handle }
  setTempTarget(path: string): void { this.tempTargetPath = path }
  setConversionDir(dir: string): void { this.conversionDir = dir }
  clearConversionDir(): void { this.conversionDir = undefined }
  clearTempTarget(): void { this.tempTargetPath = undefined }
  get tempTarget(): string | undefined { return this.tempTargetPath }
  get transferred(): boolean { return this._transferred }

  async release(service: SpreadsheetService): Promise<void> {
    if (this._transferred) return
    if (this.engineHandle) { try { await service.close(this.engineHandle) } catch {} this.engineHandle = undefined }
    if (this.snapshotPath) { try { await rm(this.snapshotPath, { force: true }) } catch {} this.snapshotPath = undefined }
    if (this.tempTargetPath) { try { await rm(this.tempTargetPath, { force: true }) } catch {} this.tempTargetPath = undefined }
    if (this.conversionDir) { try { await rm(this.conversionDir, { recursive: true, force: true }) } catch {} this.conversionDir = undefined }
  }

  transfer(): void {
    this.snapshotPath = undefined; this.engineHandle = undefined
    this.tempTargetPath = undefined; this.conversionDir = undefined
    this._transferred = true
  }
}

// ── Save commit marker ──

interface SaveCommitMarker {
  readonly version: 1
  readonly finalTarget: string
  readonly tempTarget: string
  readonly sessionId: string
}

/** Type guard: is the value a non-null, non-array object (a record)? */
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * Validate a parsed JSON value as a SaveCommitMarker.
 * Uses only runtime type guards — no `as` type assertions.
 * Returns null if the value is not a valid marker.
 */
function validateMarker(raw: unknown): SaveCommitMarker | null {
  if (!isRecord(raw)) return null
  if (raw.version !== 1) return null
  if (typeof raw.finalTarget !== 'string' || raw.finalTarget.length === 0) return null
  if (typeof raw.tempTarget !== 'string' || raw.tempTarget.length === 0) return null
  if (typeof raw.sessionId !== 'string' || raw.sessionId.length === 0) return null
  return { version: 1, finalTarget: raw.finalTarget, tempTarget: raw.tempTarget, sessionId: raw.sessionId }
}

// ── Coordinator ──

export class SheetsShellCoordinator {
  private readonly tabs = new Map<number, RendererState>()

  constructor(private readonly deps: SheetsShellCoordinatorDeps) {}

  registerRenderer(wcId: number, webContents: WebContents): void {
    if (!this.tabs.has(wcId)) {
      this.tabs.set(wcId, {
        sessions: new Map(), epoch: 0, locks: new Map(), commitStates: new Map(),
      })
    }
    webContents.once('destroyed', () => { void this.teardown(wcId) })
  }

  getSession(wcId: number, sessionId: string): ShellWorkbookSession {
    const state = this.tabs.get(wcId)
    if (!state) throw new InvalidSessionError(`Unknown renderer: ${wcId}`)
    const session = state.sessions.get(sessionId)
    if (!session) throw new InvalidSessionError(`Unknown workbook session: ${sessionId}`)
    return session
  }

  wcIdFromEvent(event: IpcMainInvokeEvent): number { return event.sender.id }

  private isAlive(wcId: number, startEpoch: number): boolean {
    const state = this.tabs.get(wcId)
    return !!state && state.epoch === startEpoch
  }

  private checkEpoch(wcId: number, startEpoch: number): void {
    if (!this.isAlive(wcId, startEpoch)) throw new InvalidSessionError(`Renderer ${wcId} was torn down during operation`)
  }

  private withSessionLock<T>(wcId: number, sessionId: string, fn: () => Promise<T>): Promise<T> {
    const state = this.tabs.get(wcId)
    if (!state) throw new InvalidSessionError(`Unknown renderer: ${wcId}`)
    const prev = state.locks.get(sessionId) ?? Promise.resolve()
    const next = prev.then(() => fn(), () => fn())
    state.locks.set(sessionId, next.catch(() => {}))
    return next
  }

  private getCommitState(wcId: number, sessionId: string): SessionCommitState {
    return this.tabs.get(wcId)?.commitStates.get(sessionId) ?? SessionCommitState.IDLE
  }

  private setCommitState(wcId: number, sessionId: string, state: SessionCommitState): void {
    this.tabs.get(wcId)?.commitStates.set(sessionId, state)
  }

  // ── Open ──

  async openWorkbook(
    wcId: number, callerWindow: BrowserWindow | undefined,
    options: { queuedPath?: string | undefined; locale: string },
  ): Promise<{ sessionId: string; session: ShellWorkbookSession } | null> {
    let state = this.tabs.get(wcId)
    if (!state) {
      state = { sessions: new Map(), epoch: 0, locks: new Map(), commitStates: new Map() }
      this.tabs.set(wcId, state)
    }
    const startEpoch = state.epoch

    let path = options.queuedPath
    if (!path) {
      const selection = callerWindow
        ? await dialog.showOpenDialog(callerWindow, { properties: ['openFile'], filters: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'csv'] }] })
        : await dialog.showOpenDialog({ properties: ['openFile'], filters: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'csv'] }] })
      this.checkEpoch(wcId, startEpoch)
      if (selection.canceled || !selection.filePaths[0]) return null
      path = selection.filePaths[0]
    }

    const prepared = await this.prepareWorkbookForOpen(path, callerWindow)
    this.checkEpoch(wcId, startEpoch)

    const owned = new OwnedResources()
    if (prepared.conversionDir) owned.setConversionDir(prepared.conversionDir)

    try {
      const snapshotPath = await this.snapshotWorkbook(prepared.openPath)
      owned.setSnapshot(snapshotPath)
      this.checkEpoch(wcId, startEpoch)

      if (prepared.conversionDir) {
        try {
          await rm(prepared.conversionDir, { recursive: true, force: true })
          owned.clearConversionDir()
        } catch { /* retained for release() retry */ }
      }

      const bytes = await readFile(snapshotPath)
      const fileName = prepared.openPath.split(/[\\/]/).pop() ?? 'workbook.xlsx'
      const openResult = await this.deps.service.open(new Uint8Array(bytes), options.locale, fileName)
      owned.setEngineHandle(openResult.engineHandle)
      this.checkEpoch(wcId, startEpoch)

      const diskFingerprint = await this.sha256File(snapshotPath)
      const restoreTargetSha = prepared.restoreTarget
        ? await this.sha256File(prepared.restoreTarget).catch(() => undefined) : undefined

      this.checkEpoch(wcId, startEpoch)
      const sessionId = randomUUID()
      const shellSession: ShellWorkbookSession = {
        sessionId, originalPath: path, snapshotPath, diskFingerprint,
        suggestSaveAs: prepared.suggestSaveAs, csvImport: prepared.csvImport,
        restoreTarget: prepared.restoreTarget, restoreTargetSha,
        engineHandle: openResult.engineHandle, domainSession: openResult.session,
        metadata: openResult.metadata, locale: options.locale, recoveryEpoch: 0,
      }
      state = this.tabs.get(wcId)!
      state.sessions.set(sessionId, shellSession)
      state.commitStates.set(sessionId, SessionCommitState.IDLE)
      owned.transfer()

      return { sessionId, session: shellSession }
    } catch (error) {
      await owned.release(this.deps.service)
      throw error
    }
  }

  // ── Legacy session adoption (Increment 5A) ──

  /**
   * Adopt a legacy-opened session into the coordinator's registry.
   *
   * This is the compatibility handoff between the legacy `workbook:select`
   * open lifecycle (which uses `XlsxSidecarClient.open()` directly) and the
   * migrated read/recalc/media/close path (which uses the coordinator-backed
   * `SpreadsheetService`).
   *
   * The caller (sheets-runtime.ts) has ALREADY:
   *   - Opened the workbook via the legacy `XlsxSidecarClient.open()`.
   *   - Built a snapshot, computed its sha256, captured sheet names.
   *   - Constructed a `ShellWorkbookSession` whose `engineHandle` was produced
   *     by `ElectronXlsxSidecarEngine.adoptExternalSession()` — a pure
   *     in-process handle wrap with NO wire call, NO file IO, NO spawn.
   *
   * The coordinator just registers the session under (wcId, sessionId) and
   * initializes its commit state to IDLE. There is NO re-open, NO re-spawn,
   * NO duplicate snapshot, NO recomputed fingerprint.
   *
   * OWNERSHIP: After adoption, the coordinator is the EXCLUSIVE owner of:
   *   - The engine handle (and the underlying sidecar session it wraps).
   *   - The snapshot path.
   *   - The recovery resources.
   * The legacy `SessionInfo` keeps a NON-OWNING reference (marked
   * `adopted: true`) for the legacy `save`/`write-recovery` paths to keep
   * working — but it MUST NOT independently `client.close(sessionId)` or
   * `rm(snapshotPath)` once adopted.
   *
   * SESSION IDENTITY: The coordinator registers the session under the SAME
   * `sessionId` the legacy open returned — preserving renderer continuity.
   * The renderer's existing `sessionId` is valid for both legacy
   * `save`/`write-recovery` AND migrated reads/closes without remapping.
   *
   * THREAD SAFETY: Acquires the per-session mutation lock. If a concurrent
   * teardown already holds the session's lock (e.g., the renderer is being
   * destroyed), adoption still completes — the session is registered but
   * will be cleaned up by the teardown that owns the lock.
   *
   * @returns the registered session (same as the input).
   */
  async adoptLegacySession(
    wcId: number,
    session: ShellWorkbookSession,
  ): Promise<ShellWorkbookSession> {
    // Lazily register the renderer if this is the first call for this wcId.
    // The legacy `workbook:select` path does not call registerRenderer()
    // explicitly — adoption is the first coordinator contact for a tab.
    let state = this.tabs.get(wcId)
    if (!state) {
      state = {
        sessions: new Map(),
        epoch: 0,
        locks: new Map(),
        commitStates: new Map(),
      }
      this.tabs.set(wcId, state)
    }
    return this.withSessionLock(wcId, session.sessionId, async () => {
      // Re-fetch inside the lock — the renderer may have been torn down
      // between the lazy register above and the lock acquisition.
      const cur = this.tabs.get(wcId)
      if (!cur) {
        throw new InvalidSessionError(`Renderer ${wcId} torn down during adoption`)
      }
      // If a session already exists under this sessionId, it means adoption
      // was called twice (a bug, or a save-reopen path that hasn't been
      // migrated yet). The previous session's resources must be cleaned up
      // by the caller BEFORE re-adoption — we throw to surface the bug.
      if (cur.sessions.has(session.sessionId)) {
        throw new InvalidSessionError(
          `Session ${session.sessionId} already adopted — caller must close before re-adopting`,
        )
      }
      cur.sessions.set(session.sessionId, session)
      cur.commitStates.set(session.sessionId, SessionCommitState.IDLE)
      return session
    })
  }

  // ── Read operations ──

  async readRange(wcId: number, sessionId: string, sheetId: string, range: string): Promise<EngineRangeResult> {
    const s = this.getSession(wcId, sessionId)
    return this.deps.service.readRange(s.domainSession, s.engineHandle, sheetId, range)
  }
  async readFormulaCells(wcId: number, sessionId: string, sheetId: string): Promise<EngineFormulaCellsResult> {
    const s = this.getSession(wcId, sessionId)
    return this.deps.service.readFormulaCells(s.domainSession, s.engineHandle, sheetId)
  }
  async recalculate(wcId: number, sessionId: string, edits: EngineRecalcEdit[], reads: EngineRecalcRead[]): Promise<EngineRecalcResult> {
    const s = this.getSession(wcId, sessionId)
    return this.deps.service.recalculate(s.domainSession, s.engineHandle, edits, reads)
  }
  async readMedia(wcId: number, sessionId: string, visualId: string): Promise<EngineMediaResult> {
    const s = this.getSession(wcId, sessionId)
    return this.deps.service.readMedia(s.domainSession, s.engineHandle, visualId)
  }

  // ── Save (commit protocol with lifecycle state) ──

  async saveWorkbook(
    wcId: number, sessionId: string, request: SaveRequest, mode: 'save' | 'save-as',
    callerWindow: BrowserWindow | undefined,
  ): Promise<SaveResult & { canceled?: boolean }> {
    return this.withSessionLock(wcId, sessionId, async () => {
      const state = this.tabs.get(wcId)
      if (!state) throw new InvalidSessionError(`Unknown renderer: ${wcId}`)
      const startEpoch = state.epoch
      const session = this.getSession(wcId, sessionId)

      // Check commit state — if teardown is in progress, abort
      if (this.getCommitState(wcId, sessionId) === SessionCommitState.TEARING_DOWN) {
        throw new InvalidSessionError(`Session ${sessionId} is being torn down`)
      }

      // 1. Resolve target + ExternalChangeStatus
      let targetPath: string
      let externalChange: ExternalChangeStatus
      if (mode === 'save-as' || session.suggestSaveAs !== undefined) {
        const opts = { defaultPath: session.suggestSaveAs ?? session.restoreTarget ?? session.originalPath, filters: [{ name: 'Excel', extensions: ['xlsx'] }] }
        const sel = callerWindow ? await dialog.showSaveDialog(callerWindow, opts) : await dialog.showSaveDialog(opts)
        this.checkEpoch(wcId, startEpoch)
        if (sel.canceled || !sel.filePath) return { ok: false, canceled: true }
        targetPath = sel.filePath.endsWith('.xlsx') ? sel.filePath : `${sel.filePath}.xlsx`
        externalChange = 'unchanged'
      } else if (session.restoreTarget !== undefined) {
        externalChange = await this.computeExternalChangeStatus(session.restoreTarget, session.restoreTargetSha ?? '')
        this.checkEpoch(wcId, startEpoch)
        targetPath = session.restoreTarget
      } else {
        externalChange = await this.computeExternalChangeStatus(session.originalPath, session.diskFingerprint)
        this.checkEpoch(wcId, startEpoch)
        targetPath = session.originalPath
      }

      // 2. service.save()
      const result = await this.deps.service.save(session.domainSession, session.engineHandle, request, externalChange)
      this.checkEpoch(wcId, startEpoch)
      if (!result.ok || !result.data) return result

      // ═══ Phase A: Prepare ═══
      const owned = new OwnedResources()
      let replacementSession: ShellWorkbookSession
      let tempTargetPath: string
      try {
        tempTargetPath = join(dirname(targetPath), `.genoffice-save-${randomUUID()}.xlsx`)
        await writeFile(tempTargetPath, result.data)
        owned.setTempTarget(tempTargetPath)
        this.checkEpoch(wcId, startEpoch)

        const newSnapshotPath = await this.snapshotWorkbook(tempTargetPath)
        owned.setSnapshot(newSnapshotPath)
        this.checkEpoch(wcId, startEpoch)

        const newBytes = await readFile(newSnapshotPath)
        const fileName = targetPath.split(/[\\/]/).pop() ?? 'workbook.xlsx'
        const newOpenResult = await this.deps.service.open(new Uint8Array(newBytes), session.locale, fileName)
        owned.setEngineHandle(newOpenResult.engineHandle)
        this.checkEpoch(wcId, startEpoch)

        const newDiskFingerprint = await this.sha256File(newSnapshotPath)
        this.checkEpoch(wcId, startEpoch)

        replacementSession = {
          sessionId, originalPath: targetPath, snapshotPath: newSnapshotPath,
          diskFingerprint: newDiskFingerprint, engineHandle: newOpenResult.engineHandle,
          domainSession: newOpenResult.session, metadata: newOpenResult.metadata,
          locale: session.locale, recoveryEpoch: session.recoveryEpoch + 1,
        }
      } catch (error) {
        await owned.release(this.deps.service)
        throw error
      }

      // ═══ Phase B: Commit ═══
      // Check epoch AND commit state before entering COMMITTING.
      // If teardown has already started, abort — do NOT commit.
      this.checkEpoch(wcId, startEpoch)
      if (this.getCommitState(wcId, sessionId) === SessionCommitState.TEARING_DOWN) {
        await owned.release(this.deps.service)
        throw new InvalidSessionError(`Session ${sessionId} is being torn down`)
      }

      // Transition to COMMITTING — teardown that arrives after this
      // MUST wait for the commit to complete (it cannot preempt).
      this.setCommitState(wcId, sessionId, SessionCommitState.COMMITTING)

      // Commit gate — injectable barrier for deterministic testing.
      // In production, this is a no-op. In tests, it pauses the save
      // between COMMITTING and the irreversible rename, allowing the
      // test to verify the state and trigger teardown.
      if (this.deps.onCommitGate) {
        await this.deps.onCommitGate(sessionId)
      }

      try {
        // Write commit marker to the DETERMINISTIC userData directory
        // (same location reconcileSaveCommit scans)
        const commitDir = this.commitMarkerDir()
        await mkdir(commitDir, { recursive: true })
        const markerPath = join(commitDir, `${sessionId}.json`)
        const marker: SaveCommitMarker = { version: 1, finalTarget: targetPath, tempTarget: tempTargetPath, sessionId }
        await writeFile(markerPath, JSON.stringify(marker))

        // Marker-written hook — injectable barrier for deterministic
        // testing of the save→reconcile crash path. Called AFTER the
        // marker is on disk but BEFORE the rename. In production, undefined.
        if (this.deps.onMarkerWritten) {
          await this.deps.onMarkerWritten(markerPath, sessionId)
        }

        // Atomically promote temp → final via rename.
        // NO copyFile fallback. If rename fails, save fails.
        try {
          await rename(tempTargetPath, targetPath)
        } catch (renameError) {
          try { await rm(markerPath, { force: true }) } catch {}
          throw new EngineError(
            `Save commit failed: cannot atomically rename temp to final target — ${renameError}`,
            'INTERNAL_ERROR',
          )
        }
        owned.clearTempTarget()

        // Install replacement session
        const currentState = this.tabs.get(wcId)
        if (currentState) currentState.sessions.set(sessionId, replacementSession)
        owned.transfer()

        // Clear commit marker
        try { await rm(markerPath, { force: true }) } catch {}

        // Clear recovery copies
        this.clearWorkbookRecovery(targetPath)
        if (session.suggestSaveAs !== undefined) this.clearWorkbookRecovery(session.suggestSaveAs)
        if (session.restoreTarget !== undefined) this.clearWorkbookRecovery(session.restoreTarget)
      } catch (error) {
        // Commit failure — release owned resources
        await owned.release(this.deps.service)
        this.setCommitState(wcId, sessionId, SessionCommitState.IDLE)
        throw error
      }

      // ═══ Phase C: Old-resource cleanup (isolated, best-effort) ═══
      this.setCommitState(wcId, sessionId, SessionCommitState.IDLE)
      try { await this.deps.service.close(session.engineHandle) } catch {}
      try { await rm(session.snapshotPath, { force: true }) } catch {}

      return result
    })
  }

  // ── Crash reconciliation ──

  /**
   * Reconcile leftover save-commit markers from a crash during save.
   *
   * Markers are stored in: userData/sheets-save-commits/<sessionId>.json
   *
   * For each marker:
   *   - If tempTarget exists: delete it (rename never happened or crashed after)
   *   - Delete the marker
   *
   * If final target already has new bytes (rename succeeded before crash),
   * the old bytes are already gone — we can't undo the rename. But the
   * temp file is cleaned up.
   */
  static async reconcileSaveCommit(userDataDir: string): Promise<void> {
    const commitDir = join(userDataDir, 'sheets-save-commits')
    let entries: string[]
    try { entries = await readdir(commitDir) } catch { return }

    for (const entry of entries) {
      if (!entry.endsWith('.json')) continue
      const markerPath = join(commitDir, entry)
      try {
        const markerText = await readFile(markerPath, 'utf8')
        const raw = JSON.parse(markerText)
        const marker = validateMarker(raw)
        if (!marker) {
          // Invalid marker — delete it
          try { await rm(markerPath, { force: true }) } catch {}
          continue
        }
        // Clean up temp target if it still exists
        if (existsSync(marker.tempTarget)) {
          try { await rm(marker.tempTarget, { force: true }) } catch {}
        }
        // Always clean up the marker
        try { await rm(markerPath, { force: true }) } catch {}
      } catch {
        // Read/parse failed — delete the marker
        try { await rm(markerPath, { force: true }) } catch {}
      }
    }
  }

  // ── Recovery ──

  async writeRecovery(wcId: number, sessionId: string, request: SaveRequest): Promise<{ ok: boolean }> {
    let startRecoveryEpoch: number
    try { startRecoveryEpoch = this.getSession(wcId, sessionId).recoveryEpoch } catch { return { ok: false } }

    return this.withSessionLock(wcId, sessionId, async () => {
      let session: ShellWorkbookSession
      try { session = this.getSession(wcId, sessionId) } catch { return { ok: false } }
      if (session.suggestSaveAs !== undefined || session.restoreTarget !== undefined) return { ok: false }
      try {
        const result = await this.deps.service.writeRecovery(session.domainSession, session.engineHandle, request)
        const cur = this.tabs.get(wcId)?.sessions.get(sessionId)
        if (!cur || cur.recoveryEpoch !== startRecoveryEpoch) return { ok: false }
        const recoveryPath = this.recoveryPathFor(session.originalPath)
        await mkdir(join(recoveryPath, '..'), { recursive: true })
        await writeFile(recoveryPath, result)
        return { ok: true }
      } catch (error) { console.warn('[sheets] recovery copy failed:', error); return { ok: false } }
    })
  }

  // ── Close ──

  async closeWorkbook(wcId: number, sessionId: string): Promise<void> {
    await this.withSessionLock(wcId, sessionId, async () => { await this.closeSession(wcId, sessionId) })
  }

  // ── Pivot definition read (INCREMENT 12, corrected 15, hardened 15A) ──

  /**
   * Read a pivot table definition from the session's workbook.
   *
   * Delegates to `service.readPivotDefinition()`, which in turn delegates
   * to `engine.readPivotDefinition()` — the SINGLE translation point
   * between the OOXML wire format and the runtime-independent
   * `WorkbookPivotDefinition` contract. The engine reads both XML parts
   * from its on-disk temp file and parses them via the canonical
   * `@genoffice/xlsx-gateway` `parsePivotDefinition()` parser.
   *
   * The coordinator owns only session lookup (wcId + sessionId).
   * All archive I/O and parsing happen below the service boundary — the
   * coordinator passes the typed `WorkbookPivotDefinition` through
   * unchanged.
   *
   * @returns the parsed pivot definition (typed contract — NOT `unknown`)
   */
  async readPivotDefinition(
    wcId: number,
    sessionId: string,
    pivotTablePath: string,
    cacheDefinitionPath: string,
  ): Promise<WorkbookPivotDefinition> {
    const session = this.getSession(wcId, sessionId)
    return this.deps.service.readPivotDefinition(
      session.domainSession,
      session.engineHandle,
      pivotTablePath,
      cacheDefinitionPath,
    )
  }

  // ── Auto-rename (INCREMENT 12, hardened 15) ──

  /**
   * Rename the workbook file on disk.
   *
   * The coordinator owns:
   *   - session lookup (wcId + sessionId → ShellWorkbookSession)
   *   - rename validation (untitled path check, name sanitization, collision)
   *   - filesystem rename (renameSync)
   *   - session path update (ShellWorkbookSession.originalPath = newPath)
   *   - renderer notification (workbook:renamed push to event.sender only)
   *
   * SESSION CONTINUITY:
   *   - sessionId remains unchanged
   *   - engineHandle remains unchanged
   *   - snapshotPath remains unchanged (the snapshot is independent of the file path)
   *   - originalPath becomes the new path
   *   - diskFingerprint remains correct (the file content didn't change)
   *
   * PUSH EVENT:
   *   workbook:renamed is sent ONLY to the initiating renderer (event.sender).
   *   No broadcast. No getFocusedWindow.
   *
   * LEGACY COMPATIBILITY:
   *   The legacy SessionInfo.path is updated via the shell's compatibility mirror
   *   (sheetsTabs). The coordinator's ShellWorkbookSession.originalPath is the
   *   AUTHORITATIVE owner after migration.
   *
   * @returns { renamed: boolean; name?: string }
   */
  async renameWorkbook(
    wcId: number,
    webContents: WebContents,
    sessionId: string,
    baseName: string,
  ): Promise<{ renamed: boolean; name?: string }> {
    return this.withSessionLock(wcId, sessionId, async () => {
      // Re-fetch the session inside the lock — it may have been modified
      // by a concurrent save or close.
      let session: ShellWorkbookSession
      try {
        session = this.getSession(wcId, sessionId)
      } catch {
        return { renamed: false }
      }

      // Only rename untitled workbooks (matching legacy: untitledWorkbookPaths check)
      if (!this.isUntitledPath(session.originalPath)) {
        return { renamed: false }
      }

      // Sanitize the base name
      const base = sanitizeAutoRenameBase(baseName)
      if (!base) return { renamed: false }

      // Compute target path (collision avoidance: name-2, name-3, ...)
      const dir = dirname(session.originalPath)
      let target = join(dir, `${base}.xlsx`)
      for (let i = 2; existsSync(target) && i < 100; i++) {
        target = join(dir, `${base}-${i}.xlsx`)
      }
      if (existsSync(target) || target === session.originalPath) {
        return { renamed: false }
      }

      // Atomic rename (no copy fallback — matching legacy)
      const oldPath = session.originalPath
      try {
        renameSync(oldPath, target)
      } catch (err) {
        console.warn('[sheets] auto-rename failed:', err)
        return { renamed: false }
      }

      // Update the coordinator's session (authoritative owner)
      const state = this.tabs.get(wcId)
      if (state) {
        const updatedSession: ShellWorkbookSession = {
          ...session,
          originalPath: target,
        }
        state.sessions.set(sessionId, updatedSession)
      }

      // Remove from untitled set
      this.untitledPaths.delete(oldPath)

      // Push event to the initiating renderer only — exactly once.
      // The legacy mirror callback below MUST NOT also push.
      const name = basename(target)
      if (!webContents.isDestroyed()) {
        webContents.send('workbook:renamed', name)
      }

      // Update the legacy `SessionInfo.path` mirror (NON-authoritative —
      // the coordinator's `ShellWorkbookSession.originalPath` is the
      // source of truth). Legacy consumers (resolveSheetsSessionPath
      // used by project:rebindChat) read from sheetsTabs.sessions[].path
      // and would see a stale path without this callback. The callback
      // is invoked exactly once after a successful rename; it MUST NOT
      // re-push the workbook:renamed event (the coordinator already did).
      const onWorkbookRenamed = this.deps.onWorkbookRenamed
      if (onWorkbookRenamed) {
        try {
          onWorkbookRenamed(wcId, oldPath, target)
        } catch (err) {
          // Best-effort — the authoritative state is already updated.
          // A failure in the legacy mirror update MUST NOT undo the
          // rename or affect the return value.
          console.warn('[sheets] onWorkbookRenamed callback failed:', err)
        }
      }

      return { renamed: true, name }
    })
  }

  /** Check if a path is in the untitled workbook set (shell-owned state). */
  private isUntitledPath(path: string): boolean {
    return this.untitledPaths.has(path)
  }

  /** Untitled workbook paths — shell-owned state for auto-rename gating. */
  private readonly untitledPaths = new Set<string>()

  /** Mark a path as untitled (called from the shell when creating a new workbook). */
  markUntitledPath(path: string): void {
    this.untitledPaths.add(path)
  }

  // ── PDF export (INCREMENT 7 / ADR-006) ──

  /**
   * Export the workbook to PDF.
   *
   * Returns:
   *   - { canceled: true } — user canceled the save dialog
   *   - { canceled: false, path: string } — success, PDF written to path
   *
   * Errors (render failure, filesystem failure) are THROWN — matching the
   * legacy exportPdf behavior.
   *
   * The coordinator owns callerWindow + save dialog + output authorization +
   * writing the PDF bytes. The PDF renderer (SpreadsheetPdfRenderer) owns
   * only the rendering context (hidden BrowserWindow + printToPDF + cleanup).
   *
   * SECURITY ORDERING (output authorization):
   *   1. Show save dialog → user selects path (or cancels)
   *   2. If canceled → return { canceled: true }
   *   3. Render HTML → PDF bytes (via the renderer port)
   *   4. If render fails → throw Error (no file written)
   *   5. Write PDF bytes to the authorized path
   *
   * INCREMENT 7A: test-only output path override. When the environment
   * variable `GENOFFICE_PDF_TEST_OUTPATH` is set, the save dialog is
   * SKIPPED and the env var's value is used as the output path. This
   * enables deterministic real Electron E2E testing without native
   * dialog interaction. Production behavior is unchanged when the env
   * var is absent.
   */
  async exportPdf(
    wcId: number,
    callerWindow: BrowserWindow | undefined,
    request: {
      fileName: string
      html: string
      landscape: boolean
      pageSize: SpreadsheetPdfOptions['pageSize']
      margins: SpreadsheetPdfOptions['margins']
      scale: number
    },
  ): Promise<{ canceled: true } | { canceled: false; path: string }> {
    const pdfRenderer = this.deps.pdfRenderer
    if (!pdfRenderer) {
      throw new Error('PDF renderer not available')
    }

    // 1. Authorize output path (save dialog or test override)
    let outputPath: string
    const testOutPath = process.env['GENOFFICE_PDF_TEST_OUTPATH']
    if (testOutPath !== undefined && testOutPath.length > 0) {
      // Test-only: skip the dialog, use the env var path directly.
      // This is NOT a security bypass — the env var is set only in test
      // environments. Production never sets it.
      outputPath = testOutPath
    } else {
      const dialogOptions = {
        defaultPath: request.fileName,
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
      }
      const selection = callerWindow
        ? await dialog.showSaveDialog(callerWindow, dialogOptions)
        : await dialog.showSaveDialog(dialogOptions)
      if (selection.canceled || !selection.filePath) return { canceled: true }
      outputPath = selection.filePath
    }

    // 2. Render HTML → PDF bytes (renderer owns hidden BrowserWindow)
    const renderResult = await pdfRenderer.renderToPdf(request.html, {
      landscape: request.landscape,
      pageSize: request.pageSize,
      margins: request.margins,
      scale: request.scale,
    })

    // 3. If render failed → throw (NO file written)
    if (!renderResult.ok) {
      throw new Error(renderResult.message)
    }

    // 4. Write PDF bytes to the authorized path
    await writeFile(outputPath, renderResult.data)
    return { canceled: false, path: outputPath }
  }

  async teardown(wcId: number): Promise<void> {
    const state = this.tabs.get(wcId)
    if (!state) return
    state.epoch++

    const sessionIds = [...state.sessions.keys()]
    await Promise.all(sessionIds.map(async (sid) => {
      await this.withSessionLock(wcId, sid, async () => {
        // If a commit is in progress, wait for it to complete.
        // The mutation lock ensures we don't close the handle while
        // the commit is using it. The commit will check for TEARING_DOWN
        // before entering COMMITTING — but if it's already COMMITTING,
        // we must let it finish (it holds the lock).
        //
        // Since we're inside the lock, the commit has already completed
        // (either successfully or with failure). We can safely close.
        this.setCommitState(wcId, sid, SessionCommitState.TEARING_DOWN)
        await this.closeSession(wcId, sid)
        this.setCommitState(wcId, sid, SessionCommitState.CLOSED)
      })
    }))
    this.tabs.delete(wcId)
  }

  // ── Internal ──

  private async closeSession(wcId: number, sessionId: string): Promise<void> {
    const state = this.tabs.get(wcId)
    if (!state) return
    const session = state.sessions.get(sessionId)
    if (!session) return
    try { await this.deps.service.close(session.engineHandle) } catch {}
    try { await rm(session.snapshotPath, { force: true }) } catch {}
    state.sessions.delete(sessionId)
  }

  private async computeExternalChangeStatus(filePath: string, storedFingerprint: string): Promise<ExternalChangeStatus> {
    try {
      const currentSha = await this.sha256File(filePath)
      if (currentSha === storedFingerprint) return 'unchanged'
      return 'changed'
    } catch { return 'unknown' }
  }

  private commitMarkerDir(): string { return join(app.getPath('userData'), 'sheets-save-commits') }

  private async prepareWorkbookForOpen(
    path: string, parent: BrowserWindow | undefined,
  ): Promise<{ openPath: string; suggestSaveAs?: string; csvImport?: boolean; restoreTarget?: string; conversionDir?: string }> {
    const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
    if (extension !== 'csv' && extension !== 'xls') {
      const recovery = this.pendingRecoveryFor(path)
      if (recovery) {
        const opts = { type: 'question' as const, buttons: ['Restore', 'Discard'], defaultId: 0, cancelId: 1, message: 'Crash recovery copy found', detail: 'Unsaved work from a previous session was found. Restore it?' }
        const answer = parent ? await dialog.showMessageBox(parent, opts) : await dialog.showMessageBox(opts)
        if (answer.response === 0) return { openPath: recovery, restoreTarget: path }
        this.clearWorkbookRecovery(path)
      }
      return { openPath: path }
    }
    const stem = path.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '')
    const directory = join(app.getPath('temp'), 'genoffice-imports', randomUUID())
    await mkdir(directory, { recursive: true })
    const openPath = join(directory, `${stem}.xlsx`)
    if (extension === 'csv') {
      const { csvToXlsxBuffer, decodeCsvBuffer } = await import('../gateway/csv-import')
      const csvBytes = await readFile(path)
      await writeFile(openPath, await csvToXlsxBuffer(decodeCsvBuffer(csvBytes)))
      return { openPath, suggestSaveAs: path.replace(/\.[^.]+$/, '.xlsx'), csvImport: true, conversionDir: directory }
    } else {
      try { await rm(directory, { recursive: true, force: true }) } catch {}
      throw new EngineError('.xls conversion not yet supported — requires SpreadsheetEngine.convertWorkbook wired through SpreadsheetService', 'INTERNAL_ERROR')
    }
  }

  private async snapshotWorkbook(path: string): Promise<string> {
    const dir = join(app.getPath('temp'), 'genoffice-sheets-sessions')
    await mkdir(dir, { recursive: true })
    const snapshotPath = join(dir, `${randomUUID()}.xlsx`)
    const { copyFile } = await import('node:fs/promises')
    await copyFile(path, snapshotPath)
    return snapshotPath
  }

  private async sha256File(path: string): Promise<string> {
    const bytes = await readFile(path)
    return createHash('sha256').update(bytes).digest('hex')
  }

  private recoveryDir(): string { return join(app.getPath('userData'), 'sheets-autosave') }
  private recoveryPathFor(filePath: string): string {
    const hash = createHash('sha1').update(filePath).digest('hex').slice(0, 16)
    return join(this.recoveryDir(), `${hash}.xlsx`)
  }
  private clearWorkbookRecovery(filePath: string): void { try { unlinkSync(this.recoveryPathFor(filePath)) } catch {} }
  private pendingRecoveryFor(filePath: string): string | null {
    const copy = this.recoveryPathFor(filePath)
    try {
      if (!existsSync(copy)) return null
      if (statSync(copy).mtimeMs <= statSync(filePath).mtimeMs) { unlinkSync(copy); return null }
      return copy
    } catch { return null }
  }
}

// ── Startup reconciliation helper ──

/**
 * Reconcile leftover save-commit markers from a previous crash.
 *
 * This must be called once during Sheets main-process startup, BEFORE
 * any migrated workbook operations can begin. It uses app.getPath('userData')
 * to locate the commit marker directory.
 *
 * The function is safe and idempotent — calling it when no markers exist
 * is a no-op.
 *
 * Usage (in sheets-main.ts startup):
 *   import { reconcileSheetsSaveCommits } from './sheets-shell-coordinator'
 *   // ... after app.whenReady():
 *   await reconcileSheetsSaveCommits()
 */
/** Sanitize an AI-provided sheet name into a safe filename base: strip illegal path chars, collapse whitespace, cap length; null if invalid. (Mirrors legacy sanitizeAutoRenameBase.) */
function sanitizeAutoRenameBase(raw: string): string | null {
  const cleaned = raw
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .trim()
  if (!cleaned) return null
  return cleaned.length > 40 ? cleaned.slice(0, 40).trim() : cleaned
}

export async function reconcileSheetsSaveCommits(): Promise<void> {
  await SheetsShellCoordinator.reconcileSaveCommit(app.getPath('userData'))
}
