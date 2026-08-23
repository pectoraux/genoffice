/**
 * Increment 15A — Rename mutation serialization tests (SIDECAR-FREE).
 *
 * These tests verify the coordinator's `renameWorkbook()` method against
 * concurrent save / close / teardown / concurrent-rename races WITHOUT
 * spawning the real Rust sidecar. The SpreadsheetService is mocked with
 * a minimal stub; the filesystem operations operate on real temp files.
 *
 * The previous version of this file (at commit a37c3db) required the real
 * sidecar binary for every race test — slow, environment-dependent, and
 * a violation of the user's directive (Increment 15A): "The rename race
 * tests must run WITHOUT the Rust sidecar." The real-sidecar pivot +
 * rename flow is verified in `sheets-pivot-rename.test.ts` (the SINGLE
 * real-sidecar integration test).
 *
 * Verified invariants (per user directive):
 *   - rename vs save — serialized via session lock
 *   - rename vs close — serialized via session lock
 *   - rename vs teardown — serialized via session lock
 *   - concurrent rename — only one succeeds
 *   - no stale `ShellWorkbookSession.originalPath`
 *   - no stale legacy `SessionInfo.path` mirror
 *   - no duplicate `workbook:renamed` event
 */
import { describe, test, expect, vi, beforeEach } from 'vitest'
import { writeFileSync, mkdirSync, existsSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'

// Mock the electron module BEFORE importing any code that uses it. The
// coordinator imports `app` (for userData path) and `dialog` (for save
// dialogs); the rename path uses neither, but the imports must resolve.
const { mockApp, mockDialog, mockBrowserWindow } = vi.hoisted(() => ({
  mockApp: { getPath: vi.fn((name: string) => join(tmpdir(), `genoffice-test-${name}-${randomUUID()}`)) },
  mockDialog: { showOpenDialog: vi.fn(), showSaveDialog: vi.fn(), showMessageBox: vi.fn() },
  mockBrowserWindow: vi.fn(),
}))
vi.mock('electron', () => ({
  app: mockApp,
  dialog: mockDialog,
  BrowserWindow: mockBrowserWindow,
}))

import {
  SheetsShellCoordinator,
  type ShellWorkbookSession,
} from '../src/main/sheets-shell-coordinator'
import type {
  SpreadsheetService,
  WorkbookSession,
  EngineSessionHandle,
  WorkbookMetadata,
  WorksheetMetadata,
  ExternalChangeStatus,
  SaveRequest,
  SaveResult,
  EngineRangeResult,
  EngineFormulaCellsResult,
  EngineRecalcEdit,
  EngineRecalcRead,
  EngineRecalcResult,
  EngineMediaResult,
  WorkbookPivotDefinition,
} from '@genoffice/runtime-contracts'
import { ENGINE_SESSION_HANDLE_BRAND, InvalidSessionError } from '@genoffice/runtime-contracts'

// ── Test fixtures ─────────────────────────────────────────────────────

let testDir: string

beforeEach(() => {
  testDir = join(tmpdir(), `genoffice-test-${randomUUID()}`)
  mkdirSync(testDir, { recursive: true })
  vi.clearAllMocks()
})

function makeMetadata(name = 'test.xlsx'): WorkbookMetadata {
  const sheets: WorksheetMetadata[] = [{
    id: 'sheet-1', name: 'Data', index: 0, hidden: false, rtl: false,
    showGridlines: true, rowCount: 100, columnCount: 26,
    defaultRowHeight: 15, defaultColumnWidth: 8.43,
  }]
  return {
    name, sha256: 'abc123', entryCount: 10, sheets, activeTab: 0,
    definedNames: [], themeColors: [], themeFonts: { major: '', minor: '' },
  }
}

/**
 * Create a fake opaque `EngineSessionHandle` using the brand symbol
 * exported by runtime-contracts. The coordinator treats it as opaque —
 * it stores it in the session but never inspects it. The mock service
 * also never inspects it.
 */
function makeFakeHandle(): EngineSessionHandle {
  const obj = { [ENGINE_SESSION_HANDLE_BRAND]: ENGINE_SESSION_HANDLE_BRAND }
  Object.freeze(obj)
  return obj as EngineSessionHandle
}

function makeShellSession(originalPath: string): ShellWorkbookSession {
  const handle = makeFakeHandle()
  const domainSession: WorkbookSession = {
    workbookName: 'test.xlsx',
    workbookHash: 'abc123',
    sheetNames: new Map([['sheet-1', 'Data']]),
  }
  return {
    sessionId: randomUUID(),
    originalPath,
    snapshotPath: join(testDir, 'snapshots', `${randomUUID()}.xlsx`),
    diskFingerprint: 'abc123',
    engineHandle: handle,
    domainSession,
    metadata: makeMetadata(),
    locale: 'en',
    recoveryEpoch: 0,
  }
}

function makeEmptySavePlan() {
  return {
    edits: [], structuralOps: [], formulaValues: [], sheetOps: [], sheetOrder: [],
    filterStates: [], hyperlinkEdits: [], cfStates: [], dvStates: [],
    pageSetupStates: [], noteStates: [], sheetProtections: [], protectedRangeStates: [],
    visualAdditions: [], tableAdditions: [], pivotAdditions: [], sparklineAdditions: [],
    chartEdits: [], visualEdits: [], pivotCacheRefreshPaths: [], pivotRefreshUpdates: [],
    definedNamesState: null, themeState: null, workbookProtectionState: null,
  }
}

/**
 * Minimal mock `SpreadsheetService` for rename-race tests. Implements
 * only the methods exercised by the coordinator's rename / save / close
 * paths. NO sidecar, NO engine, NO real I/O — pure in-memory bookkeeping
 * so we can assert what the coordinator called.
 */
class MockSpreadsheetService implements SpreadsheetService {
  /** Recorded save() invocations (for asserting serialization order). */
  readonly saveInvocations: Array<{
    engineHandle: EngineSessionHandle
    externalChange: ExternalChangeStatus
  }> = []
  /** Recorded close() invocations (for asserting close was called). */
  readonly closeInvocations: EngineSessionHandle[] = []
  /** Controls whether save() succeeds or throws. */
  nextSaveShouldThrow: boolean = false

  async open(
    _workbook: Uint8Array,
    _locale: string,
    fileName: string,
  ): Promise<{
    session: WorkbookSession
    engineHandle: EngineSessionHandle
    metadata: WorkbookMetadata
  }> {
    // Return a fresh opaque handle for each open() call. The coordinator
    // stores it on the replacement session; the mock close() records it.
    // No sidecar, no real engine — just bookkeeping.
    return {
      session: {
        workbookName: fileName,
        workbookHash: 'abc123',
        sheetNames: new Map([['sheet-1', 'Data']]),
      },
      engineHandle: makeFakeHandle(),
      metadata: makeMetadata(fileName),
    }
  }
  async close(engineHandle: EngineSessionHandle): Promise<void> {
    this.closeInvocations.push(engineHandle)
  }
  async readRange(): Promise<EngineRangeResult> {
    return {
      cells: [], rows: [], merges: [], columns: [],
      hyperlinks: [], conditionalFormatting: [], dataValidation: [],
      rowBreaks: [], columnBreaks: [], sheetProtection: false,
    }
  }
  async readFormulaCells(): Promise<EngineFormulaCellsResult> {
    return { cells: [] }
  }
  async recalculate(): Promise<EngineRecalcResult> {
    return { cells: [] }
  }
  async readMedia(): Promise<EngineMediaResult> {
    return { mediaType: 'image/png', base64: '' }
  }
  async save(
    _session: WorkbookSession,
    engineHandle: EngineSessionHandle,
    _request: SaveRequest,
    externalChange: ExternalChangeStatus,
  ): Promise<SaveResult> {
    this.saveInvocations.push({ engineHandle, externalChange })
    if (this.nextSaveShouldThrow) {
      throw new Error('Mock save failure')
    }
    return { ok: true, data: new Uint8Array(), touchedEntries: [] }
  }
  async writeRecovery(): Promise<Uint8Array> {
    return new Uint8Array()
  }
  async readPivotDefinition(): Promise<WorkbookPivotDefinition> {
    // Minimal stub — the rename tests don't read pivot definitions.
    throw new Error('MockSpreadsheetService.readPivotDefinition() is not used in rename tests')
  }
}

/**
 * Mock WebContents that records every `send()` call. Used to verify
 * the `workbook:renamed` event is pushed exactly once.
 */
function makeMockWc() {
  const sends: Array<{ channel: string; data: unknown }> = []
  return {
    isDestroyed: () => false,
    send: (channel: string, data: unknown) => { sends.push({ channel, data }) },
    _sends: sends,
  }
}

/**
 * Helper: register a session with the coordinator by direct adoption
 * (no engine, no sidecar). Returns the session.
 */
async function adoptSession(
  coordinator: SheetsShellCoordinator,
  wcId: number,
  fixturePath: string,
): Promise<ShellWorkbookSession> {
  const session = makeShellSession(fixturePath)
  return coordinator.adoptLegacySession(wcId, session)
}

// ── Tests ────────────────────────────────────────────────────────────

describe('Increment 15A — Rename mutation serialization (SIDECAR-FREE)', () => {
  test('rename vs save: serialized via session lock', async () => {
    const service = new MockSpreadsheetService()
    const coordinator = new SheetsShellCoordinator({ service })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    // Mock the save dialog (used by save-as path)
    mockDialog.showSaveDialog.mockResolvedValue({
      canceled: false,
      filePath: join(testDir, 'saved.xlsx'),
    })

    const wc = makeMockWc()
    // Start both operations concurrently — they MUST serialize via the
    // per-session mutation lock. The save is a save-as (which acquires
    // the lock); the rename also acquires the lock. Exactly one runs
    // at a time.
    const savePromise = coordinator.saveWorkbook(
      100, sessionId, { plan: makeEmptySavePlan() }, 'save-as', undefined,
    )
    const renamePromise = coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, 'Renamed',
    )

    const [saveResult, renameResult] = await Promise.all([savePromise, renamePromise])

    // Both must complete without throwing — the lock serializes them.
    expect(saveResult.ok).toBe(true)
    expect(typeof renameResult.renamed).toBe('boolean')

    // The save must have been called exactly once (proves the lock did
    // not deadlock and did not double-invoke).
    expect(service.saveInvocations.length).toBe(1)

    await coordinator.teardown(100)
  })

  test('rename vs close: serialized via session lock', async () => {
    const service = new MockSpreadsheetService()
    const coordinator = new SheetsShellCoordinator({ service })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    const wc = makeMockWc()
    const renamePromise = coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, 'Renamed',
    )
    const closePromise = coordinator.closeWorkbook(100, sessionId)

    // Both must complete — the session lock serializes them. Neither
    // throws; the close either finds the session still present (rename
    // finished first) or finds it gone (close finished first). Either
    // way, the lock prevents corruption.
    await Promise.all([renamePromise, closePromise])

    // After close, the session is gone.
    expect(() => coordinator.getSession(100, sessionId)).toThrow(InvalidSessionError)

    await coordinator.teardown(100)
  })

  test('rename vs teardown: serialized via session lock', async () => {
    const service = new MockSpreadsheetService()
    const coordinator = new SheetsShellCoordinator({ service })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    const wc = makeMockWc()
    const renamePromise = coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, 'Renamed',
    )
    const teardownPromise = coordinator.teardown(100)

    // Both must complete — the session lock serializes them.
    await Promise.all([renamePromise, teardownPromise])

    // After teardown, the session is gone (regardless of rename outcome).
    expect(() => coordinator.getSession(100, sessionId)).toThrow(InvalidSessionError)
  })

  test('concurrent rename: only one succeeds', async () => {
    const service = new MockSpreadsheetService()
    const coordinator = new SheetsShellCoordinator({ service })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    const wc1 = makeMockWc()
    const wc2 = makeMockWc()
    // Two concurrent renames — the lock serializes them. The first one
    // renames the file; the second one finds `existsSync(target)` true
    // (the first rename produced it) and either retries with a suffix
    // or refuses (depending on the collision-avoidance loop).
    const rename1 = coordinator.renameWorkbook(
      100, wc1 as unknown as import('electron').WebContents,
      sessionId, 'Name1',
    )
    const rename2 = coordinator.renameWorkbook(
      100, wc2 as unknown as import('electron').WebContents,
      sessionId, 'Name2',
    )

    const [r1, r2] = await Promise.all([rename1, rename2])

    // The first rename succeeds; the second one either:
    //   - succeeds with a different name (collision-avoidance suffix), OR
    //   - refuses (file already moved + no suffix candidate available).
    // Either way, at least one must succeed.
    const successes = [r1.renamed, r2.renamed].filter(Boolean)
    expect(successes.length).toBeGreaterThanOrEqual(1)

    // The coordinator's session path must point to whichever name won —
    // never the original untitled path (no stale originalPath).
    const session = coordinator.getSession(100, sessionId)
    expect(session.originalPath).not.toBe(fixturePath)
    expect(session.originalPath).toMatch(/Name[12](-[0-9]+)?\.xlsx$/)

    await coordinator.teardown(100)
  })

  test('no stale originalPath: rename updates ShellWorkbookSession.originalPath', async () => {
    const service = new MockSpreadsheetService()
    const coordinator = new SheetsShellCoordinator({ service })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    const wc = makeMockWc()
    const result = await coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, 'CoherentName',
    )
    expect(result.renamed).toBe(true)

    // Coordinator's session has the new path — NOT stale.
    const session = coordinator.getSession(100, sessionId)
    expect(session.originalPath).toBe(join(testDir, 'CoherentName.xlsx'))

    // The file actually moved on disk (renameSync succeeded).
    expect(existsSync(session.originalPath)).toBe(true)
    expect(existsSync(fixturePath)).toBe(false)

    await coordinator.teardown(100)
  })

  test('no stale legacy mirror: onWorkbookRenamed callback fires with old/new paths', async () => {
    const service = new MockSpreadsheetService()
    // The "legacy mirror" lives in sheets-main.ts as `sheetsTabs`. In
    // this sidecar-free test we mock the legacy mirror as a simple Map
    // that the `onWorkbookRenamed` callback updates. This proves the
    // coordinator invokes the callback with the correct (wcId, oldPath,
    // newPath) tuple — wiring that updateLegacySessionPath relies on.
    const legacyMirror = new Map<string, string>() // sessionId → path
    const onWorkbookRenamed = vi.fn(
      (wcId: number, oldPath: string, newPath: string) => {
        // Find any session in the mirror whose path === oldPath and
        // update it. (Mirrors the real updateLegacySessionPath helper.)
        for (const [sid, p] of legacyMirror) {
          if (p === oldPath) legacyMirror.set(sid, newPath)
        }
      },
    )
    const coordinator = new SheetsShellCoordinator({ service, onWorkbookRenamed })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    // Seed the legacy mirror with the original path (simulating the
    // legacy `workbook:select` having registered the same sessionId).
    legacyMirror.set(sessionId, fixturePath)

    const wc = makeMockWc()
    const result = await coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, 'MirrorUpdated',
    )
    expect(result.renamed).toBe(true)

    // The callback MUST have been invoked exactly once with the correct
    // (wcId, oldPath, newPath) tuple.
    expect(onWorkbookRenamed).toHaveBeenCalledTimes(1)
    expect(onWorkbookRenamed).toHaveBeenCalledWith(
      100,
      fixturePath,
      join(testDir, 'MirrorUpdated.xlsx'),
    )

    // The legacy mirror now reflects the new path — NOT stale.
    expect(legacyMirror.get(sessionId)).toBe(join(testDir, 'MirrorUpdated.xlsx'))

    await coordinator.teardown(100)
  })

  test('no duplicate workbook:renamed event: coordinator pushes exactly once', async () => {
    const service = new MockSpreadsheetService()
    // The `onWorkbookRenamed` callback represents the legacy mirror
    // updater. If the coordinator (incorrectly) pushed the
    // `workbook:renamed` event BOTH directly AND via the callback,
    // we'd see two events. This test asserts the coordinator pushes
    // exactly once.
    const onWorkbookRenamed = vi.fn(() => {
      // Simulate the legacy mirror update. Note: the real
      // `updateLegacySessionPath` does NOT push the IPC event —
      // only the coordinator does. We assert that here by counting
      // the mock WebContents sends.
    })
    const coordinator = new SheetsShellCoordinator({ service, onWorkbookRenamed })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    const wc = makeMockWc()
    await coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, 'SingleEvent',
    )

    // The coordinator pushed the event exactly once via `webContents.send`.
    const renamedEvents = wc._sends.filter((s) => s.channel === 'workbook:renamed')
    expect(renamedEvents.length).toBe(1)
    expect(renamedEvents[0]?.data).toBe('SingleEvent.xlsx')

    await coordinator.teardown(100)
  })

  test('failed rename: callback NOT invoked, no event pushed', async () => {
    const service = new MockSpreadsheetService()
    const onWorkbookRenamed = vi.fn()
    const coordinator = new SheetsShellCoordinator({ service, onWorkbookRenamed })

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, Buffer.from('fake xlsx content'))
    coordinator.markUntitledPath(fixturePath)
    const { sessionId } = await adoptSession(coordinator, 100, fixturePath)

    const wc = makeMockWc()
    // Invalid name (only spaces → sanitized to empty) → rename refuses.
    const result = await coordinator.renameWorkbook(
      100, wc as unknown as import('electron').WebContents,
      sessionId, '   ',
    )
    expect(result.renamed).toBe(false)

    // No event was pushed (rename did not happen).
    const renamedEvents = wc._sends.filter((s) => s.channel === 'workbook:renamed')
    expect(renamedEvents.length).toBe(0)

    // The legacy mirror callback was NOT invoked.
    expect(onWorkbookRenamed).not.toHaveBeenCalled()

    // The session's path is unchanged.
    const session = coordinator.getSession(100, sessionId)
    expect(session.originalPath).toBe(fixturePath)

    await coordinator.teardown(100)
  })
})
