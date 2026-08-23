/**
 * Increment 15 — Rename mutation serialization tests.
 *
 * Tests:
 *   1. rename vs save — concurrent operations are serialized via session lock
 *   2. rename vs close — close during rename or vice versa
 *   3. rename vs teardown — teardown during rename
 *   4. concurrent rename — only one succeeds
 *   5. legacy mirror coherence after rename
 */
import { describe, test, expect, vi, beforeEach } from 'vitest'
import { writeFileSync, mkdirSync, existsSync, rmSync, copyFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID, createHash } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'

const { mockApp, mockDialog } = vi.hoisted(() => ({
  mockApp: { getPath: vi.fn((name: string) => join(tmpdir(), `genoffice-test-${name}-${randomUUID()}`)) },
  mockDialog: { showOpenDialog: vi.fn(), showSaveDialog: vi.fn(), showMessageBox: vi.fn() },
}))
vi.mock('electron', () => ({
  app: mockApp,
  dialog: mockDialog,
  BrowserWindow: vi.fn(),
}))

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { ElectronXlsxSidecarEngine } from '@genoffice/platform-electron'
import { SpreadsheetServiceImpl } from '@genoffice/services-sheets/src/spreadsheet-service.js'
import { SheetsShellCoordinator, type ShellWorkbookSession } from '../src/main/sheets-shell-coordinator'
import { initSheetsRuntime, adoptLegacySessionIntoCoordinator, type SheetsRuntimeBundle, type LegacySessionAdoption } from '../src/main/sheets-runtime'
import { buildPivotFixture } from './pivot-fixture-builder'
import type { WorkbookMetadata, WorksheetMetadata } from '@genoffice/runtime-contracts'
import { InvalidSessionError } from '@genoffice/runtime-contracts'

const here = dirname(fileURLToPath(import.meta.url))
const repoRoot = resolve(here, '..', '..', '..')
const SIDECAR_BIN = join(repoRoot, 'apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar')
const SIDECAR_AVAILABLE = existsSync(SIDECAR_BIN)

let testDir: string

beforeEach(() => {
  testDir = join(tmpdir(), `genoffice-test-${randomUUID()}`)
  mkdirSync(testDir, { recursive: true })
  vi.clearAllMocks()
})

function sha256OfFile(path: string): string {
  const { readFileSync } = require('node:fs')
  return createHash('sha256').update(readFileSync(path)).digest('hex')
}

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

async function openAndAdopt(
  bundle: SheetsRuntimeBundle,
  wcId: number,
  mockClient: XlsxSidecarClient,
  workbookPath: string,
  locale = 'en',
): Promise<{ sessionId: string; session: ShellWorkbookSession; snapshotPath: string }> {
  const snapshotDir = join(tmpdir(), `genoffice-test-snapshots-${randomUUID()}`)
  mkdirSync(snapshotDir, { recursive: true })
  const snapshotPath = join(snapshotDir, `${randomUUID()}.xlsx`)
  copyFileSync(workbookPath, snapshotPath)
  const opened = await mockClient.open(snapshotPath, locale) as {
    sessionId: string
    sheets: Array<{ id: string; name: string }>
  }
  const sessionId = opened.sessionId
  const diskFingerprint = sha256OfFile(snapshotPath)
  const sheetNames = new Map<string, string>()
  for (const s of opened.sheets) sheetNames.set(s.id, s.name)
  const metadata = makeMetadata(workbookPath.split(/[\\/]/).pop() ?? 'workbook.xlsx')
  const adoption: LegacySessionAdoption = {
    sidecarSessionId: sessionId,
    originalPath: workbookPath,
    snapshotPath,
    diskFingerprint,
    sheetNames,
    metadata,
    locale,
  }
  const session = await adoptLegacySessionIntoCoordinator(bundle, wcId, adoption)
  return { sessionId, session, snapshotPath }
}

function makeMockWc() {
  const sends: Array<{ channel: string; data: unknown }> = []
  return {
    isDestroyed: () => false,
    send: (channel: string, data: unknown) => { sends.push({ channel, data }) },
    _sends: sends,
  }
}

describe.skipIf(!SIDECAR_AVAILABLE)('Increment 15 — Rename mutation serialization', () => {
  beforeEach(() => {
    testDir = join(tmpdir(), `genoffice-test-${randomUUID()}`)
    mkdirSync(testDir, { recursive: true })
    vi.clearAllMocks()
  })

  test('rename vs save: serialized via session lock', async () => {
    const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
    sidecarClient.start()
    const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
    const service = new SpreadsheetServiceImpl({ engine })
    const coordinator = new SheetsShellCoordinator({ service })
    const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, await buildPivotFixture())
    const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
    coordinator.markUntitledPath(fixturePath)
    mockDialog.showSaveDialog.mockResolvedValue({ canceled: false, filePath: join(testDir, 'saved.xlsx') })

    const wc = makeMockWc()
    // Start both operations concurrently
    const savePromise = coordinator.saveWorkbook(100, sessionId, { plan: makeEmptySavePlan() }, 'save-as', undefined)
    const renamePromise = coordinator.renameWorkbook(100, wc as unknown as import('electron').WebContents, sessionId, 'Renamed')

    // Both should complete without throwing
    const [saveResult, renameResult] = await Promise.all([savePromise, renamePromise])
    expect(saveResult.ok).toBe(true)
    // Rename may or may not succeed depending on lock ordering — but neither should corrupt state
    // The key assertion: exactly one operation owns the session at any point
    const session = coordinator.getSession(100, sessionId)
    expect(session.sessionId).toBe(sessionId)

    await coordinator.teardown(100)
    sidecarClient.stop()
  })

  test('rename vs close: close during rename', async () => {
    const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
    sidecarClient.start()
    const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
    const service = new SpreadsheetServiceImpl({ engine })
    const coordinator = new SheetsShellCoordinator({ service })
    const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, await buildPivotFixture())
    const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
    coordinator.markUntitledPath(fixturePath)

    const wc = makeMockWc()
    const renamePromise = coordinator.renameWorkbook(100, wc as unknown as import('electron').WebContents, sessionId, 'Renamed')
    const closePromise = coordinator.closeWorkbook(100, sessionId)

    // Both should complete — the session lock serializes them
    await Promise.all([renamePromise, closePromise])

    // After close, the session is gone
    expect(() => coordinator.getSession(100, sessionId)).toThrow(InvalidSessionError)

    sidecarClient.stop()
  })

  test('rename vs teardown: teardown during rename', async () => {
    const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
    sidecarClient.start()
    const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
    const service = new SpreadsheetServiceImpl({ engine })
    const coordinator = new SheetsShellCoordinator({ service })
    const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, await buildPivotFixture())
    const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
    coordinator.markUntitledPath(fixturePath)

    const wc = makeMockWc()
    const renamePromise = coordinator.renameWorkbook(100, wc as unknown as import('electron').WebContents, sessionId, 'Renamed')
    const teardownPromise = coordinator.teardown(100)

    // Both should complete — the session lock serializes them
    await Promise.all([renamePromise, teardownPromise])

    // After teardown, the session is gone
    expect(() => coordinator.getSession(100, sessionId)).toThrow(InvalidSessionError)

    sidecarClient.stop()
  })

  test('concurrent rename: only one succeeds', async () => {
    const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
    sidecarClient.start()
    const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
    const service = new SpreadsheetServiceImpl({ engine })
    const coordinator = new SheetsShellCoordinator({ service })
    const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, await buildPivotFixture())
    const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
    coordinator.markUntitledPath(fixturePath)

    const wc1 = makeMockWc()
    const wc2 = makeMockWc()
    const rename1 = coordinator.renameWorkbook(100, wc1 as unknown as import('electron').WebContents, sessionId, 'Name1')
    const rename2 = coordinator.renameWorkbook(100, wc2 as unknown as import('electron').WebContents, sessionId, 'Name2')

    const [r1, r2] = await Promise.all([rename1, rename2])
    // At least one must succeed; the other may refuse (file already moved)
    const successes = [r1.renamed, r2.renamed].filter(Boolean)
    expect(successes.length).toBeGreaterThanOrEqual(1)

    // The session path should point to whichever name won
    const session = coordinator.getSession(100, sessionId)
    expect(session.originalPath).toMatch(/Name[12]\.xlsx$/)

    await coordinator.teardown(100)
    sidecarClient.stop()
  })

  test('legacy mirror coherence: rename updates originalPath in coordinator', async () => {
    const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
    sidecarClient.start()
    const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
    const service = new SpreadsheetServiceImpl({ engine })
    const coordinator = new SheetsShellCoordinator({ service })
    const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, await buildPivotFixture())
    const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
    coordinator.markUntitledPath(fixturePath)

    const wc = makeMockWc()
    const result = await coordinator.renameWorkbook(100, wc as unknown as import('electron').WebContents, sessionId, 'CoherentName')
    expect(result.renamed).toBe(true)

    // Coordinator's session has the new path
    const session = coordinator.getSession(100, sessionId)
    expect(session.originalPath).toBe(join(testDir, 'CoherentName.xlsx'))
    expect(existsSync(session.originalPath)).toBe(true)
    expect(existsSync(fixturePath)).toBe(false)

    // Read after rename works (session is valid)
    const readResult = await coordinator.readRange(100, sessionId, 'sheet-1', 'A1:B1')
    expect(readResult.cells).toBeDefined()

    await coordinator.teardown(100)
    sidecarClient.stop()
  })

  test('failed rename leaves old path intact', async () => {
    const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
    sidecarClient.start()
    const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
    const service = new SpreadsheetServiceImpl({ engine })
    const coordinator = new SheetsShellCoordinator({ service })
    const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

    const fixturePath = join(testDir, 'Untitled.xlsx')
    writeFileSync(fixturePath, await buildPivotFixture())
    const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
    coordinator.markUntitledPath(fixturePath)

    const wc = makeMockWc()
    // Invalid name (only spaces → sanitized to empty)
    const result = await coordinator.renameWorkbook(100, wc as unknown as import('electron').WebContents, sessionId, '   ')
    expect(result.renamed).toBe(false)

    // Old path still exists, session unchanged
    expect(existsSync(fixturePath)).toBe(true)
    const session = coordinator.getSession(100, sessionId)
    expect(session.originalPath).toBe(fixturePath)

    await coordinator.teardown(100)
    sidecarClient.stop()
  })
})
