/**
 * Increment 14 — Real pivot + auto-rename E2E tests.
 *
 * Tests:
 *   1. Pivot fixture creation
 *   2. Real sidecar: open pivot fixture → adopt → readPivotDefinition → verify
 *   3. Auto-rename: coordinator.renameWorkbook with real file on disk
 *   4. Session continuity after rename
 *   5. Legacy mirror update after rename
 */
import { describe, test, expect, vi, beforeEach } from 'vitest'
import { writeFileSync, mkdirSync, existsSync, rmSync, copyFileSync, renameSync } from 'node:fs'
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
import {
  initSheetsRuntime,
  adoptLegacySessionIntoCoordinator,
  type SheetsRuntimeBundle,
  type LegacySessionAdoption,
} from '../src/main/sheets-runtime'
import { buildPivotFixture } from './pivot-fixture-builder'
import type { WorkbookMetadata, WorksheetMetadata } from '@genoffice/runtime-contracts'
import { InvalidSessionError } from '@genoffice/runtime-contracts'

const here = dirname(fileURLToPath(import.meta.url))
// here = .../apps/sheets/tests — go up 3 levels to reach the repo root
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

// ── Tests ────────────────────────────────────────────────────────────

describe.skipIf(!SIDECAR_AVAILABLE)('Increment 14 — Real pivot + auto-rename E2E', () => {
  beforeEach(() => {
    testDir = join(tmpdir(), `genoffice-test-${randomUUID()}`)
    mkdirSync(testDir, { recursive: true })
    vi.clearAllMocks()
  })

  // ═══ 1. Pivot fixture creation ═══

  describe('pivot fixture', () => {
    test('buildPivotFixture creates a valid XLSX', async () => {
      const buf = await buildPivotFixture()
      expect(buf.length).toBeGreaterThan(1000)
      // Write to disk for sidecar verification
      const pivotPath = join(testDir, 'pivot.xlsx')
      writeFileSync(pivotPath, buf)
      expect(existsSync(pivotPath)).toBe(true)
    })
  })

  // ═══ 2. Real sidecar pivot read ═══

  describe('real sidecar pivot read', () => {
    test('open pivot fixture → adopt → readPivotDefinition → verify', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const bundle = initSheetsRuntime({ binaryPath: SIDECAR_BIN, sidecarClient })

      // Create pivot fixture
      const pivotBuf = await buildPivotFixture()
      const pivotPath = join(testDir, 'pivot.xlsx')
      writeFileSync(pivotPath, pivotBuf)

      // Open + adopt
      const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, pivotPath)

      // Read pivot definition — returns the typed WorkbookPivotDefinition
      // contract (NOT `unknown`). The handler in sheets-migrated-handlers.ts
      // runs the value through workbookPivotDefinitionSchema.parse() as a
      // frozen-IPC sanity check; here we verify the typed fields directly.
      const pivotDefinition = await bundle.coordinator.readPivotDefinition(
        100, sessionId,
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
      )

      // Verify the parsed pivot definition has real data — typed access,
      // no `as` cast (the coordinator returns WorkbookPivotDefinition).
      expect(pivotDefinition).toBeDefined()
      expect(pivotDefinition.outputRef).toBeTruthy()
      expect(pivotDefinition.fields.length).toBe(3) // Name, Category, Value

      await bundle.coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('cross-renderer pivot read fails', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const bundle = initSheetsRuntime({ binaryPath: SIDECAR_BIN, sidecarClient })

      const pivotBuf = await buildPivotFixture()
      const pivotPath = join(testDir, 'pivot.xlsx')
      writeFileSync(pivotPath, pivotBuf)

      // Open in renderer A (wcId=100)
      const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, pivotPath)

      // Try to read from renderer B (wcId=200) — should fail
      await expect(bundle.coordinator.readPivotDefinition(
        200, sessionId,
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
      )).rejects.toThrow(InvalidSessionError)

      await bundle.coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('stale session after close → pivot read fails', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const bundle = initSheetsRuntime({ binaryPath: SIDECAR_BIN, sidecarClient })

      const pivotBuf = await buildPivotFixture()
      const pivotPath = join(testDir, 'pivot.xlsx')
      writeFileSync(pivotPath, pivotBuf)

      const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, pivotPath)

      // Close the session
      await bundle.coordinator.closeWorkbook(100, sessionId)

      // Pivot read should fail
      await expect(bundle.coordinator.readPivotDefinition(
        100, sessionId,
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
      )).rejects.toThrow(InvalidSessionError)

      sidecarClient.stop()
    })
  })

  // ═══ 3. Auto-rename with real file ═══

  describe('auto-rename', () => {
    test('successful rename: file moved, session updated, read still works', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({
        service,
      })
      const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

      // Create a fixture file
      const fixturePath = join(testDir, 'Untitled.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      // Open + adopt
      const { sessionId, session } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)

      // Mark as untitled
      coordinator.markUntitledPath(fixturePath)

      // Mock webContents
      const sendCalls: string[] = []
      const mockWc = {
        isDestroyed: () => false,
        send: (channel: string, data: unknown) => { sendCalls.push(`${channel}:${data}`) },
      }

      // Rename
      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, 'My Renamed Sheet',
      )

      expect(result.renamed).toBe(true)
      expect(result.name).toBe('My Renamed Sheet.xlsx')

      // New path exists
      const newPath = join(testDir, 'My Renamed Sheet.xlsx')
      expect(existsSync(newPath)).toBe(true)

      // Old path no longer exists
      expect(existsSync(fixturePath)).toBe(false)

      // Push event was sent to the owning renderer
      expect(sendCalls).toContain('workbook:renamed:My Renamed Sheet.xlsx')

      // Session still valid — originalPath updated
      const updatedSession = coordinator.getSession(100, sessionId)
      expect(updatedSession.sessionId).toBe(sessionId)
      expect(updatedSession.originalPath).toBe(newPath)
      expect(updatedSession.engineHandle).toBe(session.engineHandle)
      expect(updatedSession.snapshotPath).toBe(session.snapshotPath)

      // Read still works
      const readResult = await coordinator.readRange(100, sessionId, 'sheet-1', 'A1:B1')
      expect(readResult.cells).toBeDefined()

      await coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('not untitled → rename refused', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({
        service,
      })
      const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Named.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)

      // NOT marked as untitled — rename should be refused
      const mockWc = {
        isDestroyed: () => false,
        send: () => {},
      }
      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, 'New Name',
      )

      expect(result.renamed).toBe(false)
      // File still exists at original path
      expect(existsSync(fixturePath)).toBe(true)

      await coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('name collision → rename refused with suffix', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({
        service,
      })
      const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Untitled.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      // Create a colliding file
      const collisionPath = join(testDir, 'Collision.xlsx')
      writeFileSync(collisionPath, 'other')

      const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
      coordinator.markUntitledPath(fixturePath)

      const mockWc = { isDestroyed: () => false, send: () => {} }
      // Rename to "Collision" — file exists, should get "Collision-2.xlsx"
      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, 'Collision',
      )

      expect(result.renamed).toBe(true)
      expect(result.name).toBe('Collision-2.xlsx')

      const newPath = join(testDir, 'Collision-2.xlsx')
      expect(existsSync(newPath)).toBe(true)
      expect(existsSync(fixturePath)).toBe(false)
      // Original collision file is untouched
      expect(existsSync(collisionPath)).toBe(true)

      await coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('invalid name → rename refused', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({
        service,
      })
      const bundle = { engine, service, coordinator } as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Untitled.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      const { sessionId } = await openAndAdopt(bundle, 100, sidecarClient, fixturePath)
      coordinator.markUntitledPath(fixturePath)

      const mockWc = { isDestroyed: () => false, send: () => {} }
      // Name with only illegal chars → sanitized to empty → refused
      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, '   ',
      )

      expect(result.renamed).toBe(false)
      expect(existsSync(fixturePath)).toBe(true)

      await coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('unknown renderer → rename fails', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({
        service,
      })

      const mockWc = { isDestroyed: () => false, send: () => {} }
      // Unknown renderer (wcId=999) → InvalidSessionError
      await expect(coordinator.renameWorkbook(
        999, mockWc as unknown as import('electron').WebContents,
        randomUUID(), 'New Name',
      )).rejects.toThrow(InvalidSessionError)

      sidecarClient.stop()
    })
  })
})
