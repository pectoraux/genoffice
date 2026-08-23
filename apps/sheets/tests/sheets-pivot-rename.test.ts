/**
 * Increment 16 — Real pivot + auto-rename E2E tests (cutover to coordinator.openWorkbook).
 *
 * Tests:
 *   1. Pivot fixture creation
 *   2. Real sidecar: open pivot fixture via coordinator.openWorkbook → readPivotDefinition → verify
 *   3. Auto-rename: coordinator.renameWorkbook with real file on disk
 *   4. Session continuity after rename
 *   5. Legacy mirror update after rename
 *
 * Phase 2 Increment 16: the coordinator is the SOLE owner of workbook sessions.
 * Tests use `coordinator.openWorkbook(wcId, undefined, { queuedPath, locale })`
 * directly — no legacy open, no adoption bridge.
 */
import { describe, test, expect, vi, beforeEach } from 'vitest'
import { writeFileSync, mkdirSync, existsSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'
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
import { SheetsShellCoordinator } from '../src/main/sheets-shell-coordinator'
import { initSheetsRuntime, type SheetsRuntimeBundle } from '../src/main/sheets-runtime'
import { buildPivotFixture } from './pivot-fixture-builder'
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

/**
 * Open a workbook via `coordinator.openWorkbook()` with `queuedPath` —
 * the migrated open path. No legacy open, no adoption bridge.
 * The coordinator owns the entire lifecycle: snapshot, engine.open, session registration.
 */
async function openViaCoordinator(
  bundle: SheetsRuntimeBundle,
  wcId: number,
  workbookPath: string,
  locale = 'en',
): Promise<{ sessionId: string; session: ReturnType<SheetsShellCoordinator['getSession']> }> {
  const result = await bundle.coordinator.openWorkbook(wcId, undefined, {
    queuedPath: workbookPath,
    locale,
  })
  if (!result) throw new Error('openWorkbook returned null')
  return { sessionId: result.sessionId, session: result.session }
}

// ── Tests ────────────────────────────────────────────────────────────

describe.skipIf(!SIDECAR_AVAILABLE)('Increment 16 — Real pivot + auto-rename E2E (coordinator.open)', () => {
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
      const pivotPath = join(testDir, 'pivot.xlsx')
      writeFileSync(pivotPath, buf)
      expect(existsSync(pivotPath)).toBe(true)
    })
  })

  // ═══ 2. Real sidecar pivot read via coordinator.open ═══

  describe('real sidecar pivot read (coordinator.open)', () => {
    test('open pivot fixture via coordinator → readPivotDefinition → verify', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const bundle = initSheetsRuntime({ binaryPath: SIDECAR_BIN, sidecarClient })

      const pivotBuf = await buildPivotFixture()
      const pivotPath = join(testDir, 'pivot.xlsx')
      writeFileSync(pivotPath, pivotBuf)

      const { sessionId } = await openViaCoordinator(bundle, 100, pivotPath)

      const pivotDefinition = await bundle.coordinator.readPivotDefinition(
        100, sessionId,
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
      )

      expect(pivotDefinition).toBeDefined()
      expect(pivotDefinition.outputRef).toBeTruthy()
      expect(pivotDefinition.fields.length).toBe(3)

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

      const { sessionId } = await openViaCoordinator(bundle, 100, pivotPath)

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

      const { sessionId } = await openViaCoordinator(bundle, 100, pivotPath)

      await bundle.coordinator.closeWorkbook(100, sessionId)

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
      const coordinator = new SheetsShellCoordinator({ service })
      const bundle = { engine, service, coordinator } as unknown as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Untitled.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      const { sessionId, session } = await openViaCoordinator(bundle, 100, fixturePath)
      coordinator.markUntitledPath(fixturePath)

      const sendCalls: string[] = []
      const mockWc = {
        isDestroyed: () => false,
        send: (channel: string, data: unknown) => { sendCalls.push(`${channel}:${data}`) },
      }

      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, 'My Renamed Sheet',
      )

      expect(result.renamed).toBe(true)
      expect(result.name).toBe('My Renamed Sheet.xlsx')

      const newPath = join(testDir, 'My Renamed Sheet.xlsx')
      expect(existsSync(newPath)).toBe(true)
      expect(existsSync(fixturePath)).toBe(false)

      expect(sendCalls).toContain('workbook:renamed:My Renamed Sheet.xlsx')

      const updatedSession = coordinator.getSession(100, sessionId)
      expect(updatedSession.sessionId).toBe(sessionId)
      expect(updatedSession.originalPath).toBe(newPath)
      expect(updatedSession.engineHandle).toBe(session.engineHandle)
      expect(updatedSession.snapshotPath).toBe(session.snapshotPath)

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
      const coordinator = new SheetsShellCoordinator({ service })
      const bundle = { engine, service, coordinator } as unknown as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Named.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      const { sessionId } = await openViaCoordinator(bundle, 100, fixturePath)

      const mockWc = { isDestroyed: () => false, send: () => {} }
      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, 'New Name',
      )

      expect(result.renamed).toBe(false)
      expect(existsSync(fixturePath)).toBe(true)

      await coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('name collision → rename with suffix', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({ service })
      const bundle = { engine, service, coordinator } as unknown as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Untitled.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      const collisionPath = join(testDir, 'Collision.xlsx')
      writeFileSync(collisionPath, 'other')

      const { sessionId } = await openViaCoordinator(bundle, 100, fixturePath)
      coordinator.markUntitledPath(fixturePath)

      const mockWc = { isDestroyed: () => false, send: () => {} }
      const result = await coordinator.renameWorkbook(
        100, mockWc as unknown as import('electron').WebContents,
        sessionId, 'Collision',
      )

      expect(result.renamed).toBe(true)
      expect(result.name).toBe('Collision-2.xlsx')

      const newPath = join(testDir, 'Collision-2.xlsx')
      expect(existsSync(newPath)).toBe(true)
      expect(existsSync(fixturePath)).toBe(false)
      expect(existsSync(collisionPath)).toBe(true)

      await coordinator.teardown(100)
      sidecarClient.stop()
    })

    test('invalid name → rename refused', async () => {
      const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
      sidecarClient.start()
      const engine = new ElectronXlsxSidecarEngine({ binaryPath: SIDECAR_BIN, sidecarClient })
      const service = new SpreadsheetServiceImpl({ engine })
      const coordinator = new SheetsShellCoordinator({ service })
      const bundle = { engine, service, coordinator } as unknown as SheetsRuntimeBundle

      const fixturePath = join(testDir, 'Untitled.xlsx')
      const pivotBuf = await buildPivotFixture()
      writeFileSync(fixturePath, pivotBuf)

      const { sessionId } = await openViaCoordinator(bundle, 100, fixturePath)
      coordinator.markUntitledPath(fixturePath)

      const mockWc = { isDestroyed: () => false, send: () => {} }
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
      const coordinator = new SheetsShellCoordinator({ service })

      const mockWc = { isDestroyed: () => false, send: () => {} }
      await expect(coordinator.renameWorkbook(
        999, mockWc as unknown as import('electron').WebContents,
        randomUUID(), 'New Name',
      )).rejects.toThrow(InvalidSessionError)

      sidecarClient.stop()
    })
  })
})
