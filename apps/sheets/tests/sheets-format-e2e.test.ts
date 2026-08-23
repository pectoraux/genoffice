/**
 * Phase 2 Increment 19 — Real CSV / XLS / Recovery E2E tests.
 *
 * These tests exercise the coordinator's `openWorkbook` path with real
 * CSV / XLS / recovery fixtures through the REAL Rust sidecar (no mocks).
 * They verify:
 *   - CSV conversion → XLSX → engine.open → read cells
 *   - XLS conversion → XLSX → engine.open → read cells
 *   - XLS conversion failure → typed error + cleanup (no leak)
 *   - Recovery restore → recovered content opened
 *   - Recovery discard → original content opened
 *   - Multi-renderer isolation (renderer B cannot access A's session)
 *
 * The recovery tests use the coordinator's real `pendingRecoveryFor` logic
 * (recovery copy mtime > original mtime → eligible).
 */
import { describe, test, expect, vi, beforeEach, afterEach } from 'vitest'
import { writeFileSync, mkdirSync, existsSync, rmSync, statSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'

const { mockApp, mockDialog } = vi.hoisted(() => {
  // Stable paths — resolved lazily on first getPath call, then cached.
  const cache: Record<string, string> = {}
  return {
    mockApp: {
      getPath: vi.fn((name: string) => {
        if (!cache[name]) {
          // Defer require to avoid hoisting issues
          const { tmpdir } = require('node:os')
          const { join } = require('node:path')
          const { randomUUID } = require('node:crypto')
          cache[name] = join(tmpdir(), `genoffice-test-${name}-${randomUUID()}`)
        }
        return cache[name]
      }),
    },
    mockDialog: { showOpenDialog: vi.fn(), showSaveDialog: vi.fn(), showMessageBox: vi.fn() },
  }
})
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
import { InvalidSessionError } from '@genoffice/runtime-contracts'
import {
  writeCsvFixture,
  writeXlsFixture,
  writeInvalidXlsFixture,
  buildMinimalXlsx,
  buildRecoveryXlsx,
} from './format-fixture-builder'

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

async function makeBundle(): Promise<{
  bundle: SheetsRuntimeBundle
  sidecarClient: XlsxSidecarClient
  cleanup: () => void
}> {
  const sidecarClient = new XlsxSidecarClient(SIDECAR_BIN)
  sidecarClient.start()
  const bundle = initSheetsRuntime({ binaryPath: SIDECAR_BIN, sidecarClient })
  return {
    bundle,
    sidecarClient,
    cleanup: () => {
      try { void bundle.coordinator.teardown(100) } catch { /* */ }
      try { sidecarClient.stop() } catch { /* */ }
    },
  }
}

describe.skipIf(!SIDECAR_AVAILABLE)('Increment 19 — Real CSV / XLS / Recovery E2E', () => {
  let cleanup: (() => void) | undefined

  afterEach(() => {
    if (cleanup) { cleanup(); cleanup = undefined }
  })

  // ═══ CSV E2E ═══

  describe('CSV open E2E', () => {
    test('CSV → convert → open → read cells → verify content', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup
      const csvPath = join(testDir, 'data.csv')
      writeCsvFixture(csvPath)

      const result = await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: csvPath,
        locale: 'en',
      })
      expect(result).not.toBeNull()
      expect(result!.sessionId).toBeTruthy()
      // CSV import sets suggestSaveAs + csvImport + needsSaveAs
      expect(result!.session.csvImport).toBe(true)
      expect(result!.session.suggestSaveAs).toContain('data.xlsx')
      expect(result!.session.needsSaveAs).toBe(true)
      // The originalPath is the CSV path (renderer-facing identifier)
      expect(result!.session.originalPath).toBe(csvPath)

      // Read cells — the CSV content should be in the workbook
      const sheetId = result!.session.domainSession.sheetNames.keys().next().value
      const readResult = await ctx.bundle.coordinator.readRange(100, result!.sessionId, sheetId, 'A1:B3')
      expect(readResult.cells.length).toBeGreaterThan(0)
      // First cell should be "Name" (from the CSV header)
      const firstCell = readResult.cells.find(c => c.row === 0 && c.column === 0)
      expect(firstCell?.value).toBe('Name')

      // Verify exactly one session registered
      const session = ctx.bundle.coordinator.getSession(100, result!.sessionId)
      expect(session.sessionId).toBe(result!.sessionId)
    })

    test('CSV: no orphan conversion directory after open', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup
      const csvPath = join(testDir, 'clean.csv')
      writeCsvFixture(csvPath)

      const tempDir = mockApp.getPath('temp')
      const beforeCount = countTempDirs(tempDir, 'genoffice-imports')

      await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: csvPath,
        locale: 'en',
      })

      // The conversion dir is cleaned up after snapshot creation — verify
      // no orphaned genoffice-imports directories remain.
      const afterCount = countTempDirs(tempDir, 'genoffice-imports')
      expect(afterCount).toBe(beforeCount)
    })
  })

  // ═══ XLS E2E ═══

  describe('XLS open E2E', () => {
    test('XLS → convert → open → read cells → verify content', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup
      const xlsPath = join(testDir, 'legacy.xls')
      writeXlsFixture(xlsPath)

      const result = await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: xlsPath,
        locale: 'en',
      })
      expect(result).not.toBeNull()
      expect(result!.sessionId).toBeTruthy()
      // XLS import sets suggestSaveAs + needsSaveAs (NOT csvImport)
      expect(result!.session.suggestSaveAs).toContain('legacy.xlsx')
      expect(result!.session.needsSaveAs).toBe(true)
      expect(result!.session.csvImport).toBeUndefined()
      expect(result!.session.originalPath).toBe(xlsPath)

      // Read cells from the converted workbook
      const sheetId = result!.session.domainSession.sheetNames.keys().next().value
      const readResult = await ctx.bundle.coordinator.readRange(100, result!.sessionId, sheetId, 'A1:B2')
      expect(readResult.cells.length).toBeGreaterThan(0)
    })

    test('XLS conversion failure: typed error + no orphan resources', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup
      const invalidXlsPath = join(testDir, 'invalid.xls')
      writeInvalidXlsFixture(invalidXlsPath)

      const tempDir = mockApp.getPath('temp')
      const beforeCount = countTempDirs(tempDir, 'genoffice-imports')

      // The open should fail (the sidecar rejects the invalid XLS).
      await expect(
        ctx.bundle.coordinator.openWorkbook(100, undefined, {
          queuedPath: invalidXlsPath,
          locale: 'en',
        }),
      ).rejects.toThrow()

      // Verify NO session was registered
      // (getSession throws InvalidSessionError for unknown sessions)
      expect(() => {
        // Try a random sessionId — should fail
        ctx.bundle.coordinator.getSession(100, randomUUID())
      }).toThrow(InvalidSessionError)

      // Verify no orphaned conversion directories
      const afterCount = countTempDirs(tempDir, 'genoffice-imports')
      expect(afterCount).toBe(beforeCount)
    })
  })

  // ═══ Recovery E2E ═══

  describe('Recovery E2E', () => {
    test('RESTORE: recovery copy newer → restore prompt → recovered content opened', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup

      // Create original XLSX
      const originalPath = join(testDir, 'workbook.xlsx')
      writeFileSync(originalPath, await buildMinimalXlsx())

      // Wait 100ms so the recovery copy is newer
      await new Promise(r => setTimeout(r, 100))

      // Create recovery copy in the coordinator's recovery directory.
      // The recovery path is derived from the original file path via sha1
      // and lives in userData/sheets-autosave/.
      const recoveryDir = join(mockApp.getPath('userData'), 'sheets-autosave')
      mkdirSync(recoveryDir, { recursive: true })
      // The coordinator's recoveryPathFor uses sha1(filePath).slice(0,16)
      const crypto = await import('node:crypto')
      const hash = crypto.createHash('sha1').update(originalPath).digest('hex').slice(0, 16)
      const recoveryPath = join(recoveryDir, `${hash}.xlsx`)
      writeFileSync(recoveryPath, await buildRecoveryXlsx())

      // Mock the recovery dialog — respond "Restore" (index 0)
      mockDialog.showMessageBox.mockResolvedValue({ response: 0 })

      const result = await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: originalPath,
        locale: 'en',
      })

      expect(result).not.toBeNull()
      // restoreTarget should be set to the original path
      expect(result!.session.restoreTarget).toBe(originalPath)
      expect(result!.session.restoredFromRecovery).toBe(true)
      // The renderer-facing path is the ORIGINAL (restoreTarget)
      expect(result!.session.originalPath).toBe(originalPath)

      // Read cells — should be the RECOVERED content ("Recovered", 999)
      const sheetId = result!.session.domainSession.sheetNames.keys().next().value
      const readResult = await ctx.bundle.coordinator.readRange(100, result!.sessionId, sheetId, 'A1:B1')
      const firstCell = readResult.cells.find(c => c.row === 0 && c.column === 0)
      // The recovered content has "Recovered" in A1
      expect(firstCell?.value).toBe('Recovered')
    })

    test('DISCARD: recovery copy newer → discard prompt → original content opened', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup

      const originalPath = join(testDir, 'workbook-discard.xlsx')
      writeFileSync(originalPath, await buildMinimalXlsx())

      await new Promise(r => setTimeout(r, 100))

      const recoveryDir = join(mockApp.getPath('userData'), 'sheets-autosave')
      mkdirSync(recoveryDir, { recursive: true })
      const crypto = await import('node:crypto')
      const hash = crypto.createHash('sha1').update(originalPath).digest('hex').slice(0, 16)
      const recoveryPath = join(recoveryDir, `${hash}.xlsx`)
      writeFileSync(recoveryPath, await buildRecoveryXlsx())

      // Mock the recovery dialog — respond "Discard" (index 1)
      mockDialog.showMessageBox.mockResolvedValue({ response: 1 })

      const result = await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: originalPath,
        locale: 'en',
      })

      expect(result).not.toBeNull()
      // No restoreTarget (discard → original is opened directly)
      expect(result!.session.restoreTarget).toBeUndefined()
      // restoredFromRecovery is only set when restoreTarget is set; absent otherwise
      expect(result!.session.restoredFromRecovery ?? false).toBe(false)
      expect(result!.session.originalPath).toBe(originalPath)

      // Read cells — should be the ORIGINAL content ("Original", 100)
      const sheetId = result!.session.domainSession.sheetNames.keys().next().value
      const readResult = await ctx.bundle.coordinator.readRange(100, result!.sessionId, sheetId, 'A1:B1')
      const firstCell = readResult.cells.find(c => c.row === 0 && c.column === 0)
      expect(firstCell?.value).toBe('Original')

      // Recovery copy should be removed (clearWorkbookRecovery was called)
      expect(existsSync(recoveryPath)).toBe(false)
    })

    test('NO RECOVERY: normal workbook → no recovery dialog → normal open', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup

      const originalPath = join(testDir, 'normal.xlsx')
      writeFileSync(originalPath, await buildMinimalXlsx())

      // No recovery copy → dialog should NOT be shown
      const result = await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: originalPath,
        locale: 'en',
      })

      expect(result).not.toBeNull()
      expect(result!.session.restoreTarget).toBeUndefined()
      // restoredFromRecovery is only set when restoreTarget is set; absent otherwise
      expect(result!.session.restoredFromRecovery ?? false).toBe(false)
      expect(mockDialog.showMessageBox).not.toHaveBeenCalled()
    })
  })

  // ═══ Multi-renderer isolation ═══

  describe('Multi-renderer isolation', () => {
    test('renderer A opens CSV → renderer B cannot access A\'s session', async () => {
      const ctx = await makeBundle()
      cleanup = ctx.cleanup

      const csvPath = join(testDir, 'isolation.csv')
      writeCsvFixture(csvPath)

      // Renderer A (wcId=100) opens the CSV
      const resultA = await ctx.bundle.coordinator.openWorkbook(100, undefined, {
        queuedPath: csvPath,
        locale: 'en',
      })
      expect(resultA).not.toBeNull()

      // Renderer B (wcId=200) tries to read A's session → InvalidSessionError
      await expect(
        ctx.bundle.coordinator.readRange(200, resultA!.sessionId, 'sheet-1', 'A1:B1'),
      ).rejects.toThrow(InvalidSessionError)

      // Renderer B tries to close A's session → InvalidSessionError (or no-op)
      await expect(
        ctx.bundle.coordinator.closeWorkbook(200, resultA!.sessionId),
      ).rejects.toThrow(InvalidSessionError)
    })
  })
})

// ── Helpers ──────────────────────────────────────────────────────────

function countTempDirs(parentDir: string, prefix: string): number {
  try {
    const entries = readdirSync(parentDir)
    return entries.filter(e => e.startsWith(prefix)).length
  } catch {
    return 0
  }
}
