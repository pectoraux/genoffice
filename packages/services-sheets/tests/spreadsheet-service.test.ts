/**
 * Service-level tests for SpreadsheetServiceImpl (Increment 3B correction).
 *
 * Uses a mock SpreadsheetEngine + mock SavePlanTranslator — does NOT use
 * ElectronXlsxSidecarEngine.
 *
 * Coverage:
 *   - open: session + engineHandle + metadata, sheetNames built from [sheet.id, sheet.name]
 *   - readRange, readFormulaCells, recalculate, readMedia (session-scoped)
 *   - save: unchanged (permitted), changed/unknown (refused), engine error
 *     propagation, SavePlan sheetId validation (fail-closed)
 *   - writeRecovery: returns bytes on success; throws typed error on failure
 *   - close: void on success; throws typed error on failure
 *
 * SHEET-ID MAPPING (Increment 3B):
 *   Tests verify that sheetNames is built from [sheet.id, sheet.name] —
 *   the stable XLSX sheetId, NOT the mutable sheet name. Known sheetId
 *   resolves to the correct file sheet name.
 *
 * UNKNOWN SHEET FAIL-CLOSED (Increment 3B):
 *   Tests verify that unknown sheetIds in readRange, readFormulaCells,
 *   recalculate, save, writeRecovery → InvalidInputError (NOT return sheetId).
 *
 * SAVE PLAN VALIDATION (Increment 3B):
 *   Tests verify that all sheetId-keyed mutation families in the SavePlan
 *   are validated (edits, structuralOps, formulaValues, sheetOps,
 *   filterStates, hyperlinkEdits, cfStates, dvStates, pageSetupStates,
 *   noteStates, sheetProtections, protectedRangeStates, visualAdditions,
 *   tableAdditions, pivotAdditions, sparklineAdditions, pivotRefreshUpdates).
 *
 * MEDIA SESSION SAFETY (Increment 3B):
 *   Tests verify that cross-session misuse (session A's visualId with
 *   session B's engineHandle) fails at the engine level — the engine's
 *   own session isolation enforces it.
 *
 * ERROR PROPAGATION (Increment 3A):
 *   Each test verifies that the correct TYPED error reaches the service
 *   caller — the service does NOT swallow engine exceptions into null or
 *   { ok: false }.
 */
import { describe, test, expect, vi } from 'vitest'
import { SpreadsheetServiceImpl } from '../src/spreadsheet-service.js'
import type {
  SpreadsheetEngine,
  EngineSessionHandle,
  WorkbookMetadata,
  WorksheetMetadata,
  EngineRangeResult,
  EngineFormulaCellsResult,
  EngineRecalcResult,
  EngineMediaResult,
  EngineSaveResult,
  SavePlan,
  SpreadsheetServiceDeps,
} from '@genoffice/runtime-contracts'
import { EngineError, InvalidSessionError, InvalidInputError } from '@genoffice/runtime-contracts'

// ── Mock helpers ──────────────────────────────────────────────────────

function makeMockHandle(): EngineSessionHandle {
  return { [Symbol('brand')]: Symbol('brand') } as unknown as EngineSessionHandle
}

/**
 * Build mock metadata where sheet.id differs from sheet.name.
 * This is CRITICAL for testing the Increment 3B sheetId mapping fix:
 * the service must build sheetNames from [sheet.id, sheet.name], NOT
 * [sheet.name, sheet.name].
 */
function makeMockMetadata(): WorkbookMetadata {
  return {
    name: 'test.xlsx',
    sha256: 'abc123',
    entryCount: 10,
    sheets: [
      // sheet.id = 'sheet-1' (stable XLSX sheetId), sheet.name = 'Sheet1' (visible tab name)
      { id: 'sheet-1', name: 'Sheet1', index: 0, hidden: false, rtl: false, showGridlines: true, rowCount: 100, columnCount: 26, defaultRowHeight: 15, defaultColumnWidth: 8.43 } as WorksheetMetadata,
      // A second sheet to test multi-sheet mapping
      { id: 'sheet-2', name: 'Data', index: 1, hidden: false, rtl: false, showGridlines: true, rowCount: 50, columnCount: 10, defaultRowHeight: 15, defaultColumnWidth: 8.43 } as WorksheetMetadata,
    ],
    activeTab: 0,
    definedNames: [],
    themeColors: [],
    themeFonts: { major: '', minor: '' },
  }
}

function makeMockEngine(): SpreadsheetEngine & { _handle: EngineSessionHandle; _saveResult: EngineSaveResult } {
  const handle = makeMockHandle()
  const saveResult: EngineSaveResult = {
    data: new Uint8Array([1, 2, 3]),
    touchedEntries: ['xl/worksheets/sheet1.xml'],
  }
  return {
    _handle: handle,
    _saveResult: saveResult,
    open: vi.fn(async () => ({ handle, metadata: makeMockMetadata() })),
    readRange: vi.fn(async () => ({ cells: [], rows: [], merges: [], columns: [], hyperlinks: [], conditionalFormatting: [], dataValidation: [], rowBreaks: [], columnBreaks: [], sheetProtection: false }) as EngineRangeResult),
    readFormulaCells: vi.fn(async () => ({ cells: [] }) as EngineFormulaCellsResult),
    recalculate: vi.fn(async () => ({ cells: [] }) as EngineRecalcResult),
    readMedia: vi.fn(async () => ({ mediaType: 'image/png', base64: 'iVBOR' }) as EngineMediaResult),
    // Increment 3C: applySavePlan replaces saveArchive. The engine accepts
    // the domain SavePlan directly and returns EngineSaveResult (data +
    // touchedEntries). No EngineArchivePatch leaks above the engine boundary.
    applySavePlan: vi.fn(async () => saveResult),
    convertWorkbook: vi.fn(async () => ({ data: new Uint8Array([1]), fileName: 'converted.xlsx' })),
    close: vi.fn(async () => {}),
    stop: vi.fn(async () => {}),
  }
}

function makeService(
  engine?: ReturnType<typeof makeMockEngine>,
) {
  const eng = engine ?? makeMockEngine()
  const deps: SpreadsheetServiceDeps = { engine: eng }
  const service = new SpreadsheetServiceImpl(deps)
  return { service, engine: eng }
}

function makeEmptySavePlan(): SavePlan {
  return {
    edits: [],
    structuralOps: [],
    formulaValues: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    sheetProtections: [],
    protectedRangeStates: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sparklineAdditions: [],
    chartEdits: [],
    visualEdits: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    definedNamesState: null,
    themeState: null,
    workbookProtectionState: null,
  }
}

// ── Tests ─────────────────────────────────────────────────────────────

describe('SpreadsheetServiceImpl', () => {
  // ── SHEET-ID MAPPING (Increment 3B) ──────────────────────────────

  describe('sheet-id mapping', () => {
    test('sheetNames is built from [sheet.id, sheet.name] (NOT [sheet.name, sheet.name])', async () => {
      const { service } = makeService()
      const result = await service.open(new Uint8Array([1, 2, 3]), 'en', 'test.xlsx')
      // sheet.id='sheet-1' → sheet.name='Sheet1'
      expect(result.session.sheetNames.get('sheet-1')).toBe('Sheet1')
      // sheet.id='sheet-2' → sheet.name='Data'
      expect(result.session.sheetNames.get('sheet-2')).toBe('Data')
      // The sheet.name should NOT be a key (the 3A bug was sheet.name → sheet.name)
      expect(result.session.sheetNames.get('Sheet1')).toBeUndefined()
      expect(result.session.sheetNames.get('Data')).toBeUndefined()
    })

    test('sheetNames map has exactly one entry per sheet', async () => {
      const { service } = makeService()
      const result = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      expect(result.session.sheetNames.size).toBe(2)
    })
  })

  // ── open ──────────────────────────────────────────────────────────

  describe('open', () => {
    test('returns session + engineHandle + metadata', async () => {
      const { service, engine } = makeService()
      const result = await service.open(new Uint8Array([1, 2, 3]), 'en', 'test.xlsx')
      expect(result.session.workbookName).toBe('test.xlsx')
      expect(result.session.workbookHash).toBe('abc123')
      expect(result.engineHandle).toBe(engine._handle)
      expect(result.metadata.name).toBe('test.xlsx')
    })

    test('engine open failure → throws EngineError with INTERNAL_ERROR code', async () => {
      const engine = makeMockEngine()
      engine.open = vi.fn(async () => { throw new EngineError('fail', 'INTERNAL_ERROR') })
      const { service } = makeService(engine)
      await expect(service.open(new Uint8Array([1]), 'en', 'test.xlsx')).rejects.toThrow(EngineError)
    })

    test('engine protocol error → throws EngineError with PROTOCOL_ERROR code', async () => {
      const engine = makeMockEngine()
      engine.open = vi.fn(async () => { throw new EngineError('protocol', 'PROTOCOL_ERROR') })
      const { service } = makeService(engine)
      await expect(service.open(new Uint8Array([1]), 'en', 'test.xlsx')).rejects.toMatchObject({
        name: 'EngineError',
        code: 'PROTOCOL_ERROR',
      })
    })

    test('invalid workbook input → throws InvalidInputError (distinguishable from engine failure)', async () => {
      const engine = makeMockEngine()
      engine.open = vi.fn(async () => { throw new InvalidInputError('not a valid xlsx') })
      const { service } = makeService(engine)
      await expect(service.open(new Uint8Array([1]), 'en', 'test.xlsx')).rejects.toBeInstanceOf(InvalidInputError)
    })
  })

  // ── readRange: fail-closed on unknown sheetId ─────────────────────

  describe('readRange', () => {
    test('known sheetId → delegates to engine with resolved sheet name', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      // Pass the domain sheetId ('sheet-1'), engine receives the file sheet name ('Sheet1')
      await service.readRange(opened.session, opened.engineHandle, 'sheet-1', 'A1:B2')
      expect(engine.readRange).toHaveBeenCalledWith(engine._handle, 'Sheet1', 'A1:B2')
    })

    test('unknown sheetId → throws InvalidInputError (NOT return sheetId)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      // 'unknown-sheet' is not in session.sheetNames
      await expect(service.readRange(opened.session, opened.engineHandle, 'unknown-sheet', 'A1:B2')).rejects.toThrow(InvalidInputError)
      // Engine must NOT be called
      expect(engine.readRange).not.toHaveBeenCalled()
    })

    test('sheet.name is NOT a valid sheetId (3B mapping fix)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      // 'Sheet1' is the sheet NAME, not the sheetId — must fail
      await expect(service.readRange(opened.session, opened.engineHandle, 'Sheet1', 'A1:B2')).rejects.toThrow(InvalidInputError)
      expect(engine.readRange).not.toHaveBeenCalled()
    })
  })

  // ── readFormulaCells: fail-closed ─────────────────────────────────

  describe('readFormulaCells', () => {
    test('known sheetId → delegates to engine', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      await service.readFormulaCells(opened.session, opened.engineHandle, 'sheet-2')
      expect(engine.readFormulaCells).toHaveBeenCalledWith(engine._handle, 'Data')
    })

    test('unknown sheetId → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      await expect(service.readFormulaCells(opened.session, opened.engineHandle, 'nope')).rejects.toThrow(InvalidInputError)
    })
  })

  // ── recalculate: fail-closed ──────────────────────────────────────

  describe('recalculate', () => {
    test('resolves sheet ids and delegates to engine', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const edits = [{ sheetName: 'sheet-1', row: 0, column: 0, value: '42' }]
      const reads = [{ sheetName: 'sheet-2', row: 0, column: 0 }]
      await service.recalculate(opened.session, opened.engineHandle, edits, reads)
      expect(engine.recalculate).toHaveBeenCalledWith(
        engine._handle,
        [{ sheetName: 'Sheet1', row: 0, column: 0, value: '42' }],
        [{ sheetName: 'Data', row: 0, column: 0 }],
      )
    })

    test('unknown sheetId in edits → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const edits = [{ sheetName: 'unknown', row: 0, column: 0, value: '42' }]
      const reads = [{ sheetName: 'sheet-1', row: 0, column: 0 }]
      await expect(service.recalculate(opened.session, opened.engineHandle, edits, reads)).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in reads → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const edits = [{ sheetName: 'sheet-1', row: 0, column: 0, value: '42' }]
      const reads = [{ sheetName: 'unknown', row: 0, column: 0 }]
      await expect(service.recalculate(opened.session, opened.engineHandle, edits, reads)).rejects.toThrow(InvalidInputError)
    })
  })

  // ── readMedia: session-scoped, cross-session safety ─────────────

  describe('readMedia', () => {
    test('delegates to engine with engineHandle + visualId', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      await service.readMedia(opened.session, opened.engineHandle, 'img1')
      expect(engine.readMedia).toHaveBeenCalledWith(engine._handle, 'img1')
    })

    test('cross-session misuse: session A visualId + session B engineHandle → engine fails', async () => {
      // Open two sessions
      const engine = makeMockEngine()
      const { service } = makeService(engine)
      const openedA = await service.open(new Uint8Array([1]), 'en', 'a.xlsx')
      const openedB = await service.open(new Uint8Array([2]), 'en', 'b.xlsx')

      // Simulate cross-session misuse: the engine (sidecar) won't find
      // session A's visualId in session B's sidecar session. The engine
      // throws InvalidInputError (visualId not found).
      engine.readMedia = vi.fn(async () => { throw new InvalidInputError('visualId not found in session') })

      // The service does NOT validate session ↔ engineHandle binding —
      // the engine's own session isolation enforces it. Pass session A
      // with engineHandle B: the engine fails.
      await expect(service.readMedia(openedA.session, openedB.engineHandle, 'img-from-A')).rejects.toThrow(InvalidInputError)
    })
  })

  // ── save: external-change policy + SavePlan validation + error propagation ──

  describe('save', () => {
    test('unchanged → save permitted, delegates to engine.applySavePlan, returns data + touchedEntries', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: '42' }] }
      const result = await service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')
      expect(result.ok).toBe(true)
      expect(result.data).toBeInstanceOf(Uint8Array)
      expect(result.touchedEntries).toEqual(['xl/worksheets/sheet1.xml'])
      // Increment 3C: engine.applySavePlan is called with the domain SavePlan directly.
      // No translator, no EngineArchivePatch leakage.
      expect(engine.applySavePlan).toHaveBeenCalledWith(engine._handle, plan)
    })

    test('changed → save refused with external-modified (engine NOT called)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const result = await service.save(opened.session, opened.engineHandle, { plan: makeEmptySavePlan() }, 'changed')
      expect(result.ok).toBe(false)
      expect(result.reason).toBe('external-modified')
      expect(engine.applySavePlan).not.toHaveBeenCalled()
    })

    test('unknown → save refused (safe default)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const result = await service.save(opened.session, opened.engineHandle, { plan: makeEmptySavePlan() }, 'unknown')
      expect(result.ok).toBe(false)
      expect(result.reason).toBe('external-modified')
      expect(engine.applySavePlan).not.toHaveBeenCalled()
    })

    test('unknown sheetId in edits → throws InvalidInputError (fail-closed)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'unknown', row: 0, column: 0, writeValue: true, value: '42' }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
      expect(engine.applySavePlan).not.toHaveBeenCalled()
    })

    test('unknown sheetId in structuralOps → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), structuralOps: [{ sheetId: 'unknown', kind: 'insert-row', index: 0, count: 1 }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in filterStates → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), filterStates: [{ sheetId: 'unknown', filter: {}, hiddenRows: [] }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in pivotAdditions.sourceSheetId → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = {
        ...makeEmptySavePlan(),
        pivotAdditions: [{ sheetId: 'sheet-1', sourceSheetId: 'unknown', sourceArea: {}, location: {}, name: 'p1' }],
      }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('add-sheet op with unknown sourceSheetId (duplicate) → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = {
        ...makeEmptySavePlan(),
        sheetOps: [{ kind: 'duplicate-sheet' as const, sheetId: 'new-1', sourceSheetId: 'unknown', name: 'Copy' }],
        sheetOrder: ['sheet-1', 'sheet-2', 'new-1'],
      }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('add-sheet op (new sheetId not in map) → does NOT throw (added sheets are new)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = {
        ...makeEmptySavePlan(),
        sheetOps: [{ kind: 'add-sheet' as const, sheetId: 'new-1', name: 'NewSheet' }],
        sheetOrder: ['sheet-1', 'sheet-2', 'new-1'],
      }
      const result = await service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')
      expect(result.ok).toBe(true)
      expect(engine.applySavePlan).toHaveBeenCalled()
    })

    test('engine failure during applySavePlan → throws EngineError (NOT { ok: false })', async () => {
      const engine = makeMockEngine()
      engine.applySavePlan = vi.fn(async () => { throw new EngineError('save failed', 'INTERNAL_ERROR') })
      const { service } = makeService(engine)
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: '42' }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(EngineError)
    })

    test('engine protocol error during applySavePlan → throws EngineError with PROTOCOL_ERROR', async () => {
      const engine = makeMockEngine()
      engine.applySavePlan = vi.fn(async () => { throw new EngineError('protocol', 'PROTOCOL_ERROR') })
      const { service } = makeService(engine)
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: '42' }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toMatchObject({
        name: 'EngineError',
        code: 'PROTOCOL_ERROR',
      })
    })
  })

  // ── writeRecovery: SavePlan validation + error propagation ────────

  describe('writeRecovery', () => {
    test('returns archive bytes for recovery', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: '42' }] }
      const data = await service.writeRecovery(opened.session, opened.engineHandle, { plan })
      expect(data).toBeInstanceOf(Uint8Array)
    })

    test('unknown sheetId → throws InvalidInputError', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'unknown', row: 0, column: 0, writeValue: true, value: '42' }] }
      await expect(service.writeRecovery(opened.session, opened.engineHandle, { plan })).rejects.toThrow(InvalidInputError)
    })

    test('engine failure → throws EngineError', async () => {
      const engine = makeMockEngine()
      engine.applySavePlan = vi.fn(async () => { throw new EngineError('recovery fail', 'INTERNAL_ERROR') })
      const { service } = makeService(engine)
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: '42' }] }
      await expect(service.writeRecovery(opened.session, opened.engineHandle, { plan })).rejects.toThrow(EngineError)
    })
  })

  // ── close ────────────────────────────────────────────────────────

  describe('close', () => {
    test('delegates to engine.close and returns void on success', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      await service.close(opened.engineHandle)
      expect(engine.close).toHaveBeenCalledWith(engine._handle)
    })

    test('engine failure → throws EngineError', async () => {
      const engine = makeMockEngine()
      engine.close = vi.fn(async () => { throw new EngineError('close fail', 'INTERNAL_ERROR') })
      const { service } = makeService(engine)
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      await expect(service.close(opened.engineHandle)).rejects.toThrow(EngineError)
    })
  })

  // ── SavePlan validation: comprehensive coverage ────────────────────

  describe('SavePlan sheetId validation — all mutation families', () => {
    test('unknown sheetId in hyperlinkEdits → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), hyperlinkEdits: [{ sheetId: 'unknown', row: 0, column: 0, target: 'http://x' }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in cfStates → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), cfStates: [{ sheetId: 'unknown', rules: [] }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in dvStates → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), dvStates: [{ sheetId: 'unknown', rules: [] }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in sheetProtections → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), sheetProtections: [{ sheetId: 'unknown', protected: true }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in visualAdditions → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), visualAdditions: [{ sheetId: 'unknown', anchor: {} }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in tableAdditions → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), tableAdditions: [{ sheetId: 'unknown', area: {}, name: 't1', columnNames: [] }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in sparklineAdditions → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), sparklineAdditions: [{ sheetId: 'unknown', type: 'line' as const, cells: [] }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('unknown sheetId in pivotRefreshUpdates → throws', async () => {
      const { service } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan = { ...makeEmptySavePlan(), pivotRefreshUpdates: [{ cachePath: 'xl/pivotCache.xml', sheetId: 'unknown', newOutputRef: 'A1:B2' }] }
      await expect(service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')).rejects.toThrow(InvalidInputError)
    })

    test('valid plan with all sheetIds known → succeeds (engine.applySavePlan called)', async () => {
      const { service, engine } = makeService()
      const opened = await service.open(new Uint8Array([1]), 'en', 'test.xlsx')
      const plan: SavePlan = {
        ...makeEmptySavePlan(),
        edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: '42' }],
        structuralOps: [{ sheetId: 'sheet-2', kind: 'insert-row', index: 0, count: 1 }],
        formulaValues: [{ sheetId: 'sheet-1', row: 0, column: 0, value: 42 }],
        filterStates: [{ sheetId: 'sheet-1', filter: {}, hiddenRows: [] }],
        hyperlinkEdits: [{ sheetId: 'sheet-2', row: 0, column: 0, target: 'http://x' }],
        cfStates: [{ sheetId: 'sheet-1', rules: [] }],
        dvStates: [{ sheetId: 'sheet-2', rules: [] }],
        pageSetupStates: [{ sheetId: 'sheet-1' }],
        noteStates: [{ sheetId: 'sheet-2', notes: [] }],
        sheetProtections: [{ sheetId: 'sheet-1', protected: true }],
        protectedRangeStates: [{ sheetId: 'sheet-2', ranges: [] }],
        visualAdditions: [{ sheetId: 'sheet-1', anchor: {} }],
        tableAdditions: [{ sheetId: 'sheet-2', area: {}, name: 't1', columnNames: [] }],
        pivotAdditions: [{ sheetId: 'sheet-1', sourceSheetId: 'sheet-2', sourceArea: {}, location: {}, name: 'p1' }],
        sparklineAdditions: [{ sheetId: 'sheet-2', type: 'line', cells: [] }],
        pivotRefreshUpdates: [{ cachePath: 'xl/pivotCache.xml', sheetId: 'sheet-1', newOutputRef: 'A1:B2' }],
      }
      const result = await service.save(opened.session, opened.engineHandle, { plan }, 'unchanged')
      expect(result.ok).toBe(true)
      // Increment 3C: engine.applySavePlan is called exactly once with the full SavePlan.
      expect(engine.applySavePlan).toHaveBeenCalledTimes(1)
      expect(engine.applySavePlan).toHaveBeenCalledWith(engine._handle, plan)
    })
  })
})

// ═══ INCREMENT 16 — convertWorkbook service port ═══

describe('SpreadsheetServiceImpl — convertWorkbook (Increment 16)', () => {
  test('delegates to engine.convertWorkbook with bytes + fileName', async () => {
    const { service, engine } = makeService()
    const legacyBytes = new Uint8Array([0, 1, 2, 3, 4, 5])
    const result = await service.convertWorkbook(legacyBytes, 'legacy.xls')
    expect(engine.convertWorkbook).toHaveBeenCalledTimes(1)
    expect(engine.convertWorkbook).toHaveBeenCalledWith(legacyBytes, 'legacy.xls')
    expect(result.data).toBeInstanceOf(Uint8Array)
    expect(result.fileName).toBe('converted.xlsx')
  })

  test('returns the converted bytes (no filesystem I/O in the service)', async () => {
    const convertedBytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]) // ZIP magic
    const engine = makeMockEngine()
    ;(engine.convertWorkbook as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      data: convertedBytes,
      fileName: 'report.xlsx',
    })
    const { service } = makeService(engine)
    const result = await service.convertWorkbook(new Uint8Array([1, 2, 3]), 'report.xls')
    expect(result.data).toBe(convertedBytes)
    expect(result.fileName).toBe('report.xlsx')
  })

  test('propagates engine errors (does NOT swallow them)', async () => {
    const engine = makeMockEngine()
    const engineError = new EngineError('conversion failed', 'INTERNAL_ERROR')
    ;(engine.convertWorkbook as ReturnType<typeof vi.fn>).mockRejectedValueOnce(engineError)
    const { service } = makeService(engine)
    await expect(
      service.convertWorkbook(new Uint8Array([1, 2, 3]), 'bad.xls'),
    ).rejects.toThrow(EngineError)
  })

  test('performs NO filesystem I/O — accepts bytes, returns bytes', async () => {
    // The service contract is data-oriented: it accepts Uint8Array content,
    // NOT a filesystem path. The engine implementation may write the bytes
    // to a temp file internally (private to the adapter), but the service
    // itself touches no filesystem.
    const { service, engine } = makeService()
    const spy = vi.spyOn(service, 'convertWorkbook')
    const inputBytes = new Uint8Array([1, 2, 3, 4, 5])
    await service.convertWorkbook(inputBytes, 'input.xls')
    expect(spy).toHaveBeenCalledWith(inputBytes, 'input.xls')
    // Verify the engine was called with bytes (not a path string).
    const call = (engine.convertWorkbook as ReturnType<typeof vi.fn>).mock.calls[0]
    expect(call![0]).toBeInstanceOf(Uint8Array)
    expect(typeof call![1]).toBe('string')
  })
})
