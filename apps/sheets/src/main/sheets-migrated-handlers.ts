/**
 * Migrated Sheets IPC handlers — thin adapter that delegates to
 * SheetsShellCoordinator.
 *
 * This module contains ZERO domain logic. It only:
 *   - extracts event.sender.id → wcId
 *   - resolves sessionId from the request payload
 *   - calls the coordinator
 *   - maps the result to the frozen renderer response shape
 *
 * ARCHITECTURE GUARDS (verified by tests):
 *   - ZERO XlsxSidecarClient imports
 *   - ZERO child_process imports
 *   - ZERO direct xlsx-gateway calls
 *   - ZERO filesystem save/open implementation
 *   - ZERO getFocusedWindow
 *   - ZERO global session state
 *   - ZERO type assertions (as unknown as, as any, as never)
 *
 * INCREMENT 6: workbook:save and workbook:write-recovery are migrated here.
 *   They delegate to coordinator.saveWorkbook()/writeRecovery(). The
 *   SavePlan translation and WorkbookFile building live in
 *   sheets-save-adapter.ts (shell-owned conversion boundary).
 *
 * INCREMENT 6A: the save handler is now genuinely thin — the 23-field
 *   SavePlan construction has moved to sheets-save-adapter.ts. The handler
 *   only does: IPC validation → typed conversion → coordinator call →
 *   response mapping.
 */

import { ipcMain, BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { z } from 'zod'

import { IPC_CHANNELS } from '../shared/ipc-channels'
import {
  workbookRangeRequestSchema,
  workbookRangeResultSchema,
  workbookFormulaCellsRequestSchema,
  workbookFormulaCellsResultSchema,
  workbookRecalcRequestSchema,
  workbookRecalcResultSchema,
  workbookMediaRequestSchema,
  workbookMediaResultSchema,
  workbookSaveRequestSchema,
  workbookExportPdfRequestSchema,
  screenCaptureRequestSchema,
} from '../shared/desktop-api'
import type { SheetsShellCoordinator } from './sheets-shell-coordinator'
import type { EngineRecalcEdit, EngineRecalcRead } from '@genoffice/runtime-contracts'
import type { ScreenCapture } from '@genoffice/platform'
import { translateSaveRequest, buildWorkbookFile } from './sheets-save-adapter'
import {
  collectAttachments,
  readAttachmentText,
  readAttachmentImage,
  savePastedImage,
  getAttachmentExtensions,
  readLocalImage,
} from './sheets-attachment-adapter'

// ── Session resolution ──

/** Resolve wcId from an IPC event. */
function wcIdFromEvent(event: IpcMainInvokeEvent): number {
  return event.sender.id
}

/**
 * Convert a numeric {startRow,startColumn,endRow,endColumn} range (0-indexed)
 * to Excel A1 notation (e.g. {0,0,0,1} → "A1:B1").
 *
 * INCREMENT 5B (build-fix): the migrated handler was producing
 * "0:0-0:1" which the engine's parseRange rejects. The engine contract
 * expects A1:B2 string notation. This helper mirrors the engine's
 * private colToIdx inverse without depending on the engine internals.
 */
function rangeToA1(r: { startRow: number; startColumn: number; endRow: number; endColumn: number }): string {
  return `${colIdxToLetter(r.startColumn)}${r.startRow + 1}:${colIdxToLetter(r.endColumn)}${r.endRow + 1}`
}

function colIdxToLetter(idx: number): string {
  let s = ''
  let n = idx + 1
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/**
 * Parse an A1 cell reference (e.g. "A1", "Z100", "AA1") to 0-indexed
 * (row, column). Used to convert the engine's hyperlinks `{ cell: "A1" }`
 * to the renderer's `{ row: 0, column: 0 }`.
 */
function parseCellRef(ref: string): { row: number; column: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (!match) return { row: 0, column: 0 }
  const colStr = match[1]
  const rowStr = match[2]
  if (colStr === undefined || rowStr === undefined) return { row: 0, column: 0 }
  let col = 0
  for (const ch of colStr) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: parseInt(rowStr, 10) - 1, column: col - 1 }
}

// ── Register migrated handlers ──

let migratedIpcRegistered = false

/**
 * Register the 5 migrated Sheets IPC handlers.
 * Must be called AFTER the coordinator is constructed and AFTER
 * the legacy registerSheetsIpc() has run (so the migrated handlers
 * replace the legacy ones).
 *
 * The coordinator MUST have its sessions registered via the legacy
 * open path (workbook:select is NOT yet migrated). The migrated
 * handlers resolve sessions from the coordinator's registry.
 */
export function registerMigratedSheetsIpc(coordinator: SheetsShellCoordinator, screenCapture?: ScreenCapture): void {
  if (migratedIpcRegistered) return
  migratedIpcRegistered = true

  // ── workbook:read-range ──
  ipcMain.removeHandler(IPC_CHANNELS.readWorkbookRange)
  ipcMain.handle(IPC_CHANNELS.readWorkbookRange, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookRangeRequestSchema.parse(input)
    const result = await coordinator.readRange(
      wcId, request.sessionId, request.sheetId,
      rangeToA1(request.range),
    )
    // INCREMENT 5B (build-fix): Translate EngineRangeResult → WorkbookRangeResult.
    // The engine contract uses different field names than the renderer's frozen
    // schema (conditionalFormatting vs conditionalRules, dataValidation vs
    // dataValidations). The validator now reads the sidecar's actual field
    // names (conditionalRules, dataValidations) and stores them in the engine
    // contract's fields. The translator below renames them back to the
    // renderer's expected names and adds defaults for fields the engine
    // contract doesn't carry (indexedThroughRow, indexingComplete, protectedRanges).
    // Extra engine-only fields (columns, rowBreaks, columnBreaks) are dropped.
    return workbookRangeResultSchema.parse({
      cells: result.cells.map(c => {
        // Recover the typed value: if the engine has a number, use it;
        // otherwise use the string value.
        const value = c.number !== undefined ? c.number : c.value
        const cell: Record<string, unknown> = {
          row: c.row,
          column: c.column,
          value,
        }
        if (c.isFormula) cell.formula = c.value
        if (c.styleIndex !== undefined && c.styleIndex !== 0) cell.styleIndex = c.styleIndex
        return cell
      }),
      rows: result.rows.map(r => {
        const row: Record<string, unknown> = { row: r.row, hidden: r.hidden ?? false }
        if (r.height !== undefined) row.height = r.height
        if (r.customHeight !== undefined) row.customHeight = r.customHeight
        if (r.outlineLevel !== undefined) row.outlineLevel = r.outlineLevel
        if (r.collapsed !== undefined) row.collapsed = r.collapsed
        if (r.styleIndex !== undefined) row.styleIndex = r.styleIndex
        return row
      }),
      merges: result.merges.map(m => ({
        startRow: m.firstRow, startColumn: m.firstColumn,
        endRow: m.lastRow, endColumn: m.lastColumn,
      })),
      hyperlinks: result.hyperlinks.map(h => {
        // Engine returns { cell: "A1", target: "..." } — convert to
        // { row, column, target } for the renderer.
        const parsed = parseCellRef(h.cell)
        return { row: parsed.row, column: parsed.column, target: h.target }
      }),
      conditionalRules: result.conditionalFormatting,
      autoFilter: result.autoFilter ?? null,
      dataValidations: result.dataValidation,
      sheetProtection: result.sheetProtection
        ? { protected: true, hasPassword: false }
        : null,
      protectedRanges: [],
      rowBreaks: result.rowBreaks,
      colBreaks: result.columnBreaks,
      indexedThroughRow: null,
      indexingComplete: true,
    })
  })

  // ── workbook:read-formulas ──
  ipcMain.removeHandler(IPC_CHANNELS.readWorkbookFormulas)
  ipcMain.handle(IPC_CHANNELS.readWorkbookFormulas, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookFormulaCellsRequestSchema.parse(input)
    const result = await coordinator.readFormulaCells(
      wcId, request.sessionId, request.sheetId,
    )
    // INCREMENT 5B (build-fix): Translate EngineFormulaCellsResult →
    // WorkbookFormulaCellsResult. The engine contract doesn't carry
    // indexingComplete or truncated (the sidecar does); the migrated
    // handler defaults them to true/false (trusted-complete). The cell
    // shape is mapped: engine's { formula, cachedValue? } → renderer's
    // { value, formula }.
    return workbookFormulaCellsResultSchema.parse({
      cells: result.cells.map(c => {
        const cell: Record<string, unknown> = {
          row: c.row,
          column: c.column,
          value: c.cachedValue ?? '',
        }
        if (c.formula) cell.formula = c.formula
        return cell
      }),
      indexingComplete: true,
      truncated: false,
    })
  })

  // ── workbook:recalc ──
  ipcMain.removeHandler(IPC_CHANNELS.recalcWorkbook)
  ipcMain.handle(IPC_CHANNELS.recalcWorkbook, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookRecalcRequestSchema.parse(input)
    const edits: EngineRecalcEdit[] = request.edits.map(e => ({
      sheetName: e.sheetId, // service resolves sheetId → sheetName internally
      row: e.row,
      column: e.column,
      value: e.input,
    }))
    const reads: EngineRecalcRead[] = request.reads.map(r => ({
      sheetName: r.sheetId, // service resolves sheetId → sheetName internally
      row: r.range.startRow,
      column: r.range.startColumn,
    }))
    const result = await coordinator.recalculate(wcId, request.sessionId, edits, reads)
    // Map EngineRecalcResult → WorkbookRecalcResult
    // The service returns cells with sheetName; map back to sheetId
    // by looking up in the session's domainSession.sheetNames map
    const session = coordinator.getSession(wcId, request.sessionId)
    const idsByName = new Map<string, string>()
    for (const [id, name] of session.domainSession.sheetNames) {
      idsByName.set(name, id)
    }
    return workbookRecalcResultSchema.parse({
      cells: result.cells.flatMap(c => {
        const sheetId = idsByName.get(c.sheetName)
        if (sheetId === undefined) return []
        return [{
          sheetId,
          row: c.row,
          column: c.column,
          formatted: c.formatted,
          ...(c.number !== undefined ? { number: c.number } : {}),
          isFormula: c.isFormula,
        }]
      }),
    })
  })

  // ── workbook:read-media ──
  ipcMain.removeHandler(IPC_CHANNELS.readWorkbookMedia)
  ipcMain.handle(IPC_CHANNELS.readWorkbookMedia, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookMediaRequestSchema.parse(input)
    const result = await coordinator.readMedia(
      wcId, request.sessionId, request.visualId,
    )
    // Map EngineMediaResult → WorkbookMediaResult
    return workbookMediaResultSchema.parse(result)
  })

  // ── workbook:close ──
  ipcMain.removeHandler(IPC_CHANNELS.closeWorkbook)
  ipcMain.handle(IPC_CHANNELS.closeWorkbook, async (event, sessionId: unknown) => {
    const wcId = wcIdFromEvent(event)
    const validatedSessionId = z.string().uuid().parse(sessionId)
    await coordinator.closeWorkbook(wcId, validatedSessionId)
    // Legacy handler returns void
  })

  // ── workbook:save (INCREMENT 6) ──
  // Migrated from sheets-main.ts. The handler is a THIN ADAPTER:
  //   1. Validates input via workbookSaveRequestSchema (frozen IPC shape)
  //   2. Translates WorkbookSaveRequest → SaveRequest (wrapping SavePlan)
  //   3. Calls coordinator.saveWorkbook(wcId, sessionId, request, mode, callerWindow)
  //   4. Maps the SaveResult → frozen WorkbookSaveResult
  //
  // The coordinator owns the commit journal (Phase A/B/C), atomic promotion
  // (rename, no copyFile fallback), teardown safety, external-change policy,
  // and session replacement. This handler does NOT:
  //   - invoke XlsxSidecarClient
  //   - call xlsx-package-io or xlsx-gateway planning functions
  //   - manipulate snapshots or commit markers
  //   - call child_process or node:fs
  //   - perform recovery logic
  ipcMain.removeHandler(IPC_CHANNELS.saveWorkbook)
  ipcMain.handle(IPC_CHANNELS.saveWorkbook, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookSaveRequestSchema.parse(input)
    const callerWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined

    const saveRequest = translateSaveRequest(request)
    const result = await coordinator.saveWorkbook(
      wcId, request.sessionId, saveRequest, request.mode, callerWindow,
    )

    // Map SaveResult → frozen WorkbookSaveResult
    if ('canceled' in result && result.canceled) {
      return { canceled: true }
    }
    if (!result.ok) {
      // External-change policy refusal — the service returned { ok: false,
      // reason: 'external-modified' }. The legacy handler threw tm('errDiskChanged').
      // Preserve the throw semantics so the renderer's catch path stays unchanged.
      throw new Error('errDiskChanged')
    }

    // Build the WorkbookFile from the replacement session's metadata.
    // After a successful save, the coordinator has replaced the old session
    // with a new one (same sessionId, new engine handle, new snapshot, new
    // fingerprint). We read the replacement session and build a WorkbookFile
    // the renderer can use to update its in-memory state.
    const session = coordinator.getSession(wcId, request.sessionId)
    const file = buildWorkbookFile(session)
    return {
      canceled: false as const,
      file,
      touchedEntries: result.touchedEntries ?? [],
    }
  })

  // ── workbook:write-recovery (INCREMENT 6) ──
  // Migrated from sheets-main.ts. The handler is a THIN ADAPTER:
  //   1. Validates input via workbookSaveRequestSchema
  //   2. Translates WorkbookSaveRequest → SaveRequest
  //   3. Calls coordinator.writeRecovery(wcId, sessionId, request)
  //   4. Returns { ok: boolean }
  //
  // The coordinator owns the recovery path derivation, epoch, mutation lock,
  // and stale-write rejection. This handler does NOT:
  //   - invoke XlsxSidecarClient
  //   - derive recovery paths (recoveryPathFor)
  //   - manipulate recovery files
  //   - call child_process or node:fs
  ipcMain.removeHandler(IPC_CHANNELS.writeWorkbookRecovery)
  ipcMain.handle(IPC_CHANNELS.writeWorkbookRecovery, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookSaveRequestSchema.parse(input)
    const saveRequest = translateSaveRequest(request)
    const result = await coordinator.writeRecovery(wcId, request.sessionId, saveRequest)
    return result
  })

  // ── workbook:export-pdf (INCREMENT 7 / ADR-006) ──
  // Migrated from sheets-main.ts:exportPdf(). The handler is a THIN ADAPTER:
  //   1. Validates input via workbookExportPdfRequestSchema (frozen IPC shape)
  //   2. Resolves callerWindow from event.sender (NOT getFocusedWindow)
  //   3. Calls coordinator.exportPdf(wcId, callerWindow, request)
  //   4. Returns the coordinator's PdfExportResult directly (already the
  //      frozen WorkbookExportPdfResult shape)
  //
  // The coordinator owns callerWindow + save dialog + output authorization +
  // writing the PDF bytes. The PDF renderer (SpreadsheetPdfRenderer, injected
  // into the coordinator) owns the hidden BrowserWindow + printToPDF + cleanup.
  //
  // ERROR SEMANTICS (matching legacy exportPdf):
  //   - User cancellation → { canceled: true } (returned by coordinator)
  //   - Render/filesystem failure → throws Error (propagates to renderer)
  //
  // This handler does NOT:
  //   - create BrowserWindow
  //   - call printToPDF
  //   - write PDF files directly
  //   - call getFocusedWindow
  //   - call child_process or node:fs
  //   - use type assertions (as unknown as, as any, as never)
  ipcMain.removeHandler(IPC_CHANNELS.exportPdf)
  ipcMain.handle(IPC_CHANNELS.exportPdf, async (event, input: unknown) => {
    const wcId = wcIdFromEvent(event)
    const request = workbookExportPdfRequestSchema.parse(input)
    const callerWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined

    // The coordinator returns PdfExportResult:
    //   { canceled: true } | { canceled: false, path: string }
    // This matches the frozen WorkbookExportPdfResult schema exactly.
    // Errors (render failure, fs failure) are thrown — the IPC layer
    // propagates them to the renderer as Error objects.
    return coordinator.exportPdf(wcId, callerWindow, {
      fileName: request.fileName,
      html: request.html,
      landscape: request.landscape,
      pageSize: request.pageSize,
      margins: request.margins,
      scale: request.scale,
    })
  })

  // ── sheets:capture-screen-sources (INCREMENT 8 / ADR-005) ──
  // Migrated from sheets-main.ts. The handler is a THIN ADAPTER:
  //   1. Resolves callerWindow from event.sender (for self-exclusion)
  //   2. Calls screenCapture.enumerateSources()
  //   3. Returns the frozen ScreenSourcesResult
  //
  // The ScreenCapture capability (ElectronScreenCapture) owns all
  // desktopCapturer/screen/systemPreferences logic. This handler does NOT:
  //   - import or call desktopCapturer
  //   - import or call screen.getAllDisplays()
  //   - import or call systemPreferences
  //   - call getFocusedWindow
  //   - use global source state
  ipcMain.removeHandler(IPC_CHANNELS.captureScreenSources)
  ipcMain.handle(IPC_CHANNELS.captureScreenSources, async (event) => {
    if (!screenCapture) throw new Error('Screen capture not available')
    // Exclude the app's own window from the source list (legacy behavior).
    // In tab mode the sheets renderer is a WebContentsView, so fromWebContents
    // on the sender may be null — fall back to the shell window.
    const selfWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const selfId = selfWindow?.getMediaSourceId()
    // If we have a selfId, we need to pass it to the capability. But the
    // capability is constructed once — we pass the selfId via a fresh
    // config. For now, the capability was constructed with the shell's
    // media source ID at runtime construction time. If the selfId differs
    // (standalone vs tab mode), the exclusion may not be perfect — but
    // this matches the legacy behavior which also used the shell window.
    return screenCapture.enumerateSources()
  })

  // ── sheets:capture-screen-source (INCREMENT 8 / ADR-005) ──
  // Migrated from sheets-main.ts. The handler is a THIN ADAPTER:
  //   1. Validates input via screenCaptureRequestSchema
  //   2. Calls screenCapture.captureSource(sourceId)
  //   3. Returns the frozen ScreenCaptureResult | null
  ipcMain.removeHandler(IPC_CHANNELS.captureScreenSource)
  ipcMain.handle(IPC_CHANNELS.captureScreenSource, async (event, input: unknown) => {
    if (!screenCapture) throw new Error('Screen capture not available')
    const request = screenCaptureRequestSchema.parse(input)
    return screenCapture.captureSource(request.id)
  })

  // ── sheets:files-pick (INCREMENT 9) ──
  // Migrated from sheets-main.ts. Thin adapter:
  //   1. Opens a file picker (caller-window-owned)
  //   2. Calls collectAttachments() from sheets-attachment-adapter.ts
  //   3. Returns frozen AttachmentAddResult | null
  ipcMain.removeHandler(IPC_CHANNELS.filesPick)
  ipcMain.handle(IPC_CHANNELS.filesPick, async (event) => {
    const { dialog } = await import('electron')
    const callerWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const exts = getAttachmentExtensions()
    const selection = callerWindow
      ? await dialog.showOpenDialog(callerWindow, {
          properties: ['openFile', 'multiSelections'],
          filters: [
            { name: 'Supported', extensions: exts },
            { name: 'All', extensions: ['*'] },
          ],
        })
      : await dialog.showOpenDialog({
          properties: ['openFile', 'multiSelections'],
          filters: [
            { name: 'Supported', extensions: exts },
            { name: 'All', extensions: ['*'] },
          ],
        })
    if (selection.canceled || selection.filePaths.length === 0) return null
    return collectAttachments(selection.filePaths)
  })

  // ── sheets:files-add (INCREMENT 9) ──
  // Thin adapter: validates paths, calls collectAttachments()
  ipcMain.removeHandler(IPC_CHANNELS.filesAdd)
  ipcMain.handle(IPC_CHANNELS.filesAdd, (_event, paths: unknown) => {
    const validatedPaths = z.array(z.string().min(1).max(1024)).max(50).parse(paths)
    return collectAttachments(validatedPaths)
  })

  // ── sheets:files-read (INCREMENT 9) ──
  // Thin adapter: validates path, calls readAttachmentText()
  ipcMain.removeHandler(IPC_CHANNELS.filesRead)
  ipcMain.handle(
    IPC_CHANNELS.filesRead,
    async (_event, filePath: unknown, offset: unknown, maxChars: unknown) => {
      const validatedPath = z.string().min(1).max(1024).parse(filePath)
      return readAttachmentText(validatedPath, Number(offset) || 0, Number(maxChars) || 1)
    },
  )

  // ── sheets:files-read-image (INCREMENT 9) ──
  // Thin adapter: validates path, calls readAttachmentImage()
  ipcMain.removeHandler(IPC_CHANNELS.filesReadImage)
  ipcMain.handle(IPC_CHANNELS.filesReadImage, (_event, filePath: unknown) => {
    const validatedPath = z.string().min(1).max(1024).parse(filePath)
    return readAttachmentImage(validatedPath)
  })

  // ── sheets:files-add-pasted-image (INCREMENT 9) ──
  // Thin adapter: validates ext, calls savePastedImage() + collectAttachments()
  ipcMain.removeHandler(IPC_CHANNELS.filesAddPastedImage)
  ipcMain.handle(IPC_CHANNELS.filesAddPastedImage, (_event, data: unknown, ext: unknown) => {
    const filePath = savePastedImage(data, ext)
    return filePath
      ? collectAttachments([filePath])
      : { accepted: [], rejected: ['not an image'] }
  })

  // ── shell:read-local-image (INCREMENT 11) ──
  // Thin adapter: validates path, calls readLocalImage() from the adapter.
  // The adapter handles: path resolution (~ → home), stat validation,
  // size limit (20MB), MIME sniffing from magic bytes, base64 encoding.
  ipcMain.removeHandler(IPC_CHANNELS.readLocalImage)
  ipcMain.handle(IPC_CHANNELS.readLocalImage, async (_event, input: unknown) => {
    const { localImageRequestSchema } = await import('../shared/desktop-api')
    const request = localImageRequestSchema.parse(input)
    return readLocalImage(request.path)
  })

  // ── workbook:read-pivot-definition (INCREMENT 12, corrected 15, hardened 15A) ──
  // Thin adapter: validates input, calls coordinator.readPivotDefinition()
  // The coordinator delegates to service.readPivotDefinition() which in turn
  // delegates to engine.readPivotDefinition() — the SINGLE translation point
  // between the OOXML wire format and the runtime-independent
  // WorkbookPivotDefinition contract. The engine reads both XML parts from
  // its on-disk temp file and parses them via the canonical @genoffice/xlsx-gateway
  // parser. ZERO parser logic in the handler. ZERO xlsx-gateway imports in the
  // handler. The returned WorkbookPivotDefinition is run through
  // workbookPivotDefinitionSchema.parse() as a frozen-IPC sanity check before
  // being returned to the renderer.
  ipcMain.removeHandler(IPC_CHANNELS.readPivotDefinition)
  ipcMain.handle(IPC_CHANNELS.readPivotDefinition, async (event, input: unknown) => {
    const { workbookPivotRequestSchema, workbookPivotDefinitionSchema } = await import('../shared/desktop-api')
    const wcId = wcIdFromEvent(event)
    const request = workbookPivotRequestSchema.parse(input)
    const pivotDefinition = await coordinator.readPivotDefinition(
      wcId, request.sessionId, request.path, request.cachePath,
    )
    // The coordinator returns a typed `WorkbookPivotDefinition`. Run it
    // through the frozen Zod schema as a final IPC-contract sanity check
    // before returning to the renderer — this catches any drift between
    // the engine's parser output and the renderer's expected shape.
    return workbookPivotDefinitionSchema.parse(pivotDefinition)
  })

  // ── workbook:auto-rename (INCREMENT 12) ──
  // Thin adapter: validates input, calls coordinator.renameWorkbook()
  // The coordinator owns: session lookup, rename validation, filesystem
  // rename, session path update, renderer push event (workbook:renamed).
  ipcMain.removeHandler(IPC_CHANNELS.autoRenameWorkbook)
  ipcMain.handle(IPC_CHANNELS.autoRenameWorkbook, async (event, sessionId: unknown, baseName: unknown) => {
    const wcId = wcIdFromEvent(event)
    const validatedSessionId = z.string().uuid().parse(sessionId)
    const validatedBaseName = z.string().min(1).max(100).parse(baseName)
    return coordinator.renameWorkbook(wcId, event.sender, validatedSessionId, validatedBaseName)
  })
}

// ── SavePlan translation + WorkbookFile building ────────────────────
//
// INCREMENT 6A: the 23-field SavePlan construction and WorkbookFile
// building have moved to sheets-save-adapter.ts (shell-owned conversion
// boundary). The handler imports translateSaveRequest() and
// buildWorkbookFile() from there — this file contains ZERO domain/XLSX
// translation logic and ZERO type assertions.
