/**
 * Architecture-boundary test for @genoffice/sheets (Increment 3I/5/5A — AST-based).
 *
 * Uses the TypeScript compiler API (via dep-scanner.ts) to parse source files
 * into an AST and extract actual module specifiers. This GUARANTEES:
 *   - Comments and JSDoc are NOT reported
 *   - String literals that are NOT import specifiers are NOT reported
 *   - ALL import forms ARE detected
 *
 * Enforces:
 *   - ZERO imports of @genoffice/platform-electron EXCEPT in:
 *     sheets-runtime.ts (constructs ElectronXlsxSidecarEngine)
 *     sheets-migrated-handlers.ts (thin IPC adapter, no domain logic)
 *   - package.json declares @genoffice/platform-electron (runtime construction)
 *   - package.json declares @genoffice/xlsx-gateway (canonical planner)
 *   - DOES import @genoffice/xlsx-gateway
 *
 * INCREMENT 5A — Legacy session adoption guards:
 *   - sheets-shell-coordinator.ts MUST NOT import the legacy sidecar client
 *     (XlsxSidecarClient) or child_process — the coordinator is pure to
 *     SpreadsheetService.
 *   - sheets-runtime.ts MUST NOT import the legacy sidecar client — adoption
 *     goes through engine.adoptExternalSession (no legacy client dependency).
 *   - No file under src/main/ except xlsx-sidecar-client.ts may import
 *     child_process (the legacy client is the ONLY module that spawns the
 *     sidecar binary; the engine accepts an injectable SidecarProtocolLike).
 */
import { describe, test, expect } from 'vitest'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'
import { extractImports, listSourceFiles } from './dep-scanner.js'

const SRC = join(__dirname, '..', 'src')
const PACKAGE_JSON = join(__dirname, '..', 'package.json')

function scanForImports(rootDir: string, forbidden: string[]): Array<{ file: string; line: number; specifier: string; kind: string }> {
  const hits: Array<{ file: string; line: number; specifier: string; kind: string }> = []
  for (const file of listSourceFiles(rootDir)) {
    const imports = extractImports(file)
    for (const imp of imports) {
      for (const f of forbidden) {
        if (imp.specifier === f || imp.specifier.startsWith(f + '/')) {
          hits.push({ file, line: imp.line, specifier: imp.specifier, kind: imp.kind })
        }
      }
    }
  }
  return hits
}

describe('@genoffice/sheets architecture boundary (Increment 3I/5/5A — AST-based)', () => {
  test('ZERO imports of @genoffice/platform-electron in production source (except runtime/handlers)', () => {
    // Exception: sheets-runtime.ts and sheets-migrated-handlers.ts are
    // the runtime construction + thin IPC adapter — they are the ONLY
    // modules permitted to import from platform-electron, because they
    // construct the ElectronXlsxSidecarEngine. All other source files
    // must NOT import platform-electron.
    const hits = scanForImports(SRC, ['@genoffice/platform-electron'])
    const violations = hits.filter((h) =>
      !h.file.endsWith('sheets-runtime.ts') &&
      !h.file.endsWith('sheets-migrated-handlers.ts'),
    )
    if (violations.length > 0) {
      console.error('Found @genoffice/platform-electron imports in apps/sheets source:')
      for (const h of violations) {
        console.error(`  ${h.file}:${h.line}: ${h.kind} '${h.specifier}'`)
      }
    }
    expect(violations).toEqual([])
  })

  test('package.json declares @genoffice/platform-electron as dependency (runtime construction)', () => {
    const pkg = JSON.parse(readFileSync(PACKAGE_JSON, 'utf8'))
    const deps = pkg.dependencies ?? {}
    expect(deps).toHaveProperty('@genoffice/platform-electron')
  })

  test('package.json DOES declare @genoffice/xlsx-gateway as a dependency (canonical planner)', () => {
    const pkg = JSON.parse(readFileSync(PACKAGE_JSON, 'utf8'))
    const deps = pkg.dependencies ?? {}
    expect(deps).toHaveProperty('@genoffice/xlsx-gateway')
  })

  test('DOES import @genoffice/xlsx-gateway (the canonical planner)', () => {
    const hits = scanForImports(SRC, ['@genoffice/xlsx-gateway'])
    expect(hits.length).toBeGreaterThan(0)
  })

  // ═══ INCREMENT 5A — Legacy session adoption architecture guards ═══

  test('sheets-shell-coordinator.ts does NOT import the legacy sidecar client', () => {
    // The coordinator receives ONLY SpreadsheetService — it must not
    // depend on XlsxSidecarClient, child_process, or any sidecar-specific code.
    const hits = scanForImports(join(SRC, 'main'), ['./xlsx-sidecar-client'])
    const violations = hits.filter((h) => h.file.endsWith('sheets-shell-coordinator.ts'))
    expect(violations).toEqual([])
  })

  test('sheets-runtime.ts does NOT import the legacy sidecar client', () => {
    // sheets-runtime.ts constructs the engine via @genoffice/platform-electron.
    // It accepts an injectable SidecarProtocolLike (the legacy client is
    // passed in via ElectronXlsxSidecarEngineConfig.sidecarClient), but
    // does NOT directly import XlsxSidecarClient.
    const hits = scanForImports(join(SRC, 'main'), ['./xlsx-sidecar-client'])
    const violations = hits.filter((h) => h.file.endsWith('sheets-runtime.ts'))
    expect(violations).toEqual([])
  })

  test('ONLY xlsx-sidecar-client.ts imports child_process in src/main/', () => {
    // The legacy client is the ONLY module that spawns the sidecar binary.
    // The engine accepts an injectable SidecarProtocolLike — when the
    // legacy client is injected, no second child_process spawn happens.
    const hits = scanForImports(join(SRC, 'main'), ['node:child_process'])
    const violations = hits.filter((h) => !h.file.endsWith('xlsx-sidecar-client.ts'))
    if (violations.length > 0) {
      console.error('Found child_process imports outside xlsx-sidecar-client.ts:')
      for (const h of violations) {
        console.error(`  ${h.file}:${h.line}: ${h.kind}`)
      }
    }
    expect(violations).toEqual([])
  })

  test('no src/main/ file calls getFocusedWindow', () => {
    // getFocusedWindow is a UI concern — adoption must not depend on the
    // currently-focused window. The wcId is sourced from event.sender.id
    // at the IPC boundary, not from focus state.
    const mainSrc = join(SRC, 'main')
    for (const file of listSourceFiles(mainSrc)) {
      const src = readFileSync(file, 'utf8')
      // Allow references inside string literals or comments by checking
      // for the call pattern only.
      expect(src).not.toMatch(/getFocusedWindow\s*\(/)
    }
  })

  test('no src/main/ file creates a second engine handle for the same sidecar session', () => {
    // The ElectronXlsxSidecarEngine.adoptExternalSession method is the
    // ONLY way to wrap an existing sidecar sessionId into an opaque
    // EngineSessionHandle. No code should manually construct handles.
    const mainSrc = join(SRC, 'main')
    for (const file of listSourceFiles(mainSrc)) {
      const src = readFileSync(file, 'utf8')
      // The handle creation is via createHandle() — private to the engine.
      // No file in apps/sheets/src/main should reference createHandle.
      expect(src).not.toMatch(/\bcreateHandle\b/)
    }
  })

  test('no global session state — wcId is always a parameter', () => {
    // The coordinator's `tabs` Map is keyed by wcId — there is no
    // "currentWcId" or "activeSession" module-level state. Sessions are
    // always resolved by (wcId, sessionId) pair.
    const coordinatorSrc = readFileSync(
      join(SRC, 'main', 'sheets-shell-coordinator.ts'),
      'utf8',
    )
    expect(coordinatorSrc).not.toMatch(/^(let|var|const)\s+(currentWcId|activeSession|globalSession)\b/m)
  })

  // ═══ INCREMENT 6A — Save adapter architecture guards ═══

  test('sheets-migrated-handlers.ts has ZERO type assertions', () => {
    // The save handler must not use `as unknown as`, `as any`, `as never`.
    // The SavePlan translation and WorkbookFile building live in
    // sheets-save-adapter.ts. (Note: the read-range/read-formulas translators
    // use `Record<string, unknown>` for cell/row construction — those are
    // already validated via workbookRangeResultSchema.parse() before return,
    // so they're not unchecked. This test guards the SAVE path only.)
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    // Strip JSDoc/block comments AND line comments
    const stripped = src
      .replace(/\/\*\*?[\s\S]*?\*\//g, '') // block comments
      .replace(/^\s*\/\/.*$/gm, '')          // line comments
    expect(stripped).not.toMatch(/\bas\s+unknown\s+as\b/)
    expect(stripped).not.toMatch(/\bas\s+any\b/)
    expect(stripped).not.toMatch(/\bas\s+never\b/)
  })

  test('sheets-save-adapter.ts has ZERO type assertions', () => {
    // The save adapter must not use `as unknown as`, `as any`, `as never`.
    // All conversions use explicit per-family typed mappers with fresh
    // object literals (assignable to readonly interfaces).
    const src = readFileSync(join(SRC, 'main', 'sheets-save-adapter.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '')
    expect(stripped).not.toMatch(/\bas\s+unknown\s+as\b/)
    expect(stripped).not.toMatch(/\bas\s+any\b/)
    expect(stripped).not.toMatch(/\bas\s+never\b/)
  })

  test('sheets-save-adapter.ts does NOT import XlsxSidecarClient or child_process', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-save-adapter.ts'), 'utf8')
    expect(src).not.toMatch(/from\s+['"]\.\/xlsx-sidecar-client['"]/)
    expect(src).not.toMatch(/^import.*child_process/m)
  })

  test('sheets-save-adapter.ts does NOT import xlsx-gateway', () => {
    // The adapter is a pure data conversion module — no gateway planning.
    const src = readFileSync(join(SRC, 'main', 'sheets-save-adapter.ts'), 'utf8')
    expect(src).not.toMatch(/from\s+['"]@genoffice\/xlsx-gateway['"]/)
  })

  test('sheets-save-adapter.ts does NOT call filesystem APIs', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-save-adapter.ts'), 'utf8')
    expect(src).not.toMatch(/^import.*node:fs/m)
    expect(src).not.toMatch(/^import.*node:path/m)
  })

  test('sheets-save-adapter.ts uses workbookFileSchema.parse() for validation', () => {
    // The buildWorkbookFile function must validate the candidate via the
    // frozen Zod schema — no raw `unknown` return.
    const src = readFileSync(join(SRC, 'main', 'sheets-save-adapter.ts'), 'utf8')
    expect(src).toMatch(/workbookFileSchema\.parse\(/)
  })

  test('sheets-save-adapter.ts buildWorkbookFile returns WorkbookFile (not unknown)', () => {
    // The return type must be the frozen WorkbookFile — not unknown or
    // Record<string, unknown>.
    const src = readFileSync(join(SRC, 'main', 'sheets-save-adapter.ts'), 'utf8')
    expect(src).toMatch(/function buildWorkbookFile\([^)]*\):\s*WorkbookFile/)
  })

  test('sheets-migrated-handlers.ts imports translateSaveRequest + buildWorkbookFile from sheets-save-adapter', () => {
    // The handler must delegate to the adapter — not contain the logic inline.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/from\s+['"]\.\/sheets-save-adapter['"]/)
    expect(src).toMatch(/translateSaveRequest/)
    expect(src).toMatch(/buildWorkbookFile/)
  })

  // ═══ INCREMENT 7 — PDF export migration architecture guards ═══

  test('migrated export-pdf handler has ZERO BrowserWindow construction', () => {
    // The handler must NOT create BrowserWindow — the PDF renderer
    // (ElectronSpreadsheetPdfRenderer) owns the hidden window.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '')
    expect(stripped).not.toMatch(/new\s+BrowserWindow\b/)
  })

  test('migrated export-pdf handler has ZERO printToPDF calls', () => {
    // The handler must NOT call printToPDF — that's the renderer's job.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    // Strip JSDoc/block comments AND line comments
    const stripped = src
      .replace(/\/\*\*?[\s\S]*?\*\//g, '') // block comments
      .replace(/^\s*\/\/.*$/gm, '')          // line comments
    expect(stripped).not.toMatch(/printToPDF/)
  })

  test('migrated export-pdf handler has ZERO direct filesystem PDF writes', () => {
    // The handler must NOT write PDF files directly — the coordinator
    // owns output-path authorization + writing.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '')
    expect(stripped).not.toMatch(/writeFile.*pdf/i)
    expect(stripped).not.toMatch(/^import.*node:fs/m)
  })

  test('migrated export-pdf handler has ZERO getFocusedWindow calls', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).not.toMatch(/getFocusedWindow\s*\(/)
  })

  test('migrated export-pdf handler delegates to coordinator.exportPdf', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/coordinator\.exportPdf\b/)
  })

  test('migrated export-pdf handler replaces the legacy handler', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.exportPdf\)/)
  })

  // ═══ INCREMENT 8 — Screen capture migration architecture guards ═══

  test('migrated screen-capture handlers have ZERO desktopCapturer imports', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/desktopCapturer/)
  })

  test('migrated screen-capture handlers have ZERO screen.getAllDisplays calls', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/screen\.getAllDisplays/)
  })

  test('migrated screen-capture handlers have ZERO getFocusedWindow calls', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).not.toMatch(/getFocusedWindow\s*\(/)
  })

  test('migrated screen-capture handlers delegate to ScreenCapture capability', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/screenCapture\.enumerateSources/)
    expect(src).toMatch(/screenCapture\.captureSource/)
  })

  test('migrated screen-capture handlers replace legacy handlers', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.captureScreenSources\)/)
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.captureScreenSource\)/)
  })

  test('ScreenCapture contract (platform) has ZERO Electron/node imports', () => {
    const src = readFileSync(
      join(__dirname, '..', '..', '..', 'packages', 'platform', 'src', 'capabilities', 'screen-capture.ts'),
      'utf8',
    )
    // Check import statements only (JSDoc may mention desktopCapturer for documentation)
    expect(src).not.toMatch(/from\s+['"]electron['"]/)
    expect(src).not.toMatch(/from\s+['"]node:/)
  })

  test('ElectronScreenCapture is the ONLY implementation with desktopCapturer', () => {
    const src = readFileSync(
      join(__dirname, '..', '..', '..', 'packages', 'platform-electron', 'src', 'capabilities', 'electron-screen-capture.ts'),
      'utf8',
    )
    expect(src).toMatch(/desktopCapturer/)
    expect(src).toMatch(/class ElectronScreenCapture/)
  })

  test('migrated screen-capture handlers have ZERO global capture state', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^(let|var|const)\s+(activeSource|activeDisplay|currentCapture|activeRenderer|globalCaptureSource)\b/m)
  })

  // ═══ INCREMENT 9 — Files/attachments migration architecture guards ═══

  test('migrated file handlers have ZERO parseFileToText imports', () => {
    // The parser is used by the attachment adapter, not the handler.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/parseFileToText/)
  })

  test('migrated file handlers have ZERO node:fs imports', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^import.*node:fs/m)
  })

  test('migrated file handlers delegate to sheets-attachment-adapter', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/from\s+['"]\.\/sheets-attachment-adapter['"]/)
    expect(src).toMatch(/collectAttachments/)
    expect(src).toMatch(/readAttachmentText/)
    expect(src).toMatch(/readAttachmentImage/)
    expect(src).toMatch(/savePastedImage/)
  })

  test('migrated file handlers replace legacy handlers', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.filesPick\)/)
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.filesAdd\)/)
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.filesRead\)/)
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.filesReadImage\)/)
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.filesAddPastedImage\)/)
  })

  test('migrated file handlers have ZERO getFocusedWindow calls', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).not.toMatch(/getFocusedWindow\s*\(/)
  })

  test('migrated file handlers have ZERO global attachment state', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^(let|var|const)\s+(activeAttachment|currentPicker|currentPath|activeFile)\b/m)
  })

  // ═══ INCREMENT 10 — AI migration architecture guards ═══

  test('migrated AI handlers are in a separate module (sheets-ai-handlers.ts)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    expect(src).toMatch(/registerMigratedSheetsAiIpc/)
    expect(src).toMatch(/abortStreamsForRenderer/)
  })

  test('migrated AI handlers replace legacy handlers via removeHandler', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    const channels = ['aiGetSettings', 'aiSetSettings', 'aiGskStatus', 'aiGskLogin', 'aiChat', 'aiStream', 'aiStreamCancel', 'aiGenerateImage']
    for (const ch of channels) {
      expect(src).toMatch(new RegExp(`removeHandler\\(IPC_CHANNELS\\.${ch}\\)`))
    }
    // Web search, image search, fetch image use raw channel strings
    expect(src).toMatch(/removeHandler\('ai:web-search'\)/)
    expect(src).toMatch(/removeHandler\('ai:image-search'\)/)
    expect(src).toMatch(/removeHandler\('ai:fetch-image'\)/)
  })

  test('migrated AI handlers have ZERO getFocusedWindow calls', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    expect(src).not.toMatch(/getFocusedWindow\s*\(/)
  })

  test('migrated AI stream tracking is renderer-scoped (Map<wcId, Map>)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    // Must use event.sender.id for renderer scoping
    expect(src).toMatch(/event\.sender\.id/)
    // Must track streams per-renderer
    expect(src).toMatch(/streamState\s*=\s*new\s+Map<number,\s*Map<string,\s*AbortController>>/)
  })

  test('migrated AI handlers have ZERO global stream state (no activeStream/currentStream)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^(let|var|const)\s+(activeStream|currentStream|currentRenderer|activeRequestId|currentChatId)\b/m)
  })

  test('migrated AI stream push routes to event.sender only', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    // Stream chunks are sent via event.sender.send, not via broadcast
    expect(src).toMatch(/sender\.send\(IPC_CHANNELS\.aiStreamChunk/)
    expect(src).toMatch(/sender\.isDestroyed\(\)/)
  })

  test('migrated AI handlers have ZERO BrowserWindow construction', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-ai-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/new\s+BrowserWindow\b/)
  })

  test('migrated main handlers import registerMigratedSheetsAiIpc', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    expect(src).toMatch(/from\s+['"]\.\/sheets-ai-handlers['"]/)
    expect(src).toMatch(/registerMigratedSheetsAiIpc/)
    expect(src).toMatch(/abortStreamsForRenderer/)
  })

  test('renderer teardown calls abortStreamsForRenderer', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    expect(src).toMatch(/abortStreamsForRenderer\(webContents\.id\)/)
  })

  // ═══ INCREMENT 12 — Pivot read + auto-rename architecture guards ═══

  test('migrated pivot handler delegates to coordinator.readPivotDefinition', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/coordinator\.readPivotDefinition/)
  })

  test('migrated pivot handler replaces legacy handler', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.readPivotDefinition\)/)
  })

  test('migrated pivot handler has ZERO XlsxSidecarClient imports', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/from\s+['"]\.\/xlsx-sidecar-client['"]/)
  })

  test('migrated pivot handler delegates to coordinator.readPivotDefinition', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/coordinator\.readPivotDefinition/)
    // The parser now lives in the service, not the handler
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/parsePivotDefinition/)
  })

  test('migrated pivot handler has ZERO xlsx-gateway imports (Increment 15A)', () => {
    // The handler must NOT import xlsx-gateway directly — neither statically
    // nor dynamically. The parser lives in the engine (platform-electron),
    // which is the single translation point between OOXML wire format and
    // the runtime-independent WorkbookPivotDefinition contract.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/@genoffice\/xlsx-gateway/)
    expect(stripped).not.toMatch(/import\s*\(\s*['"]@genoffice\/xlsx-gateway/)
  })

  test('migrated pivot handler has ZERO sidecarClient references (Increment 15A)', () => {
    // The handler must not reference the sidecar client in any form —
    // neither as a coordinator dep nor as a direct import.
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/sidecarClient/)
  })

  test('migrated pivot handler has ZERO filesystem implementation (Increment 15A)', () => {
    // The handler must not perform filesystem operations — the engine
    // owns the on-disk temp file (private to the adapter).
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^import.*node:fs/m)
    expect(stripped).not.toMatch(/readFileSync|writeFileSync|mkdirSync|rmSync|existsSync/)
  })

  test('migrated rename handler delegates to coordinator.renameWorkbook', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/coordinator\.renameWorkbook/)
  })

  test('migrated rename handler replaces legacy handler', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.autoRenameWorkbook\)/)
  })

  test('coordinator has readPivotDefinition method', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/async readPivotDefinition/)
  })

  test('coordinator readPivotDefinition returns WorkbookPivotDefinition (NOT unknown) — Increment 15A', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    // The coordinator's readPivotDefinition must declare a typed return —
    // `Promise<WorkbookPivotDefinition>`, not `Promise<unknown>`.
    expect(src).toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/m)
    expect(src).not.toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<unknown>/m)
  })

  test('coordinator has renameWorkbook method', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/async renameWorkbook/)
  })

  test('coordinator renameWorkbook sends push event to event.sender only', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/webContents\.send\('workbook:renamed'/)
    expect(src).toMatch(/webContents\.isDestroyed\(\)/)
  })

  test('coordinator renameWorkbook does NOT broadcast', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    // No BrowserWindow.getAllWindows() or broadcast pattern
    expect(stripped).not.toMatch(/getAllWindows\(\)/)
  })

  test('coordinator has renameWorkbookFromShell method (Increment 17)', () => {
    // Phase 2 Increment 17: the manual rename path (shell → sheetsFileRenamed)
    // delegates to `coordinator.renameWorkbookFromShell(wcId, oldPath, newPath)`.
    // The coordinator finds the session by `originalPath` and updates it.
    // There is NO legacy SessionInfo mirror — the coordinator is the SOLE owner.
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/renameWorkbookFromShell\s*\(/)
  })

  test('coordinator has NO onWorkbookRenamed dep callback (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/onWorkbookRenamed\??\s*:/)
    expect(stripped).not.toMatch(/this\.deps\.onWorkbookRenamed/)
  })

  test('coordinator has ZERO raw sidecar command construction (Increment 15A)', () => {
    // The coordinator must NOT construct `{ command: '...' }` sidecar
    // payloads directly — all sidecar wire-protocol construction lives
    // behind the engine boundary (ElectronXlsxSidecarEngine).
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/command:\s*['"]read_entries['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]read_range['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]open['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]close['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]save_archive['"]/)
  })

  test('coordinator has ZERO global caller state (Increment 15A)', () => {
    // The coordinator must not keep module-level mutable state keyed by
    // renderer id or session id — all session state lives inside the
    // `tabs` Map (per-renderer, lazily registered).
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^(let|var|const)\s+(currentWcId|activeSession|globalSession|currentRenderer)\b/m)
  })

  test('sheets-runtime has NO onWorkbookRenamed in SheetsCoordinatorConfig (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-runtime.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/onWorkbookRenamed/)
  })

  test('sheets-main sheetsFileRenamed delegates to coordinator.renameWorkbookFromShell (Increment 17)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    expect(src).toMatch(/coordinator\.renameWorkbookFromShell\b/)
  })

  test('sheets-main has NO updateLegacySessionPath (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/export function updateLegacySessionPath/)
    expect(stripped).not.toMatch(/updateLegacySessionPath\b/)
  })

  test('sheets-main has NO SessionInfo interface (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/interface SessionInfo\b/)
  })

  test('sheets-main has NO closeAllSessions function (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/async function closeAllSessions\b/)
    expect(stripped).not.toMatch(/closeAllSessions\b/)
  })

  test('sheets-main has NO XlsxSidecarClient import (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/from\s+['"]\.\/xlsx-sidecar-client['"]/)
    expect(stripped).not.toMatch(/new XlsxSidecarClient\b/)
  })

  test('sheets-main has NO registerSheetsAiIpc (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/export function registerSheetsAiIpc\b/)
  })

  test('sheets-main has NO writeWorkbookTo (Increment 17 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/async function writeWorkbookTo\b/)
  })

  // ═══ INCREMENT 16 — Legacy open cutover architecture guards ═══

  test('migrated selectWorkbook handler delegates to coordinator.openWorkbook (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/coordinator\.openWorkbook\b/)
  })

  test('migrated selectWorkbook handler replaces legacy handler (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    expect(src).toMatch(/removeHandler\(IPC_CHANNELS\.selectWorkbook\)/)
  })

  test('migrated selectWorkbook handler has ZERO XlsxSidecarClient imports (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/from\s+['"]\.\/xlsx-sidecar-client['"]/)
  })

  test('migrated selectWorkbook handler has ZERO filesystem implementation (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-migrated-handlers.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/readFileSync|writeFileSync|mkdirSync|rmSync|existsSync|copyFileSync/)
  })

  test('coordinator has openWorkbook method (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/async openWorkbook\b/)
  })

  test('coordinator has NO adoptLegacySession method (Increment 16 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/async adoptLegacySession\b/)
    expect(stripped).not.toMatch(/\.adoptLegacySession\b/)
  })

  test('sheets-runtime has NO adoptLegacySessionIntoCoordinator export (Increment 16 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-runtime.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/export function adoptLegacySessionIntoCoordinator\b/)
    expect(stripped).not.toMatch(/export interface LegacySessionAdoption\b/)
  })

  test('sheets-main has NO legacy openWorkbookSession function (Increment 16 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^async function openWorkbookSession\b/m)
  })

  test('sheets-main has NO legacy prepareWorkbookForOpen function (Increment 16 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^async function prepareWorkbookForOpen\b/m)
  })

  test('sheets-main has NO adoptLegacySessionFromWorkbookFile function (Increment 16 — removed)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/adoptLegacySessionFromWorkbookFile\b/)
  })

  test('coordinator openWorkbook uses service.convertWorkbook for .xls (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/this\.deps\.service\.convertWorkbook\b/)
  })

  test('coordinator has recoveryDialogText dep callback (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/recoveryDialogText\??\s*:\s*\(\)\s*=>\s*\{/)
  })

  test('coordinator has onWorkbookOpened dep callback (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/onWorkbookOpened\??\s*:\s*\(wcId:\s*number,\s*openedPath:\s*string\)\s*=>\s*void/)
  })

  test('coordinator has consumeQueuedWorkbookPath dep callback (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-shell-coordinator.ts'), 'utf8')
    expect(src).toMatch(/consumeQueuedWorkbookPath\??\s*:\s*\(\)\s*=>\s*string\s*\|\s*undefined/)
  })

  test('sheets-main resolveSheetsSessionPath reads from coordinator (NOT legacy mirror) (Increment 16)', () => {
    const src = readFileSync(join(SRC, 'main', 'sheets-main.ts'), 'utf8')
    expect(src).toMatch(/getMigratedRuntime\(\)\.coordinator\.getSession/)
  })

  test('SpreadsheetService contract has convertWorkbook (Increment 16)', () => {
    const src = readFileSync(
      join(__dirname, '..', '..', '..', 'packages', 'runtime-contracts', 'src', 'services', 'sheets.ts'),
      'utf8',
    )
    expect(src).toMatch(/convertWorkbook\s*\(/)
    expect(src).toMatch(/Promise<\{\s*data:\s*Uint8Array;\s*fileName:\s*string\s*\}>/)
  })

  // ═══ INCREMENT 18 — Zod version alignment guard ═══
  //
  // apps/sheets re-exports schemas from @genoffice/xlsx-gateway (e.g.
  // workbookOperationSchema). xlsx-gateway uses zod ^3.23.8. If apps/sheets
  // declares zod ^4.x, the z.array(workbookOperationSchema).parse() call
  // in renderer/ai/tools.ts fails at runtime + compile time (zod 4's
  // ZodArray expects a schema with `_zod`, which zod 3 schemas don't have).
  //
  // This guard verifies apps/sheets declares the SAME zod major version
  // as xlsx-gateway — preventing a silent mismatch from reappearing.

  test('apps/sheets declares zod ^3.x (matching xlsx-gateway) — Increment 18', () => {
    const pkg = JSON.parse(readFileSync(PACKAGE_JSON, 'utf8'))
    const sheetsZod = pkg.dependencies?.zod ?? ''
    // Must start with ^3. or ~3. (zod 3.x).
    expect(sheetsZod).toMatch(/^[~^]3\./)
  })

  test('xlsx-gateway declares zod ^3.x (canonical schema owner) — Increment 18', () => {
    const gatewayPkg = JSON.parse(
      readFileSync(join(__dirname, '..', '..', '..', 'packages', 'xlsx-gateway', 'package.json'), 'utf8'),
    )
    const gatewayZod = gatewayPkg.dependencies?.zod ?? ''
    expect(gatewayZod).toMatch(/^[~^]3\./)
  })
})
