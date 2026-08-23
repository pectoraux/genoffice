/**
 * Architecture-boundary test for @genoffice/services-sheets.
 *
 * Enforces:
 *   - ZERO imports of electron
 *   - ZERO imports of node:* (no node:fs, node:crypto, node:path, node:buffer)
 *   - ZERO references to BrowserWindow / webContents / wcId
 *   - ZERO references to child_process / Rust / stdio
 *   - ZERO references to snapshotPath / sidecarSessionId / engineSessionId
 *   - Does NOT import platform-electron
 *   - DOES import runtime-contracts (dependency direction)
 *
 * DOMAIN-EVENT PURITY (Increment 3A correction):
 *   - ZERO references to SheetsEventBus
 *   - ZERO references to onOpened / onRenamed / onTeardown
 *   - ZERO references to oldPath / newPath
 *   The shell coordinator owns renderer/event routing — the domain service
 *   must NOT.
 */
import { describe, test, expect } from 'vitest'
import { join } from 'node:path'
import { readFileSync, readdirSync, statSync } from 'node:fs'

const SRC = join(__dirname, '..', 'src')

function listSourceFiles(rootDir: string): string[] {
  const out: string[] = []
  const walk = (dir: string): void => {
    for (const name of readdirSync(dir)) {
      const full = join(dir, name)
      const st = statSync(full)
      if (st.isDirectory()) {
        if (name === 'node_modules' || name === 'dist' || name === '.git') continue
        walk(full)
      } else if (st.isFile() && (full.endsWith('.ts') || full.endsWith('.tsx'))) {
        out.push(full)
      }
    }
  }
  walk(rootDir)
  return out
}

function scanForImports(
  rootDir: string,
  forbidden: Array<string | RegExp>,
): Array<{ file: string; line: number; text: string }> {
  const hits: Array<{ file: string; line: number; text: string }> = []
  const importPattern = /(?:from\s+|require\s*\(\s*)['"]([^'"]+)['"]/g
  for (const file of listSourceFiles(rootDir)) {
    const text = readFileSync(file, 'utf8')
    let m: RegExpExecArray | null
    while ((m = importPattern.exec(text)) !== null) {
      const mod = m[1]
      const lineNum = text.slice(0, m.index).split('\n').length
      for (const f of forbidden) {
        const isHit = typeof f === 'string' ? mod === f || mod.startsWith(f + '/') : f.test(mod)
        if (isHit) {
          hits.push({ file, line: lineNum, text: `import ... from '${mod}'` })
        }
      }
    }
  }
  return hits
}

function scanForTokens(rootDir: string, forbidden: string[]): Array<{ file: string; line: number; text: string }> {
  const hits: Array<{ file: string; line: number; text: string }> = []
  for (const file of listSourceFiles(rootDir)) {
    const lines = readFileSync(file, 'utf8').split('\n')
    lines.forEach((line, i) => {
      for (const token of forbidden) {
        if (line.includes(token)) {
          hits.push({ file, line: i + 1, text: line.trim() })
        }
      }
    })
  }
  return hits
}

describe('@genoffice/services-sheets architecture boundary', () => {
  test('ZERO imports of electron', () => {
    const hits = scanForImports(SRC, ['electron'])
    expect(hits).toEqual([])
  })

  test('ZERO imports of node:*', () => {
    const hits = scanForImports(SRC, [/^node:/])
    expect(hits).toEqual([])
  })

  test('ZERO references to BrowserWindow / webContents / wcId', () => {
    const hits = scanForTokens(SRC, [
      'BrowserWindow',
      'webContents',
      'wcId',
      'WebContentsView',
    ]).filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to child_process / Rust / stdio', () => {
    const hits = scanForTokens(SRC, [
      'child_process',
      'Rust',
      'stdio',
    ]).filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to snapshotPath / sidecarSessionId / engineSessionId', () => {
    const hits = scanForTokens(SRC, [
      'snapshotPath',
      'sidecarSessionId',
      'engineSessionId',
    ]).filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('does NOT import platform-electron', () => {
    const hits = scanForImports(SRC, ['@genoffice/platform-electron'])
    expect(hits).toEqual([])
  })

  test('does NOT import from apps/sheets', () => {
    const hits = scanForImports(SRC, [/apps\/sheets/])
    expect(hits).toEqual([])
  })

  test('does NOT import XlsxSidecarClient', () => {
    const hits = scanForTokens(SRC, ['XlsxSidecarClient'])
    expect(hits).toEqual([])
  })

  test('imports runtime-contracts (dependency direction)', () => {
    const hits = scanForImports(SRC, ['@genoffice/runtime-contracts'])
    expect(hits.length).toBeGreaterThan(0)
  })

  // ── DOMAIN-EVENT PURITY (Increment 3A correction) ──────────────────
  //
  // The domain service must NOT own renderer/event routing. The shell
  // coordinator owns `docs/workbook opened`, `renamed`, `teardown` and
  // dispatches renderer notifications. The runtime-independent service
  // contract must remain domain-only.

  test('ZERO references to SheetsEventBus (shell owns event routing)', () => {
    const hits = scanForTokens(SRC, ['SheetsEventBus'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to onOpened / onRenamed / onTeardown', () => {
    const hits = scanForTokens(SRC, ['onOpened', 'onRenamed', 'onTeardown'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to oldPath / newPath (no filesystem-specific event payloads)', () => {
    const hits = scanForTokens(SRC, ['oldPath', 'newPath'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to WebContents / BrowserWindow / wcId (shell-layer concerns)', () => {
    const hits = scanForTokens(SRC, ['WebContents', 'BrowserWindow', 'wcId'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to workbookPath (renamed to workbookName in 3A)', () => {
    const hits = scanForTokens(SRC, ['workbookPath'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  // ── SAVE DOMAIN MODEL (Increment 3B/3C correction) ────────────────
  //
  // The service must NOT leak EngineArchivePatch into its source — the
  // engine-specific archive type is PRIVATE to the engine implementation
  // (packages/platform-electron/). The service accepts a domain SavePlan
  // (preserving all mutation families) and delegates to engine.applySavePlan
  // (Increment 3C: no SavePlanTranslator, no SavePlanTranslation, no
  // EngineArchivePatch in runtime-contracts or services-sheets).

  test('ZERO references to EngineArchivePatch in the service implementation source (Increment 3C)', () => {
    // EngineArchivePatch is an engine-internal type defined only in
    // packages/platform-electron/. It must NOT appear in services-sheets.
    const hits = scanForTokens(SRC, ['EngineArchivePatch'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to SavePlanTranslator / SavePlanTranslation (Increment 3C removed them)', () => {
    const hits = scanForTokens(SRC, ['SavePlanTranslator', 'SavePlanTranslation'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO references to xlsx-gateway / xlsx-package-io (engine boundary translation is injected)', () => {
    // The service must NOT statically OR dynamically import from
    // xlsx-gateway. The engine implementation (platform-electron) is
    // the single translation point between OOXML wire format and the
    // runtime-independent WorkbookPivotDefinition contract.
    //
    // INCREMENT 15A: the previous version of this test only matched
    // `from '...'` and `require('...')` patterns — it missed dynamic
    // `await import('...')`. The service used to use a dynamic import
    // to call `parsePivotDefinition` directly, which slipped through.
    // This guard now matches ALL forms: static, dynamic, require.
    //
    // Comments are stripped before matching so JSDoc rationale that
    // references `xlsx-gateway` does not produce false positives.
    const hits: Array<{ file: string; line: number; text: string }> = []
    for (const file of listSourceFiles(SRC)) {
      const text = readFileSync(file, 'utf8')
      // Strip block comments and line comments — only CODE counts.
      const stripped = text
        .replace(/\/\*\*?[\s\S]*?\*\//g, '')
        .replace(/^\s*\/\/.*$/gm, '')
      const lines = stripped.split('\n')
      lines.forEach((line, i) => {
        // Static import: from 'xlsx-gateway' / require('xlsx-gateway')
        // Dynamic import: await import('xlsx-gateway')
        if (
          /(?:from\s+|require\s*\(\s*)['"]@genoffice\/xlsx-gateway/.test(line) ||
          /import\s*\(\s*['"]@genoffice\/xlsx-gateway/.test(line) ||
          /(?:from\s+|require\s*\(\s*)['"][^'"]*xlsx-package-io/.test(line) ||
          /import\s*\(\s*['"][^'"]*xlsx-package-io/.test(line)
        ) {
          hits.push({ file, line: i + 1, text: line.trim() })
        }
      })
    }
    expect(hits).toEqual([])
  })

  test('ZERO references to XlsxSidecarClient / sidecar (no direct sidecar coupling)', () => {
    const hits = scanForTokens(SRC, ['XlsxSidecarClient', 'sidecar'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits).toEqual([])
  })

  test('ZERO Promise<unknown> return types for pivot (Increment 15A)', () => {
    // The service's readPivotDefinition method must declare a typed
    // return — `Promise<WorkbookPivotDefinition>`, NOT `Promise<unknown>`.
    // This guards against regression: the prior version returned
    // Promise<unknown> and the typed contract was lost.
    const src = readFileSync(join(SRC, 'spreadsheet-service.ts'), 'utf8')
    expect(src).toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/m)
    expect(src).not.toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<unknown>/m)
  })

  test('ZERO raw sidecar protocol construction in services-sheets (Increment 15A)', () => {
    // The service must NOT construct `{ command: '...' }` sidecar
    // payloads — all sidecar wire-protocol construction lives behind
    // the engine boundary (ElectronXlsxSidecarEngine).
    const src = readFileSync(join(SRC, 'spreadsheet-service.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/command:\s*['"]read_entries['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]read_range['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]open['"]/)
    expect(stripped).not.toMatch(/command:\s*['"]save_archive['"]/)
  })

  test('ZERO filesystem references in services-sheets (Increment 15A)', () => {
    // The service must not perform filesystem operations — all archive
    // I/O is private to the engine adapter. The service operates on
    // Uint8Array content + opaque handles only.
    const src = readFileSync(join(SRC, 'spreadsheet-service.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/^import.*node:fs/m)
    expect(stripped).not.toMatch(/readFileSync|writeFileSync|mkdirSync|rmSync|existsSync|renameSync/)
  })

  test('ZERO Electron references in services-sheets (Increment 15A)', () => {
    // The service must not import Electron — it is runtime-independent.
    const src = readFileSync(join(SRC, 'spreadsheet-service.ts'), 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    expect(stripped).not.toMatch(/from\s+['"]electron['"]/)
    expect(stripped).not.toMatch(/BrowserWindow|WebContents|webContents/)
  })

  test('SpreadsheetServiceDeps is referenced (engine-only dependency, no translator)', () => {
    // Increment 3C: SpreadsheetServiceDeps contains ONLY the engine —
    // the SavePlanTranslator dependency was removed. The engine accepts
    // the domain SavePlan directly via applySavePlan.
    const hits = scanForTokens(SRC, ['SpreadsheetServiceDeps'])
      .filter((h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'))
    expect(hits.length).toBeGreaterThan(0)
  })
})
