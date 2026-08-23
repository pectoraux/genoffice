/**
 * Dependency-direction architecture test for @genoffice/runtime-contracts.
 *
 * Verifies that runtime-contracts does NOT import from any @genoffice/*-shared
 * alias (which points to apps/apps/apps/*lt;starapps/*gt;/src/shared/lt;starapps/apps/*lt;starapps/*gt;/src/shared/gt;/src/shared/) or any app package.
 */
import { describe, test, expect } from 'vitest'
import { join } from 'node:path'
import { readFileSync, readdirSync, statSync } from 'node:fs'

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

describe('dependency direction: runtime-contracts must not import from app layer', () => {
  test('runtime-contracts has ZERO imports from @genoffice/*-shared', () => {
    const SRC = join(__dirname, '..', 'src')
    const hits = scanForImports(SRC, [
      '@genoffice/docs-shared',
      '@genoffice/sheets-shared',
      '@genoffice/slides-shared',
      '@genoffice/pdf-shared',
      '@genoffice/markdown-shared',
      '@genoffice/shell-home-shared',
      '@genoffice/shell-tabs-shared',
      '@genoffice/shell-update-shared',
    ])
    expect(hits).toEqual([])
  })

  test('runtime-contracts has ZERO imports from app packages (@genoffice/docs, etc.)', () => {
    const SRC = join(__dirname, '..', 'src')
    const hits = scanForImports(SRC, [
      '@genoffice/docs',
      '@genoffice/sheets',
      '@genoffice/slides',
      '@genoffice/pdf',
      '@genoffice/markdown',
      '@genoffice/shell',
    ])
    expect(hits).toEqual([])
  })

  // Increment 2F: runtime-contracts (Layer 1) MUST NOT depend on @genoffice/platform
  // (Layer 3). The previous Increment 2E violated this by importing `DialogParent`
  // from platform into services/docs.ts. This test enforces the frozen dependency
  // direction: runtime-contracts → (nothing below it except @genoffice/ai-provider
  // and @genoffice/font-metrics, which are workspace packages, not platform layers).
  //
  // EXCEPTION: src/runtime.ts imports capability INTERFACES (Storage, Files, AI,
  // Printing, Clipboard, Notifications, Windowing, Settings) from @genoffice/platform
  // because RuntimeContext AGGREGATES them. This is a pre-existing frozen contract
  // (ADR-001) — the RuntimeContext type MUST reference the capability interfaces.
  // The exception is narrowly scoped: ONLY runtime.ts may import from platform,
  // and ONLY type-only imports (no runtime values). No other file may import
  // from @genoffice/platform — especially not services/docs.ts (which 2E
  // incorrectly imported DialogParent from).
  test('runtime-contracts has ZERO imports from @genoffice/platform EXCEPT runtime.ts (frozen RuntimeContext aggregate)', () => {
    const SRC = join(__dirname, '..', 'src')
    const hits = scanForImports(SRC, ['@genoffice/platform'])
    // Allow ONLY src/runtime.ts (the frozen RuntimeContext aggregate)
    const violations = hits.filter((h) => !h.file.endsWith('runtime.ts'))
    expect(violations).toEqual([])
  })

  // ═══ INCREMENT 15A — WorkbookPivotDefinition contract purity ═══
  //
  // The `WorkbookPivotDefinition` contract lives in src/services/pivot-definition.ts.
  // It MUST be runtime-independent: ZERO Electron imports, ZERO node:* imports,
  // ZERO @genoffice/xlsx-gateway imports (runtime-contracts cannot depend on
  // xlsx-gateway — runtime-contracts is Layer 1, xlsx-gateway is a peer of
  // platform-electron at Layer 4a). The shape mirrors xlsx-gateway's
  // `PivotDefinition` structurally; the engine impl is the single translation
  // point that bridges the two.
  test('runtime-contracts pivot-definition.ts has ZERO Electron/node:/xlsx-gateway imports (Increment 15A)', () => {
    const SRC = join(__dirname, '..', 'src')
    const hits = scanForImports(SRC, [
      'electron',
      /^node:/,
      '@genoffice/xlsx-gateway',
    ])
    // No file in runtime-contracts may import any of these.
    expect(hits).toEqual([])
  })

  test('runtime-contracts exports WorkbookPivotDefinition (Increment 15A)', () => {
    const SRC = join(__dirname, '..', 'src')
    const pivotFile = join(SRC, 'services', 'pivot-definition.ts')
    const text = readFileSync(pivotFile, 'utf8')
    expect(text).toMatch(/export interface WorkbookPivotDefinition\b/)
    expect(text).toMatch(/export type WorkbookPivotFilterDef\b/)
    expect(text).toMatch(/export type WorkbookPivotFieldGrouping\b/)
    expect(text).toMatch(/export interface WorkbookPivotCacheField\b/)
    expect(text).toMatch(/export interface WorkbookPivotDataField\b/)
    expect(text).toMatch(/export interface WorkbookPivotLayoutLine\b/)
    expect(text).toMatch(/export type WorkbookPivotSharedItem\b/)
    expect(text).toMatch(/export interface WorkbookPivotFieldItem\b/)
  })

  test('runtime-contracts SpreadsheetEngine.readPivotDefinition returns Promise<WorkbookPivotDefinition> (Increment 15A)', () => {
    const SRC = join(__dirname, '..', 'src')
    const engineFile = join(SRC, 'services', 'spreadsheet-engine.ts')
    const text = readFileSync(engineFile, 'utf8')
    expect(text).toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/m)
    // The generic readArchiveEntry method MUST NOT exist on the contract.
    // Strip JSDoc/block comments before checking — the contract may
    // reference `readArchiveEntry` in JSDoc rationale, but must NOT
    // declare it as a method.
    const stripped = text.replace(/\/\*\*?[\s\S]*?\*\//g, '')
    expect(stripped).not.toMatch(/\breadArchiveEntry\s*\(/)
  })
})
