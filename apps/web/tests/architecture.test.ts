/**
 * Architecture tests for apps/web.
 *
 * Enforces the browser purity boundary:
 *   - apps/web/src/ has ZERO imports of `electron`, `node:*`, `fs`,
 *     `child_process`, `@genoffice/platform-electron`.
 *   - The office API client uses ONLY `fetch` (no Node APIs).
 *
 * Purity is critical: the browser bundle must never pull in Node-only
 * modules. The office engines (`@genoffice/xlsx-gateway`,
 * `@genoffice/docx-engine`) are server-only; the browser may import their
 * TYPES (erased at compile time) but must not import their runtime code.
 *
 * Reads actual source files. Skips comment lines to avoid false positives.
 */

import { describe, it, expect } from 'vitest'
import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs'
import { join, relative, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
// apps/web/tests/architecture.test.ts — up 1 level to apps/web/
const WEB_ROOT = resolve(__dirname, '..')

function walkTs(dir: string): string[] {
  const files: string[] = []
  if (!existsSync(dir)) return files
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry)
    if (statSync(full).isDirectory()) {
      files.push(...walkTs(full))
    } else if (entry.endsWith('.ts') || entry.endsWith('.tsx')) {
      files.push(full)
    }
  }
  return files
}

function readFiles(dir: string): { rel: string; content: string }[] {
  return walkTs(dir).map((p) => ({ rel: relative(WEB_ROOT, p), content: readFileSync(p, 'utf8') }))
}

function nonCommentLines(content: string): string[] {
  return content.split('\n').filter((line) => {
    const t = line.trim()
    return !t.startsWith('//') && !t.startsWith('*') && !t.startsWith('/*')
  })
}

// ── Path-resolution guard ───────────────────────────────────────────────────

describe('architecture: apps/web source directory is non-empty', () => {
  it('apps/web/src exists and contains source files', () => {
    const src = join(WEB_ROOT, 'src')
    const files = walkTs(src)
    expect(
      files.length,
      `apps/web/src should contain source files (found 0 — path resolution broken?)`,
    ).toBeGreaterThan(0)
  })
})

// ── Forbidden imports ───────────────────────────────────────────────────────

describe('architecture: apps/web has zero Electron / Node API imports', () => {
  const forbidden = [
    // Electron runtime
    /from\s+['"]electron['"]/,
    /from\s+['"]@genoffice\/electron-utils['"]/,
    /from\s+['"]@genoffice\/platform-electron['"]/,
    /from\s+['"]@genoffice\/project-store['"]/,
    /from\s+['"]apps\/shell['"]/,
    // Frozen desktop/doc surfaces — the web shell must NOT reach into the
    // Electron apps' renderer or the docs app source. (Phase 4 §J.)
    /['"][^'"]*apps\/sheets\/src[^'"]*['"]/,
    /['"][^'"]*apps\/docs\/src[^'"]*['"]/,
    /['"][^'"]*apps\/sheets\/(renderer|shared|preload)[^'"]*['"]/,
    // Node built-ins (require / import)
    /from\s+['"]node:[a-z]+['"]/,
    /from\s+['"]fs['"]/,
    /from\s+['"]fs\/promises['"]/,
    /from\s+['"]child_process['"]/,
    /from\s+['"]path['"]/,
    /from\s+['"]os['"]/,
    /from\s+['"]crypto['"]/,
    /from\s+['"]stream['"]/,
    /from\s+['"]http['"]/,
    /from\s+['"]https['"]/,
    /from\s+['"]net['"]/,
    /from\s+['"]url['"]/,
    /from\s+['"]util['"]/,
    /from\s+['"]buffer['"]/,
    // DB drivers
    /from\s+['"]pg['"]/,
    /from\s+['"]@electric-sql\/pglite['"]/,
    // Contractor core persistence / service / storage (server-only)
    /from\s+['"]@contractor\/core\/persistence['"]/,
    /from\s+['"]@contractor\/core\/service['"]/,
    /from\s+['"]@contractor\/core\/storage['"]/,
  ]

  it('apps/web/src does NOT import any forbidden Node / Electron modules', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    expect(webFiles.length, 'apps/web/src should have source files to scan').toBeGreaterThan(0)
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => forbidden.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src does NOT contain `require(` calls', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => /\brequire\s*\(/.test(line))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src does NOT reference process.env at runtime (browser must not read Node env)', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    // Allow `process.env` only inside `import.meta.env`-style Vite types; in
    // practice we just ban `process.env` outright here.
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => /\bprocess\.env\b/.test(line))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  // Phase 4 Increment 4 (Data → Filter): the browser is a thin typed client.
  // It must NEVER parse or serialize OOXML itself — the canonical
  // xlsx-gateway owns all XML work, and the browser only exchanges typed
  // SheetFilterState / CellEdit / structural-op payloads.
  it('apps/web/src does NOT import JSZip (no browser-side archive handling)', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => /from\s+['"]jszip['"]/.test(line))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src does NOT do raw OOXML work (no XML parsing/serialization of sheet parts)', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    // Patterns that indicate hand-rolled OOXML handling in the browser:
    // XML-mode DOM parsing, sheet-part tags, or building worksheet XML
    // strings. (HTML-mode DOMParser is the Word editor's block pipeline and
    // is not OOXML work.)
    const ooxmlPatterns = [
      /new\s+DOMParser\s*\(\s*\)\s*\.\s*parseFromString\s*\([^)]*['"](?:application|text)\/xml['"]/,
      /\bXMLSerializer\b/,
      /<autoFilter\b/,
      /<sheetData\b/,
      /<worksheet\b/,
      /<mergeCells\b/,
      /<customFilters\b/,
      /<filterColumn\b/,
    ]
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => ooxmlPatterns.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src filter surfaces use ONLY the canonical typed SheetFilterState', () => {
    // The filter journal/save code must reference the canonical gateway type,
    // not a locally-declared duplicate of the filter model.
    const editorPath = join(WEB_ROOT, 'src', 'screens', 'ExcelEditor.tsx')
    expect(existsSync(editorPath), `${editorPath} should exist`).toBe(true)
    const content = readFileSync(editorPath, 'utf8')
    expect(content).toContain('SheetFilterState')
    // ...and must import it from the canonical package, not declare it.
    expect(
      /import type \{[^}]*SheetFilterState[^}]*\} from '@genoffice\/xlsx-gateway'/.test(content),
    ).toBe(true)
  })
})

// ── Office API client purity ────────────────────────────────────────────────

describe('architecture: office API client uses only fetch', () => {
  it('apps/web/src/api/office-client.ts uses ONLY fetch for HTTP', () => {
    const clientPath = join(WEB_ROOT, 'src', 'api', 'office-client.ts')
    expect(existsSync(clientPath), `${clientPath} should exist`).toBe(true)
    const content = readFileSync(clientPath, 'utf8')
    const lines = nonCommentLines(content)
    // The client must call fetch() at least once (to make requests).
    expect(
      lines.some((l) => /\bfetch\s*\(/.test(l)),
      'office-client.ts should call fetch()',
    ).toBe(true)
    // And must not import any HTTP/Node library.
    const forbiddenHttp = [
      /from\s+['"]node:http['"]/,
      /from\s+['"]node:https['"]/,
      /from\s+['"]axios['"]/,
      /from\s+['"]got['"]/,
      /from\s+['"]node-fetch['"]/,
      /from\s+['"]undici['"]/,
      /\brequire\s*\(/,
    ]
    const violations = lines.filter((line) => forbiddenHttp.some((re) => re.test(line)))
    expect(
      violations,
      'office-client.ts must not import any HTTP library other than fetch',
    ).toEqual([])
  })

  it('apps/web/src/api/office-client.ts does NOT use `as any` or `as unknown as`', () => {
    const clientPath = join(WEB_ROOT, 'src', 'api', 'office-client.ts')
    const content = readFileSync(clientPath, 'utf8')
    const lines = nonCommentLines(content)
    const violations = lines.filter(
      (l) => /\bas\s+any\b/.test(l) || /\bas\s+unknown\s+as\b/.test(l),
    )
    expect(violations, 'office-client.ts must not use `as any` or `as unknown as`').toEqual([])
  })

  it('apps/web/src/api/office-client.ts imports from @genoffice/* are type-only', () => {
    const clientPath = join(WEB_ROOT, 'src', 'api', 'office-client.ts')
    const content = readFileSync(clientPath, 'utf8')
    // A type-only import is `import type { ... } from '@genoffice/...'`.
    // A non-type import is `import { ... } from '@genoffice/...'` (no `type`).
    // We strip comments first so we don't false-positive on commented-out imports.
    const stripped = content.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '')
    // Match every `from '@genoffice/...'` import statement (possibly multi-line).
    // We capture whether the preceding `import` keyword was followed by `type`.
    const importRe = /import\s+(type\s+)?[^;]*?\bfrom\s+['"]@genoffice\/[^'"]+['"]/gs
    const matches = [...stripped.matchAll(importRe)]
    const nonTypeImports = matches.filter((m) => m[1] === undefined)
    expect(
      nonTypeImports.map((m) => m[0].replace(/\s+/g, ' ').trim()),
      '@genoffice/* imports in office-client.ts must be type-only (use `import type { ... }`)',
    ).toEqual([])
  })
})

// ── Regression: patterns that would catch future bypasses ──────────────────

describe('architecture: regression patterns', () => {
  it('would catch a direct electron import', () => {
    const re = /from\s+['"]electron['"]/
    expect(re.test("import { ipcRenderer } from 'electron'")).toBe(true)
  })

  it('would catch a node:fs import', () => {
    const re = /from\s+['"]node:fs['"]/
    expect(re.test("import { readFile } from 'node:fs'")).toBe(true)
  })

  it('would catch a require call', () => {
    expect(/\brequire\s*\(/.test("const fs = require('fs')")).toBe(true)
  })

  it('would NOT flag a type-only import', () => {
    const re = /from\s+['"]@genoffice\//
    const line = "import type { CellEdit } from '@genoffice/xlsx-gateway'"
    expect(re.test(line)).toBe(true)
    expect(/^\s*import\s+type\b/.test(line)).toBe(true)
  })

  it('would NOT flag a fetch call', () => {
    const re = /\bfetch\s*\(/
    expect(
      re.test("const res = await fetch('/api/office/workbooks/open', { method: 'POST' })"),
    ).toBe(true)
  })
})
