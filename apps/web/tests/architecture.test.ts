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

// ── EXCEL-018: Remove Duplicates canonical-path guards ─────────────────────
//
// Remove Duplicates is implemented as a value-level rewrite through the
// EXISTING cell-edit mutation family (sheet.mutation.set-range-values →
// cellEditFromMutation → savePlan.edits → applyCellEditsToXlsx). The
// guards below enforce that no future "shortcut" introduces:
//   - a parallel XLSX engine in the browser (jszip / OOXML construction
//     inside the dedupe path),
//   - a new save-plan family or wire field for "dedupe ops",
//   - a non-canonical write path that bypasses FWorksheet.setValues /
//     FRange.setValue (which would skip the set-range-values journal
//     subscription and lose save/reopen fidelity).
//
// The dedupe module is a PURE function — it must not import Univer, the
// gateway, the save plan, or any host API. The wiring happens in
// useExcelRuntime, which calls the dedupe with the values it read from
// the live Univer range.

describe('architecture: EXCEL-018 Remove Duplicates uses the canonical cell-edit path', () => {
  const dedupePath = join(WEB_ROOT, 'src', 'office', 'dedupe.ts')
  const runtimePath = join(WEB_ROOT, 'src', 'screens', 'excel', 'useExcelRuntime.ts')
  const ribbonPath = join(WEB_ROOT, 'src', 'screens', 'excel', 'Ribbon.tsx')

  it('apps/web/src/office/dedupe.ts exists', () => {
    expect(existsSync(dedupePath), `${dedupePath} should exist`).toBe(true)
  })

  it('dedupe.ts is a PURE module — no Univer, gateway, save-plan, or host imports', () => {
    expect(existsSync(dedupePath), `${dedupePath} should exist`).toBe(true)
    const content = readFileSync(dedupePath, 'utf8')
    const lines = nonCommentLines(content)
    // The dedupe module must not import anything that would make it
    // non-browser-safe or that would let it bypass the canonical
    // cell-edit channel. It's a pure value-transformation function.
    const forbidden = [
      /from\s+['"]@univerjs/,
      /from\s+['"]@genoffice\/xlsx-gateway['"]/,
      /from\s+['"]@genoffice\/docx-engine['"]/,
      /from\s+['"]electron['"]/,
      /from\s+['"]node:/,
      /from\s+['"]jszip['"]/,
      /from\s+['"]fs['"]/,
    ]
    const violations = lines.filter((line) => forbidden.some((re) => re.test(line)))
    expect(
      violations,
      'dedupe.ts must be a pure module — no Univer/gateway/electron/node imports',
    ).toEqual([])
  })

  it('dedupe.ts exports the dedupeRows function', () => {
    expect(existsSync(dedupePath), `${dedupePath} should exist`).toBe(true)
    const content = readFileSync(dedupePath, 'utf8')
    expect(content).toMatch(/export\s+function\s+dedupeRows\b/)
  })

  it('useExcelRuntime.ts wires removeDuplicates through FWorksheet.getRange().setValues', () => {
    // The canonical write path is FWorksheet.getRange(row, col, numRows,
    // numCols).setValues(matrix) — the SAME facade Sort uses, firing
    // sheet.mutation.set-range-values which is journaled by the existing
    // subscription. A shortcut that writes via setValueForCell per cell
    // would still be canonical (same mutation id), but a shortcut that
    // writes via a private worksheet cellDataMatrix bypass would NOT.
    // The guard below catches the latter.
    expect(existsSync(runtimePath), `${runtimePath} should exist`).toBe(true)
    const content = readFileSync(runtimePath, 'utf8')
    expect(content).toMatch(/removeDuplicates/)
    expect(content).toMatch(/\.setValues\(/)
    // The dedupe algorithm is delegated to the pure module — no inline
    // reimplementation of the comparison key.
    expect(content).toMatch(/from\s+['"][^'"]*office\/dedupe['"]/)
  })

  it('useExcelRuntime.removeDuplicates reads computed values via getValues()', () => {
    expect(existsSync(runtimePath), `${runtimePath} should exist`).toBe(true)
    const content = readFileSync(runtimePath, 'utf8')
    // The dedupe MUST read computed values (the result of formulas) —
    // matching the desktop's `range.getValues()` call. Reading raw
    // cell data would compare formula TEXT instead of results, which
    // is a different semantic.
    expect(content).toMatch(/\.getValues\(\)/)
  })

  it('Ribbon.tsx enables the Remove Duplicates button (no longer disabled)', () => {
    expect(existsSync(ribbonPath), `${ribbonPath} should exist`).toBe(true)
    const content = readFileSync(ribbonPath, 'utf8')
    // The button must NOT carry a `disabled` attribute that depends on
    // a missing feature. It must call `api.removeDuplicates(...)` (or
    // open the dedupe dialog that does).
    // Find the Remove Duplicates button block — it must NOT contain
    // a top-level `disabled` prop (without a value).
    const buttonBlockMatch = content.match(/label="Remove Duplicates"[\s\S]*?\/>/)
    expect(buttonBlockMatch, 'Remove Duplicates button block present').not.toBeNull()
    const buttonBlock = buttonBlockMatch![0]
    // The button block must NOT contain a bare `disabled` prop (the
    // pattern that the previous placeholder used: `<RibbonButton ...
    // disabled\n/>`). A conditional `disabled={disabled}` is fine —
    // that just disables while the runtime is booting.
    expect(buttonBlock).not.toMatch(/^\s*disabled\s*$/m)
    expect(buttonBlock).not.toMatch(/\sdisabled\s+onClick/)
    // The button must call removeDuplicates somewhere in its handler —
    // either `api.removeDuplicates(...)` (after a null guard) or
    // `api?.removeDuplicates(...)` (optional-chained).
    expect(
      /api\??\.removeDuplicates\(/.test(content),
      'Ribbon must call api.removeDuplicates(...) or api?.removeDuplicates(...)',
    ).toBe(true)
  })

  it('apps/web/src has NO raw OOXML or JSZip construction for dedupe', () => {
    // The dedupe path must not introduce any OOXML / JSZip / XML
    // construction in the browser. The browser only exchanges typed
    // CellEdit payloads with the canonical gateway.
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    const ooxmlPatterns = [
      /new\s+DOMParser\s*\(\s*\)\s*\.\s*parseFromString\s*\([^)]*['"](?:application|text)\/xml['"]/,
      /\bXMLSerializer\b/,
      /from\s+['"]jszip['"]/,
      /<sheetData\b/,
      /<worksheet\b/,
      /<row\s+r=/,
    ]
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => ooxmlPatterns.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('ExcelEditor save plan does NOT introduce a "dedupeOps" or "removeDuplicates" family', () => {
    // The canonical save plan (BrowserWorkbookSavePlan) exposes the
    // families: edits, structuralOps, pageSetupStates, filterStates,
    // dvStates, noteStates. Remove Duplicates must NOT add a new
    // family — it emits CellEdits through the EXISTING edits channel.
    const editorPath = join(WEB_ROOT, 'src', 'screens', 'ExcelEditor.tsx')
    expect(existsSync(editorPath), `${editorPath} should exist`).toBe(true)
    const content = readFileSync(editorPath, 'utf8')
    const lines = nonCommentLines(content)
    // A new save-plan family would appear as `dedupeOps: ...` or
    // `removeDuplicates: ...` in the savePlan object literal. The guard
    // catches that pattern.
    const newFamilyPatterns = [
      /\bdedupeOps\b/,
      /\bremoveDuplicatesState\b/,
      /\bdedupeStates\b/,
      // The save plan assembles these fields; a removeDuplicates
      // family would be a new top-level key. Match a `removeDuplicates:`
      // or `dedupe:` key inside a savePlan literal — but NOT inside
      // the runtime API surface (api.removeDuplicates(...)).
      /savePlan\s*:\s*\{[^}]*\bremoveDuplicates:/s,
      /savePlan\s*:\s*\{[^}]*\bdedupe:/s,
    ]
    const violations = lines.filter((line) => newFamilyPatterns.some((re) => re.test(line)))
    expect(
      violations,
      'ExcelEditor save plan must NOT introduce a new dedupe family — emit CellEdits through the existing edits channel',
    ).toEqual([])
  })
})
