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

  // EXCEL-022 (Insert → Picture / image edit): the browser is a thin typed
  // client for the drawing surface too — it must NEVER parse or construct
  // drawing XML, relationship XML, or touch media parts directly; all
  // anchor/picture/relationship/media work stays in the canonical
  // xlsx-gateway.
  it('apps/web/src does NOT do drawing/relationship XML work (no image OOXML)', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    const drawingPatterns = [
      /<xdr:(?:wsDr|twoCellAnchor|oneCellAnchor|absoluteAnchor|pic)\b/,
      /<a:blip\b/,
      /relationships\/image/,
      /relationships\/drawing/,
      /xl\/media\//,
      /xl\/drawings\//,
      /Target="\.\.\//,
    ]
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => drawingPatterns.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src imports xlsx-gateway for TYPES only (no image mutation imports)', () => {
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    // A VALUE import from the gateway would bundle the engine into the
    // browser — only `import type` (erased at build time) is allowed.
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) =>
        /^import\s+\{[^}]*\}\s+from\s+'@genoffice\/xlsx-gateway'/.test(line.trim()),
      )
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src image surfaces use ONLY the canonical typed visual families', () => {
    // The image journal/save code must reference the canonical gateway
    // types (SheetVisualAddition / WorkbookVisualEdit on the wire,
    // SheetImageInfo on the read side), not locally-declared duplicates.
    const clientPath = join(WEB_ROOT, 'src', 'api', 'office-client.ts')
    expect(existsSync(clientPath), `${clientPath} should exist`).toBe(true)
    const clientContent = readFileSync(clientPath, 'utf8')
    expect(clientContent).toContain('SheetVisualAddition')
    expect(clientContent).toContain('WorkbookVisualEdit')
    expect(
      /import type \{[^}]*SheetVisualAddition[^}]*\} from '@genoffice\/xlsx-gateway'/.test(
        clientContent,
      ),
    ).toBe(true)
    const imagesPath = join(WEB_ROOT, 'src', 'office', 'sheet-images.ts')
    expect(existsSync(imagesPath), `${imagesPath} should exist`).toBe(true)
    const imagesContent = readFileSync(imagesPath, 'utf8')
    expect(imagesContent).toContain('SheetImageInfo')
    expect(
      /import type \{[^}]*SheetImageInfo[^}]*\} from '@genoffice\/xlsx-gateway'/.test(
        imagesContent,
      ),
    ).toBe(true)
  })

  // ARCHITECT REVIEW (PR #20, blocker 1): the browser image module must
  // reach Univer only through the PUBLIC facade surface — no `as unknown
  // as` casts and no private internals (`_image`, private transform
  // fields). Geometry reads go through toBuilder().buildAsync(), the
  // same public surface the facade's own setters build commands on.
  it('apps/web/src/office/sheet-images.ts has NO private-Univer-internals access', () => {
    const imagesPath = join(WEB_ROOT, 'src', 'office', 'sheet-images.ts')
    expect(existsSync(imagesPath), `${imagesPath} should exist`).toBe(true)
    const lines = readFileSync(imagesPath, 'utf8')
      .split('\n')
      .map((line, index) => ({ line, number: index + 1 }))
      .filter(({ line }) => !line.trim().startsWith('//') && !line.trim().startsWith('*'))
    const violations = lines.filter(({ line }) => /\bas unknown as\b|._image\b/.test(line))
    expect(
      violations.map((v) => `${v.number}: ${v.line.trim()}`),
      'sheet-images.ts must use only the public facade surface',
    ).toEqual([])
    // The public read adapter must be present (toBuilder + buildAsync).
    const content = readFileSync(imagesPath, 'utf8')
    expect(content).toContain('toBuilder()')
    expect(content).toContain('buildAsync()')
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

// ── EXCEL-018: Remove Duplicates canonical structural path guards ─────────
//
// Remove Duplicates is implemented as a sequence of canonical
// `remove-rows` structural ops — one per duplicate row, in DESCENDING
// offset order so earlier deletes don't shift later indices. Each
// delete fires `sheet.mutation.remove-rows` (Univer's row-deletion
// mutation), journaled by ExcelEditor's existing
// `STRUCTURAL_MUTATION_IDS` subscription as a `{ kind: 'remove-rows',
// index, count: 1 }` structural op in the save plan. The gateway's
// `applyStructuralOps` applies each op atomically:
//   - `transformSheetRows` renumbers `<row>` r= and inner `<c>` r=
//     (cell contents — value, formula text, style ref, hyperlink,
//     comment pointer, shared-formula si= — travel UNTOUCHED inside
//     their `<c>` elements).
//   - `transformFormulas` rewrites `<f>` bodies via `shiftFormulaText`
//     (relative + absolute + mixed references all track the moved
//     cells — the `$` markers are preserved by `shiftReferenceToken`'s
//     colDollar/rowDollar capture groups; references to deleted rows
//     throw `StructuralShiftError` — fail-closed).
//   - `transformRangedFeatures` shifts merges, autoFilter, hyperlink
//     sqref, dataValidation sqref, and conditionalFormatting sqref.
//
// This is the EXACT canonical path `excel-structural.spec.ts` already
// proves for Insert/Delete Rows — no value-rewrite, no formula loss.
// The guards below enforce that no future "shortcut" re-introduces:
//   - a value-level rewrite via FWorksheet.setValues that DESTROYS
//     formulas on moved rows (the architect's EXPLICIT objection to
//     the prior implementation — "moved rows become computed values"
//     is explicitly NOT accepted as formula preservation),
//   - a parallel XLSX engine in the browser (jszip / OOXML construction
//     inside the dedupe path),
//   - a new save-plan family or wire field for "dedupe ops",
//   - a non-canonical write path that bypasses ws.deleteRows /
//     FWorksheet.deleteRows (which would skip the structural-ops
//     journal subscription and lose save/reopen fidelity).
//
// The dedupe module is a PURE function — it must not import Univer, the
// gateway, the save plan, or any host API. The wiring happens in
// useExcelRuntime, which calls the dedupe with the values it read from
// the live Univer range.

describe('architecture: EXCEL-018 Remove Duplicates uses the canonical structural remove-rows path', () => {
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

  it('dedupe.ts exports both dedupeRows (legacy) and dedupeRowIndices (canonical)', () => {
    expect(existsSync(dedupePath), `${dedupePath} should exist`).toBe(true)
    const content = readFileSync(dedupePath, 'utf8')
    expect(content).toMatch(/export\s+function\s+dedupeRows\b/)
    // The canonical entry point returns { keptIndices, duplicateIndices,
    // removed } — the runtime uses this to issue `ws.deleteRows(...)` per
    // duplicate (structural path), NOT a value-rewrite via setValues.
    expect(content).toMatch(/export\s+function\s+dedupeRowIndices\b/)
  })

  it('useExcelRuntime.ts wires removeDuplicates through ws.deleteRows (structural path)', () => {
    // The canonical structural write path is FWorksheet.deleteRows(row,
    // count) — fires sheet.mutation.remove-rows, journaled by the
    // existing STRUCTURAL_MUTATION_IDS subscription as a remove-rows
    // structural op. A value-level shortcut via FWorksheet.getRange().
    // setValues would fire sheet.mutation.set-range-values and DESTROY
    // formulas on moved rows (a moved `=B6` becomes the literal
    // computed value `30`). The architect explicitly rejected this
    // path; this guard enforces the structural path.
    expect(existsSync(runtimePath), `${runtimePath} should exist`).toBe(true)
    const content = readFileSync(runtimePath, 'utf8')
    expect(content).toMatch(/removeDuplicates/)
    expect(content).toMatch(/\.deleteRows\(/)
    // The dedupe algorithm is delegated to the pure module — no inline
    // reimplementation of the comparison key. The canonical entry point
    // is dedupeRowIndices (returns indices, not values).
    expect(content).toMatch(/from\s+['"][^'"]*office\/dedupe['"]/)
    expect(content).toMatch(/dedupeRowIndices/)
  })

  it('useExcelRuntime.removeDuplicates does NOT rewrite moved rows via setValues', () => {
    // The prior implementation rewrote moved rows with their computed
    // values via FWorksheet.getRange(...).setValues(...), which DESTROYS
    // formulas on moved rows. The architect explicitly rejected this:
    // "The current behavior 'moved rows become computed values' is
    // explicitly NOT accepted as formula preservation." The guard below
    // catches any reintroduction of a setValues call inside the
    // removeDuplicates function body.
    expect(existsSync(runtimePath), `${runtimePath} should exist`).toBe(true)
    const content = readFileSync(runtimePath, 'utf8')
    // Locate the removeDuplicates callback body — from
    // `const removeDuplicates = useCallback(` to its closing `)`.
    const startMatch = content.match(/const\s+removeDuplicates\s+=\s+useCallback\(/)
    expect(startMatch, 'removeDuplicates useCallback present').not.toBeNull()
    const startIdx = startMatch!.index! + startMatch![0].length
    // Find the matching close `)` of useCallback. The body is bracket-
    // balanced; we walk forward counting ( / ) and stop at the first
    // standalone `)` that closes the useCallback call.
    let depth = 1
    let i = startIdx
    while (i < content.length && depth > 0) {
      const ch = content[i]
      if (ch === '(') depth++
      else if (ch === ')') depth--
      i++
    }
    const body = content.slice(startIdx, i - 1)
    // The body MUST call deleteRows (the structural path).
    expect(body, 'removeDuplicates body must call ws.deleteRows').toMatch(/\.deleteRows\(/)
    // The body MUST NOT call .setValues( — that path destroys formulas.
    expect(body, 'removeDuplicates body must NOT call .setValues(').not.toMatch(/\.setValues\(/)
    // The body MUST NOT pad with nulls and rewrite via getRange — the
    // structural path deletes rows atomically; no padding is needed.
    expect(body, 'removeDuplicates body must NOT pad with null rows').not.toMatch(/padded/)
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
    // dvStates, noteStates. Remove Duplicates emits `remove-rows` ops
    // through the EXISTING structuralOps channel — it MUST NOT add a
    // new top-level family.
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
      'ExcelEditor save plan must NOT introduce a new dedupe family — emit remove-rows ops through the existing structuralOps channel',
    ).toEqual([])
  })
})

// ── EXCEL-020: Protection canonical wire-family guards ────────────────────
//
// Sheet/Workbook Protection is implemented through the CANONICAL save-plan
// families the desktop already ships — no new engine path, no browser-side
// OOXML, no parallel protection model:
//   - Review → Protect Sheet journals per-sheet desired states and emits
//     `savePlan.sheetProtections: [{ sheetName, protected }]` (the wire
//     type + strict validation live in contractor-core's office-routes;
//     the gateway's applySheetProtection writes/removes the worksheet's
//     `<sheetProtection>` element).
//   - Review → Protect Workbook emits `savePlan.workbookProtectionState:
//     { lockStructure }` (applyWorkbookProtection writes workbook.xml).
//   - Review → Lock/Unlock Cell journals style-only CellEdits carrying
//     `style.protectionLocked` — the SAME canonical WorkbookStyleEdit
//     family formatting uses (the desktop's neutral-delta path for cell
//     protection, since Univer's OSS presets carry no cell-protection
//     model).
//   - Password semantics are fail-closed EVERYWHERE: the browser refuses
//     to unprotect a password-protected sheet/structure up front (the
//     file state's hasPassword flag comes from the gateway reader), the
//     wire rejects password-bearing payloads as unknown fields, and the
//     gateway throws SheetProtectionError on the write.
//
// The guards below enforce that no future shortcut re-introduces:
//   - browser-side OOXML/JSZip construction for protection elements,
//   - a protection write that bypasses the save plan (e.g. calling a
//     gateway function directly from the browser),
//   - the disabled Protect Sheet stub (the feature must stay wired),
//   - password fields on the wire (the canonical family carries none).

describe('architecture: EXCEL-020 Protection uses the canonical wire families', () => {
  const editorPath = join(WEB_ROOT, 'src', 'screens', 'ExcelEditor.tsx')
  const ribbonPath = join(WEB_ROOT, 'src', 'screens', 'excel', 'Ribbon.tsx')
  const clientPath = join(WEB_ROOT, 'src', 'api', 'office-client.ts')

  it('ExcelEditor emits the sheetProtections save-plan family', () => {
    expect(existsSync(editorPath), `${editorPath} should exist`).toBe(true)
    const content = readFileSync(editorPath, 'utf8')
    expect(content).toContain('sheetProtectionJournalRef')
    expect(
      /sheetProtections\.length > 0 \? \{ sheetProtections \} : \{\}/.test(content),
      'handleSave must conditionally emit the sheetProtections family',
    ).toBe(true)
  })

  it('ExcelEditor emits the workbookProtectionState save-plan family', () => {
    const content = readFileSync(editorPath, 'utf8')
    expect(content).toContain('workbookProtectionJournalRef')
    expect(
      /workbookProtectionState !== null \? \{ workbookProtectionState \} : \{\}/.test(content),
      'handleSave must conditionally emit the workbookProtectionState family',
    ).toBe(true)
  })

  it('ExcelEditor toggles implement desktop recordSheetProtection semantics', () => {
    const content = readFileSync(editorPath, 'utf8')
    // The journal entry is DROPPED when the desired state matches the
    // file's original — the desktop's recordSheetProtection contract.
    expect(content).toContain('desired === original')
    // Password fail-closed guard BEFORE any journal write.
    expect(content).toContain('hasPassword')
    expect(content).toContain('removing its protection is not supported')
  })

  it('Ribbon Protection group carries NO disabled stub and wires all four commands', () => {
    expect(existsSync(ribbonPath), `${ribbonPath} should exist`).toBe(true)
    const content = readFileSync(ribbonPath, 'utf8')
    // The old stub documented the missing wire family — it must be gone.
    expect(
      content.includes('does not yet expose the sheetProtections family'),
      'the disabled Protect Sheet stub must be removed',
    ).toBe(false)
    expect(content.includes("'Protect Sheet'")).toBe(true)
    expect(content.includes("'Unprotect Sheet'")).toBe(true)
    expect(content.includes("'Protect Workbook'")).toBe(true)
    expect(content.includes("'Unprotect Workbook'")).toBe(true)
    expect(content.includes(`label="Lock Cell"`)).toBe(true)
    expect(content.includes(`label="Unlock Cell"`)).toBe(true)
    // The buttons must call back into the shell-owned journal handlers.
    expect(content).toContain('protection?.onToggleSheetProtection()')
    expect(content).toContain('protection?.onToggleWorkbookProtection()')
    expect(content).toContain('protection?.onSetCellsLocked(true)')
    expect(content).toContain('protection?.onSetCellsLocked(false)')
  })

  it('Lock/Unlock Cell journals canonical protectionLocked style-only CellEdits', () => {
    const content = readFileSync(editorPath, 'utf8')
    expect(
      content.includes('style: { protectionLocked: locked }'),
      'setCellsLocked must journal WorkbookStyleEdit.protectionLocked deltas',
    ).toBe(true)
    // Style-only edits never rewrite cell content.
    expect(content).toContain('writeValue: false')
  })

  it('apps/web/src has NO raw OOXML, JSZip, or direct-gateway protection writes', () => {
    // The browser must never construct protection XML or call the
    // gateway's applySheetProtection/applyWorkbookProtection directly —
    // protection travels ONLY through the typed save-plan families.
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    const forbidden = [
      /applySheetProtection/,
      /applyWorkbookProtection/,
      /applyProtectedRanges/,
      /<sheetProtection\b/,
      /<workbookProtection\b/,
      /<protectedRanges\b/,
      /from\s+['"]jszip['"]/,
    ]
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => forbidden.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('office-client carries the typed protection families (no passwords)', () => {
    expect(existsSync(clientPath), `${clientPath} should exist`).toBe(true)
    const content = readFileSync(clientPath, 'utf8')
    expect(content).toContain('sheetProtections?:')
    expect(content).toContain('workbookProtectionState?:')
    // The wire family carries no password field — fail-closed by design.
    const passwordFields = nonCommentLines(content).filter((line) => /password\s*[?:]/i.test(line))
    expect(passwordFields).toEqual([])
  })
})

// ── EXCEL-021: Tables canonical wire-family guards ────────────────────────
//
// Excel tables (ListObjects) are implemented through the CANONICAL save
// families — no new engine path, no browser-side OOXML, no parallel
// table model:
//   - READ: the gateway's xlsx-table-read resolves each worksheet's
//     <tableParts> through its rels into WorksheetState.tables (metadata
//     + PRE-RESOLVED banding colors — theme accents, Excel's HSL tint
//     transform, custom tableStyle dxfs). The browser paints the banding
//     into the cell matrix through the pure table-banding.ts module and
//     registers the VISUAL Univer table with a muted plain theme; a
//     headerless table skips registration (Univer synthesizes headers).
//   - WRITE: Insert → Table journals SheetTableAddition entries and
//     emits `savePlan.tableAdditions` (routeOffice validates with the
//     desktop preload's exact shape: sheetName, 0-based ordered area,
//     name 1-255, columnNames 1-1000 × ≤255, built-in style name,
//     bandedRows; ≤50 entries) → applyCellEditsToXlsx's trailing
//     parameter → applyTableAdditions writes the table part +
//     <tableParts> + rel + [Content_Types] override, failing closed on
//     name collisions, overlaps, and bad column names.
//   - DELETE: convert-to-range for session tables (journal splice —
//     nothing reaches the file; baked cells stay). File-native tables
//     refuse with the desktop's exact message.
//   - FILTER: a sheet whose filter belongs to a table (no worksheet
//     <autoFilter> — the filter lives in the table part) refuses filter
//     commands through a BeforeCommandExecute gate (desktop
//     FILTER_COMMAND_PATTERN parity).
//   - SPLIT-SAVE: new tables + row/column changes on the same save hold
//     the tables back (desktop heldTables parity): phase 1 saves the
//     structure, phase 2 saves the tables alone against the phase-1
//     bytes.
//
// The guards below enforce that no future shortcut re-introduces:
//   - browser-side OOXML/JSZip construction for table parts,
//   - a table write that bypasses the save plan,
//   - the disabled Table stub (the feature must stay wired),
//   - a filter-origin gate that lets table-owned filters be edited.

describe('architecture: EXCEL-021 Tables uses the canonical wire families', () => {
  const editorPath = join(WEB_ROOT, 'src', 'screens', 'ExcelEditor.tsx')
  const ribbonPath = join(WEB_ROOT, 'src', 'screens', 'excel', 'Ribbon.tsx')
  const clientPath = join(WEB_ROOT, 'src', 'api', 'office-client.ts')
  const bandingPath = join(WEB_ROOT, 'src', 'office', 'table-banding.ts')

  it('ExcelEditor emits the tableAdditions save-plan family', () => {
    expect(existsSync(editorPath), `${editorPath} should exist`).toBe(true)
    const content = readFileSync(editorPath, 'utf8')
    expect(content).toContain('tableAddsRef')
    expect(
      /tableAdditions\.length > 0 \? \{ tableAdditions \} : \{\}/.test(content),
      'handleSave must conditionally emit the tableAdditions family',
    ).toBe(true)
  })

  it('ExcelEditor implements the desktop split-save (heldTables) semantics', () => {
    const content = readFileSync(editorPath, 'utf8')
    expect(
      content.includes('heldTables'),
      'row/column changes + new tables must hold the tables into a phase-2 save',
    ).toBe(true)
    expect(content.includes('tableAdditions: heldTables')).toBe(true)
  })

  it('Ribbon Insert → Tables is wired (no disabled stub)', () => {
    expect(existsSync(ribbonPath), `${ribbonPath} should exist`).toBe(true)
    const content = readFileSync(ribbonPath, 'utf8')
    // The old stub documented the missing wire family — it must be gone.
    expect(
      content.includes('does not yet expose the tableAdditions family'),
      'the disabled Table stub must be removed',
    ).toBe(false)
    expect(content).toContain(`label="Table"`)
    expect(content).toContain(`label="Delete Table"`)
    // The buttons must call back into the shell-owned journal handlers.
    expect(content).toContain('tables?.onInsertTable()')
    expect(content).toContain('tables?.onDeleteTable()')
  })

  it('table-banding.ts is a pure value transformation (no imports)', () => {
    expect(existsSync(bandingPath), `${bandingPath} should exist`).toBe(true)
    const content = readFileSync(bandingPath, 'utf8')
    const imports = nonCommentLines(content).filter((line) => /^\s*import\s/.test(line))
    expect(
      imports,
      'table banding must stay import-free — colors arrive pre-resolved from the gateway',
    ).toEqual([])
  })

  it('ExcelEditor implements the table-owned filter-origin gate', () => {
    const content = readFileSync(editorPath, 'utf8')
    expect(
      content.includes('FILTER_COMMAND_PATTERN'),
      'the desktop FILTER_COMMAND_PATTERN gate must exist',
    ).toBe(true)
    expect(content.includes('BeforeCommandExecute')).toBe(true)
    expect(content.includes('filter belongs to an Excel table')).toBe(true)
  })

  it('ExcelEditor implements desktop delete semantics (session-only)', () => {
    const content = readFileSync(editorPath, 'utf8')
    // Convert-to-range: the journal entry is spliced (never persisted).
    expect(content.includes('tableAddsRef.current.splice')).toBe(true)
    // File-native tables refuse with the desktop's exact message.
    expect(content.includes('tables already in the file cannot be deleted yet')).toBe(true)
  })

  it('ExcelEditor seeds file tables from WorksheetState.tables (read path)', () => {
    const content = readFileSync(editorPath, 'utf8')
    expect(content).toContain('tablesFileRef')
    expect(content).toContain('applyTableBandingToMatrix')
    expect(content).toContain('addTableTheme')
    // Headerless tables skip Univer registration (desktop parity).
    expect(content.includes('table.headerRowCount === 0')).toBe(true)
  })

  it('apps/web/src has NO raw OOXML, JSZip, or direct-gateway table writes', () => {
    // The browser must never construct table XML, zip parts, or call the
    // gateway's applyTableAdditions directly — tables travel ONLY through
    // the typed save-plan family.
    const webFiles = readFiles(join(WEB_ROOT, 'src'))
    const forbidden = [
      /applyTableAdditions/,
      /<tableParts\b/,
      /<tableStyleInfo\b/,
      /<tableColumn\b/,
      /from\s+['"]jszip['"]/,
    ]
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => forbidden.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('office-client carries the typed tableAdditions family', () => {
    expect(existsSync(clientPath), `${clientPath} should exist`).toBe(true)
    const content = readFileSync(clientPath, 'utf8')
    expect(content).toContain('tableAdditions?:')
    expect(content).toContain('SheetTableAddition')
  })
})

// ── EXCEL-023 (Insert → Chart / chart edit): the browser is a thin typed
//    client for the chart surface too — it renders the gateway's canonical
//    ChartVisualState and journals through the chartEdits /
//    visualAdditions / visualEdits wire families; ALL chart OOXML work
//    stays in the xlsx-gateway. Charts render through the app's own SVG
//    components floated via Univer's PUBLIC registerComponent +
//    addFloatDomToRange facades (Univer 0.25.1 ships no chart plugin —
//    the desktop's exact rendering architecture).
describe('architecture: EXCEL-023 Charts uses the canonical wire families', () => {
  const webFiles = readFiles(join(WEB_ROOT, 'src'))

  it('apps/web/src has NO chart OOXML, JSZip, or direct-gateway chart writes', () => {
    // (applyCellEditsToXlsx itself is deliberately absent from this list —
    // the file-level comments legitimately name the engine's entry point;
    // the root-import guard below blocks every actual value import.)
    const forbidden = [
      /<c:chartSpace\b/,
      /<c:barChart\b/,
      /<c:ser\b/,
      /xl\/charts\//,
      /relationships\/chart\b/,
      /from\s+['"]jszip['"]/,
      /applyChartEdit/,
      /applyVisualAdditions/,
    ]
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) => forbidden.some((re) => re.test(line)))
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })

  it('apps/web/src chart surfaces use ONLY the canonical typed families', () => {
    const chartsPath = join(WEB_ROOT, 'src', 'office', 'sheet-charts.tsx')
    expect(existsSync(chartsPath), `${chartsPath} should exist`).toBe(true)
    const chartsContent = readFileSync(chartsPath, 'utf8')
    expect(chartsContent).toContain('SheetChartInfo')
    expect(chartsContent).toContain('WorkbookChartEdit')
    expect(
      /import type \{[^}]*SheetChartInfo[^}]*\} from '@genoffice\/xlsx-gateway'/.test(
        chartsContent,
      ),
    ).toBe(true)
    const clientContent = readFileSync(
      join(WEB_ROOT, 'src', 'api', 'office-client.ts'),
      'utf8',
    )
    expect(clientContent).toContain('chartEdits?:')
    expect(clientContent).toContain('WorkbookChartEdit')
  })

  it('apps/web/src/office/sheet-charts.tsx has NO private-Univer-internals access', () => {
    const chartsPath = join(WEB_ROOT, 'src', 'office', 'sheet-charts.tsx')
    expect(existsSync(chartsPath), `${chartsPath} should exist`).toBe(true)
    const lines = readFileSync(chartsPath, 'utf8')
      .split('\n')
      .map((line, index) => ({ line, number: index + 1 }))
      .filter(({ line }) => !line.trim().startsWith('//') && !line.trim().startsWith('*'))
    const violations = lines.filter(({ line }) => /\bas unknown as\b|._image\b/.test(line))
    expect(
      violations.map((v) => `${v.number}: ${v.line.trim()}`),
      'sheet-charts.tsx must use only the public facade surface (structural interfaces, no casts)',
    ).toEqual([])
  })

  it('the gateway pure-domain submodule is the ONLY value import from xlsx-gateway', () => {
    // Value imports from the gateway package ROOT remain forbidden (the
    // engine must never bundle into the browser). EXCEL-023 permits
    // exactly one subpath — the pure domain module the desktop also
    // consumes (apps/sheets/src/domain/chart-visual.ts re-exports it) —
    // which carries no archive/XML/relationship code.
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some((line) =>
        /^import\s+\{[^}]*\}\s+from\s+'@genoffice\/xlsx-gateway[^']*'/.test(line.trim()),
      )
    })
    const offenders = violations.filter(
      (f) =>
        !f.content.includes("from '@genoffice/xlsx-gateway/src/domain/chart-visual.js'") ||
        nonCommentLines(f.content).some(
          (line) =>
            /^import\s+\{[^}]*\}\s+from\s+'@genoffice\/xlsx-gateway'/.test(line.trim()),
        ),
    )
    expect(offenders.map((v) => v.rel)).toEqual([])
    // The pure-domain re-export must exist and carry no engine imports.
    const domainPath = join(WEB_ROOT, 'src', 'office', 'chart-domain.ts')
    expect(existsSync(domainPath), `${domainPath} should exist`).toBe(true)
    const domainContent = readFileSync(domainPath, 'utf8')
    expect(domainContent).toContain('chart-visual.js')
  })

  it('Ribbon Insert → Charts is wired (no disabled stub)', () => {
    const ribbonPath = join(WEB_ROOT, 'src', 'screens', 'excel', 'Ribbon.tsx')
    expect(existsSync(ribbonPath)).toBe(true)
    const content = readFileSync(ribbonPath, 'utf8')
    expect(content).toContain('onInsertChart')
    expect(content).not.toContain('Chart — disabled')
  })

  it('ExcelEditor seeds file charts from WorksheetState.charts (read path) and emits the chartEdits family', () => {
    const editorPath = join(WEB_ROOT, 'src', 'screens', 'ExcelEditor.tsx')
    expect(existsSync(editorPath)).toBe(true)
    const content = readFileSync(editorPath, 'utf8')
    expect(content).toContain('sheet.charts')
    expect(content).toContain('collectChartEdits')
    expect(content).toContain('collectChartAdditions')
    expect(content).toContain('collectChartVisualEdits')
    expect(content).toContain('chartEdits')
  })
})
