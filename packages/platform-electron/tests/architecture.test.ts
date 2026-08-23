/**
 * Architecture-boundary test for @genoffice/platform-electron (Increment 3I — AST-based).
 *
 * Uses the TypeScript compiler API (via dep-scanner.ts) to parse source files
 * into an AST and extract actual module specifiers. This GUARANTEES:
 *   - Comments and JSDoc are NOT reported
 *   - String literals that are NOT import specifiers are NOT reported
 *   - ALL import forms ARE detected
 *
 * Enforces:
 *   - ZERO imports of apps/sheets (no upward dependency on the application)
 *   - ZERO imports of @genoffice/sheets-shared (app-layer IPC contract)
 *   - DOES import @genoffice/xlsx-gateway (the canonical gateway package)
 *   - DOES import @genoffice/runtime-contracts
 */
import { describe, test, expect } from 'vitest'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'
import { scanForForbiddenImports } from './dep-scanner.js'

const SRC = join(__dirname, '..', 'src')
const ENGINE_SRC = join(SRC, 'capabilities', 'electron-xlsx-sidecar-engine.ts')
const VALIDATORS_SRC = join(SRC, 'capabilities', 'sidecar-validators.ts')

describe('@genoffice/platform-electron architecture boundary (Increment 3I — AST-based)', () => {
  test('ZERO imports of apps/sheets (no upward dependency on the application)', () => {
    const hits = scanForForbiddenImports(SRC, [/apps\/sheets/])
    if (hits.length > 0) {
      console.error('Found apps/sheets imports in platform-electron source:')
      for (const h of hits) {
        console.error(`  ${h.file}:${h.line}: ${h.kind} '${h.specifier}'`)
      }
    }
    expect(hits).toEqual([])
  })

  test('ZERO imports of @genoffice/sheets-shared (app-layer IPC contract)', () => {
    const hits = scanForForbiddenImports(SRC, ['@genoffice/sheets-shared'])
    expect(hits).toEqual([])
  })

  test('DOES import @genoffice/xlsx-gateway (the canonical gateway package)', () => {
    const hits = scanForForbiddenImports(SRC, ['@genoffice/xlsx-gateway'])
    expect(hits.length).toBeGreaterThan(0)
  })

  test('DOES import @genoffice/runtime-contracts', () => {
    const hits = scanForForbiddenImports(SRC, ['@genoffice/runtime-contracts'])
    expect(hits.length).toBeGreaterThan(0)
  })

  // ═══ INCREMENT 15A — Engine contract hardening ═══
  //
  // The engine's `readPivotDefinition()` method must:
  //   1. Declare a typed `Promise<WorkbookPivotDefinition>` return
  //      (NOT `Promise<string>` for raw XML, NOT `Promise<unknown>`).
  //   2. Use ZERO unchecked `as Record` / `as Array` / `as unknown as`
  //      casts on the sidecar response. All response validation goes
  //      through type guards in `sidecar-validators.ts`.
  //   3. Generate the work directory ONCE via `mkdtempSync` and clean
  //      up exactly that path in `finally` (no race between two
  //      randomUUID() calls).
  //   4. The generic `readArchiveEntry()` method MUST NOT exist on the
  //      engine contract — it was a ZIP-entry escape-hatch that
  //      invited callers above the engine boundary to pluck arbitrary
  //      OOXML parts. The contract now exposes only the Sheets-specific
  //      `readPivotDefinition()`.

  test('ElectronXlsxSidecarEngine declares readPivotDefinition (NOT readArchiveEntry) — Increment 15A', () => {
    const src = readFileSync(ENGINE_SRC, 'utf8')
    expect(src).toMatch(/async readPivotDefinition\s*\(/)
    // The generic readArchiveEntry method MUST NOT exist.
    expect(src).not.toMatch(/async readArchiveEntry\s*\(/)
  })

  test('ElectronXlsxSidecarEngine.readPivotDefinition returns Promise<WorkbookPivotDefinition> — Increment 15A', () => {
    const src = readFileSync(ENGINE_SRC, 'utf8')
    expect(src).toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/m)
    // NOT a Promise<string> (raw XML) or Promise<unknown>.
    expect(src).not.toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<string>/m)
    expect(src).not.toMatch(/readPivotDefinition[\s\S]*?:\s*Promise<unknown>/m)
  })

  test('ElectronXlsxSidecarEngine readPivotDefinition has ZERO unchecked as casts on sidecar response — Increment 15A', () => {
    // The sidecar response arrives as `unknown`. The engine impl MUST
    // validate it via type guards (in `validateReadEntriesResponse`)
    // before accessing any field. ZERO unchecked `as Record` / `as Array`
    // / `as unknown as` casts on the response.
    const src = readFileSync(ENGINE_SRC, 'utf8')
    const stripped = src
      .replace(/\/\*\*?[\s\S]*?\*\//g, '') // block comments
      .replace(/^\s*\/\/.*$/gm, '') // line comments
    // The validator-based path uses type guards, not `as` casts.
    // The previous readArchiveEntry impl used `raw as Record<string, unknown>`
    // and `obj.entries as Array<Record<string, unknown>>` — both forbidden.
    expect(stripped).not.toMatch(/\bas\s+Record<string,\s*unknown>\b/)
    expect(stripped).not.toMatch(/\bas\s+Array<Record<string,\s*unknown>>\b/)
    expect(stripped).not.toMatch(/\bas\s+unknown\s+as\b/)
    expect(stripped).not.toMatch(/\bas\s+any\b/)
  })

  test('ElectronXlsxSidecarEngine readPivotDefinition generates workDir once via mkdtempSync — Increment 15A', () => {
    // The previous impl used `mkdtempSync(tmpdir(), ...)` for the workDir
    // but a different `randomUUID()` for the input/output file paths — a
    // race that left orphaned temp dirs on failure paths. The new impl
    // generates ONE workDir via mkdtempSync and uses it for both the
    // sidecar outputDir AND the cleanup in finally.
    const src = readFileSync(ENGINE_SRC, 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    // readPivotDefinition body must contain mkdtempSync (workDir creation).
    const readPivotBody = stripped.match(/async readPivotDefinition[\s\S]*?^  }/m)?.[0] ?? ''
    expect(readPivotBody.length).toBeGreaterThan(0)
    expect(readPivotBody).toMatch(/mkdtempSync/)
    // The cleanup in finally must rmSync the SAME workDir.
    expect(readPivotBody).toMatch(/rmSync\(workDir/)
  })

  test('ElectronXlsxSidecarEngine readPivotDefinition delegates parsing to parsePivotDefinition from xlsx-gateway — Increment 15A', () => {
    // The engine is the SINGLE translation point between OOXML wire format
    // and the runtime-independent WorkbookPivotDefinition contract. It
    // must call the canonical parser from xlsx-gateway (no duplication).
    const src = readFileSync(ENGINE_SRC, 'utf8')
    expect(src).toMatch(/parsePivotDefinition\s*\(/)
    // The import must be present.
    expect(src).toMatch(/parsePivotDefinition/)
  })

  test('sidecar-validators exports validateReadEntriesResponse — Increment 15A', () => {
    // The validator for the sidecar read_entries response lives in
    // sidecar-validators.ts (alongside the other response validators).
    // It uses type guards (ZERO `as` casts) and returns a typed
    // `ReadArchiveEntriesResult`.
    const src = readFileSync(VALIDATORS_SRC, 'utf8')
    expect(src).toMatch(/export function validateReadEntriesResponse/)
    expect(src).toMatch(/export interface ReadArchiveEntriesResult/)
    expect(src).toMatch(/export interface ReadArchiveEntry/)
  })

  test('sidecar-validators validateReadEntriesResponse has ZERO unchecked as casts — Increment 15A', () => {
    const src = readFileSync(VALIDATORS_SRC, 'utf8')
    const stripped = src.replace(/\/\*\*?[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
    // The validator body must use type guards, not `as` casts.
    const validatorBody = stripped.match(
      /export function validateReadEntriesResponse[\s\S]*?^}/m,
    )?.[0] ?? ''
    expect(validatorBody.length).toBeGreaterThan(0)
    expect(validatorBody).not.toMatch(/\bas\s+Record<string,\s*unknown>\b/)
    expect(validatorBody).not.toMatch(/\bas\s+Array<Record<string,\s*unknown>>\b/)
    expect(validatorBody).not.toMatch(/\bas\s+unknown\s+as\b/)
    expect(validatorBody).not.toMatch(/\bas\s+any\b/)
  })
})
