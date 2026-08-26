/**
 * Architecture-boundary test for @genoffice/xlsx-gateway (Increment 3I — AST-based).
 *
 * Uses the TypeScript compiler API (via dep-scanner.ts) to parse source files
 * into an AST and extract actual module specifiers. This GUARANTEES:
 *   - Comments and JSDoc are NOT reported
 *   - String literals that are NOT import specifiers are NOT reported
 *   - ALL import forms ARE detected
 *
 * Enforces ZERO imports of:
 *   - node:* (ANY Node builtin — the package must be pure)
 *   - electron
 *   - child_process
 *   - apps/sheets (no upward dependency on the application)
 *   - @genoffice/platform-electron
 *   - @genoffice/platform
 *   - @genoffice/runtime-contracts
 *   - BrowserWindow / WebContents / ipcMain / ipcRenderer tokens
 *   - sidecar / ArchiveClient / saveWorkbookViaSidecar tokens
 */
import { describe, test, expect } from 'vitest'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'
import { scanForForbiddenImports, listSourceFiles } from './dep-scanner.js'

const SRC = join(__dirname, '..', 'src')

function scanForTokens(
  rootDir: string,
  forbidden: string[],
): Array<{ file: string; line: number; text: string }> {
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

describe('@genoffice/xlsx-gateway architecture boundary (Increment 3I — AST-based)', () => {
  test('ZERO imports of node:* (the package must be pure — no Node builtins)', () => {
    const hits = scanForForbiddenImports(SRC, [/^node:/])
    if (hits.length > 0) {
      console.error('Found node:* imports in xlsx-gateway source:')
      for (const h of hits) {
        console.error(`  ${h.file}:${h.line}: ${h.kind} '${h.specifier}'`)
      }
    }
    expect(hits).toEqual([])
  })

  test('ZERO imports of electron', () => {
    const hits = scanForForbiddenImports(SRC, ['electron'])
    expect(hits).toEqual([])
  })

  test('ZERO imports of child_process', () => {
    const hits = scanForForbiddenImports(SRC, ['child_process'])
    expect(hits).toEqual([])
  })

  test('ZERO imports of apps/sheets (no upward dependency on the application)', () => {
    const hits = scanForForbiddenImports(SRC, [
      /apps\/sheets/,
      /\.\.\/\.\.\/apps\/sheets/,
      /\.\.\/\.\.\/\.\.\/apps\/sheets/,
    ])
    if (hits.length > 0) {
      console.error('Found apps/sheets imports in xlsx-gateway source:')
      for (const h of hits) {
        console.error(`  ${h.file}:${h.line}: ${h.kind} '${h.specifier}'`)
      }
    }
    expect(hits).toEqual([])
  })

  test('ZERO imports of @genoffice/platform-electron', () => {
    const hits = scanForForbiddenImports(SRC, ['@genoffice/platform-electron'])
    expect(hits).toEqual([])
  })

  test('ZERO imports of @genoffice/platform', () => {
    const hits = scanForForbiddenImports(SRC, ['@genoffice/platform'])
    expect(hits).toEqual([])
  })

  test('ZERO imports of @genoffice/runtime-contracts (gateway is below the contract layer)', () => {
    const hits = scanForForbiddenImports(SRC, ['@genoffice/runtime-contracts'])
    expect(hits).toEqual([])
  })

  test('ZERO references to BrowserWindow / WebContents / wcId / ipcMain / ipcRenderer (token scan)', () => {
    const hits = scanForTokens(SRC, [
      'BrowserWindow',
      'WebContents',
      'wcId',
      'ipcMain',
      'ipcRenderer',
    ]).filter(
      (h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'),
    )
    expect(hits).toEqual([])
  })

  test('ZERO references to sidecar / ArchiveClient / saveWorkbookViaSidecar (token scan)', () => {
    const hits = scanForTokens(SRC, [
      'ArchiveClient',
      'saveWorkbookViaSidecar',
      'readArchiveEntryText',
      'sidecar',
    ]).filter(
      (h) => !h.text.startsWith('*') && !h.text.startsWith('//') && !h.text.startsWith('/*'),
    )
    expect(hits).toEqual([])
  })

  test('DOES import jszip (the in-memory archive library)', () => {
    const hits = scanForForbiddenImports(SRC, ['jszip'])
    expect(hits.length).toBeGreaterThan(0)
  })

  test('DOES export planCellEditsToXlsx (the canonical planner)', () => {
    const gatewayFile = join(SRC, 'gateway', 'xlsx-gateway.ts')
    const text = readFileSync(gatewayFile, 'utf8')
    expect(text).toContain('export async function planCellEditsToXlsx')
  })
})
