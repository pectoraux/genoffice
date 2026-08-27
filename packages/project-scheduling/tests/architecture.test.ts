import { describe, expect, it } from 'vitest'
import pkgRaw from '../package.json?raw'
import architectureLockRaw from '../../../spec/project/architecture-lock.md?raw'
import acrRaw from '../../../spec/project/architecture-changes/ACR-001-project-file-adapter-boundary.md?raw'

/**
 * Roadmap reconciliation increment — architecture-boundary self-scan for
 * `@genoffice/project-scheduling` (the scheduling engine).
 *
 * Documents and enforces the clarified architecture-lock §13 rule (ACR-001):
 * the scheduling engine MUST NOT import `project-file`, MSPDI parser
 * internals, or MPP parser internals — the scheduling engine must not
 * acquire file-format knowledge. File-format implementations remain behind
 * the `project-file` adapter boundary; the scheduling engine's runtime
 * dependencies are exactly the contracts and engine packages (the accepted
 * dependency direction: scheduling → engine → contracts).
 *
 * Mechanism: source-level import-specifier scans over every TypeScript
 * module under `src/` (comments stripped — prose may mention format names,
 * imports are the boundary evidence). This suite uses ONLY vitest + `?raw`
 * module sources, satisfying the same CI foundation boundary grep as every
 * other foundation test.
 */

const srcModules = import.meta.glob('../src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>

const srcFiles = Object.entries(srcModules)

/** Strip comments (prose may mention format names; imports are the evidence). */
function stripComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '')
}

/** Extract every import specifier from a TypeScript source (comment-free). */
function importSpecifiers(source: string): string[] {
  const clean = stripComments(source)
  const out: string[] = []
  const re =
    /(?:^|\n)\s*(?:import|export)\s[^;'"()]*?from\s*['"]([^'"]+)['"]|(?:^|\n)\s*import\s*['"]([^'"]+)['"]/g
  let match: RegExpExecArray | null
  while ((match = re.exec(clean)) !== null) {
    out.push((match[1] ?? match[2]) as string)
  }
  return out
}

const ALLOWED_EXTERNAL = new Set(['@genoffice/project-contracts', '@genoffice/project-engine'])
const FORMAT_INTERNALS_RE = /mspdi|gproj|\bmpp\b|project-file/i

describe('PROJECT-scheduling architecture — the scheduling engine has no file-format knowledge', () => {
  it('ships exactly the contracts + engine runtime dependencies', () => {
    const pkg = JSON.parse(pkgRaw) as {
      dependencies?: Record<string, string>
      devDependencies?: Record<string, string>
    }
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
    expect(Object.keys(pkg.devDependencies ?? {}).sort()).toEqual(['typescript', 'vitest'])
  })

  it('never imports project-file, MSPDI parser internals, or MPP parser internals in src', () => {
    expect(srcFiles.length).toBeGreaterThan(0)
    for (const [file, source] of srcFiles) {
      const specifiers = importSpecifiers(source)
      for (const specifier of specifiers) {
        if (specifier.startsWith('.')) continue
        if (!ALLOWED_EXTERNAL.has(specifier)) {
          throw new Error(
            `${file} imports forbidden module ${specifier} (allowed: ${[...ALLOWED_EXTERNAL].join(', ')})`,
          )
        }
        if (FORMAT_INTERNALS_RE.test(specifier)) {
          throw new Error(`${file} imports file-format knowledge (${specifier})`)
        }
      }
    }
  })

  it('never imports the renderer or host packages (dependency direction)', () => {
    for (const [file, source] of srcFiles) {
      const clean = stripComments(source)
      // Determinism markers are scanned comment-free: doc comments may say
      // "never Date.now()" — only code is boundary evidence.
      expect(clean, `${file} uses Date.now`).not.toContain('Date.now(')
      expect(clean, `${file} uses Math.random`).not.toContain('Math.random(')
      expect(clean, `${file} uses localeCompare`).not.toContain('localeCompare')
      expect(clean, `${file} imports the renderer package`).not.toMatch(
        /from ['"]@genoffice\/project-renderer-core['"]/,
      )
      expect(clean, `${file} imports the MPP host package`).not.toMatch(
        /from ['"]@genoffice\/project-mpp-host['"]/,
      )
      expect(clean, `${file} imports the file adapter package`).not.toMatch(
        /from ['"]@genoffice\/project-file['"]/,
      )
    }
  })

  it('documents the clarified §13 rule (ACR-001) it is governed by', () => {
    expect(architectureLockRaw).toContain(
      'No `project-engine`, `project-scheduling`, `project-renderer-core`, or host package may directly import format-specific parser internals.',
    )
    expect(architectureLockRaw).toContain('Status: FROZEN')
    expect(acrRaw).toContain('## 4. New interpretation')
  })
})
