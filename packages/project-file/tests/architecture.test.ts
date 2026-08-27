import { describe, expect, it } from 'vitest'
import ownPkgRaw from '../package.json?raw'
import enginePkgRaw from '../../project-engine/package.json?raw'
import schedulingPkgRaw from '../../project-scheduling/package.json?raw'
import rendererPkgRaw from '../../project-renderer-core/package.json?raw'
import architectureLockRaw from '../../../spec/project/architecture-lock.md?raw'
import acrRaw from '../../../spec/project/architecture-changes/ACR-001-project-file-adapter-boundary.md?raw'

/**
 * Roadmap reconciliation increment — architecture-boundary guards for the
 * clarified architecture-lock §13 rule (ACR-001):
 *
 *   project-file MAY contain:
 *     .gproj implementation, MSPDI implementation, MPP adapter implementation
 *
 *   project-engine / project-scheduling / project-renderer-core MUST NOT
 *   import:
 *     project-file, MSPDI parser internals, MPP parser internals
 *
 * The renderer must not acquire file-format knowledge. The scheduling engine
 * must not acquire file-format knowledge. The domain engine must not acquire
 * file-format knowledge. The host package may import the project-file PUBLIC
 * adapter surface only — never deep parser internals.
 *
 * Mechanism: source-level import-specifier scans over every TypeScript module
 * under `src/` of every foundation package (cross-package vitest `?raw` globs —
 * no runtime package, no `node:` imports, satisfying the same CI boundary
 * grep as every other foundation test). Prose comments may legitimately
 * mention format names; only import statements are boundary evidence, so
 * comments are stripped before scanning.
 */

const engineSources = import.meta.glob('../../project-engine/src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>
const schedulingSources = import.meta.glob('../../project-scheduling/src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>
const rendererSources = import.meta.glob('../../project-renderer-core/src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>
const hostSources = import.meta.glob('../../project-mpp-host/src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>
const ownSources = import.meta.glob('../src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>

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

function externalSpecifiers(sources: Record<string, string>): Record<string, string[]> {
  const byFile: Record<string, string[]> = {}
  for (const [file, source] of Object.entries(sources)) {
    byFile[file] = importSpecifiers(source).filter((s) => !s.startsWith('.'))
  }
  return byFile
}

const FORMAT_INTERNALS_RE = /mspdi|gproj|\bmpp\b/i

describe('ACR-001 boundary — project-file is the sanctioned file-adapter boundary', () => {
  it('project-file MAY contain the .gproj, MSPDI, and MPP adapter implementations', () => {
    const files = Object.keys(ownSources)
    // The `.gproj` implementation is the top-level adapter/serialize/
    // deserialize/envelope module family (PROJECT-014).
    expect(
      files.some((f) => /src\/serialize\.ts$/.test(f)),
      'no .gproj serializer implementation',
    ).toBe(true)
    expect(
      files.some((f) => /src\/deserialize\.ts$/.test(f)),
      'no .gproj deserializer implementation',
    ).toBe(true)
    // The MSPDI implementation is src/mspdi (PROJECT-015/016).
    expect(
      files.some((f) => /src\/mspdi\/xml-parser\.ts$/.test(f)),
      'no MSPDI parser implementation',
    ).toBe(true)
    expect(
      files.some((f) => /src\/mspdi\/xml-writer\.ts$/.test(f)),
      'no MSPDI serializer implementation',
    ).toBe(true)
    // The MPP adapter implementation is src/mpp (PROJECT-018).
    expect(
      files.some((f) => /src\/mpp\//.test(f)),
      'no MPP adapter implementation',
    ).toBe(true)
    // The implementations are behind the public adapter surface.
    const index = ownSources['../src/index.ts'] ?? ''
    expect(index).toContain("from './adapter.js'")
    expect(index).toContain("from './mspdi/index.js'")
    expect(index).toContain("from './mpp/index.js'")
  })

  it('project-file declares exactly the accepted runtime dependencies (no external parser libraries)', () => {
    const pkg = JSON.parse(ownPkgRaw) as { dependencies?: Record<string, string> }
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
  })

  it('the format implementations are in-package pure TypeScript (no external parser/process imports)', () => {
    for (const [file, specifiers] of Object.entries(externalSpecifiers(ownSources))) {
      for (const specifier of specifiers) {
        if (specifier.startsWith('@genoffice/')) continue
        throw new Error(`${file} imports non-foundation module ${specifier}`)
      }
    }
  })
})

describe('ACR-001 boundary — semantic/runtime packages MUST NOT import file-format knowledge', () => {
  const forbiddenFor = (label: string, sources: Record<string, string>): void => {
    for (const [file, specifiers] of Object.entries(externalSpecifiers(sources))) {
      for (const specifier of specifiers) {
        if (
          specifier === '@genoffice/project-contracts' ||
          specifier === '@genoffice/project-engine'
        ) {
          continue
        }
        throw new Error(`${label}: ${file} imports forbidden module ${specifier}`)
      }
      // The file-format-internals prohibition, independent of the allowlist:
      // even a hypothetical allowed-relative or deep-path import carrying
      // MSPDI/MPP/.gproj internals is rejected.
      const all = importSpecifiers(sources[file] ?? '')
      for (const specifier of all) {
        if (FORMAT_INTERNALS_RE.test(specifier)) {
          throw new Error(`${label}: ${file} imports file-format internals (${specifier})`)
        }
      }
    }
  }

  it('project-engine never imports project-file, MSPDI parser internals, or MPP parser internals', () => {
    expect(Object.keys(engineSources).length).toBeGreaterThan(0)
    forbiddenFor('project-engine', engineSources)
  })

  it('project-scheduling never imports project-file, MSPDI parser internals, or MPP parser internals', () => {
    expect(Object.keys(schedulingSources).length).toBeGreaterThan(0)
    forbiddenFor('project-scheduling', schedulingSources)
  })

  it('project-renderer-core never imports project-file, MSPDI parser internals, or MPP parser internals', () => {
    expect(Object.keys(rendererSources).length).toBeGreaterThan(0)
    forbiddenFor('project-renderer-core', rendererSources)
  })

  it('the foundation runtime dependency sets are exactly the accepted edges', () => {
    const enginePkg = JSON.parse(enginePkgRaw) as { dependencies?: Record<string, string> }
    expect(Object.keys(enginePkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
    ])
    const schedulingPkg = JSON.parse(schedulingPkgRaw) as { dependencies?: Record<string, string> }
    expect(Object.keys(schedulingPkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
    const rendererPkg = JSON.parse(rendererPkgRaw) as { dependencies?: Record<string, string> }
    expect(Object.keys(rendererPkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
  })
})

describe('ACR-001 boundary — the host package may use the project-file PUBLIC surface only', () => {
  it('project-mpp-host imports the public adapter surface, never deep parser internals', () => {
    expect(Object.keys(hostSources).length).toBeGreaterThan(0)
    for (const [file, specifiers] of Object.entries(externalSpecifiers(hostSources))) {
      for (const specifier of specifiers) {
        if (specifier.startsWith('@genoffice/')) {
          if (specifier.startsWith('@genoffice/project-file/')) {
            throw new Error(`${file} imports project-file parser internals (${specifier})`)
          }
          continue
        }
        if (specifier.startsWith('node:')) continue
        throw new Error(`${file} imports unexpected module ${specifier}`)
      }
    }
  })
})

describe('ACR-001 boundary — the clarified lock and the change record stay in lockstep', () => {
  it('architecture-lock §13 carries the clarified rule (ACR-001 referenced, still FROZEN)', () => {
    expect(architectureLockRaw).toContain(
      'Foundation semantic/runtime packages (`project-contracts`, `project-engine`, `project-scheduling`, `project-renderer-core`) must not import external MSPDI/MPP parser implementations.',
    )
    expect(architectureLockRaw).toContain(
      '`packages/project-file` is the sanctioned file-adapter boundary and may contain format-specific parser/serializer implementations.',
    )
    expect(architectureLockRaw).toContain(
      'No `project-engine`, `project-scheduling`, `project-renderer-core`, or host package may directly import format-specific parser internals.',
    )
    expect(architectureLockRaw).toContain(
      'File-format implementations remain behind the `project-file` adapter boundary.',
    )
    expect(architectureLockRaw).toContain('ACR-001')
    expect(architectureLockRaw).toContain('Status: FROZEN')
  })

  it('ACR-001 records the normative rule and the Principal Architect approval reference', () => {
    expect(acrRaw).toContain('## 9. Principal Architect approval reference')
    expect(acrRaw).toContain(
      '`packages/project-file` is the sanctioned file-adapter boundary and may contain format-specific parser/serializer implementations.',
    )
  })
})
