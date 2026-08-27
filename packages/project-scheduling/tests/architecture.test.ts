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

// ===========================================================================
// PROJECT-026 — the single canonical allocation kernel.
//
// The Principal Architect's review of PR #28 (CHANGES REQUIRED): a
// higher-level semantic concept must have ONE canonical authority — the
// resource demand/capacity semantics (demand contribution, demand-interval
// construction, resource-calendar resolution, availability-capacity
// resolution, calendar-aware segmentation, over-allocation predicate) may be
// IMPLEMENTED only in `src/allocation-kernel.ts`. Both
// `resourceAllocations()` (allocation.ts) and `levelResources()` (leveling.ts)
// must CONSUME that kernel; neither may carry an independent implementation
// of the shared rules. These guards fail if the leveler (or the allocation
// projection) reintroduces an independent implementation.
// ===========================================================================

describe('PROJECT-026 architecture — the single canonical allocation kernel (one authority)', () => {
  const kernel = srcModules['../src/allocation-kernel.ts'] ?? ''
  const leveling = srcModules['../src/leveling.ts'] ?? ''
  const allocation = srcModules['../src/allocation.ts'] ?? ''
  const nonKernelModules = srcFiles.filter(([file]) => file !== '../src/allocation-kernel.ts')

  it('the kernel module defines the shared allocation semantics', () => {
    // The named semantic primitives exist exactly here.
    expect(kernel).toContain('export const isLeaf')
    expect(kernel).toContain('export const contributesToDemand')
    expect(kernel).toContain('export const effectiveMaxUnits')
    expect(kernel).toContain('export const demandIntervalsForResource')
    expect(kernel).toContain('export const resourceCalendarFor')
    expect(kernel).toContain('export const allocationSegments')
    // The kernel owns the over-allocation predicate — the single
    // classification both consumers echo (never re-derive).
    expect(kernel).toContain('overallocated: demand > capacity')
  })

  it('no other module defines an independent implementation of the shared allocation rules', () => {
    // The kernel's semantic primitives may be DEFINED only in the kernel —
    // a second definition anywhere else in this package (e.g. a leveler
    // that reintroduces its own capacity sweep) fails here.
    const kernelPrimitives = [
      'isLeaf',
      'contributesToDemand',
      'effectiveMaxUnits',
      'effectiveCapacity',
      'demandIntervalsForResource',
      'assignmentIntervalsForResource',
      'resourceCalendarFor',
      'allocationSegments',
      'allocationSegmentsForResource',
    ]
    for (const [file, source] of nonKernelModules) {
      const clean = stripComments(source)
      for (const name of kernelPrimitives) {
        expect(
          clean,
          `${file} defines the kernel primitive ${name} outside allocation-kernel.ts`,
        ).not.toMatch(new RegExp(`(?:const|function)\\s+${name}\\s*[(=]`))
      }
    }
  })

  it('the leveler consumes the kernel (no independent capacity sweep survives in leveling.ts)', () => {
    const clean = stripComments(leveling)
    // The leveler builds its conflicts from the kernel's tiling and its
    // demand intervals from the kernel's construction.
    expect(clean).toContain("from './allocation-kernel.js'")
    expect(clean).toContain('allocationSegments(')
    expect(clean).toContain('demandIntervalsForResource(')
    expect(clean).toContain('resourceCalendarFor(')
    // The sweep's structural primitives no longer exist in the leveler: an
    // independent capacity sweep would need boundary collection and
    // working-time evaluation. (nextWorkingInstant/addWorkingTime stay free
    // — delay PLACEMENT is leveling policy, not allocation semantics.)
    expect(clean, 'leveling.ts collects sweep boundaries').not.toContain('boundaries')
    expect(clean, 'leveling.ts evaluates working time').not.toContain('workingIntervals')
    expect(clean, 'leveling.ts evaluates working status').not.toContain('isWorking')
    expect(clean, 'leveling.ts resolves capacity itself').not.toContain('effectiveMaxUnits')
    expect(clean, 'leveling.ts re-implements the demand rule').not.toContain('contributesToDemand')
  })

  it('the allocation projection consumes the kernel (no independent capacity sweep survives in allocation.ts)', () => {
    const clean = stripComments(allocation)
    expect(clean).toContain("from './allocation-kernel.js'")
    expect(clean).toContain('demandIntervalsForResource(')
    expect(clean).toContain('allocationSegments(')
    expect(clean).toContain('resourceCalendarFor(')
    expect(clean, 'allocation.ts collects sweep boundaries').not.toContain('boundaries')
    expect(clean, 'allocation.ts evaluates working time').not.toContain('workingIntervals')
    expect(clean, 'allocation.ts evaluates working status').not.toContain('isWorking')
    expect(clean, 'allocation.ts resolves capacity itself').not.toContain('effectiveMaxUnits')
    expect(clean, 'allocation.ts re-implements the demand rule').not.toContain(
      'contributesToDemand',
    )
  })

  it('the kernel stays a lower layer than its consumers (no import cycles)', () => {
    const clean = stripComments(kernel)
    expect(clean).not.toContain("from './leveling.js'")
    expect(clean).not.toContain("from './allocation.js'")
  })
})
