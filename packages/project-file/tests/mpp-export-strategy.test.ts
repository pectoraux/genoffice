/**
 * PROJECT-019 — MPP export strategy evidence/discipline suite.
 *
 * PROJECT-019 is rescoped to a strategy/product-decision increment (019A):
 * the deliverable is `spec/project/mpp-export-strategy.md` plus synchronized
 * spec updates — NO production MPP writer, NO MPP export API, NO commercial
 * SDK dependency, NO COM bridge, and zero changes to accepted semantics.
 * This suite enforces exactly that, and validates the report's structure
 * against the PROJECT-019 brief:
 *
 *   - the report exists and carries all 19 mandated sections;
 *   - it ends with exactly one of the five allowed strategy decisions from
 *     the brief's A–E outcome set (the chosen one being outcome E — MPP
 *     export deferred, MSPDI/.gproj is the supported write path);
 *   - it documents the required content: candidate strategies (all ten
 *     axes), runtime matrix, licensing matrix, security analysis, fidelity
 *     tiers A–D with an explicit statement of what "MPP export supported"
 *     would mean, determinism analysis incl. the round-trip oracle, the
 *     MSPDI-intermediary/COM/commercial-SDK analyses, rejected
 *     alternatives, the 019A/019B/019C rescope with the blocked-until-
 *     approval rule, and numbered source links for external claims;
 *   - NO production MPP writer/export API was introduced: the project-file
 *     package's runtime dependencies are exactly contracts + engine, no
 *     commercial SDK/COM package appears anywhere in project-file or the
 *     MPP host package, and the mspdi/mpp adapter surfaces are unchanged
 *     in shape (MSPDI export remains the sole sanctioned export path);
 *   - the spec set is in lockstep (requirements, work-items, dependency
 *     graph, verification matrix all carry the PROJECT-019 rescope);
 *   - the frozen architecture lock still forbids foundation packages from
 *     importing MPP parser implementation code (the investigation may not
 *     silently weaken it).
 *
 * The dynamic claims themselves (live web sources re-verified for this
 * increment) were read as primary sources and are recorded in the report
 * §19 [E1]–[E20] — they are intentionally NOT re-fetched here (no
 * network-dependent, non-hermetic tests in CI).
 */
import { describe, expect, it } from 'vitest'
import report from '../../../spec/project/mpp-export-strategy.md?raw'
import packageJson from '../package.json?raw'
import hostPackageJson from '../../../packages/project-mpp-host/package.json?raw'
import indexSource from '../src/index.ts?raw'
import mppIndexSource from '../src/mpp/index.ts?raw'
import mspdiIndexSource from '../src/mspdi/index.ts?raw'
import architectureLock from '../../../spec/project/architecture-lock.md?raw'
import requirements from '../../../spec/project/requirements.md?raw'
import workItems from '../../../spec/project/work-items.md?raw'
import dependencyGraph from '../../../spec/project/dependency-graph.md?raw'
import verificationMatrix from '../../../spec/project/verification-matrix.md?raw'
import feasibilityReport from '../../../spec/project/mpp-feasibility.md?raw'

const REQUIRED_SECTIONS = [
  '## 1. Executive conclusion',
  '## 2. Current PROJECT-017 evidence',
  '## 3. Candidate strategies',
  '## 4. Runtime matrix',
  '## 5. Licensing matrix',
  '## 6. Security analysis',
  '## 7. Fidelity matrix',
  '## 8. Determinism analysis',
  '## 9. MSPDI intermediary analysis',
  '## 10. COM analysis',
  '## 11. Commercial SDK analysis',
  '## 12. Browser/server implications',
  '## 13. Product/distribution implications',
  '## 14. Architecture impact',
  '## 15. Recommended strategy',
  '## 16. Explicit rejected alternatives',
  '## 17. Proposed rescope of PROJECT-019',
  '## 18. Open questions',
  '## 19. Evidence / source links',
]

const ALLOWED_DECISIONS = [
  'MPP EXPORT FEASIBLE — APPROVED EXTERNAL CONVERTER STRATEGY',
  'MPP EXPORT FEASIBLE — COMMERCIAL SDK STRATEGY',
  'MPP EXPORT FEASIBLE — LIMITED SUBSET',
  'MPP EXPORT NOT FEASIBLE UNDER CURRENT ARCHITECTURE',
  'MPP EXPORT DEFERRED — MSPDI/.gproj IS THE SUPPORTED WRITE PATH',
]

describe('PROJECT-019 — strategy report structure', () => {
  it('the report file exists and is substantial', () => {
    expect(report.length).toBeGreaterThan(10_000)
  })

  it('contains all 19 mandated sections in order', () => {
    let position = -1
    for (const section of REQUIRED_SECTIONS) {
      const next = report.indexOf(section)
      expect(next).toBeGreaterThan(position) // present AND ordered
      position = next
    }
  })

  it('ends with exactly one allowed strategy decision — outcome E (deferred)', () => {
    const decisions = ALLOWED_DECISIONS.filter((d) => report.includes(`**${d}**`))
    expect(decisions).toEqual(['MPP EXPORT DEFERRED — MSPDI/.gproj IS THE SUPPORTED WRITE PATH'])
    const decisionLine = report.split('\n').find((l) => l.includes('STRATEGY DECISION:'))
    expect(decisionLine).toBeDefined()
    expect(decisionLine).toContain('MPP EXPORT DEFERRED — MSPDI/.gproj IS THE SUPPORTED WRITE PATH')
  })

  it('documents all ten candidate strategy axes the brief mandates', () => {
    for (const marker of [
      'MPXJ', // #1 extend the existing sidecar
      'MSPDI intermediary', // #2 intermediary + third-party converter
      'Commercial SDK', // #3 Aspose
      'COM automation', // #4 Microsoft Project COM
      'Windows-only automation', // #5 desktop-only scoping
      'Server-side conversion service', // #6 cloud/server conversion
      'Native in-house MPP writer', // #7 native implementation
      'Limited MPP subset', // #8 subset writer
      'User-installed Microsoft Project workflow', // #9 user-driven save-as
      'MSPDI/.gproj-only supported write model', // #10 the recommended model
    ]) {
      expect(report).toContain(marker)
    }
  })

  it('carries the mandated runtime matrix shape', () => {
    // (prettier pads the table cells — normalize each row's cells before
    // comparing, so the assertion is reflow-stable)
    const cells = (line: string): string[] =>
      line
        .split('|')
        .slice(1, -1)
        .map((c) => c.trim())
    const header = report
      .split('\n')
      .find((l) => cells(l).length === 7 && cells(l)[0] === 'Strategy')
    expect(header).toBeDefined()
    expect(cells(header as string)).toEqual([
      'Strategy',
      'Desktop',
      'Web',
      'Server',
      'Offline',
      'Cross-platform',
      'License',
    ])
    // The matrix must cover the decisive rows:
    for (const row of [
      'Aspose.Tasks export sidecar',
      'Microsoft Project COM automation',
      'Cloud conversion service',
      'Status quo',
    ]) {
      expect(report).toContain(row)
    }
  })

  it('carries the licensing matrix with the commercial figures', () => {
    for (const marker of [
      'GNU LGPL',
      'US$2,397', // current Aspose Developer OEM first-year price
      'US$1,797', // current OEM renewal price
      'Per-seat Microsoft Project license', // COM licensing reality
      'EULA',
    ]) {
      expect(report).toContain(marker)
    }
  })

  it('documents the export fidelity tiers and what MPP export would mean', () => {
    for (const tier of ['Tier A', 'Tier B', 'Tier C', 'Tier D']) {
      expect(report).toContain(tier)
    }
    expect(report).toContain('What "MPP export supported" would mean')
    // The MSPDI column is grounded in accepted evidence; the MPP column is
    // explicitly conditional/unproven:
    expect(report).toContain('conditional and unproven')
  })

  it('documents the determinism analysis incl. the round-trip oracle', () => {
    expect(report).toContain('round-trip oracle')
    expect(report).toContain('unproven') // SDK writer determinism
    // COM non-determinism + the no-suppression finding:
    expect(report).toContain('save timestamps')
    // The oracle is the mandated instrument for any future licensed spike:
    expect(report).toContain('ProjectDocument → MSPDI → converter → .mpp')
  })

  it('documents the security analysis markers', () => {
    for (const marker of [
      'Process isolation',
      'network',
      'data-exfiltration boundary',
      'metering service', // Aspose metered licensing phones home
      'closed source', // cannot statically audit the commercial writer
    ]) {
      expect(report).toContain(marker)
    }
  })

  it('documents the COM analysis per the brief', () => {
    for (const marker of [
      'Windows-only',
      'Microsoft Project installation requirement',
      'server-side Automation of Office', // the KB 257757 prohibition
      'KB 257757',
      'PjFileFormat', // the interop save-format enumeration
      'pjMPP',
    ]) {
      expect(report).toContain(marker)
    }
  })

  it('documents the commercial SDK analysis per the brief', () => {
    for (const marker of [
      'licensed versions', // MPP save gated on a paid license
      'Developer Small Business',
      'Developer OEM',
      'Metered',
      'Long-term viability',
      'closed-source',
    ]) {
      expect(report).toContain(marker)
    }
  })

  it('documents the MSPDI intermediary analysis without lossless claims', () => {
    expect(report).toContain('cannot be called lossless')
    expect(report).toContain('interchange endpoint')
  })

  it('explicitly rejects the forbidden alternatives (numbered)', () => {
    expect(report).toContain('1. **In-house MPP writer')
    expect(report).toContain('Aspose.Tasks as a default dependency')
    expect(report).toContain('Microsoft Project COM automation as a product feature')
    expect(report).toContain('A fake "export test" suite')
  })

  it('proposes the rescope with 019A/019B/019C and the blocked-until-approval rule', () => {
    for (const marker of ['PROJECT-019A', 'PROJECT-019B', 'PROJECT-019C']) {
      expect(report).toContain(marker)
    }
    expect(report).toContain('BLOCKED')
    expect(report).toContain('this document proposes, it does not decide')
  })

  it('carries a source-link section with numbered external evidence', () => {
    expect(report).toContain('[E1] MPXJ FAQ')
    expect(report).toContain('[E20]')
    expect(report).toContain('https://www.mpxj.org/faq/')
    expect(report).toContain('https://support.microsoft.com/')
    expect(report).toContain('https://learn.microsoft.com/')
    expect(report).toContain('https://purchase.aspose.com/pricing/tasks/net')
    expect(report).toContain('https://docs.aspose.com/tasks/net/licensing/')
  })
})

describe('PROJECT-019 — no production MPP writer / no export API introduced', () => {
  it('the project-file package declares exactly the accepted runtime dependencies', () => {
    const pkg = JSON.parse(packageJson) as {
      dependencies?: Record<string, string>
      devDependencies?: Record<string, string>
    }
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
    // No MPP-writer/commercial-SDK/COM package may appear even as a
    // devDependency of the foundation package.
    const allDeps = [
      ...Object.keys(pkg.dependencies ?? {}),
      ...Object.keys(pkg.devDependencies ?? {}),
    ]
    expect(allDeps.filter((d) => /mpp|mpxj|aspose|com|interop|writer/i.test(d))).toEqual([])
  })

  it('the MPP host package carries no commercial SDK / COM dependency', () => {
    const pkg = JSON.parse(hostPackageJson) as {
      dependencies?: Record<string, string>
      devDependencies?: Record<string, string>
    }
    // The accepted host package depends on GenOffice foundation packages
    // only (host → foundation direction) — no Aspose, no interop, no new
    // runtime of any kind was added by the strategy increment.
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-file',
      '@genoffice/project-scheduling',
    ])
    const allDeps = [
      ...Object.keys(pkg.dependencies ?? {}),
      ...Object.keys(pkg.devDependencies ?? {}),
    ]
    expect(allDeps.filter((d) => /aspose|interop|com[/.]/i.test(d))).toEqual([])
  })

  it('no MPP export surface exists on the foundation public index', () => {
    // The foundation exports the .gproj adapter, the MSPDI adapter
    // (import + the sanctioned MSPDI export), and the MPP import contract —
    // and nothing that writes MPP.
    expect(indexSource).toContain("from './mspdi/index.js'")
    expect(indexSource).toContain("from './mpp/index.js'")
    expect(indexSource).not.toContain('exportMpp')
    expect(indexSource).not.toMatch(/export\s+\w*[Mm]pp\w*[Ee]xport/)
  })

  it('the MPP foundation surface remains an import-only contract', () => {
    // PROJECT-018's contract: types, diagnostics, normalizations, and the
    // canonical import entry point — no writer, no bytes-out API.
    expect(mppIndexSource).not.toContain('exportMpp')
    expect(mppIndexSource).not.toMatch(/export\s+\w*[Mm]pp\w*[Ee]xport/)
    expect(mppIndexSource).toContain('importMppFromMspdi')
    // The contract still carries no process/binary code (architecture-lock
    // §13 — unchanged by this investigation):
    expect(mppIndexSource).not.toMatch(/child_process|node:|spawn|exec/)
  })

  it('the MSPDI export surface is unchanged — MSPDI remains the sole sanctioned export path', () => {
    expect(mspdiIndexSource).toContain('exportMspdi')
    expect(mspdiIndexSource).not.toContain('exportMpp')
    expect(mspdiIndexSource).not.toMatch(/mpp/i)
  })

  it('the frozen architecture lock still forbids foundation MPP parser imports', () => {
    // The investigation must not weaken architecture-lock.md §13 — the guard
    // text itself must still be present verbatim.
    expect(architectureLock).toContain(
      'Foundation packages must not import React/React DOM, Electron, Node filesystem/process APIs, browser globals, HTTP clients/server route modules, Excel renderer packages, or `.mpp`/MSPDI parser implementation code.',
    )
    expect(architectureLock).toContain('Status: FROZEN')
  })
})

describe('PROJECT-019 — spec lockstep (the rescope is recorded everywhere)', () => {
  it('requirements.md carries the PROJECT-019 section with the decision and the blocked rule', () => {
    expect(requirements).toContain('## PROJECT-019 — MPP export strategy / rescope')
    expect(requirements).toContain('MPP EXPORT DEFERRED — MSPDI/.gproj IS THE SUPPORTED WRITE PATH')
    expect(requirements).toContain('BLOCKED-until-acceptance rule')
    expect(requirements).toContain('Zero production code')
  })

  it('work-items.md rescopes the 019 row to the strategy increment', () => {
    expect(workItems).toContain('| PROJECT-019 | MPP export strategy / rescope (019A) |')
    // The original (unscoped) definition is gone:
    expect(workItems).not.toMatch(/\|\s*PROJECT-019\s*\|\s*MPP export\s*\|/)
    expect(workItems).toContain(
      'MPP writer, commercial SDK, COM bridge, any MPP export implementation',
    )
  })

  it('dependency-graph.md records the PROJECT-019 edges section', () => {
    expect(dependencyGraph).toContain('## Package dependency edges (PROJECT-019)')
    expect(dependencyGraph).toContain('DECISION RECORD')
  })

  it('verification-matrix.md carries the PROJECT-019 evidence requirements', () => {
    expect(verificationMatrix).toContain('## PROJECT-019 evidence requirements')
    expect(verificationMatrix).toContain('round-trip oracle')
    expect(verificationMatrix).toContain('zero skipped')
  })

  it('the accepted PROJECT-017 report remains intact (its decision line untouched)', () => {
    // The investigation must not silently rewrite the accepted record it
    // builds upon.
    const decisionLine = feasibilityReport
      .split('\n')
      .find((l) => l.includes('FEASIBILITY DECISION:'))
    expect(decisionLine).toBeDefined()
    expect(decisionLine).toContain('FEASIBLE — MSPDI/INTERMEDIARY ADAPTER')
    expect(decisionLine).toContain('import only')
  })
})
