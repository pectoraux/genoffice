/**
 * PROJECT-017 — MPP feasibility evidence/discipline suite.
 *
 * PROJECT-017 is an investigation: the deliverable is
 * `spec/project/mpp-feasibility.md` plus spec updates — NO production MPP
 * parser/writer code, NO new runtime dependency, NO scheduling/adapter
 * behavior change. This suite enforces exactly that, and validates the
 * report's structure against the PROJECT-017 brief:
 *
 *   - the report exists and carries all 20 mandated sections;
 *   - it ends with exactly one of the four allowed feasibility decisions
 *     (the chosen one being the MSPDI/intermediary conclusion);
 *   - it documents the required content: version matrix, candidate/library/
 *     licensing/runtime matrices, security analysis, fidelity tiers, the
 *     MSPDI-intermediary loss analysis (N1–N5 normalizations), the corpus
 *     plan, determinism, performance, and source links for external claims;
 *   - NO production MPP implementation was introduced: the project-file
 *     package has no MPP source module, its runtime dependencies are exactly
 *     contracts + engine, and the public export surface gained nothing MPP;
 *   - the frozen architecture lock still governs MPP parser boundaries under
 *     the clarified §13 rule (ACR-001 — external MSPDI/MPP parser
 *     implementations forbidden for semantic/runtime foundation packages;
 *     project-file the sanctioned in-package adapter boundary; the
 *     investigation may not silently weaken it).
 *
 * The dynamic feasibility claims themselves (MPXJ installability, real-file
 * processing, mapping to canonical structures) were validated by the isolated
 * spike documented in the report §11 and the worklog — they are intentionally
 * NOT re-run here (no network-dependent, non-hermetic tests in CI).
 */
import { describe, expect, it } from 'vitest'
import report from '../../../spec/project/mpp-feasibility.md?raw'
import packageJson from '../package.json?raw'
import architectureLock from '../../../spec/project/architecture-lock.md?raw'

const REQUIRED_SECTIONS = [
  '## 1. Executive conclusion',
  '## 2. Format overview',
  '## 3. Version matrix',
  '## 4. Candidate approaches',
  '## 5. Library/tool matrix',
  '## 6. Licensing matrix',
  '## 7. Runtime matrix',
  '## 8. Security analysis',
  '## 9. Fidelity matrix',
  '## 10. MSPDI intermediary analysis',
  '## 11. Representative MPP corpus plan',
  '## 12. Determinism analysis',
  '## 13. Performance considerations',
  '## 14. Architecture impact',
  '## 15. Recommended strategy',
  '## 16. PROJECT-018 proposed scope',
  '## 17. PROJECT-019 proposed scope',
  '## 18. Explicit rejected alternatives',
  '## 19. Open questions',
  '## 20. Evidence / source links',
]

const ALLOWED_DECISIONS = [
  'FEASIBLE — DIRECT NATIVE ADAPTER',
  'FEASIBLE — MSPDI/INTERMEDIARY ADAPTER',
  'FEASIBLE — LIMITED SUBSET ONLY',
  'NOT FEASIBLE UNDER CURRENT ARCHITECTURE',
]

describe('PROJECT-017 — MPP feasibility report structure', () => {
  it('the report file exists and is substantial', () => {
    expect(report.length).toBeGreaterThan(10_000)
  })

  it('contains all 20 mandated sections in order', () => {
    let position = -1
    for (const section of REQUIRED_SECTIONS) {
      const next = report.indexOf(section)
      expect(next).toBeGreaterThan(position) // present AND ordered
      position = next
    }
  })

  it('ends with exactly one allowed feasibility decision', () => {
    const decisions = ALLOWED_DECISIONS.filter((d) => report.includes(`**${d}**`))
    expect(decisions).toEqual(['FEASIBLE — MSPDI/INTERMEDIARY ADAPTER'])
    const decisionLine = report.split('\n').find((l) => l.includes('FEASIBILITY DECISION:'))
    expect(decisionLine).toBeDefined()
    expect(decisionLine).toContain('FEASIBLE — MSPDI/INTERMEDIARY ADAPTER')
    expect(decisionLine).toContain('import only')
  })

  it('documents the version matrix for MPP8/9/12/14', () => {
    for (const version of ['MPP8', 'MPP9', 'MPP12', 'MPP14']) {
      expect(report).toContain(version)
    }
  })

  it('documents the five MSPDI-intermediary normalization requirements', () => {
    for (const id of ['N1', 'N2', 'N3', 'N4', 'N5']) {
      expect(report).toContain(id)
    }
    // The concrete findings from the spike:
    expect(report).toContain('-1')
    expect(report).toContain('-65535')
    expect(report).toContain('OutlineLevel 0')
    expect(report).toContain('00:00:00')
  })

  it('documents the fidelity tiers A–D', () => {
    for (const tier of ['Tier A', 'Tier B', 'Tier C', 'Tier D']) {
      expect(report).toContain(tier)
    }
  })

  it('carries a source-link section with numbered external evidence', () => {
    expect(report).toContain('[S1] MPXJ')
    expect(report).toContain('[S10] Aspose.Tasks')
    expect(report).toContain('[S17] Spike primary evidence')
    expect(report).toContain('https://www.mpxj.org/')
    expect(report).toContain('https://learn.microsoft.com/')
  })

  it('documents the corpus plan with external consumption and provenance', () => {
    expect(report).toContain('pinned')
    expect(report).toContain('618')
    expect(report).toContain('LGPL')
  })

  it('documents the MPP-export rejection basis', () => {
    expect(report).toContain('MPP export is NOT feasible')
    expect(report).toContain('Aspose')
  })

  it('documents performance expectations and PROJECT-048 benchmark handoff', () => {
    expect(report).toContain('100 tasks')
    expect(report).toContain('100,000 tasks')
    expect(report).toContain('PROJECT-048')
  })
})

describe('PROJECT-017 — no production MPP implementation introduced', () => {
  it('the project-file package declares exactly the accepted runtime dependencies', () => {
    const pkg = JSON.parse(packageJson) as {
      dependencies?: Record<string, string>
      devDependencies?: Record<string, string>
    }
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
    // No MPP-related package may appear even as a devDependency (the spike ran
    // entirely outside the repository).
    const allDeps = [
      ...Object.keys(pkg.dependencies ?? {}),
      ...Object.keys(pkg.devDependencies ?? {}),
    ]
    expect(allDeps.filter((d) => /mpp|mpxj|aspose/i.test(d))).toEqual([])
  })

  it('the public surface gained no MPP parser/PROCESS module (the PROJECT-018 foundation contract is host-neutral)', async () => {
    // PROJECT-017 delivered NO production MPP code. PROJECT-018 (subsequently
    // authorized) added the host-neutral foundation contract (types, N1–N5
    // MSPDI normalization, staged diagnostics, importMppFromMspdi) under
    // src/mpp/** — the superseding-authorization update of this assertion,
    // exactly as PROJECT-016 superseded the PROJECT-015 no-export test.
    // The guard still enforces what remains true: no MPP PARSER or PROCESS
    // code — the package must never spawn processes or parse MPP bytes.
    const indexSource = (await import('../src/index.ts?raw')).default as string
    expect(indexSource).toContain("from './mspdi/index.js'")
    // The foundation MPP surface is contract-only (no child_process/fs/
    // process spawning — that lives in the host package):
    const mppIndexSource = (await import('../src/mpp/index.ts?raw')).default as string
    expect(mppIndexSource).not.toMatch(/child_process|node:|spawn|exec/)
  })

  it('the frozen architecture lock still governs MPP parser boundaries (clarified by ACR-001, never weakened)', () => {
    // The roadmap reconciliation increment (ACR-001) clarified §13: external
    // MSPDI/MPP parser implementations stay forbidden for the semantic/runtime
    // foundation packages, and packages/project-file is the sanctioned
    // in-package adapter boundary — the pure-TS MSPDI tokenizer this report
    // recommended keeps its sanctioned home, and MPXJ remains confined to the
    // host sidecar. The guard text must still be present verbatim.
    expect(architectureLock).toContain(
      'Foundation semantic/runtime packages (`project-contracts`, `project-engine`, `project-scheduling`, `project-renderer-core`) must not import external MSPDI/MPP parser implementations.',
    )
    expect(architectureLock).toContain(
      '`packages/project-file` is the sanctioned file-adapter boundary and may contain format-specific parser/serializer implementations.',
    )
    expect(architectureLock).toContain(
      'File-format implementations remain behind the `project-file` adapter boundary.',
    )
    expect(architectureLock).toContain('Status: FROZEN')
  })
})
