/**
 * PROJECT-018 — REAL-CORPUS end-to-end goldens I01–I12.
 *
 * Runs the full production pipeline over the 8 real MPP corpus files
 * (MPP8/MPP9/MPP12/MPP14 — pinned downloads verified by corpus.test.ts)
 * through the actual MPXJ sidecar:
 *
 *   MPP bytes → MPXJ sidecar → MSPDI → N1–N5 normalization →
 *   accepted PROJECT-015 importer → validateProjectDocument → schedule()
 *
 * Required-test mapping (30 items): #1–#4 version imports (I01–I04),
 * #5/#29 repeated determinism + full-pipeline determinism (I12),
 * #6–#10 N1–N5 (I05–I09, against REAL files with raw-payload negative
 * controls), #16 scheduling determinism (I12), #21 baselines (I01/I02),
 * #22 resources/assignments (I04/I10), #23 calendars (I10),
 * #24 dependencies (I03/I04/I10), #25 constraints (recorded: the corpus
 * carries none — the constraint surface is proven by the synthetic suites),
 * #26 diagnostics provenance (I10), #30 .gproj save/reopen (I12).
 * Items #11–#15/#17–#20 (sidecar failure/security classes) live in
 * launcher.test.ts + architecture.test.ts (java-independent injection).
 *
 * Performance evidence (brief §performance): representative wall times and
 * sizes are console-recorded per file for the verification matrix.
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import {
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_UNSUPPORTED_FORMAT,
  deserializeGproj,
  emptyProjectDocument,
  importMspdi,
  serializeGproj,
  type MppDiagnostic,
} from '@genoffice/project-file'
import { schedule } from '@genoffice/project-scheduling'
import { MppSidecarLauncher, importMppFromFile } from '../src/index.js'

const DEPS = join(import.meta.dirname, '..', '.sidecar-deps')
const CORPUS = join(DEPS, 'corpus')

interface ManifestEntry {
  filename: string
  format: string
  expectedNormalizations: Record<'N1' | 'N2' | 'N3' | 'N4' | 'N5', number>
  expectedErrors: number
  canonical: {
    tasks: number
    dependencies: number
    resources: number
    assignments: number
    calendars: number
    baselines: number
    customFields: number
    projectFinish: string | null
  }
}

const manifest = JSON.parse(
  readFileSync(join(import.meta.dirname, '..', 'corpus', 'corpus-manifest.json'), 'utf8'),
) as { corpus: ManifestEntry[] }

const byName = new Map(manifest.corpus.map((entry) => [entry.filename, entry]))
const launcher = new MppSidecarLauncher({ mpxjHome: join(DEPS, 'mpxj-16.7.0') })

let scratch: string
beforeAll(() => {
  scratch = mkdtempSync(join(tmpdir(), 'mpp-e2e-'))
})
afterAll(() => {
  rmSync(scratch, { recursive: true, force: true })
})

/** Import a corpus file end-to-end and assert the full manifest contract. */
async function importCorpusFile(filename: string) {
  const entry = byName.get(filename)
  expect(entry, `manifest entry for ${filename}`).toBeDefined()
  const start = Date.now()
  const result = await importMppFromFile(join(CORPUS, filename), { launcher })
  const elapsed = Date.now() - start
  const document = result.document
  const derived = schedule(document)
  const gproj = serializeGproj(document)
  // ── manifest contract ──
  expect(document.tasks, `${filename} tasks`).toHaveLength(entry!.canonical.tasks)
  expect(document.dependencies, `${filename} deps`).toHaveLength(entry!.canonical.dependencies)
  expect(document.resources, `${filename} resources`).toHaveLength(entry!.canonical.resources)
  expect(document.assignments, `${filename} assignments`).toHaveLength(entry!.canonical.assignments)
  expect(document.calendars, `${filename} calendars`).toHaveLength(entry!.canonical.calendars)
  expect(document.baselines, `${filename} baselines`).toHaveLength(entry!.canonical.baselines)
  expect(document.customFields, `${filename} customFields`).toHaveLength(
    entry!.canonical.customFields,
  )
  expect(derived.projectFinish ?? null).toBe(entry!.canonical.projectFinish)
  const errors = result.diagnostics.filter((d) => d.severity === 'error')
  expect(errors, `${filename} error diagnostics`).toHaveLength(entry!.expectedErrors)
  // ── normalization counts ──
  expect(count(result.diagnostics, MPP_NORMALIZED_SENTINEL_REFERENCE)).toBe(
    entry!.expectedNormalizations.N1,
  )
  expect(count(result.diagnostics, MPP_NORMALIZED_BASE_CALENDAR_SENTINEL)).toBe(
    entry!.expectedNormalizations.N2,
  )
  expect(count(result.diagnostics, MPP_NORMALIZED_PLACEHOLDER_RECORD)).toBe(
    entry!.expectedNormalizations.N3,
  )
  expect(count(result.diagnostics, MPP_NORMALIZED_MIDNIGHT_PERIOD)).toBe(
    entry!.expectedNormalizations.N4,
  )
  expect(count(result.diagnostics, MPP_DROPPED_UNASSIGNED_ASSIGNMENT)).toBe(
    entry!.expectedNormalizations.N5,
  )
  console.log(
    `[perf] ${filename}: pipeline ${elapsed} ms; tasks ${document.tasks.length}; gproj ${gproj.byteLength} bytes; schedule ok`,
  )
  return { result, document, derived, gproj, entry: entry! }
}

function count(diagnostics: readonly MppDiagnostic[], code: string): number {
  return diagnostics.filter((d) => d.code === code).length
}

/** Convert a corpus file to raw MSPDI via the launcher (for negative
 * controls: the UN-normalized payload must fail in the accepted importer). */
async function rawMspdi(filename: string): Promise<Uint8Array> {
  const output = join(scratch, `raw-${filename}.mspdi`)
  const conversion = await launcher.convert(join(CORPUS, filename), output, `raw-${filename}`)
  expect(conversion.ok).toBe(true)
  return (conversion as { mspdiBytes: Uint8Array }).mspdiBytes
}

// ── I01–I04: version imports ───────────────────────────────────────────

describe('I01 — MPP8 import (DurationTest8.mpp, Project 98)', () => {
  it('imports deterministically to the manifest contract', async () => {
    const { document, derived } = await importCorpusFile('DurationTest8.mpp')
    // MPP8-specific evidence: 8 real tasks, one baseline, milestone-free
    expect(document.tasks.every((t) => t.milestone === false)).toBe(true)
    expect(document.baselines[0]?.id).toBe('b0')
    expect(derived.projectFinish).toBe('2006-04-10T17:00:00.000Z')
    // The MPP8 documented fidelity caveat: custom flags are not asserted
    // here (the MPXJ MPP8 Flag-field defect — PROJECT-017 §3).
  })
})

describe('I02 — MPP9 import (SubprojectA-9.mpp, Project 2000/2002)', () => {
  it('imports the local subproject tasks, dependency, and calendar', async () => {
    const { document, derived } = await importCorpusFile('SubprojectA-9.mpp')
    expect(document.dependencies).toHaveLength(1)
    expect(document.calendars.map((c) => c.id)).toContain('c1')
    expect(derived.projectFinish).toBe('2006-08-29T17:00:00.000Z')
    // Cross-file subproject expansion is out of the canonical model
    // (documented Tier-C boundary): only the local tasks imported.
  })
})

describe('I03 — MPP12 import (mpp12relations.mpp, Project 2003/2007)', () => {
  it('imports with the same structural contract as its MPP14 sibling', async () => {
    const { document, derived } = await importCorpusFile('mpp12relations.mpp')
    expect(document.dependencies).toHaveLength(4)
    expect(document.tasks).toHaveLength(5)
    expect(derived.projectFinish).toBe('2006-09-26T17:00:00.000Z')
  })
})

describe('I04 — MPP14 import (mpp14relations.mpp, Project 2010+)', () => {
  it('imports tasks and predecessor links', async () => {
    const { document } = await importCorpusFile('mpp14relations.mpp')
    expect(document.dependencies).toHaveLength(4)
    expect(document.dependencies.every((d) => d.type.length === 2)).toBe(true)
  })
})

// ── I05–I09: the five normalizations against REAL files ────────────────

describe('I05 — N1 sentinel CalendarUID strip (SubprojectA-9.mpp)', () => {
  it('every stripped sentinel is diagnosed with a task entity id', async () => {
    const { result } = await importCorpusFile('SubprojectA-9.mpp')
    const n1 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_SENTINEL_REFERENCE)
    expect(n1).toHaveLength(3)
    expect(n1.every((d) => /^t\d+$/.test(d.entityId ?? ''))).toBe(true)
  })

  it('negative control: the RAW payload fails in the accepted importer', async () => {
    const raw = await rawMspdi('SubprojectA-9.mpp')
    expect(
      importMspdi(raw).diagnostics.some(
        (d) => d.code === 'INVALID_MSPDI_REFERENCE' && d.severity === 'error',
      ),
    ).toBe(true)
  })
})

describe('I06 — N2 sentinel BaseCalendarUID strip (SubprojectA-9.mpp)', () => {
  it('the base-calendar sentinel is diagnosed with a calendar entity id', async () => {
    const { result } = await importCorpusFile('SubprojectA-9.mpp')
    const n2 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_BASE_CALENDAR_SENTINEL)
    expect(n2).toHaveLength(1)
    expect(n2[0].entityId).toMatch(/^c\d+$/)
  })

  it('negative control: the RAW payload reports MISSING_BASE_CALENDAR', async () => {
    const raw = await rawMspdi('SubprojectA-9.mpp')
    expect(importMspdi(raw).diagnostics.some((d) => d.code === 'MISSING_BASE_CALENDAR')).toBe(true)
  })
})

describe('I07 — N3 hidden placeholder records (mpp14relations.mpp)', () => {
  it('the uid-0 placeholder task AND the null-name placeholder resource are filtered', async () => {
    const { result, document } = await importCorpusFile('mpp14relations.mpp')
    const n3 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_PLACEHOLDER_RECORD)
    expect(n3).toHaveLength(2)
    expect(n3.map((d) => d.entityId).sort()).toEqual(['r0', 't0'])
    expect(document.tasks.some((t) => t.id === 't0')).toBe(false)
  })

  it('negative control: the RAW payload reports INVALID_OUTLINE_LEVEL', async () => {
    const raw = await rawMspdi('mpp14relations.mpp')
    expect(importMspdi(raw).diagnostics.some((d) => d.code === 'INVALID_OUTLINE_LEVEL')).toBe(true)
  })
})

describe('I08 — N4 until-midnight working periods (sample.mpp)', () => {
  it('the Night Shift calendar keeps its 23:00→24:00 period as endMinute 1440', async () => {
    const { document } = await importCorpusFile('sample.mpp')
    const nightShift = document.calendars.find((c) => c.name === 'Night Shift')
    expect(nightShift).toBeDefined()
    // Some weekday of the night-shift calendar carries {1380, 1440}:
    const allPeriods = Object.values(nightShift!.workingWeek).flat()
    expect(allPeriods).toContainEqual({ startMinute: 1380, endMinute: 1440 })
  })

  it('negative control: the RAW payload drops the period with INVALID_MSPDI_CALENDAR', async () => {
    const raw = await rawMspdi('sample.mpp')
    expect(
      importMspdi(raw).diagnostics.some(
        (d) => d.code === 'INVALID_MSPDI_CALENDAR' && d.severity === 'error',
      ),
    ).toBe(true)
  })
})

describe('I09 — N5 unassigned assignments (mpp14relations.mpp)', () => {
  it('all five -65535 placeholder assignments are dropped with warnings; none survive', async () => {
    const { result, document } = await importCorpusFile('mpp14relations.mpp')
    const n5 = result.diagnostics.filter((d) => d.code === MPP_DROPPED_UNASSIGNED_ASSIGNMENT)
    expect(n5).toHaveLength(5)
    expect(n5.every((d) => d.severity === 'warning')).toBe(true)
    expect(document.assignments).toHaveLength(0)
  })

  it('negative control: the RAW payload fails with reference errors', async () => {
    const raw = await rawMspdi('mpp14relations.mpp')
    expect(
      importMspdi(raw).diagnostics.some(
        (d) => d.code === 'INVALID_MSPDI_REFERENCE' && d.severity === 'error',
      ),
    ).toBe(true)
  })
})

// ── I10: comprehensive import (sample.mpp) ──────────────────────────────

describe('I10 — comprehensive MPP import (sample.mpp)', () => {
  it('maps every canonical surface with correct provenance-staged diagnostics', async () => {
    const { result, document, derived } = await importCorpusFile('sample.mpp')

    // Tasks: names, uids, hierarchy, WBS, milestones, summaries:
    expect(document.tasks.map((t) => t.name).slice(0, 3)).toEqual([
      'First Task',
      'Second Task',
      'Third task',
    ])
    expect(document.tasks.find((t) => t.id === 't1')).toMatchObject({
      uid: 1,
      name: 'First Task',
      wbs: '1',
      priority: 500,
    })
    expect(document.tasks.find((t) => t.id === 't2')?.parentTaskId).toBe('t1')
    expect(document.tasks.find((t) => t.id === 't2')?.wbs).toBe('1.1')
    expect(document.tasks.filter((t) => t.milestone)).toHaveLength(1)
    expect(document.tasks.filter((t) => t.summary)).toHaveLength(4)

    // Dependencies: all four relationship types + non-trivial lags:
    const depTypes = new Set(document.dependencies.map((d) => d.type))
    expect(depTypes).toEqual(new Set(['FS', 'SS', 'FF', 'SF']))
    expect(document.dependencies.map((d) => d.lagMinutes)).toContain(480)
    expect(document.dependencies.map((d) => d.lagMinutes)).toContain(-480)

    // Constraints: the corpus carries none (recorded honestly — the
    // constraint surface is proven by the synthetic suites):
    expect(document.tasks.filter((t) => t.constraint !== undefined)).toHaveLength(0)

    // Resources + assignments:
    expect(document.resources.map((r) => r.kind)).toEqual(['work', 'work'])
    expect(document.assignments).toHaveLength(4)
    expect(new Set(document.assignments.map((a) => a.resourceId))).toEqual(new Set(['r1', 'r2']))

    // Calendars: base + derived with inheritance preserved:
    expect(document.calendars.map((c) => [c.id, c.name, c.baseCalendarId])).toEqual([
      ['c1', 'Standard', undefined],
      ['c2', 'Night Shift', undefined],
      ['c3', 'First Resource', 'c1'],
      ['c4', 'Second Resource', 'c2'],
    ])

    // Scheduling: exact derived values (hand-verified against the scheduler):
    expect(derived.projectFinish).toBe('2003-05-26T17:00:00.000Z')
    expect(Object.values(derived.taskSchedules).filter((s) => s.critical)).toHaveLength(2)
    const t1 = derived.taskSchedules['t1']
    expect(t1?.earlyStart).toBe('2003-01-07T08:00:00.000Z')
    expect(t1?.earlyFinish).toBe('2003-05-26T17:00:00.000Z')
    expect(t1?.critical).toBe(true)

    // Diagnostics provenance (test item #26): every stage is one of the five
    // documented values; normalization + mspdi stages are present; zero errors.
    const stages = new Set(result.diagnostics.map((d) => d.stage))
    expect(
      [...stages].every((s) =>
        ['sidecar', 'normalization', 'mspdi', 'canonical', 'scheduling'].includes(s),
      ),
    ).toBe(true)
    expect(stages.has('normalization')).toBe(true)
    expect(stages.has('mspdi')).toBe(true)
    expect(result.diagnostics.filter((d) => d.severity === 'error')).toHaveLength(0)
  })
})

// ── I11: malformed / unsupported inputs (through the REAL sidecar) ─────

describe('I11 — malformed/unsupported inputs', () => {
  it('garbage bytes are deterministically refused as MPP_UNSUPPORTED_FORMAT', async () => {
    const result = await (async () => {
      const path = join(scratch, 'garbage.mpp')
      writeFileSync(path, Buffer.from('this is definitely not a CFB container'))
      return importMppFromFile(path, { launcher })
    })()
    expect(result.document).toEqual(emptyProjectDocument())
    const unsupported = result.diagnostics.filter((d) => d.code === MPP_UNSUPPORTED_FORMAT)
    expect(unsupported).toHaveLength(1)
    expect(unsupported[0].severity).toBe('error')
    expect(unsupported[0].stage).toBe('sidecar')
    expect(result.diagnostics.every((d) => d.stage === 'sidecar')).toBe(true)
  })

  it('a ZIP-magic file (not an OLE2 container) is refused the same way', async () => {
    const path = join(scratch, 'zipped.mpp')
    writeFileSync(path, Buffer.from([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]))
    const result = await importMppFromFile(path, { launcher })
    expect(result.diagnostics.filter((d) => d.code === MPP_UNSUPPORTED_FORMAT)).toHaveLength(1)
    expect(result.document).toEqual(emptyProjectDocument())
  })

  it('an empty file is refused the same way', async () => {
    const path = join(scratch, 'empty.mpp')
    writeFileSync(path, Buffer.alloc(0))
    const result = await importMppFromFile(path, { launcher })
    expect(result.diagnostics.filter((d) => d.code === MPP_UNSUPPORTED_FORMAT)).toHaveLength(1)
  })
})

// ── I12: determinism + .gproj round-trip (test items #5/#29/#30) ───────

describe('I12 — determinism: repeated full-pipeline runs + .gproj save/reopen', () => {
  const ALL_FILES = [
    'DurationTest8.mpp',
    'SubprojectA-9.mpp',
    'task-baselines-project2003-mpp9.mpp',
    'mpp14relations.mpp',
    'mpp14assignmentcustom.mpp',
    'sample.mpp',
    'ResourceIdAndUniqueId-project2010-mpp14.mpp',
    'mpp12relations.mpp',
  ]

  for (const filename of ALL_FILES) {
    it(`${filename}: MPP → MPXJ → MSPDI → canonical is byte-identical across runs`, async () => {
      const first = await importMppFromFile(join(CORPUS, filename), { launcher })
      const second = await importMppFromFile(join(CORPUS, filename), { launcher })
      // Canonical-document determinism (the sidecar's <CurrentDate> save
      // stamp is non-semantic — the accepted importer ignores it):
      expect(serializeGproj(second.document)).toEqual(serializeGproj(first.document))
      // Scheduling determinism:
      expect(JSON.stringify(schedule(second.document))).toBe(
        JSON.stringify(schedule(first.document)),
      )
      // Diagnostic determinism (same codes, same order):
      expect(second.diagnostics).toEqual(first.diagnostics)
    })

    it(`${filename}: MPP → .gproj → reopen preserves canonical semantics and the schedule`, async () => {
      const imported = await importMppFromFile(join(CORPUS, filename), { launcher })
      const gproj = serializeGproj(imported.document)
      const reopened = deserializeGproj(gproj)
      // Reopen is byte-stable (the accepted PROJECT-014 invariant):
      expect(serializeGproj(reopened.document)).toEqual(gproj)
      // The schedule over the reopened document is identical:
      expect(JSON.stringify(schedule(reopened.document))).toBe(
        JSON.stringify(schedule(imported.document)),
      )
      expect(reopened.diagnostics.filter((d) => d.severity === 'error')).toEqual([])
    })
  }
})
