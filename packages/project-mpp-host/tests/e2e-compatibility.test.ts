/**
 * PROJECT-020 — REAL-CORPUS compatibility integration.
 *
 * Reuses the PROJECT-018 real MPP corpus (the same pinned, SHA-256-verified
 * downloads — NO corpus duplication) and runs every file through the FULL
 * production pipeline with the PROJECT-020 compatibility layer:
 *
 *   real MPP → MPXJ sidecar → N1–N5 normalization → accepted MSPDI
 *   importer → canonical validation → schedule() → CompatibilityReport
 *
 * For every corpus file the test asserts:
 *   - the manifest's expected N1–N5 diagnostics (as classified
 *     normalization-stage entries — never silent);
 *   - the manifest's expected error count (as classified error entries);
 *   - the three status dimensions, authority, and save eligibility;
 *   - the honest sourceVersion (the sidecar's detected MPP format);
 *   - determinism: a second full run produces a byte-identical report
 *     (JSON), canonical document (`serializeGproj`), and DerivedSchedule.
 *
 * Network-isolation posture follows the accepted e2e-real-corpus suite:
 * the production 'required' posture wherever the host provides the
 * mechanism, with an explicit recorded opt-out otherwise.
 */
import { mkdtempSync, rmSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import {
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  serializeGproj,
} from '@genoffice/project-file'
import { schedule } from '@genoffice/project-scheduling'
import {
  MppSidecarLauncher,
  importMppFromFileWithCompatibility,
  probeNetworkIsolation,
  type NetworkIsolationPolicy,
} from '../src/index.js'

const DEPS = join(import.meta.dirname, '..', '.sidecar-deps')
const CORPUS = join(DEPS, 'corpus')

interface ManifestEntry {
  filename: string
  format: string
  /** Byte-true MPP container generation detected via MPXJ (PROJECT-020) —
   * may differ from the provenance `format` label (e.g. SubprojectA-9.mpp
   * is an MPP9-era name but an MPP14 container). */
  detectedFormat: string
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

const isolationCapability = await probeNetworkIsolation()
const e2eIsolationPolicy: NetworkIsolationPolicy = isolationCapability.supported
  ? 'required'
  : 'off'
if (isolationCapability.supported) {
  console.info(
    `[compatibility] MPP sidecar network isolation: REQUIRED (${isolationCapability.mechanism}) — every conversion runs inside a kernel network namespace`,
  )
} else {
  console.warn(
    `[compatibility] MPP sidecar network isolation: this host cannot provide the mechanism (${isolationCapability.reason}); running with the explicit 'off' opt-out — the fail-closed refusal is proven in network-isolation.test.ts`,
  )
}
const launcher = new MppSidecarLauncher({
  mpxjHome: join(DEPS, 'mpxj-16.7.0'),
  networkIsolation: e2eIsolationPolicy,
})

let scratch: string
beforeAll(() => {
  scratch = mkdtempSync(join(tmpdir(), 'mpp-compat-'))
})
afterAll(() => {
  rmSync(scratch, { recursive: true, force: true })
})

const json = (value: unknown): string => JSON.stringify(value)

/** Detected formats across the per-file loop (for the family-coverage
 * assertion — no additional conversions). */
const detectedFormats = new Set<string>()

describe('PROJECT-020 — real-corpus compatibility reports', () => {
  for (const entry of manifest.corpus) {
    // Each test runs TWO full pipeline passes (2× JVM conversion) under the
    // enforced isolation wrapper — a generous explicit timeout keeps slow
    // shared CI runners honest (the files themselves run sequentially —
    // see vitest.config.ts).
    it(
      `${entry.filename}: full-pipeline compatibility report (N1–N5 classified, manifest-conformant, deterministic)`,
      { timeout: 120_000 },
      async () => {
        const first = await importMppFromFileWithCompatibility(join(CORPUS, entry.filename), {
          launcher,
        })
        const report = first.report

        // The honest source version is the sidecar's DETECTED format
        // (byte-true container generation — not the filename label):
        expect(report.format).toBe('mpp')
        expect(report.sourceVersion).toBe(entry.detectedFormat)
        detectedFormats.add(entry.detectedFormat)

        // Status dimensions: every corpus file imports (possibly with
        // degradation), passes canonical validation, and schedules:
        expect(report.status.validation).toBe('success')
        expect(report.status.scheduling).toBe('success')
        expect(report.authoritative).toBe(true)
        expect(report.saveEligibility).toBe('allowed')
        expect(report.status.import).not.toBe('failure')
        if (entry.expectedErrors > 0) {
          expect(report.status.import).toBe('success-with-errors')
        }

        // The manifest's expected N1–N5 counts as classified normalization
        // diagnostics (warnings for N5 — expected loss; infos for N1–N4):
        const count = (code: string): number =>
          report.diagnostics.filter((d) => d.code === code).length
        expect(count(MPP_NORMALIZED_SENTINEL_REFERENCE)).toBe(entry.expectedNormalizations.N1)
        expect(count(MPP_NORMALIZED_BASE_CALENDAR_SENTINEL)).toBe(entry.expectedNormalizations.N2)
        expect(count(MPP_NORMALIZED_PLACEHOLDER_RECORD)).toBe(entry.expectedNormalizations.N3)
        expect(count(MPP_NORMALIZED_MIDNIGHT_PERIOD)).toBe(entry.expectedNormalizations.N4)
        expect(count(MPP_DROPPED_UNASSIGNED_ASSIGNMENT)).toBe(entry.expectedNormalizations.N5)
        // Every N-family entry is classified, never silent:
        for (const d of report.diagnostics) {
          if (d.stage === 'normalization') expect(d.loss).not.toBe('none')
        }

        // The manifest's expected error count, as classified error entries
        // (entity-level partial recovery — the document stays authoritative):
        expect(report.errorCount).toBe(entry.expectedErrors)
        expect(
          report.diagnostics
            .filter((d) => d.severity === 'error')
            .every((d) => d.recoverability === 'partial' && d.stage !== 'scheduling'),
        ).toBe(true)

        // The canonical shape contract (the accepted I-golden invariant):
        expect(first.document.tasks).toHaveLength(entry.canonical.tasks)

        // Determinism: a second full pipeline run gives the byte-identical
        // report, canonical document, and derived schedule:
        const second = await importMppFromFileWithCompatibility(join(CORPUS, entry.filename), {
          launcher,
        })
        expect(json(second.report)).toBe(json(report))
        expect(serializeGproj(second.document)).toEqual(serializeGproj(first.document))
        expect(json(schedule(second.document))).toBe(json(schedule(first.document)))
      },
    )
  }

  it('the corpus covers every MPP format family through the compatibility layer (from the per-file reports)', () => {
    // Derived from the loop above — no additional conversions:
    expect(detectedFormats.has('MPP8')).toBe(true)
    expect(detectedFormats.has('MPP9')).toBe(true)
    expect(detectedFormats.has('MPP12')).toBe(true)
    expect(detectedFormats.has('MPP14')).toBe(true)
    expect([...detectedFormats].sort()).toEqual(
      [...new Set(manifest.corpus.map((entry) => entry.detectedFormat))].sort(),
    )
  })
})
