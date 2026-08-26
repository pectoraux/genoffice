/**
 * PROJECT-020 — compatibility pipeline entry points.
 *
 * Each entry point runs the ACCEPTED import pipeline unchanged and augments
 * the result with the canonical `CompatibilityReport`:
 *
 *   `importGprojWithCompatibility` — deserializeGproj → (engine validation
 *     is already embedded) → optional injected scheduling.
 *   `importMspdiWithCompatibility` — importMspdi → same.
 *   `importMppWithCompatibility`   — importMppFromMspdi (sidecar → N1–N5
 *     normalization → accepted importer → engine validation, staged) →
 *     optional injected scheduling.
 *
 * NO parser is rewritten: the pipelines consume the accepted entry points
 * verbatim and classify/aggregate what they produce. The ONLY parser-side
 * additions this increment makes are two MSPDI provenance diagnostics
 * (PhysicalPercentComplete drop, baseline capturedAt approximation) — see
 * `../mspdi/diagnostics.ts`.
 *
 * Scheduling: `@genoffice/project-file` deliberately does NOT depend on
 * the scheduling package (the accepted dependency edge since
 * PROJECT-014). The scheduler is INJECTED as a structurally-typed runner —
 * hosts/tests pass `schedule` from the scheduling package (its
 * `DerivedSchedule` is structurally assignable), exactly like the accepted
 * test-layer precedent. A runner is only invoked over an authoritative
 * document (import not failed AND canonical validation passed) — a
 * rejected document is never scheduled.
 */
import type {
  ImportDiagnostic,
  ProjectDocument,
  ProjectFileMetadata,
} from '@genoffice/project-contracts'
import { deserializeGproj, inspectGproj } from '../deserialize.js'
import { INVALID_GPROJ } from '../diagnostics.js'
import { importMspdi, inspectMspdi } from '../mspdi/importer.js'
import { INVALID_MSPDI } from '../mspdi/diagnostics.js'
import { importMppFromMspdi } from '../mpp/contract.js'
import type { MppConversionOutcome } from '../mpp/types.js'
import { buildCompatibilityReport, isFatalImport } from './aggregate.js'
import { resolveCompatibilityStage } from './classification.js'
import type { CompatibilityReport, CompatibilityFormat } from './model.js'

/**
 * The injected scheduling runner. `schedule` from
 * the scheduling package satisfies this type structurally
 * (`DerivedSchedule` carries `diagnostics: ImportDiagnostic[]`); the
 * foundation package itself stays scheduling-free.
 */
export type CompatibilityScheduleRunner = (document: ProjectDocument) => {
  diagnostics: readonly ImportDiagnostic[]
}

/** Options for the compatibility pipelines. */
export interface CompatibilityOptions {
  /** The injected scheduler. Absent → the report's scheduling status is
   * 'not-attempted'. */
  readonly schedule?: CompatibilityScheduleRunner
}

/** The result of a compatibility pipeline run. */
export interface CompatibilityPipelineResult {
  /** The document the accepted pipeline returned (the atomic empty
   * document whenever no authoritative import exists). */
  readonly document: ProjectDocument
  readonly report: CompatibilityReport
}

/** Run the injected scheduler when — and only when — it is eligible. */
function runSchedulerIfEligible(
  format: CompatibilityFormat,
  diagnostics: readonly ImportDiagnostic[],
  document: ProjectDocument,
  options: CompatibilityOptions | undefined,
): readonly ImportDiagnostic[] | undefined {
  const runner = options?.schedule
  if (runner === undefined) return undefined
  if (isFatalImport(format, diagnostics)) return undefined
  if (
    diagnostics.some(
      (d) => resolveCompatibilityStage(format, d) === 'canonical' && d.severity === 'error',
    )
  ) {
    return undefined
  }
  return runner(document).diagnostics
}

/** `.gproj` → canonical + compatibility report. */
export function importGprojWithCompatibility(
  input: Uint8Array,
  metadata?: ProjectFileMetadata,
  options?: CompatibilityOptions,
): CompatibilityPipelineResult {
  const result = deserializeGproj(input, metadata)
  const fatal = isFatalImport('gproj', result.diagnostics)
  const sourceVersion = gprojSourceVersion(input, fatal, result.diagnostics)
  const schedulingDiagnostics = runSchedulerIfEligible(
    'gproj',
    result.diagnostics,
    result.document,
    options,
  )
  const report = buildCompatibilityReport({
    format: 'gproj',
    ...(sourceVersion !== undefined ? { sourceVersion } : {}),
    diagnostics: result.diagnostics,
    ...(schedulingDiagnostics !== undefined ? { schedulingDiagnostics } : {}),
  })
  return { document: result.document, report }
}

/** MSPDI XML → canonical + compatibility report. */
export function importMspdiWithCompatibility(
  input: Uint8Array,
  metadata?: ProjectFileMetadata,
  options?: CompatibilityOptions,
): CompatibilityPipelineResult {
  const result = importMspdi(input, metadata)
  const fatal = isFatalImport('mspdi', result.diagnostics)
  const sourceVersion = mspdiSourceVersion(input, fatal, result.diagnostics)
  const schedulingDiagnostics = runSchedulerIfEligible(
    'mspdi',
    result.diagnostics,
    result.document,
    options,
  )
  const report = buildCompatibilityReport({
    format: 'mspdi',
    ...(sourceVersion !== undefined ? { sourceVersion } : {}),
    diagnostics: result.diagnostics,
    ...(schedulingDiagnostics !== undefined ? { schedulingDiagnostics } : {}),
  })
  return { document: result.document, report }
}

/** MPP (via a host-managed sidecar conversion outcome) → canonical +
 * compatibility report. */
export function importMppWithCompatibility(
  outcome: MppConversionOutcome,
  metadata?: ProjectFileMetadata,
  options?: CompatibilityOptions,
): CompatibilityPipelineResult {
  const result = importMppFromMspdi(outcome, metadata)
  const schedulingDiagnostics = runSchedulerIfEligible(
    'mpp',
    result.diagnostics,
    result.document,
    options,
  )
  const report = buildCompatibilityReport({
    format: 'mpp',
    ...(outcome.frame.format !== undefined ? { sourceVersion: outcome.frame.format } : {}),
    diagnostics: result.diagnostics,
    ...(schedulingDiagnostics !== undefined ? { schedulingDiagnostics } : {}),
  })
  return { document: result.document, report }
}

/**
 * The honest `.gproj` source version: the envelope's `formatVersion` when
 * the bytes parsed as an envelope at all, and `undefined` when they did not
 * (the inspect fallback would otherwise fabricate a version for garbage
 * input — an unreadable file has no honest source version).
 */
function gprojSourceVersion(
  input: Uint8Array,
  fatal: boolean,
  diagnostics: readonly ImportDiagnostic[],
): string | undefined {
  if (fatal && diagnostics.some((d) => d.code === INVALID_GPROJ)) return undefined
  return inspectGproj(input).version
}

/** The honest MSPDI source version (same rule as `.gproj`). */
function mspdiSourceVersion(
  input: Uint8Array,
  fatal: boolean,
  diagnostics: readonly ImportDiagnostic[],
): string | undefined {
  if (fatal && diagnostics.some((d) => d.code === INVALID_MSPDI)) return undefined
  return inspectMspdi(input).version
}
