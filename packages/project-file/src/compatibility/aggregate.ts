/**
 * PROJECT-020 — deterministic compatibility aggregation.
 *
 * `buildCompatibilityReport` turns the diagnostics an accepted import
 * pipeline produced (flat `ImportDiagnostic[]` for `.gproj`/MSPDI, staged
 * `MppDiagnostic[]` for MPP) plus the scheduling-stage outcome into the
 * canonical `CompatibilityReport`.
 *
 * Determinism contract (same input ⇒ same report, byte-identical):
 *   - classification is a pure table lookup (no message parsing);
 *   - the report sort is the canonical key
 *       stage → severity → code → entityType → entityId → message
 *     with a stable tie-break on the producer's original order (so source
 *     order survives ONLY among otherwise-identical keys — semantically
 *     meaningful ordering such as per-entity grouping is never destroyed,
 *     and reordering source elements cannot reorder the report);
 *   - NO dependence on object iteration order (explicit key sequences),
 *     `localeCompare`, filesystem enumeration, randomness, or timestamps.
 *
 * De-duplication policy (deterministic): the aggregation layer performs NO
 * lossy de-duplication — identical diagnostics are distinct, countable
 * occurrences (e.g. five N4 midnight rewrites) whose multiplicity carries
 * information. Declaration-level uniqueness is a PRODUCER contract, already
 * in force in the accepted pipeline (the lazy lag-factor validation emits
 * one diagnostic per malformed declaration regardless of how many
 * dependencies use it); entity-scoped diagnostics are distinct by their
 * `entityId`. Accidental de-duplication of meaningful distinct warnings is
 * therefore impossible by construction.
 *
 * Aggregation cost is O(n log n) in the diagnostic count (one sort) with
 * O(1) per-diagnostic table lookups — no quadratic de-duplication scan
 * exists (none is performed).
 */
import type { ImportDiagnostic } from '@genoffice/project-contracts'
import { classifyImportDiagnostic, resolveCompatibilityStage } from './classification.js'
import {
  COMPATIBILITY_SEVERITY_ORDER,
  COMPATIBILITY_STAGE_ORDER,
  type CompatibilityDiagnostic,
  type CompatibilityFormat,
  type CompatibilityImportStatus,
  type CompatibilityReport,
  type CompatibilityStatus,
} from './model.js'

/** The input to compatibility aggregation. */
export interface CompatibilityReportInput {
  readonly format: CompatibilityFormat
  /** The honest source version (`.gproj` formatVersion, MSPDI
   * `<SaveVersion>`, the sidecar's detected MPP format); omit when the
   * input never parsed far enough to know. */
  readonly sourceVersion?: string
  /** Diagnostics as produced by the accepted import pipelines (flat or
   * staged — staged diagnostics keep their explicit stage). */
  readonly diagnostics: readonly ImportDiagnostic[]
  /** Scheduling-stage diagnostics. PRESENCE of the array (even empty)
   * signals that a scheduler ran; `undefined` means not attempted. */
  readonly schedulingDiagnostics?: readonly ImportDiagnostic[]
}

/** The read-sentinel code per format: its presence means the structural
 * import COMPLETED (the accepted adapters emit it only at the end of a
 * successful construction — file-level failures return before it). */
const READ_SENTINELS: Readonly<Record<CompatibilityFormat, string>> = {
  gproj: 'GPROJ_READ',
  mspdi: 'MSPDI_READ',
  mpp: 'MSPDI_READ',
}

/** Stages that belong to the structural import (not validation/scheduling). */
const IMPORT_STAGES: ReadonlySet<string> = new Set(['sidecar', 'normalization', 'gproj', 'mspdi'])

/** The explicit stage of a diagnostic when it carries one. */
function explicitStageOf(diagnostic: ImportDiagnostic): string | undefined {
  return 'stage' in diagnostic
    ? ((diagnostic as { stage?: unknown }).stage as string | undefined)
    : undefined
}

/** The stage of a diagnostic as seen by THIS aggregator: the explicit stage
 * when present, otherwise the classification table's stage (flat
 * `.gproj`/MSPDI diagnostics get their stage resolved by code BEFORE any
 * status derivation — a canonical-stage engine code embedded in a flat
 * adapter result is 'canonical', an adapter code is the format's import
 * stage). */
function stageOf(format: CompatibilityFormat, diagnostic: ImportDiagnostic): string {
  return explicitStageOf(diagnostic) ?? resolveCompatibilityStage(format, diagnostic)
}

/** True when the structural import failed to produce any document. */
export function isFatalImport(
  format: CompatibilityFormat,
  diagnostics: readonly ImportDiagnostic[],
): boolean {
  const readSentinel = READ_SENTINELS[format]
  if (format === 'mpp') {
    // The MPP pipeline is atomic on every fatal layer: a sidecar-stage
    // error, a fatally failed MSPDI import (no MSPDI_READ among the
    // mspdi-stage diagnostics), or the canonical rejection (empty-document
    // atomicity) all mean NO authoritative document.
    const sidecarError = diagnostics.some(
      (d) => stageOf(format, d) === 'sidecar' && d.severity === 'error',
    )
    const mspdiSawRead = diagnostics.some(
      (d) => d.code === 'MSPDI_READ' && stageOf(format, d) === 'mspdi',
    )
    const canonicalError = diagnostics.some(
      (d) => stageOf(format, d) === 'canonical' && d.severity === 'error',
    )
    return sidecarError || !mspdiSawRead || canonicalError
  }
  const sawRead = diagnostics.some((d) => d.code === readSentinel)
  return !sawRead
}

/** Derive the structural-import status before scheduling is considered. */
function deriveImportStatus(
  format: CompatibilityFormat,
  diagnostics: readonly ImportDiagnostic[],
): { importStatus: CompatibilityImportStatus; fatal: boolean } {
  const fatal = isFatalImport(format, diagnostics)
  if (fatal) return { importStatus: 'failure', fatal }
  let sawError = false
  let sawWarning = false
  for (const d of diagnostics) {
    const stage = stageOf(format, d)
    if (IMPORT_STAGES.has(stage)) {
      if (d.severity === 'error') sawError = true
      else if (d.severity === 'warning') sawWarning = true
    }
  }
  if (sawError) return { importStatus: 'success-with-errors', fatal }
  if (sawWarning) return { importStatus: 'success-with-warnings', fatal }
  return { importStatus: 'success', fatal }
}

/** Derive the canonical-validation status. */
function deriveValidationStatus(
  fatal: boolean,
  diagnostics: readonly ImportDiagnostic[],
  format: CompatibilityFormat,
): CompatibilityStatus['validation'] {
  const canonicalError = diagnostics.some(
    (d) => stageOf(format, d) === 'canonical' && d.severity === 'error',
  )
  if (canonicalError) return 'failure'
  if (fatal) return 'not-attempted'
  return 'success'
}

/** The canonical report sort key fields, in order (documented contract). */
const STAGE_INDEX: ReadonlyMap<string, number> = new Map(
  COMPATIBILITY_STAGE_ORDER.map((stage, index) => [stage, index]),
)
const SEVERITY_INDEX: ReadonlyMap<string, number> = new Map(
  COMPATIBILITY_SEVERITY_ORDER.map((severity, index) => [severity, index]),
)

/** Locale-independent code-unit comparison — NOT `localeCompare` (which is
 * locale-dependent and therefore forbidden by the determinism contract). */
function compareStrings(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

/** The canonical sort: stage → severity → code → entityType → entityId →
 * message, with a stable tie-break on the producer's original order. */
function compareDiagnostics(
  a: CompatibilityDiagnostic,
  b: CompatibilityDiagnostic,
  aIndex: number,
  bIndex: number,
): number {
  return (
    (STAGE_INDEX.get(a.stage) ?? 0) - (STAGE_INDEX.get(b.stage) ?? 0) ||
    (SEVERITY_INDEX.get(a.severity) ?? 0) - (SEVERITY_INDEX.get(b.severity) ?? 0) ||
    compareStrings(a.code, b.code) ||
    compareStrings(a.entityType ?? '', b.entityType ?? '') ||
    compareStrings(a.entityId ?? '', b.entityId ?? '') ||
    compareStrings(a.message, b.message) ||
    aIndex - bIndex
  )
}

/**
 * Build the deterministic compatibility report.
 *
 * Pure function: the same `{format, sourceVersion, diagnostics,
 * schedulingDiagnostics}` always produce the same report (same diagnostics,
 * same canonical ordering, same severity/code/entity association, same
 * counts, same status/authority/save-eligibility).
 */
export function buildCompatibilityReport(input: CompatibilityReportInput): CompatibilityReport {
  const staged: ImportDiagnostic[] = [...input.diagnostics]
  if (input.schedulingDiagnostics !== undefined) {
    for (const d of input.schedulingDiagnostics) {
      staged.push({ ...d, stage: 'scheduling' } as ImportDiagnostic & { stage: 'scheduling' })
    }
  }

  const { importStatus, fatal } = deriveImportStatus(input.format, staged)
  const validationStatus = deriveValidationStatus(fatal, staged, input.format)

  // Scheduling is derived state: it is only "attempted" over an
  // authoritative document. (The host MPP pipeline trivially schedules the
  // atomic empty document after a fatal conversion — that is NOT an
  // attempt, and the report says so.)
  let schedulingStatus: CompatibilityStatus['scheduling']
  if (importStatus === 'failure' || validationStatus === 'failure') {
    schedulingStatus = 'not-attempted'
  } else if (input.schedulingDiagnostics === undefined) {
    schedulingStatus = 'not-attempted'
  } else {
    schedulingStatus = input.schedulingDiagnostics.some((d) => d.severity === 'error')
      ? 'failure'
      : 'success'
  }

  const authoritative = importStatus !== 'failure' && validationStatus === 'success'

  const classified = staged.map((d) => classifyImportDiagnostic(input.format, d, fatal))
  // Stable sort: the canonical key first; exact key ties keep the
  // producer's original order (source order preserved where it is the only
  // distinguishing information — never the reverse).
  const indexed = classified.map((d, index) => ({ d, index }))
  indexed.sort((a, b) => compareDiagnostics(a.d, b.d, a.index, b.index))
  const diagnostics = indexed.map((entry) => entry.d)

  let preservedCount = 0
  let normalizedCount = 0
  let approximatedCount = 0
  let droppedCount = 0
  let unsupportedCount = 0
  let errorCount = 0
  let warningCount = 0
  for (const d of diagnostics) {
    switch (d.loss) {
      case 'none':
        preservedCount++
        break
      case 'normalized':
        normalizedCount++
        break
      case 'approximated':
        approximatedCount++
        break
      case 'dropped':
        droppedCount++
        break
      case 'unsupported':
        unsupportedCount++
        break
      // 'invalid' rejections are counted in errorCount only — they claim
      // no compatibility.
      case 'invalid':
        break
    }
    if (d.severity === 'error') errorCount++
    else if (d.severity === 'warning') warningCount++
  }

  const report: CompatibilityReport = {
    format: input.format,
    ...(input.sourceVersion !== undefined ? { sourceVersion: input.sourceVersion } : {}),
    status: { import: importStatus, validation: validationStatus, scheduling: schedulingStatus },
    authoritative,
    saveEligibility: authoritative ? 'allowed' : 'prohibited',
    diagnostics,
    preservedCount,
    normalizedCount,
    approximatedCount,
    droppedCount,
    unsupportedCount,
    errorCount,
    warningCount,
  }
  return report
}
