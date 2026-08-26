/**
 * PROJECT-020 — canonical import-compatibility diagnostic model.
 *
 * The single compatibility model for EVERY import path:
 *
 *   .gproj : gproj → canonical → scheduling
 *   MSPDI  : mspdi  → canonical → scheduling
 *   MPP    : sidecar → normalization → mspdi → canonical → scheduling
 *
 * It UNIFIES the diagnostic provenance that already exists across the
 * accepted increments (PROJECT-014 `.gproj` codes, PROJECT-015/016 MSPDI
 * codes, PROJECT-018 staged `MppDiagnostic`) — it does NOT invent a second
 * canonical model, does NOT move diagnostics into renderer/host code, and
 * does NOT rename any existing code (compatibility with existing callers is
 * structural: `CompatibilityDiagnostic` is an `ImportDiagnostic` superset).
 *
 * Every diagnostic carries:
 *   - `code` / `severity` / `message` / `entityId` — the accepted contract.
 *   - `format`   — which import path produced it.
 *   - `stage`    — which pipeline stage (deterministic names; the accepted
 *                  `MppDiagnosticStage` values are preserved verbatim, with
 *                  `'gproj'` added for the native read stage).
 *   - `loss`     — the data-loss classification (NONE / NORMALIZED /
 *                  APPROXIMATED / DROPPED / UNSUPPORTED / INVALID).
 *   - `recoverability` — what remains possible after the diagnostic
 *                  (fatal / partial / canonical / preserved).
 *   - `entityType` — best-effort canonical entity kind derived from the
 *                  deterministic identity prefixes (`t`/`r`/`a`/`c`/`b`/
 *                  `d-`/numeric custom-field ids). Omitted when not
 *                  derivable — the layer never invents provenance the
 *                  parsers do not have (no `field` or `sourceLocation`
 *                  properties exist for the same reason).
 */
import type { ImportDiagnostic } from '@genoffice/project-contracts'

/** The import path a compatibility report describes. */
export type CompatibilityFormat = 'gproj' | 'mspdi' | 'mpp'

/**
 * The pipeline stage a diagnostic originates from. Deterministic names,
 * never collapsed into a generic "import failed": the accepted
 * `MppDiagnosticStage` values (`'sidecar' | 'normalization' | 'mspdi' |
 * 'canonical' | 'scheduling'`) are preserved verbatim; `'gproj'` names the
 * native `.gproj` read stage (the `GPROJ_READ` provenance point).
 */
export type CompatibilityStage =
  'sidecar' | 'normalization' | 'gproj' | 'mspdi' | 'canonical' | 'scheduling'

/** Canonical pipeline order used by the deterministic report sort. */
export const COMPATIBILITY_STAGE_ORDER: readonly CompatibilityStage[] = [
  'sidecar',
  'normalization',
  'gproj',
  'mspdi',
  'canonical',
  'scheduling',
] as const

/**
 * The stage sequence each format's pipeline runs (documented contract —
 * `MPP_SIDEcar → NORMALIZATION → MSPDI_IMPORT → CANONICAL → SCHEDULING`,
 * `MSPDI_IMPORT → CANONICAL → SCHEDULING`, `GPROJ_READ → CANONICAL →
 * SCHEDULING`).
 */
export const COMPATIBILITY_STAGE_PIPELINE: Readonly<
  Record<CompatibilityFormat, readonly CompatibilityStage[]>
> = {
  gproj: ['gproj', 'canonical', 'scheduling'],
  mspdi: ['mspdi', 'canonical', 'scheduling'],
  mpp: ['sidecar', 'normalization', 'mspdi', 'canonical', 'scheduling'],
}

/** Severity order inside one stage (most actionable first). */
export const COMPATIBILITY_SEVERITY_ORDER: readonly ImportDiagnostic['severity'][] = [
  'error',
  'warning',
  'info',
] as const

/**
 * Data-loss classification. Every compatibility issue maps to exactly one:
 *
 *   'none'        — nothing lost (bookkeeping info: a successful read).
 *   'normalized'  — a mechanical normalization preserved the semantics
 *                   (N1–N4: sentinel strips, placeholder filters, the
 *                   midnight→day-end rewrite).
 *   'approximated'— a canonical value was approximated from the best
 *                   available source (e.g. baseline `capturedAt` fallback —
 *                   MSPDI carries no per-baseline capture date).
 *   'dropped'     — parseable source data was dropped because the canonical
 *                   model has no representation for it (N5 unassigned
 *                   assignments, MSPDI `PhysicalPercentComplete`).
 *   'unsupported' — the source feature has no faithful canonical
 *                   representation at all (elapsed durations, percentage
 *                   lags, unsupported format/version).
 *   'invalid'     — malformed/unresolvable input that was rejected (a
 *                   malformed envelope, a dangling reference, an engine
 *                   validation failure).
 */
export type CompatibilityLoss =
  'none' | 'normalized' | 'approximated' | 'dropped' | 'unsupported' | 'invalid'

/**
 * What remains possible after a diagnostic — the recoverability ladder
 * (never "best effort"):
 *
 *   'fatal'    — no authoritative ProjectDocument exists; nothing downstream
 *                is permitted (file-level failure, sidecar failure, the MPP
 *                pipeline's atomic canonical rejection).
 *   'partial'  — permits partial structural import only: the affected
 *                entity/feature was dropped, the surviving document
 *                continues (entity-level error recovery); canonical
 *                creation may still succeed — or fail (an engine rejection
 *                surfaced through `.gproj`/MSPDI keeps the constructed —
 *                invalid — document, still 'partial').
 *   'canonical'— permits canonical ProjectDocument creation: the document
 *                is canonically valid, schedulable, and save-eligible; the
 *                diagnostic records degradation (warnings) or derived-state
 *                failure (a scheduling-stage error never invalidates the
 *                document).
 *   'preserved'— fully preserved semantics (informational bookkeeping:
 *                read sentinels and the lossless N1–N4 normalizations).
 */
export type CompatibilityRecoverability = 'fatal' | 'partial' | 'canonical' | 'preserved'

/** Best-effort canonical entity kind derived from identity prefixes. */
export type CompatibilityEntityType =
  'task' | 'resource' | 'assignment' | 'calendar' | 'baseline' | 'dependency' | 'custom-field'

/** An `ImportDiagnostic` stamped with its compatibility semantics. */
export interface CompatibilityDiagnostic extends ImportDiagnostic {
  /** The import path that produced the diagnostic. */
  readonly format: CompatibilityFormat
  /** The pipeline stage of origin (never flattened). */
  readonly stage: CompatibilityStage
  /** The data-loss classification. */
  readonly loss: CompatibilityLoss
  /** What remains possible after this diagnostic. */
  readonly recoverability: CompatibilityRecoverability
  /** Canonical entity kind when derivable from the identity prefix. */
  readonly entityType?: CompatibilityEntityType
}

/**
 * The structural import outcome. Four honest states, never a single boolean:
 * 'failure' means NO authoritative document was produced (file-level or —
 * for MPP — the atomic canonical rejection); entity-level errors are
 * 'success-with-errors' (a document WAS constructed, degraded and
 * diagnosed); warnings alone are 'success-with-warnings'.
 */
export type CompatibilityImportStatus =
  'success' | 'success-with-warnings' | 'success-with-errors' | 'failure'

/** Validation outcome ('not-attempted' when the import failed first). */
export type CompatibilityValidationStatus = 'success' | 'failure' | 'not-attempted'

/** Scheduling outcome ('not-attempted' when no authoritative document exists). */
export type CompatibilitySchedulingStatus = 'success' | 'failure' | 'not-attempted'

/**
 * The three independent status dimensions. A document can be e.g.
 * `import: 'success-with-warnings' / validation: 'failure' / scheduling:
 * 'not-attempted'` — the dimensions are never compressed into one value.
 */
export interface CompatibilityStatus {
  readonly import: CompatibilityImportStatus
  readonly validation: CompatibilityValidationStatus
  readonly scheduling: CompatibilitySchedulingStatus
}

/** Whether the imported document may be saved as `.gproj`. */
export type CompatibilitySaveEligibility = 'allowed' | 'prohibited'

/**
 * The deterministic compatibility summary. Counts are over the report's own
 * (canonically sorted) diagnostics: the five loss-class counts plus the raw
 * error/warning counts. 'invalid'-class diagnostics (rejections) are
 * deliberately counted only in `errorCount` — they claim no compatibility.
 */
export interface CompatibilityReport {
  readonly format: CompatibilityFormat
  /** The source format version when it could be determined honestly
   * (`.gproj` formatVersion, MSPDI `<SaveVersion>`, the sidecar's detected
   * MPP format). Omitted when the input never parsed far enough to know. */
  readonly sourceVersion?: string
  readonly status: CompatibilityStatus
  /** True iff an authoritative canonical document exists (import not a
   * failure AND canonical validation succeeded). */
  readonly authoritative: boolean
  /** 'allowed' iff authoritative — a canonically valid document may be
   * saved with its degradation recorded in this report (never presented as
   * lossless); canonical validation errors prohibit save. */
  readonly saveEligibility: CompatibilitySaveEligibility
  readonly diagnostics: readonly CompatibilityDiagnostic[]
  readonly preservedCount: number
  readonly normalizedCount: number
  readonly approximatedCount: number
  readonly droppedCount: number
  readonly unsupportedCount: number
  readonly errorCount: number
  readonly warningCount: number
}

/**
 * Derive the canonical entity kind from a diagnostic `entityId`.
 *
 * The accepted identity mapping is deterministic and prefix-disjoint
 * (`t`/`r`/`a`/`c`/`b`/`d-`; MSPDI custom-field ids are bare integers), so
 * prefix derivation is exact for adapter-produced diagnostics. Hand-authored
 * `.gproj` documents may carry arbitrary id strings — derivation is
 * best-effort and simply omits `entityType` when the prefix is unknown.
 */
export function deriveCompatibilityEntityType(
  entityId: string | undefined,
): CompatibilityEntityType | undefined {
  if (entityId === undefined || entityId === '') return undefined
  if (/^t\d+$/.test(entityId)) return 'task'
  if (/^r\d+$/.test(entityId)) return 'resource'
  if (/^a\d+$/.test(entityId)) return 'assignment'
  if (/^c\d+$/.test(entityId)) return 'calendar'
  if (/^b\d+$/.test(entityId)) return 'baseline'
  if (entityId.startsWith('d-')) return 'dependency'
  if (/^\d+$/.test(entityId)) return 'custom-field'
  return undefined
}
