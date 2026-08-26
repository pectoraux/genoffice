/**
 * PROJECT-020 — the compatibility classification knowledge base.
 *
 * A deterministic, total mapping from every KNOWN diagnostic code to its
 * compatibility semantics (stage + data-loss class). The tables cover the
 * complete accepted code families:
 *
 *   - PROJECT-014 `.gproj` codes            (stage 'gproj')
 *   - PROJECT-015 MSPDI import codes        (stage 'mspdi')
 *   - PROJECT-018 MPP normalization codes   (stage 'normalization')
 *   - PROJECT-018 sidecar/protocol codes    (stage 'sidecar')
 *   - the canonical engine validation codes (stage 'canonical')
 *   - the scheduling-engine failure codes   (stage 'scheduling')
 *
 * Codes are keyed by string and looked up in one merged table; stage
 * resolution prefers the diagnostic's OWN explicit stage (the accepted
 * `MppDiagnostic` provenance) and falls back to the table. Unknown codes
 * (forward compatibility) fall back deterministically to the format's
 * import stage with a severity-derived loss class — never to a crash and
 * never to a silently different semantic.
 *
 * A lockstep test reads the engine + scheduling sources and asserts every
 * code they can emit has a table entry, so the table cannot drift stale.
 */
import type { ImportDiagnostic } from '@genoffice/project-contracts'
import {
  INVALID_ASSIGNMENT,
  INVALID_BASELINE,
  INVALID_CALENDAR,
  INVALID_GPROJ,
  INVALID_IDENTITY,
  INVALID_REFERENCE,
  INVALID_RESOURCE,
  INVALID_TASK,
  MISSING_REQUIRED_FIELD,
  SCHEMA_INVALID,
  UNSUPPORTED_GPROJ_VERSION,
} from '../diagnostics.js'
import {
  INVALID_MSPDI,
  INVALID_MSPDI_ASSIGNMENT,
  INVALID_MSPDI_CALENDAR,
  INVALID_MSPDI_CONSTRAINT,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_REFERENCE,
  INVALID_MSPDI_RESOURCE,
  MISSING_MSPDI_FIELD,
  MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED,
  MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED,
  MSPDI_READ,
  UNSUPPORTED_MSPDI_FEATURE,
  UNSUPPORTED_MSPDI_VERSION,
} from '../mspdi/diagnostics.js'
import {
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_INPUT_TOO_LARGE,
  MPP_INPUT_UNREADABLE,
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  MPP_OUTPUT_TOO_LARGE,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_UNSUPPORTED_FORMAT,
} from '../mpp/diagnostics.js'
import type {
  CompatibilityDiagnostic,
  CompatibilityEntityType,
  CompatibilityFormat,
  CompatibilityLoss,
  CompatibilityStage,
} from './model.js'
import { deriveCompatibilityEntityType } from './model.js'

/** The table entry for one known code. */
export interface CompatibilityCodeClassification {
  /** The pipeline stage the code belongs to when the diagnostic does not
   * carry its own explicit stage. */
  readonly stage: CompatibilityStage
  /** The data-loss classification for the code. */
  readonly loss: CompatibilityLoss
}

// ── PROJECT-014 `.gproj` family (stage 'gproj') ───────────────────────────
//
// All envelope/schema/entity rejections are 'invalid' (malformed input that
// was refused or dropped); the unsupported-version refusal is 'unsupported'
// (the version itself is well-formed, just outside the supported read set);
// GPROJ_READ is pure bookkeeping ('none').
const GPROJ_CLASSIFICATIONS: Readonly<Record<string, CompatibilityCodeClassification>> = {
  [INVALID_GPROJ]: { stage: 'gproj', loss: 'invalid' },
  [UNSUPPORTED_GPROJ_VERSION]: { stage: 'gproj', loss: 'unsupported' },
  [SCHEMA_INVALID]: { stage: 'gproj', loss: 'invalid' },
  [MISSING_REQUIRED_FIELD]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_IDENTITY]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_REFERENCE]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_CALENDAR]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_BASELINE]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_ASSIGNMENT]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_TASK]: { stage: 'gproj', loss: 'invalid' },
  [INVALID_RESOURCE]: { stage: 'gproj', loss: 'invalid' },
  GPROJ_READ: { stage: 'gproj', loss: 'none' },
}

// ── PROJECT-015 MSPDI import family (stage 'mspdi') ───────────────────────
//
// Same shape as the `.gproj` family (malformed input → 'invalid'; the
// unsupported save-version refusal → 'unsupported'; the umbrella
// unsupported-feature warning → 'unsupported'), plus the two
// PROJECT-020 provenance additions: the PhysicalPercentComplete drop
// ('dropped' — parseable source data with no canonical reconstruction on
// import) and the baseline capturedAt fallback ('approximated' — a
// canonical value derived from the best available MSPDI carrier).
const MSPDI_CLASSIFICATIONS: Readonly<Record<string, CompatibilityCodeClassification>> = {
  [INVALID_MSPDI]: { stage: 'mspdi', loss: 'invalid' },
  [UNSUPPORTED_MSPDI_VERSION]: { stage: 'mspdi', loss: 'unsupported' },
  [UNSUPPORTED_MSPDI_FEATURE]: { stage: 'mspdi', loss: 'unsupported' },
  [INVALID_MSPDI_REFERENCE]: { stage: 'mspdi', loss: 'invalid' },
  [INVALID_MSPDI_DATE]: { stage: 'mspdi', loss: 'invalid' },
  [INVALID_MSPDI_DURATION]: { stage: 'mspdi', loss: 'invalid' },
  [INVALID_MSPDI_CALENDAR]: { stage: 'mspdi', loss: 'invalid' },
  [INVALID_MSPDI_RESOURCE]: { stage: 'mspdi', loss: 'invalid' },
  [INVALID_MSPDI_ASSIGNMENT]: { stage: 'mspdi', loss: 'invalid' },
  [INVALID_MSPDI_CONSTRAINT]: { stage: 'mspdi', loss: 'invalid' },
  [MISSING_MSPDI_FIELD]: { stage: 'mspdi', loss: 'invalid' },
  [MSPDI_READ]: { stage: 'mspdi', loss: 'none' },
  [MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED]: { stage: 'mspdi', loss: 'dropped' },
  [MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED]: { stage: 'mspdi', loss: 'approximated' },
}

// ── PROJECT-018 MPP normalization family (stage 'normalization') ──────────
//
// N1–N4 are the lossless mechanical rewrites ('normalized' — semantics
// preserved by construction). N5 DROPS the unassigned placeholder
// assignment ('dropped' — its code, message, and warning severity all
// state the drop; classifying it 'normalized' would hide true semantic
// loss behind a lossless label).
const MPP_NORMALIZATION_CLASSIFICATIONS: Readonly<Record<string, CompatibilityCodeClassification>> =
  {
    [MPP_NORMALIZED_SENTINEL_REFERENCE]: { stage: 'normalization', loss: 'normalized' },
    [MPP_NORMALIZED_BASE_CALENDAR_SENTINEL]: { stage: 'normalization', loss: 'normalized' },
    [MPP_NORMALIZED_PLACEHOLDER_RECORD]: { stage: 'normalization', loss: 'normalized' },
    [MPP_NORMALIZED_MIDNIGHT_PERIOD]: { stage: 'normalization', loss: 'normalized' },
    [MPP_DROPPED_UNASSIGNED_ASSIGNMENT]: { stage: 'normalization', loss: 'dropped' },
  }

// ── PROJECT-018 sidecar/protocol family (stage 'sidecar') ─────────────────
//
// Conversion failures classify 'invalid' (the operation failed; NO
// compatibility claim is made — nothing was imported), except the
// unsupported/unrecognized format refusal ('unsupported').
const MPP_SIDECAR_CLASSIFICATIONS: Readonly<Record<string, CompatibilityCodeClassification>> = {
  [MPP_SIDECAR_UNAVAILABLE]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_SIDECAR_TIMEOUT]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_SIDECAR_EXIT]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_SIDECAR_RESPONSE_INVALID]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_INPUT_UNREADABLE]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_INPUT_TOO_LARGE]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_OUTPUT_TOO_LARGE]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE]: { stage: 'sidecar', loss: 'invalid' },
  [MPP_UNSUPPORTED_FORMAT]: { stage: 'sidecar', loss: 'unsupported' },
}

// ── Canonical engine validation codes (stage 'canonical') ─────────────────
//
// The complete accepted `validateProjectDocument` code set: every entry is a
// rejection reason for the constructed document (severity 'error' in every
// accepted adapter mapping) → 'invalid'. A lockstep test keeps this list in
// sync with the engine source.
const ENGINE_VALIDATION_CODES = [
  'CALENDAR_CYCLE',
  'CALENDAR_PERIOD_MALFORMED',
  'CONSTRAINT_DATE_NOT_ALLOWED',
  'DEPENDENCY_CYCLE',
  'DUPLICATE_ASSIGNMENT_ID',
  'DUPLICATE_ASSIGNMENT_PAIR',
  'DUPLICATE_BASELINE_ID',
  'DUPLICATE_CALENDAR_ID',
  'DUPLICATE_CUSTOM_FIELD_ID',
  'DUPLICATE_DEPENDENCY_ID',
  'DUPLICATE_DEPENDENCY_LINK',
  'DUPLICATE_RESOURCE_ID',
  'DUPLICATE_RESOURCE_UID',
  'DUPLICATE_TASK_ID',
  'DUPLICATE_TASK_UID',
  'INCONSISTENT_COST',
  'INCONSISTENT_OUTLINE_LEVEL',
  'INCONSISTENT_SUMMARY_FLAG',
  'INCONSISTENT_WORK',
  'INVALID_ACTUAL_COST',
  'INVALID_ACTUAL_WORK',
  'INVALID_ASSIGNMENT_UNITS',
  'INVALID_AVAILABILITY_RANGE',
  'INVALID_AVAILABILITY_UNITS',
  'INVALID_COST',
  'INVALID_COST_PER_USE',
  'INVALID_DATE',
  'INVALID_LAG',
  'INVALID_MAX_UNITS',
  'INVALID_OUTLINE_LEVEL',
  'INVALID_PERCENT_COMPLETE',
  'INVALID_RATE',
  'INVALID_REMAINING_COST',
  'INVALID_REMAINING_WORK',
  'INVALID_RESOURCE_KIND',
  'INVALID_WORK',
  'MISSING_BASELINE_REFERENCE',
  'MISSING_BASE_CALENDAR',
  'MISSING_CALENDAR',
  'MISSING_CONSTRAINT_DATE',
  'MISSING_CUSTOM_FIELD_REFERENCE',
  'MISSING_PARENT',
  'MISSING_RESOURCE_NAME',
  'MISSING_RESOURCE_REFERENCE',
  'MISSING_TASK_REFERENCE',
  'NEGATIVE_DURATION',
  'PARENT_CYCLE',
  'SELF_DEPENDENCY',
  'SELF_PARENT',
  'SUMMARY_DEPENDENCY',
] as const

// ── Scheduling-engine failure codes (stage 'scheduling') ──────────────────
//
// `schedule()` surfaces engine validation codes (already tabled above) plus
// its own `DependencyGraphError`/`CalendarError` codes. These scheduling-only
// codes default to the 'scheduling' stage when they arrive without an
// explicit stage. A scheduling failure is derived-state failure — the
// canonical document itself stays valid and save-eligible.
const SCHEDULING_ONLY_CODES = ['CALENDAR_SEARCH_EXHAUSTED', 'CYCLE'] as const

/** The merged, total classification table over every known code. */
export const COMPATIBILITY_CODE_CLASSIFICATIONS: Readonly<
  Record<string, CompatibilityCodeClassification>
> = {
  ...GPROJ_CLASSIFICATIONS,
  ...MSPDI_CLASSIFICATIONS,
  ...MPP_NORMALIZATION_CLASSIFICATIONS,
  ...MPP_SIDECAR_CLASSIFICATIONS,
  ...Object.fromEntries(
    ENGINE_VALIDATION_CODES.map((code) => [code, { stage: 'canonical', loss: 'invalid' }]),
  ),
  ...Object.fromEntries(
    SCHEDULING_ONLY_CODES.map((code) => [code, { stage: 'scheduling', loss: 'invalid' }]),
  ),
}

/** The import-stage fallback per format (for codes without a table entry). */
const FORMAT_IMPORT_STAGE: Readonly<Record<CompatibilityFormat, CompatibilityStage>> = {
  gproj: 'gproj',
  mspdi: 'mspdi',
  mpp: 'mspdi',
}

/**
 * Resolve the pipeline stage of a diagnostic WITHOUT classifying it: the
 * explicit `stage` (accepted `MppDiagnostic` provenance) wins; otherwise the
 * classification table's stage; otherwise the format's import stage.
 * (Exported for the aggregator, which must resolve stages BEFORE deriving
 * statuses from flat adapter diagnostics.)
 */
export function resolveCompatibilityStage(
  format: CompatibilityFormat,
  diagnostic: ImportDiagnostic,
): CompatibilityStage {
  const explicitStage = (
    'stage' in diagnostic ? (diagnostic as { stage?: unknown }).stage : undefined
  ) as CompatibilityStage | undefined
  if (explicitStage !== undefined) return explicitStage
  return COMPATIBILITY_CODE_CLASSIFICATIONS[diagnostic.code]?.stage ?? FORMAT_IMPORT_STAGE[format]
}

/**
 * Stamp one produced diagnostic with its compatibility semantics.
 *
 * Stage resolution: an explicit `stage` (the accepted `MppDiagnostic`
 * provenance) always wins; otherwise the table's stage; otherwise the
 * format's import stage. Loss resolution: the table's class; otherwise a
 * deterministic severity-derived fallback ('invalid' for errors — an
 * unknown rejection is still a rejection; 'unsupported' for warnings —
 * unknown degradations in this codebase are unsupported-feature warnings;
 * 'none' for info). Recoverability follows the severity ladder, with
 * 'fatal' reserved for errors on a fatal import (see `aggregate.ts`).
 *
 * `entityType` is derived best-effort from the identity prefix — never
 * invented.
 */
export function classifyImportDiagnostic(
  format: CompatibilityFormat,
  diagnostic: ImportDiagnostic,
  importFailed: boolean,
): CompatibilityDiagnostic {
  const stage = resolveCompatibilityStage(format, diagnostic)
  const table = COMPATIBILITY_CODE_CLASSIFICATIONS[diagnostic.code]
  const loss =
    table?.loss ??
    (diagnostic.severity === 'error'
      ? 'invalid'
      : diagnostic.severity === 'warning'
        ? 'unsupported'
        : 'none')
  const recoverability = recoverabilityOf(stage, diagnostic.severity, importFailed)
  const entityType: CompatibilityEntityType | undefined = deriveCompatibilityEntityType(
    diagnostic.entityId,
  )
  const classified: CompatibilityDiagnostic = {
    code: diagnostic.code,
    severity: diagnostic.severity,
    message: diagnostic.message,
    ...(diagnostic.entityId !== undefined ? { entityId: diagnostic.entityId } : {}),
    format,
    stage,
    loss,
    recoverability,
    ...(entityType !== undefined ? { entityType } : {}),
  }
  return classified
}

/**
 * The stage+severity → recoverability ladder:
 *   info    → 'preserved'  (lossless bookkeeping / normalizations).
 *   warning → 'canonical'  (valid document; degradation recorded).
 *   error   → 'partial'    (entity-level recovery — or an invalid document
 *                            that was still constructed and returned).
 *   error on a fatal import → 'fatal' (no authoritative document).
 *   EXCEPTION — a scheduling-stage error is a DERIVED-STATE failure: the
 *   canonical document stays valid and save-eligible, so it classifies
 *   'canonical' (scheduling is recomputed state, never document validity).
 */
function recoverabilityOf(
  stage: CompatibilityStage,
  severity: ImportDiagnostic['severity'],
  importFailed: boolean,
): CompatibilityDiagnostic['recoverability'] {
  if (severity === 'info') return 'preserved'
  if (stage === 'scheduling') return 'canonical'
  if (severity === 'warning') return 'canonical'
  return importFailed ? 'fatal' : 'partial'
}
