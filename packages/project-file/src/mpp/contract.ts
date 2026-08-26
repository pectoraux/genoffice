/**
 * PROJECT-018 — MPP→canonical import entry point (foundation, host-neutral).
 *
 * `importMppFromMspdi` consumes the outcome of a HOST-MANAGED MPXJ sidecar
 * conversion (MSPDI XML bytes + the validated protocol status frame) and
 * runs the foundation side of the pipeline:
 *
 *   1. N1–N5 normalization (stage 'normalization' diagnostics)
 *   2. the accepted PROJECT-015 `importMspdi` (stage 'mspdi' diagnostics)
 *   3. `validateProjectDocument` (stage 'canonical' diagnostics)
 *
 * The sidecar stage diagnostics carried by the outcome are passed through
 * unchanged (stage 'sidecar'). Semantic authority is NEVER in the sidecar or
 * the normalizer — it stays with the accepted importer and the engine.
 *
 * Error atomicity: a fatal conversion failure (sidecar error, malformed
 * MSPDI, or a canonical-validation rejection) yields `emptyProjectDocument()`
 * — never a partially authoritative document. Recoverable importer-level
 * errors keep the accepted PROJECT-015 semantics (recovered document +
 * error diagnostics; the caller treats any error-level diagnostic as
 * "not usable").
 */
import type { ProjectFileMetadata } from '@genoffice/project-contracts'
import { validateProjectDocument } from '@genoffice/project-engine'
import { emptyProjectDocument } from '../serialize.js'
import { importMspdi } from '../mspdi/importer.js'
import { normalizeMspdiForCanonicalImport } from './normalize.js'
import type { MppConversionOutcome, MppDiagnostic, MppImportResult } from './types.js'

/** Run the foundation MPP pipeline over a successful sidecar conversion. */
export function importMppFromMspdi(
  outcome: MppConversionOutcome,
  metadata?: ProjectFileMetadata,
): MppImportResult {
  // Stage 'sidecar': pass the launcher's diagnostics through unchanged. If
  // the sidecar stage already reported a fatal error, do not touch the
  // (unusable) payload — return the empty document with full provenance.
  const diagnostics: MppDiagnostic[] = [...outcome.sidecarDiagnostics]
  if (diagnostics.some((d) => d.stage === 'sidecar' && d.severity === 'error')) {
    return { document: emptyProjectDocument(), diagnostics }
  }

  // Stage 'normalization': the five PROJECT-017-approved rewrites.
  const normalized = normalizeMspdiForCanonicalImport(outcome.mspdiBytes)
  diagnostics.push(...normalized.diagnostics)

  // Stage 'mspdi': the accepted PROJECT-015 importer remains the sole
  // semantic authority (its own diagnostics are staged, never flattened).
  const imported = importMspdi(normalized.bytes, metadata)
  diagnostics.push(...imported.diagnostics.map((d): MppDiagnostic => ({ ...d, stage: 'mspdi' })))

  // Stage 'canonical': engine validation. A rejected document is atomic
  // failure — the empty document is returned (never a rejected payload).
  // Engine `ValidationDiagnostic` carries no severity: rejection reasons map
  // to 'error' (same mapping `schedule()` applies); diagnostics accompanying
  // an ACCEPTED document map to 'warning'.
  const validation = validateProjectDocument(imported.document)
  diagnostics.push(
    ...validation.diagnostics.map((d): MppDiagnostic => ({
      code: d.code,
      severity: validation.accepted ? 'warning' : 'error',
      message: d.message,
      stage: 'canonical',
    })),
  )
  if (!validation.accepted) {
    return { document: emptyProjectDocument(), diagnostics }
  }
  return { document: imported.document, diagnostics }
}
