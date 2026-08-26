/**
 * PROJECT-018 — MPP public surface (foundation, host-neutral).
 *
 * The MPP import path reuses the accepted `ProjectFileAdapter` philosophy:
 * this package owns the adapter CONTRACT, the N1–N5 MSPDI normalization,
 * the staged diagnostics, and the canonical mapping. The MPXJ conversion
 * PROCESS is host-managed (the host-side launcher package) — no Java, binary, or
 * process code lives in a foundation package (architecture-lock §13).
 */
export {
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_INPUT_UNREADABLE,
  MPP_INPUT_TOO_LARGE,
  MPP_OUTPUT_TOO_LARGE,
  MPP_UNSUPPORTED_FORMAT,
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_NORMALIZATION_DIAGNOSTIC_CODES,
  MPP_SIDECAR_DIAGNOSTIC_CODES,
  MPP_DIAGNOSTIC_CODES,
} from './diagnostics.js'
export {
  MPP_SIDECAR_PROTOCOL_VERSION,
  MPXJ_PINNED_VERSION,
  MPP_SUPPORTED_FORMAT_VERSIONS,
  MPP_MAX_INPUT_BYTES,
  MPP_MAX_MSPDI_OUTPUT_BYTES,
  type MppDiagnostic,
  type MppDiagnosticStage,
  type MppSidecarCounts,
  type MppSidecarFrame,
  type MppSidecarFrameError,
  type MppConversionOutcome,
  type MppImportResult,
} from './types.js'
export {
  normalizeMspdiForCanonicalImport,
  serializeXmlNode,
  type MppNormalizationResult,
} from './normalize.js'
export { importMppFromMspdi } from './contract.js'
