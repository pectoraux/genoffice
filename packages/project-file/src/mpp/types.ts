/**
 * PROJECT-018 — MPP import shared contract types (foundation, host-neutral).
 *
 * The MPP import path is layered so that diagnostic provenance survives:
 *
 *   MPP bytes
 *     → host-managed MPXJ conversion sidecar   (stage: 'sidecar')
 *     → N1–N5 MSPDI normalization               (stage: 'normalization')
 *     → accepted PROJECT-015 MSPDI importer     (stage: 'mspdi')
 *     → validateProjectDocument                 (stage: 'canonical')
 *     → schedule()                              (stage: 'scheduling')
 *
 * The foundation package owns every stage EXCEPT the execution of the
 * sidecar process itself (host-managed, per the frozen architecture: no
 * Java/native/executable dependencies may live inside foundation packages).
 * This module defines only types and constants — it is fully host-neutral.
 */
import type { ImportDiagnostic, ProjectDocument } from '@genoffice/project-contracts'

/** The pipeline stage a diagnostic originates from. Provenance is layered,
 * never flattened into a single "MPP import failed" code. */
export type MppDiagnosticStage = 'sidecar' | 'normalization' | 'mspdi' | 'canonical' | 'scheduling'

/** An `ImportDiagnostic` extended with its pipeline stage of origin. */
export interface MppDiagnostic extends ImportDiagnostic {
  readonly stage: MppDiagnosticStage
}

/** The wire-protocol version of the MPXJ conversion sidecar. The host
 * launcher and the Java sidecar must agree on this number; a mismatch is a
 * `MPP_SIDECAR_RESPONSE_INVALID` error. */
export const MPP_SIDECAR_PROTOCOL_VERSION = 1

/** The pinned MPXJ version the sidecar must run (PROJECT-017 spike evidence
 * base). The host-side fetch script downloads exactly this distribution. */
export const MPXJ_PINNED_VERSION = '16.7.0'

/** MPP format versions supported for import (PROJECT-017 feasibility
 * report §3): MPP8 (Project 98), MPP9 (Project 2000/2002/2003 down-level),
 * MPP12 (Project 2003/2007), MPP14 (Project 2010 and all later versions —
 * the underlying format version remains MPP14). */
export const MPP_SUPPORTED_FORMAT_VERSIONS = ['MPP8', 'MPP9', 'MPP12', 'MPP14'] as const

/** Maximum accepted MPP input size (checked before the sidecar is ever
 * spawned). Aligned with the accepted MSPDI importer's input cap. */
export const MPP_MAX_INPUT_BYTES = 100 * 1024 * 1024 // 100 MiB

/** Maximum accepted sidecar MSPDI output size (checked before the bytes are
 * handed to the importer — the accepted importer would reject anything
 * larger anyway; this cap names the failure as a sidecar-stage error). */
export const MPP_MAX_MSPDI_OUTPUT_BYTES = 100 * 1024 * 1024 // 100 MiB

/** Entity counts reported by the sidecar status frame (conversion summary —
 * informational only; the canonical document is produced by the accepted
 * importer, never by the sidecar). */
export interface MppSidecarCounts {
  readonly tasks: number
  readonly resources: number
  readonly calendars: number
  readonly predecessorLinks: number
  readonly assignments: number
}

/** The sidecar status frame's error object (conversion refused/failed). */
export interface MppSidecarFrameError {
  readonly code: string
  readonly message: string
}

/**
 * The validated sidecar protocol envelope (one JSON object, one line, on
 * stdout). The MSPDI payload itself travels ONLY via the output file —
 * stdout can never contaminate MSPDI parsing.
 */
export interface MppSidecarFrame {
  readonly version: number
  readonly requestId: string
  readonly ok: boolean
  /** Detected source format, e.g. `"MPP14"` (present on success). */
  readonly format?: string
  readonly counts?: MppSidecarCounts
  /** Present when `ok === false`. */
  readonly error?: MppSidecarFrameError
}

/**
 * The outcome of a successful host-managed conversion: the MSPDI XML bytes
 * plus the validated sidecar status frame. This is what crosses the
 * host→foundation boundary (the foundation never spawns a process).
 */
export interface MppConversionOutcome {
  readonly mspdiBytes: Uint8Array
  readonly frame: MppSidecarFrame
  /** Sidecar-stage diagnostics produced while obtaining this outcome
   * (normally empty on success). */
  readonly sidecarDiagnostics: readonly MppDiagnostic[]
}

/** The staged result of importing an MPP-origin MSPDI through the
 * foundation pipeline. */
export interface MppImportResult {
  readonly document: ProjectDocument
  readonly diagnostics: readonly MppDiagnostic[]
}
