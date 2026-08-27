/**
 * PROJECT-014 / PROJECT-015 / PROJECT-016 — Native `.gproj` adapter + MSPDI
 * XML import/export.
 *
 * Public surface:
 *   - `ProjectFileAdapter`     — the file-adapter contract (host-neutral).
 *   - `gprojFileAdapter`       — the canonical `.gproj` adapter singleton.
 *   - `serializeGproj`/`deserializeGproj`/`inspectGproj` — low-level `.gproj`
 *     entry points for callers that operate on raw `Uint8Array` bytes.
 *   - `mspdiFileAdapter`       — the MSPDI XML adapter (import + inspect +
 *     export — PROJECT-016 added deterministic `ProjectDocument` → MSPDI XML
 *     export behind the same adapter boundary).
 *   - `importMspdi`/`inspectMspdi`/`exportMspdi` — low-level MSPDI entry points.
 *   - envelope + diagnostics constants for tests and consumers.
 *
 * The package depends ONLY on `@genoffice/project-contracts` (types + brand
 * helpers) and `@genoffice/project-engine` (the canonical document validator).
 * It has NO React, Electron, Node, browser, HTTP, or MPP/MSPDI-runtime
 * dependencies (the MSPDI parser and writer are pure-TypeScript modules
 * shipped here).
 */
import type {
  ImportDiagnostic,
  ProjectDocument,
  ProjectFileMetadata,
  ProjectSavePlan,
} from '@genoffice/project-contracts'

/** Host-neutral file-adapter contract. Implemented by `gprojFileAdapter`. */
export interface ProjectFileAdapter {
  readonly format: ProjectSavePlan['format']
  inspect(input: Uint8Array, metadata?: ProjectFileMetadata): ProjectFileMetadata
  import(
    input: Uint8Array,
    metadata?: ProjectFileMetadata,
  ): { document: ProjectDocument; diagnostics: ImportDiagnostic[] }
  export(document: ProjectDocument): { bytes: Uint8Array; diagnostics: ImportDiagnostic[] }
}

export { gprojFileAdapter } from './adapter.js'
export { serializeGproj, deserializeGproj, inspectGproj } from './deserialize.js'
export { buildGprojEnvelope, emptyProjectDocument } from './serialize.js'
export type { GprojEnvelope } from './serialize.js'
export type { GprojImportResult } from './deserialize.js'
export {
  GPROJ_FORMAT,
  GPROJ_FORMAT_VERSION,
  GPROJ_SUPPORTED_READ_VERSIONS,
  GPROJ_MAX_INPUT_BYTES,
  GPROJ_MAX_PARSE_DEPTH,
  GPROJ_FORBIDDEN_KEYS,
  isSupportedGprojVersion,
} from './envelope.js'
export {
  INVALID_GPROJ,
  UNSUPPORTED_GPROJ_VERSION,
  SCHEMA_INVALID,
  MISSING_REQUIRED_FIELD,
  INVALID_IDENTITY,
  INVALID_REFERENCE,
  INVALID_CALENDAR,
  INVALID_BASELINE,
  INVALID_ASSIGNMENT,
  INVALID_TASK,
  INVALID_RESOURCE,
  GPROJ_DIAGNOSTIC_CODES,
} from './diagnostics.js'

// PROJECT-015 — MSPDI XML importer (import + inspect only; no export).
export * from './mspdi/index.js'

// PROJECT-018 — MPP import foundation contract (host-neutral types, the
// N1–N5 MSPDI normalizations, staged diagnostics, and the canonical import
// entry point consuming a host-managed MPXJ sidecar conversion outcome).
// No process/binary/Java code lives here (architecture-lock §13).
export * from './mpp/index.js'

// PROJECT-020 — canonical import-compatibility layer: the diagnostic model
// (stages, severity, data-loss classification, recoverability, entity
// provenance), the deterministic classification/aggregation, and pipeline
// entry points pairing every accepted import path with a
// CompatibilityReport. No new package dependency (the scheduler is injected
// by the caller).
export * from './compatibility/index.js'
