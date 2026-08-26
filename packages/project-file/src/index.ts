/**
 * PROJECT-014 — Native `.gproj` file adapter package.
 *
 * Public surface:
 *   - `ProjectFileAdapter`     — the file-adapter contract (host-neutral).
 *   - `gprojFileAdapter`       — the canonical `.gproj` adapter singleton.
 *   - `serializeGproj`/`deserializeGproj`/`inspectGproj` — low-level entry
 *     points for callers that operate on raw `Uint8Array` bytes.
 *   - envelope + diagnostics constants for tests and consumers.
 *
 * The package depends ONLY on `@genoffice/project-contracts` (types + brand
 * helpers) and `@genoffice/project-engine` (the canonical document validator).
 * It has NO React, Electron, Node, browser, HTTP, or MPP/MSPDI dependencies.
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
