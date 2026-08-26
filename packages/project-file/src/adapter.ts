/**
 * PROJECT-014 — The canonical `.gproj` `ProjectFileAdapter`.
 *
 * `gprojFileAdapter` is a stateless singleton implementing the existing
 * `ProjectFileAdapter` contract (defined in `./index.ts`, against
 * `@genoffice/project-contracts`). It is the sole native-format adapter for
 * the canonical `ProjectDocument`. It does NOT depend on MPP or MSPDI
 * internals (those are PROJECT-015..019) and it carries no renderer/Electron/
 * web state (those are PROJECT-021..031).
 *
 * Architecture (`spec/project/architecture-lock.md` §10 + §13):
 *   ProjectDocument → ProjectFileAdapter → .gproj
 *   .gproj → ProjectFileAdapter → ProjectDocument
 *
 * The adapter is host-independent: it operates on `Uint8Array` (no Node `fs`,
 * no browser `File`, no Electron dialog state).
 */
import type { ProjectDocument, ProjectFileMetadata } from '@genoffice/project-contracts'
import type { ProjectFileAdapter } from './index.js'
import { GPROJ_FORMAT } from './envelope.js'
import { deserializeGproj, inspectGproj } from './deserialize.js'
import { serializeGproj } from './serialize.js'

/** The canonical `.gproj` file adapter. Stateless; safe to share. */
export const gprojFileAdapter: ProjectFileAdapter = {
  format: GPROJ_FORMAT,
  inspect(input: Uint8Array, metadata?: ProjectFileMetadata): ProjectFileMetadata {
    return inspectGproj(input, metadata)
  },
  import(input: Uint8Array, metadata?: ProjectFileMetadata) {
    return deserializeGproj(input, metadata)
  },
  export(document: ProjectDocument) {
    return {
      bytes: serializeGproj(document),
      diagnostics: [] as import('@genoffice/project-contracts').ImportDiagnostic[],
    }
  },
}
