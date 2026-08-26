/**
 * @genoffice/project-mpp-host — public surface.
 *
 * The host-managed side of MPP import (PROJECT-018):
 *   - `MppSidecarLauncher` — one-shot MPXJ conversion process management
 *     (isolation, timeout, memory/size caps, no-shell argument arrays).
 *   - `parseSidecarFrame` — the sidecar wire-protocol frame validator.
 *   - `importMppFromFile` / `importMppFromBytes` — the full staged pipeline
 *     (sidecar → N1–N5 normalization → accepted MSPDI importer → canonical
 *     validation → scheduling) with provenance-preserving diagnostics.
 *
 * This is a HOST package, not a foundation package: it uses Node child
 * process and filesystem APIs by design (architecture-lock §13 keeps
 * process code out of foundation packages; the foundation contract lives in
 * @genoffice/project-file `src/mpp/**`).
 */
export {
  MppSidecarLauncher,
  buildSidecarCommand,
  MPP_DEFAULT_TIMEOUT_MS,
  DEFAULT_MAX_STDERR_LENGTH,
  DEFAULT_MAX_HEAP,
  SPAWN_OPTIONS,
  type MppSidecarLauncherConfig,
  type MppConversionResult,
  type MppConversionSuccess,
  type MppConversionFailure,
} from './launcher.js'
export { parseSidecarFrame } from './protocol.js'
export {
  importMppFromFile,
  importMppFromBytes,
  stageSchedulingDiagnostics,
  type MppFullImportResult,
  type MppImportOptions,
} from './import-mpp.js'
