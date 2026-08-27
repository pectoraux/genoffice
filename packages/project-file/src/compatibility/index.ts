/**
 * PROJECT-020 — compatibility public surface.
 *
 * The canonical import-compatibility layer for `.gproj`, MSPDI, and MPP:
 * the diagnostic model (stages, severity, data-loss classification,
 * recoverability, entity provenance), the deterministic classification
 * table and aggregator, and the three pipeline entry points that pair the
 * accepted import pipelines with a `CompatibilityReport`.
 *
 * The layer lives at the project-file/import boundary — NEVER in
 * React/Electron/web code (architecture-lock §10/§13) — and adds no
 * package dependency: the scheduling package is injected by the
 * caller as a structurally-typed runner.
 */
export {
  type CompatibilityFormat,
  type CompatibilityStage,
  type CompatibilityLoss,
  type CompatibilityRecoverability,
  type CompatibilityEntityType,
  type CompatibilityDiagnostic,
  type CompatibilityImportStatus,
  type CompatibilityValidationStatus,
  type CompatibilitySchedulingStatus,
  type CompatibilityStatus,
  type CompatibilitySaveEligibility,
  type CompatibilityReport,
  COMPATIBILITY_STAGE_ORDER,
  COMPATIBILITY_STAGE_PIPELINE,
  COMPATIBILITY_SEVERITY_ORDER,
  deriveCompatibilityEntityType,
} from './model.js'
export {
  type CompatibilityCodeClassification,
  COMPATIBILITY_CODE_CLASSIFICATIONS,
  classifyImportDiagnostic,
  resolveCompatibilityStage,
} from './classification.js'
export {
  type CompatibilityReportInput,
  buildCompatibilityReport,
  isFatalImport,
} from './aggregate.js'
export {
  type CompatibilityScheduleRunner,
  type CompatibilityOptions,
  type CompatibilityPipelineResult,
  importGprojWithCompatibility,
  importMspdiWithCompatibility,
  importMppWithCompatibility,
} from './pipeline.js'
