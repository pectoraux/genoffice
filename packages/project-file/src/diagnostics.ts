/**
 * PROJECT-014 — `.gproj` import/export diagnostic codes.
 *
 * These extend the existing `ImportDiagnostic` contract from
 * `@genoffice/project-contracts` (whose `code` field is a plain `string`, so
 * adding codes does NOT change the contract surface — the smallest compatible
 * extension per the PROJECT-014 brief: "Do not invent duplicated diagnostic
 * systems").
 *
 * The codes distinguish file-format failures (envelope/version/schema) from
 * entity-level failures (task/resource/assignment/dependency/calendar/baseline
 * identity + reference integrity). After the adapter constructs a canonical
 * `ProjectDocument`, it delegates semantic validation to the engine's canonical
 * `validateProjectDocument` and surfaces those diagnostics too (mapped to
 * `severity: 'error'`); the engine's codes (e.g. `DUPLICATE_TASK_ID`,
 * `MISSING_TASK_REFERENCE`, `CALENDAR_PERIOD_MALFORMED`, `CALENDAR_CYCLE`)
 * are passed through verbatim because they are already valid
 * `ImportDiagnostic.code` strings.
 */
export const INVALID_GPROJ = 'INVALID_GPROJ' as const
/** The file parsed as JSON but the envelope `format` is not `"gproj"` or the
 * top-level structure is not a JSON object. */
export const UNSUPPORTED_GPROJ_VERSION = 'UNSUPPORTED_GPROJ_VERSION' as const
/** The envelope parsed but `formatVersion` is not in the supported read set. */
export const SCHEMA_INVALID = 'SCHEMA_INVALID' as const
/** The envelope parsed but a structural slot has the wrong type or an unknown
 * field appears where a specific shape is required. */
export const MISSING_REQUIRED_FIELD = 'MISSING_REQUIRED_FIELD' as const
/** A required field on an entity is absent. */
export const INVALID_IDENTITY = 'INVALID_IDENTITY' as const
/** A branded identity field (`TaskId`/`ResourceId`/…) has the wrong primitive
 * type or is empty. */
export const INVALID_REFERENCE = 'INVALID_REFERENCE' as const
/** A reference field (dependency predecessor/successor, assignment task/resource,
 * calendar base, baseline task snapshot key, …) points to a non-existent entity. */
export const INVALID_CALENDAR = 'INVALID_CALENDAR' as const
/** A `Calendar` has a malformed working-week day key, a malformed exception, or
 * a period with non-integer/out-of-range/empty/overlapping bounds. */
export const INVALID_BASELINE = 'INVALID_BASELINE' as const
/** A `Baseline` has a malformed `capturedAt` or a snapshot with invalid types. */
export const INVALID_ASSIGNMENT = 'INVALID_ASSIGNMENT' as const
/** An `Assignment` has non-numeric work/cost/units or a malformed id. */
export const INVALID_TASK = 'INVALID_TASK' as const
/** A `Task` has a non-numeric duration/priority/percentComplete, an unknown
 * `taskType`/`constraintType`, or a malformed id. */
export const INVALID_RESOURCE = 'INVALID_RESOURCE' as const

/** The full set of PROJECT-014 adapter-emitted diagnostic codes. */
export const GPROJ_DIAGNOSTIC_CODES = [
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
] as const
