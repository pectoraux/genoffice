/**
 * PROJECT-015 — MSPDI import diagnostic codes.
 *
 * These extend the existing `ImportDiagnostic` contract from
 * `@genoffice/project-contracts` (whose `code` field is a plain `string`, so
 * adding codes does NOT change the contract surface — the smallest compatible
 * extension, mirroring the PROJECT-014 `.gproj` diagnostics module).
 *
 * The codes distinguish MSPDI-format failures (envelope/version/malformed-XML)
 * from entity-level failures (task/resource/assignment/dependency/calendar/
 * baseline/constraint identity, reference, date, and duration integrity). The
 * importer delegates semantic validation to the engine's canonical
 * `validateProjectDocument` and surfaces those diagnostics too (mapped to
 * `severity: 'error'`); the engine's codes (`DUPLICATE_TASK_ID`,
 * `MISSING_TASK_REFERENCE`, `PARENT_CYCLE`, `DEPENDENCY_CYCLE`,
 * `CALENDAR_CYCLE`, …) pass through verbatim because they are already valid
 * `ImportDiagnostic.code` strings.
 *
 * Per the PROJECT-015 brief: every dropped or approximated MSPDI feature is
 * named by a diagnostic — never silently discarded. `UNSUPPORTED_MSPDI_FEATURE`
 * covers semantically-significant MSPDI constructs that have no faithful
 * canonical representation (elapsed durations, percentage lags, recurring
 * calendar exceptions, future MSPDI features); the importer still constructs a
 * valid canonical document and emits this warning/error so the loss is
 * explicit and actionable.
 */
export const INVALID_MSPDI = 'INVALID_MSPDI' as const
/** The input is not well-formed XML, or the root is not a `<Project>`
 * element, or a byte/depth/entity-expansion guard tripped. File-level
 * failure: the importer returns the empty document + a single error. */
export const UNSUPPORTED_MSPDI_VERSION = 'UNSUPPORTED_MSPDI_VERSION' as const
/** `<SaveVersion>` is present but not in the supported read set (no silent
 * forward-read of a future MSPDI schema). */
export const UNSUPPORTED_MSPDI_FEATURE = 'UNSUPPORTED_MSPDI_FEATURE' as const
/** An MSPDI feature that has no faithful canonical representation appeared
 * (elapsed duration, percentage lag, recurring/multi-day calendar exception,
 * unsupported constraint, …). The importer constructs a valid canonical
 * document and emits this so the loss is explicit. */
export const INVALID_MSPDI_REFERENCE = 'INVALID_MSPDI_REFERENCE' as const
/** An MSPDI cross-entity reference (PredecessorUID, TaskUID, ResourceUID,
 * CalendarUID, BaseCalendarUID) does not resolve to a declared entity. */
export const INVALID_MSPDI_DATE = 'INVALID_MSPDI_DATE' as const
/** An MSPDI date/time is not a valid ISO-8601 value and cannot be normalized
 * to canonical UTC. */
export const INVALID_MSPDI_DURATION = 'INVALID_MSPDI_DURATION' as const
/** An MSPDI ISO-8601 duration or `<LinkLag>` cannot be faithfully converted
 * to integer `WorkingMinutes` (e.g. sub-minute remainder, ambiguous elapsed
 * vs working time). */
export const INVALID_MSPDI_CALENDAR = 'INVALID_MSPDI_CALENDAR' as const
/** An MSPDI `<Calendar>` is structurally malformed (bad DayType, bad working
 * period, circular base-calendar chain). */
export const INVALID_MSPDI_RESOURCE = 'INVALID_MSPDI_RESOURCE' as const
/** An MSPDI `<Resource>` is structurally malformed (bad kind, bad max-units,
 * bad rate, bad availability window). */
export const INVALID_MSPDI_ASSIGNMENT = 'INVALID_MSPDI_ASSIGNMENT' as const
/** An MSPDI `<Assignment>` is structurally malformed (bad units, bad work/
 * cost values). */
export const INVALID_MSPDI_CONSTRAINT = 'INVALID_MSPDI_CONSTRAINT' as const
/** An MSPDI `<ConstraintType>` is outside the supported 0–7 enum, or a
 * date-bounded constraint has no/invalid `<ConstraintDate>`. */
export const MISSING_MSPDI_FIELD = 'MISSING_MSPDI_FIELD' as const
/** A required MSPDI element is absent (e.g. a `<Task>` with no `<UID>`,
 * a `<Resource>` with no `<Name>`). */
export const MSPDI_READ = 'MSPDI_READ' as const
/** Info-level "read succeeded" diagnostic emitted at the end of a successful
 * import, mirroring `GPROJ_READ` from the native adapter. */

/** The full set of PROJECT-015 adapter-emitted diagnostic codes. */
export const MSPDI_DIAGNOSTIC_CODES = [
  INVALID_MSPDI,
  UNSUPPORTED_MSPDI_VERSION,
  UNSUPPORTED_MSPDI_FEATURE,
  INVALID_MSPDI_REFERENCE,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_CALENDAR,
  INVALID_MSPDI_RESOURCE,
  INVALID_MSPDI_ASSIGNMENT,
  INVALID_MSPDI_CONSTRAINT,
  MISSING_MSPDI_FIELD,
  MSPDI_READ,
] as const
