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

// ---- PROJECT-020 — MSPDI import provenance additions -----------------------
//
// Two known PROJECT-015 import limitations previously produced NO diagnostic
// (silent behavior gaps). PROJECT-020 exposes their provenance — the minimal
// parser-side addition sanctioned by "unless required to expose accurate
// diagnostic provenance". Both are additive: no accepted code is renamed, no
// mapping changes, only warnings that were previously silent become explicit.

export const MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED =
  'MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED' as const
/** PROJECT-020 — a task carries a non-zero `<PhysicalPercentComplete>` which
 * the canonical import does not reconstruct (the accepted PROJECT-015
 * importer never read it; PROJECT-016's exporter emits it for fidelity and
 * warns about the same round-trip gap). Warning: the value is dropped on
 * import — an honest, diagnosed loss, no longer silent. A zero value loses
 * nothing (the canonical default) and is not diagnosed. */

export const MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED =
  'MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED' as const
/** PROJECT-020 — MSPDI carries no per-baseline captured date, so a created
 * baseline's canonical `capturedAt` is approximated from the documented
 * deterministic fallback chain (`<LastSaved>` → `<CreationDate>` → the
 * project `<StartDate>` → the epoch default). One warning per created
 * baseline slot, naming the fallback source actually used. */

// ---- PROJECT-016 — MSPDI EXPORT diagnostic codes -------------------------
//
// Export has its own code family (the importer's codes describe MSPDI→canonical
// failures; export describes canonical→MSPDI ones). Like the import codes, they
// are plain-string `ImportDiagnostic.code` values — the smallest compatible
// diagnostic extension (the `ImportDiagnostic` contract is unchanged).

export const INVALID_MSPDI_EXPORT = 'INVALID_MSPDI_EXPORT' as const
/** The canonical `ProjectDocument` failed `validateProjectDocument`, so export
 * is REFUSED (zero bytes returned; the engine's diagnostics are surfaced as
 * error-level entries alongside this code). Also used when a canonical task /
 * resource `uid` is not a non-negative integer — MSPDI UIDs are — and the
 * exporter had to synthesize a deterministic replacement uid. */
export const INVALID_MSPDI_EXPORT_LAG = 'INVALID_MSPDI_EXPORT_LAG' as const
/** A canonical `lagMinutes` cannot be represented exactly as
 * `LinkLag = lagMinutes × 10` (outside the safe-integer range). The dependency
 * is retained at lag 0 with this error diagnostic — lag semantics are never
 * silently changed. */
export const UNREPRESENTABLE_MSPDI_VALUE = 'UNREPRESENTABLE_MSPDI_VALUE' as const
/** A canonical value has no faithful MSPDI round-trip through the accepted
 * PROJECT-015 importer: a calendar/assignment/baseline identity that does not
 * follow the deterministic import mapping (remapped consistently, semantics
 * preserved but the id string changes), a working period ending at 24:00
 * (dropped by the importer's whole-minute HH:MM:SS rule), divergent baseline
 * `capturedAt` values (MSPDI carries a single `<LastSaved>` carrier), an empty
 * name (the importer substitutes a placeholder), a non-integer or negative
 * duration, or a string custom-field value the importer will re-parse as a
 * number/boolean. The value is still emitted honestly (except where noted);
 * the diagnostic makes the round-trip limitation explicit. */
export const UNSUPPORTED_MSPDI_EXPORT_FEATURE = 'UNSUPPORTED_MSPDI_EXPORT_FEATURE' as const
/** Canonical state that is emitted for MSPDI fidelity but is NOT reconstructed
 * by the accepted PROJECT-015 importer (round-trip limitation, warning): task
 * `physicalPercentComplete`, multiple notes collapsed into the single MSPDI
 * `<Notes>` field, an inconsistent `task.baseline` reverse index, view/table/
 * filter/group definitions with no MSPDI representation, or a task array order
 * canonicalized to hierarchical DFS order. */
export const MSPDI_EXPORT_NORMALIZED = 'MSPDI_EXPORT_NORMALIZED' as const
/** Info-level note that a derived-calendar weekday set was materialized from
 * the inheritance chain (the accepted importer materializes all seven weekday
 * keys, so a partial canonical `workingWeek` is normalized to its resolved
 * form; the resolved semantics are exactly recoverable). */
export const MSPDI_WRITTEN = 'MSPDI_WRITTEN' as const
/** Info-level "write succeeded" diagnostic emitted at the end of a successful
 * export, mirroring `MSPDI_READ` / `GPROJ_READ`. */

/** The full set of PROJECT-015 import adapter-emitted diagnostic codes
 * (including the two PROJECT-020 provenance additions). */
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
  MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED,
  MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED,
] as const

/** The full set of PROJECT-016 export adapter-emitted diagnostic codes. */
export const MSPDI_EXPORT_DIAGNOSTIC_CODES = [
  INVALID_MSPDI_EXPORT,
  INVALID_MSPDI_EXPORT_LAG,
  UNREPRESENTABLE_MSPDI_VALUE,
  UNSUPPORTED_MSPDI_EXPORT_FEATURE,
  MSPDI_EXPORT_NORMALIZED,
  MSPDI_WRITTEN,
] as const
