/**
 * PROJECT-018 — MPP import diagnostic codes (foundation contract).
 *
 * Two families, both staged (see `contract.ts` for the stage model):
 *
 *   1. Normalization codes (`stage: 'normalization'`) — the five
 *      PROJECT-017-approved mechanical normalizations N1–N5 applied to the
 *      MPXJ-generated MSPDI before it enters the accepted PROJECT-015
 *      importer. Every rewrite/drop is diagnosed; nothing is silent.
 *
 *   2. Sidecar/protocol codes (`stage: 'sidecar'`) — the wire-protocol error
 *      conditions of the host-managed MPXJ conversion sidecar. The
 *      foundation DEFINES the codes (they are part of the adapter contract);
 *      the host-side launcher package PRODUCES them.
 *
 * MSPDI-level diagnostics keep their accepted PROJECT-015 codes
 * (`stage: 'mspdi'`), canonical validation keeps its engine codes
 * (`stage: 'canonical'`), and scheduling keeps its scheduling-engine codes
 * (`stage: 'scheduling'`) — provenance is layered, never flattened.
 */

// ── Normalization (N1–N5) ──────────────────────────────────────────────

/** N1 — a `-1` sentinel `CalendarUID` ("no calendar") on a task/resource was
 * stripped (it means "inherit the default calendar", not a dangling
 * reference). */
export const MPP_NORMALIZED_SENTINEL_REFERENCE = 'MPP_NORMALIZED_SENTINEL_REFERENCE'

/** N2 — a `-1` sentinel `BaseCalendarUID` ("no base calendar") on a calendar
 * was stripped. Root-cause twin of N1, handled in the same pass (PROJECT-017
 * feasibility report §10) but diagnosed separately for traceability. */
export const MPP_NORMALIZED_BASE_CALENDAR_SENTINEL = 'MPP_NORMALIZED_BASE_CALENDAR_SENTINEL'

/** N3 — a hidden placeholder record was filtered: the MPP summary artifact
 * task (`UID 0` / `OutlineLevel 0` / `WBS "0"`) or the analogous null-name
 * placeholder resource (documented MPXJ behavior). */
export const MPP_NORMALIZED_PLACEHOLDER_RECORD = 'MPP_NORMALIZED_PLACEHOLDER_RECORD'

/** N4 — a `WorkingTime` period running "until midnight" (`ToTime 00:00:00`,
 * the Microsoft/MPXJ convention) was rewritten to the ISO-8601 day-end
 * expression `24:00:00` (canonical `endMinute: 1440`). */
export const MPP_NORMALIZED_MIDNIGHT_PERIOD = 'MPP_NORMALIZED_MIDNIGHT_PERIOD'

/** N5 — an "unassigned" placeholder assignment (`ResourceUID -65535`) was
 * dropped: the canonical model has no unassigned assignment (expected,
 * diagnosed loss — severity warning, never silent). */
export const MPP_DROPPED_UNASSIGNED_ASSIGNMENT = 'MPP_DROPPED_UNASSIGNED_ASSIGNMENT'

/** All PROJECT-018 MSPDI-normalization diagnostic codes (N1–N5). */
export const MPP_NORMALIZATION_DIAGNOSTIC_CODES = [
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
] as const

// ── Sidecar / protocol ─────────────────────────────────────────────────

/** The MPXJ conversion sidecar could not be started (java executable or the
 * pinned MPXJ distribution is missing/not executable). */
export const MPP_SIDECAR_UNAVAILABLE = 'MPP_SIDECAR_UNAVAILABLE'

/** The sidecar exceeded the conversion timeout and was terminated. */
export const MPP_SIDECAR_TIMEOUT = 'MPP_SIDECAR_TIMEOUT'

/** The sidecar exited with a nonzero exit code (conversion failure). */
export const MPP_SIDECAR_EXIT = 'MPP_SIDECAR_EXIT'

/** The sidecar's stdout status frame is not a valid protocol envelope. */
export const MPP_SIDECAR_RESPONSE_INVALID = 'MPP_SIDECAR_RESPONSE_INVALID'

/** The input exceeds `MPP_MAX_INPUT_BYTES` (checked before the sidecar is
 * ever spawned). */
export const MPP_INPUT_TOO_LARGE = 'MPP_INPUT_TOO_LARGE'

/** The sidecar's MSPDI output exceeds `MPP_MAX_MSPDI_OUTPUT_BYTES`
 * (checked before the bytes are handed to the importer). */
export const MPP_OUTPUT_TOO_LARGE = 'MPP_OUTPUT_TOO_LARGE'

/** The sidecar reported the input as an unsupported/unknown project format
 * (not one of MPP8/MPP9/MPP12/MPP14 — e.g. a non-MPP file or an
 * unrecognizable/encrypted container). Reported through the protocol frame. */
export const MPP_UNSUPPORTED_FORMAT = 'MPP_UNSUPPORTED_FORMAT'

/** All PROJECT-018 sidecar/protocol diagnostic codes. */
export const MPP_SIDECAR_DIAGNOSTIC_CODES = [
  MPP_SIDECAR_UNAVAILABLE,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_INPUT_TOO_LARGE,
  MPP_OUTPUT_TOO_LARGE,
  MPP_UNSUPPORTED_FORMAT,
] as const

/** All PROJECT-018 MPP diagnostic codes (normalization + sidecar). */
export const MPP_DIAGNOSTIC_CODES = [
  ...MPP_NORMALIZATION_DIAGNOSTIC_CODES,
  ...MPP_SIDECAR_DIAGNOSTIC_CODES,
] as const
