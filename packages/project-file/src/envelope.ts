/**
 * PROJECT-014 — Native `.gproj` format envelope.
 *
 * The `.gproj` file is a deterministic, versioned, self-describing JSON envelope
 * wrapping a canonical `ProjectDocument` payload. JSON is parsed with the
 * standard `JSON.parse` (no `eval`, no `Function`, no reviver, no prototype
 * deserialization, no arbitrary class instantiation) and every field is
 * explicitly schema-validated before the canonical `ProjectDocument` is
 * constructed. The format is host-independent: it never stores React/Electron/
 * renderer/UI state, scroll position, canvas coordinates, or host-specific undo
 * history (see `spec/project/requirements.md` PROJECT-014).
 *
 * Envelope shape (canonical key order — alphabetical, enforced by the
 * serializer's stable stringify):
 *
 * ```json
 * {
 *   "document":     { ... ProjectDocument ... },
 *   "format":       "gproj",
 *   "formatVersion": 1,
 *   "metadata":     { "format": "gproj", "version": "1" }
 * }
 * ```
 *
 * - `format` is the magic identifier. It MUST be `"gproj"`; anything else is
 *   rejected as `INVALID_GPROJ`.
 * - `formatVersion` is the FILE-format version (an integer). Bumped when the
 *   envelope or schema evolves. The current version is `1`.
 * - `metadata` is a `ProjectFileMetadata` block (`{ format, version, sourceName? }`).
 *   `metadata.format` echoes `format`; `metadata.version` echoes `formatVersion`
 *   as a string. `sourceName` is host-supplied (the filename) and is NOT stored
 *   in the file bytes — it is passed in via the `metadata` parameter to
 *   `inspect`/`import` so the host can report it back.
 * - `document` is the canonical `ProjectDocument` payload. It is the sole
 *   authoritative persisted state; no `DerivedSchedule`, `TaskSchedule`,
 *   `AssignmentSchedule`, `BaselineVariance`, or `LevelingResult` is persisted
 *   (all are re-derived from the canonical document via the scheduling engine).
 *
 * `ProjectDocument.schemaVersion` (the payload's own schema marker, currently
 * `1`) is distinct from `formatVersion` (the envelope's version). The payload
 * schema version is owned by `@genoffice/project-contracts`; the envelope format
 * version is owned by this package.
 */

/** Magic format identifier. Anything other than this value is `INVALID_GPROJ`. */
export const GPROJ_FORMAT = 'gproj' as const

/** Current `.gproj` file-format version. */
export const GPROJ_FORMAT_VERSION = 1 as const

/** Versions this parser can read. Future versions are rejected as
 * `UNSUPPORTED_GPROJ_VERSION` (no silent forward-read). */
export const GPROJ_SUPPORTED_READ_VERSIONS: readonly number[] = [1] as const

/** Maximum accepted input size (bytes). Guards against oversized payloads. */
export const GPROJ_MAX_INPUT_BYTES = 100 * 1024 * 1024 // 100 MiB

/** Maximum accepted depth of a parsed JSON value (guards against
 * pathologically nested structures). */
export const GPROJ_MAX_PARSE_DEPTH = 64

/** Keys that are never semantically valid on a parsed `.gproj` object and are
 * rejected to harden against prototype-pollution-style payloads. */
export const GPROJ_FORBIDDEN_KEYS: readonly string[] = [
  '__proto__',
  'constructor',
  'prototype',
] as const

/** Type-narrowing helper for the supported-format-version set. */
export function isSupportedGprojVersion(version: number): boolean {
  return (GPROJ_SUPPORTED_READ_VERSIONS as readonly number[]).includes(version)
}
