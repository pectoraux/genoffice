/**
 * PROJECT-014 — Canonical `.gproj` serializer.
 *
 * `serializeGproj(document)` produces the deterministic byte representation of
 * a canonical `ProjectDocument`. The same `ProjectDocument` always produces
 * byte-identical output (no `Date.now()`, no random UUID, no `localeCompare`,
 * no object-insertion-order dependence — see `canonical.ts`).
 *
 * The serializer does NOT persist derived state. The only persisted authority
 * is the `ProjectDocument` payload; `DerivedSchedule`, `TaskSchedule`,
 * `AssignmentSchedule`, `BaselineVariance`, `LevelingResult`, and the command
 * journal are all re-derived from the canonical document (see
 * `spec/project/requirements.md` PROJECT-014 — Derived state).
 *
 * The serializer is a pure function with no host dependencies.
 */
import { asCalendarId, asISODateTime } from '@genoffice/project-contracts'
import type { ProjectDocument, ProjectFileMetadata } from '@genoffice/project-contracts'
import { canonicalJson } from './canonical.js'
import { GPROJ_FORMAT, GPROJ_FORMAT_VERSION } from './envelope.js'

/** The shape of the on-disk `.gproj` envelope. */
export interface GprojEnvelope {
  readonly document: ProjectDocument
  readonly format: typeof GPROJ_FORMAT
  readonly formatVersion: typeof GPROJ_FORMAT_VERSION
  readonly metadata: ProjectFileMetadata
}

/** Build the canonical `.gproj` envelope for a `ProjectDocument`. */
export function buildGprojEnvelope(document: ProjectDocument): GprojEnvelope {
  // `sourceName` is host context (the filename on disk); it is NOT stored in
  // the file bytes. The host passes it in to `inspect`/`import` and the
  // adapter reports it back via the returned `ProjectFileMetadata`.
  const metadata: ProjectFileMetadata = {
    format: GPROJ_FORMAT,
    version: String(GPROJ_FORMAT_VERSION),
  }
  return {
    document,
    format: GPROJ_FORMAT,
    formatVersion: GPROJ_FORMAT_VERSION,
    metadata,
  }
}

/**
 * Serialize a `ProjectDocument` to canonical `.gproj` bytes.
 *
 * The output is deterministic UTF-8 JSON. Two semantically-equivalent
 * `ProjectDocument` values (including ones whose internal `Record<...>` maps
 * were assembled in different key orders) produce byte-identical output.
 */
export function serializeGproj(document: ProjectDocument): Uint8Array {
  const envelope = buildGprojEnvelope(document)
  return canonicalJson(envelope)
}

/**
 * The canonical "empty" `.gproj` document, returned by `deserializeGproj` when
 * a file-level error (bad JSON, wrong magic, unsupported version) makes it
 * impossible to recover a usable payload. The caller MUST treat the returned
 * document as invalid when an error-level diagnostic is present.
 *
 * The empty document carries a deterministic placeholder `startDate` and an
 * empty default-calendar id so it remains a well-typed `ProjectDocument`
 * without referencing any real calendar (the engine's validator would flag
 * `MISSING_CALENDAR`, which is the correct signal — the document is NOT usable).
 */
export function emptyProjectDocument(): ProjectDocument {
  return {
    schemaVersion: 1,
    properties: {
      id: '',
      name: '',
      startDate: asISODateTime('1970-01-01T00:00:00.000Z'),
      defaultCalendarId: asCalendarId(''),
    },
    tasks: [],
    resources: [],
    assignments: [],
    dependencies: [],
    calendars: [],
    baselines: [],
    customFields: [],
    views: [],
    tables: [],
    filters: [],
    groups: [],
  }
}
