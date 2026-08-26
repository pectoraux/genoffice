/**
 * PROJECT-015 — MSPDI format envelope constants.
 *
 * MSPDI (Microsoft Project Data Interchange) is an XML format. Unlike the
 * native `.gproj` envelope, MSPDI does not carry an explicit single "format
 * version" integer on its root; instead the root element is `<Project>` in
 * the MSPDI namespace and carries a `<SaveVersion>` child integer (12 =
 * Project 2007, 14 = 2010, 15 = 2013, 16 = 2016/2019, …). The importer treats
 * `<SaveVersion>` as the version signal: known save versions are read; an
 * unknown future `<SaveVersion>` is rejected as `UNSUPPORTED_MSPDI_VERSION`
 * (no silent forward-read of an unrecognised schema — same discipline as the
 * `.gproj` `formatVersion` gate). An absent `<SaveVersion>` is tolerated
 * (very old or third-party exports) and imported against the current schema.
 *
 * The envelope is otherwise the same host-independent, byte-safe shape as the
 * `.gproj` adapter: a byte-size guard, a depth guard, and a decoded-text-size
 * guard (entity-expansion guard) live in `./xml-parser.ts`.
 */
import type { ProjectSavePlan } from '@genoffice/project-contracts'

/** Magic format identifier for the MSPDI adapter (reserved by `ProjectSavePlan.format`). */
export const MSPDI_FORMAT = 'mspdi' as const satisfies ProjectSavePlan['format']

/** The importer's own schema version (analogous to `.gproj` `formatVersion`).
 * Bumped when the MSPDI→canonical mapping evolves. */
export const MSPDI_FORMAT_VERSION = 1 as const

/** MSPDI `<SaveVersion>` values this importer can read. Future values are
 * rejected as `UNSUPPORTED_MSPDI_VERSION` unless explicitly added here. */
export const MSPDI_SUPPORTED_SAVE_VERSIONS: readonly number[] = [12, 14, 15, 16] as const

/** The expected MSPDI root local element name. Anything else is `INVALID_MSPDI`. */
export const MSPDI_ROOT_ELEMENT = 'Project'

/** The MSPDI namespace URI. The importer does not require the namespace to be
 * declared (it parses by local element name) but reports it for transparency. */
export const MSPDI_NS = 'http://schemas.microsoft.com/project'

/** Type-narrowing helper for the supported-save-version set. */
export function isSupportedMspdiSaveVersion(version: number): boolean {
  return (MSPDI_SUPPORTED_SAVE_VERSIONS as readonly number[]).includes(version)
}
