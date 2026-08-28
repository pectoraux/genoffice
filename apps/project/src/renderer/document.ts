/**
 * PROJECT-027 — the host document module: the new-document template and the
 * open/save composition over the CANONICAL file adapters.
 *
 * The template is document DATA, never semantics: a standard Mon–Fri
 * 9:00–17:00 working week (the same shape the scheduling fixtures use) and
 * a fixed Monday project start — the host computes no dates, and the
 * scheduler derives everything else.
 *
 * Open/save ALWAYS run through `@genoffice/project-file`'s accepted
 * adapters (`.gproj` canonical, MSPDI XML interchange). The host adds NO
 * serialization of its own: documents and bytes cross through
 * `adapter.import` / `adapter.export` verbatim, diagnostics included.
 */
import type {
  Calendar,
  CalendarPeriod,
  ImportDiagnostic,
  ProjectDocument,
} from '@genoffice/project-contracts'
import { asCalendarId, asISODateTime } from '@genoffice/project-contracts'
import { emptyProjectDocument, gprojFileAdapter, mspdiFileAdapter } from '@genoffice/project-file'
import type { ProjectFileAdapter } from '@genoffice/project-file'

/** The canonical save formats this host round-trips (PROJECT-014/015/016). */
export type HostFileFormat = 'gproj' | 'mspdi'

export interface ImportedProject {
  readonly document: ProjectDocument
  readonly diagnostics: readonly ImportDiagnostic[]
  readonly format: HostFileFormat
}

export type ImportOutcome =
  | { readonly kind: 'imported'; readonly imported: ImportedProject }
  | {
      readonly kind: 'error'
      readonly message: string
      readonly diagnostics: readonly ImportDiagnostic[]
    }

/** The standard working week: Mon–Fri, 09:00–17:00, in UTC minute offsets
 * (the canonical time model, architecture-lock §5). Pure data. */
const STANDARD_WORKING_WEEK: Record<number, CalendarPeriod[]> = {
  0: [],
  1: [{ startMinute: 540, endMinute: 1020 }],
  2: [{ startMinute: 540, endMinute: 1020 }],
  3: [{ startMinute: 540, endMinute: 1020 }],
  4: [{ startMinute: 540, endMinute: 1020 }],
  5: [{ startMinute: 540, endMinute: 1020 }],
  6: [],
}

export const STANDARD_CALENDAR_ID = asCalendarId('standard')
/** A fixed Monday project start (2026-01-05) — deterministic template
 * data; the host computes no dates. */
const TEMPLATE_START = '2026-01-05T00:00:00.000Z'

/**
 * The new-document template: the canonical empty document plus the standard
 * calendar wired as the default and a fixed Monday start. Deterministic —
 * the same call always produces the same document (no wall clock).
 */
export function newProjectDocument(name = 'Project1'): ProjectDocument {
  const calendar: Calendar = {
    id: STANDARD_CALENDAR_ID,
    name: 'Standard',
    workingWeek: STANDARD_WORKING_WEEK,
    exceptions: [],
  }
  const base = emptyProjectDocument()
  return {
    ...base,
    properties: {
      ...base.properties,
      name,
      startDate: asISODateTime(TEMPLATE_START),
      defaultCalendarId: STANDARD_CALENDAR_ID,
    },
    calendars: [calendar],
  }
}

/** The canonical adapter for a host format. */
export function adapterForFormat(format: HostFileFormat): ProjectFileAdapter {
  return format === 'gproj' ? gprojFileAdapter : mspdiFileAdapter
}

/** The host format for a file path (by extension); null when unsupported. */
export function formatForPath(path: string): HostFileFormat | null {
  const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
  if (extension === 'gproj') return 'gproj'
  if (extension === 'xml') return 'mspdi'
  return null
}

/** The conventional extension for a host format. */
export function extensionForFormat(format: HostFileFormat): string {
  return format === 'gproj' ? 'gproj' : 'xml'
}

/**
 * Imports raw bytes through the canonical adapter chosen by the file path's
 * extension. NEVER throws for bad input: the accepted adapters return
 * error-severity diagnostics instead (the `.gproj` contract), so a failed
 * open is an `error` outcome carrying the adapter's own diagnostics; a
 * successful open may still carry warning/info diagnostics (shown by the
 * host's diagnostics surface, verbatim).
 */
export function importDocumentBytes(path: string, bytes: Uint8Array): ImportOutcome {
  const format = formatForPath(path)
  if (format === null) {
    return { kind: 'error', message: `Unsupported project file: ${path}`, diagnostics: [] }
  }
  const adapter = adapterForFormat(format)
  let result: { document: ProjectDocument; diagnostics: ImportDiagnostic[] }
  try {
    result = adapter.import(bytes)
  } catch (error) {
    return {
      kind: 'error',
      message: error instanceof Error ? error.message : String(error),
      diagnostics: [],
    }
  }
  const errors = result.diagnostics.filter((diagnostic) => diagnostic.severity === 'error')
  if (errors.length > 0) {
    return {
      kind: 'error',
      message: errors[0]?.message ?? 'The project file could not be imported.',
      diagnostics: result.diagnostics,
    }
  }
  return {
    kind: 'imported',
    imported: { document: result.document, diagnostics: result.diagnostics, format },
  }
}

/**
 * Exports a canonical document through the canonical adapter. The bytes are
 * the adapter's own deterministic output — the host serializes nothing.
 */
export function exportDocumentBytes(
  document: ProjectDocument,
  format: HostFileFormat,
): { bytes: Uint8Array; diagnostics: readonly ImportDiagnostic[] } {
  const adapter = adapterForFormat(format)
  return adapter.export(document)
}

/** The default file name for the save dialog (presentation only). */
export function defaultFileNameFor(document: ProjectDocument, format: HostFileFormat): string {
  const name = document.properties.name.trim() || 'Untitled'
  return `${name}.${extensionForFormat(format)}`
}
