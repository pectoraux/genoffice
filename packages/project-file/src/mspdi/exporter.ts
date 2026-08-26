/**
 * PROJECT-016 — Canonical `ProjectDocument` → deterministic MSPDI XML export.
 *
 * Architectural flow (PROJECT-016 brief):
 *
 *   ProjectDocument → MSPDIAdapter.export → MSPDI XML
 *
 * The canonical `ProjectDocument` remains the source of truth. MSPDI-specific
 * XML structures never become canonical domain state. The exporter lives
 * behind the same `ProjectFileAdapter` boundary as the accepted `.gproj`
 * adapter (PROJECT-014) and MSPDI importer (PROJECT-015) — no second
 * file-adapter abstraction. No React/Electron/browser/Node imports; the XML
 * writer is the pure-TypeScript `./xml-writer.ts` (the serialization analog
 * of the accepted pure-TS parser).
 *
 * ROUND-TRIP INVARIANT (the primary acceptance standard):
 *
 *   canonical ProjectDocument → exportMspdi → importMspdi → canonical
 *   ProjectDocument
 *
 * …must be semantically equivalent for all supported PROJECT-014/015 fields:
 * properties, tasks, hierarchy, identity, dependencies, lagMinutes,
 * constraints, deadlines, progress, calendars, resources, assignments,
 * baselines, custom fields — and the re-derived schedule must match.
 * Byte-identical MSPDI XML is NOT required by the format (many equivalent
 * serializations exist); what IS required — and delivered — is that the SAME
 * `ProjectDocument` always serializes to byte-identical canonical XML.
 *
 * IDENTITY (architecture-lock §4 — "UIDs are persistent source/interoperability
 * identifiers; local IDs are canonical application identities"):
 *
 *   - Task.uid / Resource.uid → MSPDI `<UID>` verbatim (the accepted
 *     PROJECT-015 identity mapping in reverse). `TaskId`/`ResourceId` are
 *     NEVER exported as MSPDI UIDs.
 *   - Calendar / Assignment / Baseline have no numeric uid in the canonical
 *     contract, so their MSPDI identifiers are recovered from the accepted
 *     deterministic import mapping (`'c'+uid` / `'a'+uid` / `'b'+slot`) when
 *     the canonical id parses, and otherwise SYNTHESIZED deterministically
 *     (smallest unused non-negative integer in canonical order). No random
 *     IDs, no clock reads, no array position as identity.
 *   - A task/resource uid that is not a non-negative integer is replaced by a
 *     synthesized uid with an `INVALID_MSPDI_EXPORT` error diagnostic.
 *
 * DETERMINISM (canonical XML):
 *
 *   - Fixed element order everywhere (documented in
 *     `spec/project/requirements.md` PROJECT-016).
 *   - Semantically-ordered collections follow canonical meaning: tasks in
 *     hierarchical DFS order (parents before children; sibling order = the
 *     canonical task array order — sibling order is the semantic content of
 *     task ordering). Order-insignificant collections (calendars, resources,
 *     assignments, dependencies, custom fields, calendar exceptions,
 *     availability periods, working periods) are explicitly sorted so
 *     reordered-but-equivalent canonical inputs produce identical XML bytes.
 *   - No locale-aware comparison (code-point `<` comparisons only), no
 *     randomness, no clock reads, no host timezone, no insertion-order reads
 *     (record keys are always re-sorted).
 *
 * DERIVED STATE: the exporter never consults the derived schedule and never runs
 * a scheduler. The only derived projections it computes are format-mandated
 * and documented: the MSPDI outline number (a projection of `parentTaskId` +
 * sibling order — exactly what the accepted importer reconstructs parents
 * from) and the derived-calendar weekday materialization (the accepted
 * importer materializes all seven weekday keys, so a partial canonical
 * `workingWeek` is normalized to its chain-resolved form — semantics exactly
 * recoverable; no dates or durations are ever computed).
 *
 * NO SILENT LOSS: every dropped, re-projected, or non-reconstructible piece
 * of canonical state is named by an `ImportDiagnostic` (see
 * `./diagnostics.ts` PROJECT-016 codes). A document that fails
 * `validateProjectDocument` is REFUSED (zero bytes + error diagnostics).
 */
import type {
  Baseline,
  Calendar,
  CalendarPeriod,
  ConstraintType,
  Dependency,
  DependencyType,
  ImportDiagnostic,
  ProjectDocument,
  Resource,
  Task,
} from '@genoffice/project-contracts'
import { validateProjectDocument } from '@genoffice/project-engine'
import { encodeUtf8 } from '../utf8.js'
import { XmlWriter } from './xml-writer.js'
import { MSPDI_FORMAT_VERSION, MSPDI_NS } from './envelope.js'
import { LAG_FORMAT_MINUTE } from './conversions.js'
import {
  INVALID_MSPDI_EXPORT,
  INVALID_MSPDI_EXPORT_LAG,
  MSPDI_EXPORT_NORMALIZED,
  MSPDI_WRITTEN,
  UNSUPPORTED_MSPDI_EXPORT_FEATURE,
  UNREPRESENTABLE_MSPDI_VALUE,
} from './diagnostics.js'

export interface MspdiExportResult {
  bytes: Uint8Array
  diagnostics: ImportDiagnostic[]
}

type Sink = { push(d: ImportDiagnostic): void }

function diag(
  sink: Sink,
  code: string,
  severity: ImportDiagnostic['severity'],
  message: string,
  entityId?: string,
): void {
  const entry: ImportDiagnostic = { code, severity, message }
  if (entityId !== undefined) entry.entityId = entityId
  sink.push(entry)
}

// ---- enumeration maps (inverse of the accepted PROJECT-015 import maps) ----

const CONSTRAINT_TYPE_CODES: Record<ConstraintType, number> = {
  asSoonAsPossible: 0,
  asLateAsPossible: 1,
  mustStartOn: 2,
  mustFinishOn: 3,
  startNoEarlierThan: 4,
  finishNoEarlierThan: 5,
  startNoLaterThan: 6,
  finishNoLaterThan: 7,
}

const LINK_TYPE_CODES: Record<DependencyType, number> = {
  FS: 0,
  FF: 1,
  SS: 2,
  SF: 3,
}

const RESOURCE_KIND_CODES: Record<Resource['kind'], number> = {
  work: 1,
  material: 2,
  cost: 3,
}

const TASK_TYPE_CODES: Record<Task['taskType'], number> = {
  fixedUnits: 0,
  fixedDuration: 1,
  fixedWork: 2,
}

/** MSPDI baseline slot element names, index 0..10 (accepted importer set). */
const BASELINE_SLOT_NAMES = [
  'Baseline',
  'Baseline1',
  'Baseline2',
  'Baseline3',
  'Baseline4',
  'Baseline5',
  'Baseline6',
  'Baseline7',
  'Baseline8',
  'Baseline9',
  'Baseline10',
] as const

// ---- deterministic ordering helpers ----------------------------------------

/** Code-point comparison (never locale-aware). */
function compareStrings(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

/** Stable sort: ties keep input order (original index), so the result is a
 * pure function of the input sequence — no sort-implementation dependence. */
function stableSort<T>(items: T[], compare: (a: T, b: T) => number): T[] {
  return items
    .map((item, index) => ({ item, index }))
    .sort((x, y) => compare(x.item, y.item) || x.index - y.index)
    .map((entry) => entry.item)
}

// ---- value formatting --------------------------------------------------------

/**
 * Format `WorkingMinutes` as an MSPDI ISO-8601 duration (`PT8H0M0S`).
 * Non-integer or negative values are emitted in their exact (fractional or
 * signed) form — never rounded — and the caller warns that the accepted
 * importer will reject them (`INVALID_MSPDI_DURATION`, duration 0).
 */
function minutesToIsoDuration(minutes: number): string {
  if (!Number.isFinite(minutes)) return 'PT0H0M0S'
  const sign = minutes < 0 ? '-' : ''
  const abs = Math.abs(minutes)
  const h = Math.floor(abs / 60)
  const m = abs - h * 60
  return `${sign}PT${h}H${m}M0S`
}

/** Whole minutes → `HH:MM:SS` (the accepted importer's whole-minute rule).
 * `1440` (24:00) exceeds the HH:MM:SS hour range and is emitted as
 * `24:00:00` with an `UNREPRESENTABLE_MSPDI_VALUE` warning — the importer
 * drops such a period on re-import rather than rounding it. */
function minutesToTime(minutes: number): string {
  const h = Math.floor(minutes / 60)
  const m = minutes % 60
  const pad = (n: number): string => String(n).padStart(2, '0')
  return `${pad(h)}:${pad(m)}:00`
}

/** Format a number deterministically (`String(n)`). */
function numberText(value: number): string {
  return String(value)
}

// ---- identity synthesis -------------------------------------------------------

/** Parse a canonical id produced by the accepted import mapping
 * (`'c12'` / `'a3'` / `'b0'`) back into its integer. Returns `undefined` for
 * every other shape (native ids like `'standard'`). */
function parsePrefixedId(id: string, prefix: string): number | undefined {
  if (!id.startsWith(prefix)) return undefined
  const rest = id.slice(prefix.length)
  if (!/^\d+$/.test(rest)) return undefined
  const n = Number(rest)
  return Number.isSafeInteger(n) ? n : undefined
}

interface UidTable {
  /** entityId → MSPDI uid. */
  uids: Map<string, number>
}

/**
 * Deterministic uid synthesis for an entity family without a canonical
 * numeric uid (calendars, assignments, baselines): ids that parse under the
 * accepted import mapping keep their parsed uid (first-parsed wins on
 * collisions — the loser is synthesized with a warning); remaining ids are
 * assigned the smallest unused non-negative integers in the given canonical
 * iteration order. Pure function of the document content — no randomness,
 * no array position as identity.
 */
function synthesizeUids(entityIds: string[], prefix: string, label: string, sink: Sink): UidTable {
  const uids = new Map<string, number>()
  const used = new Set<number>()
  const unmapped: string[] = []
  for (const id of entityIds) {
    const parsed = parsePrefixedId(id, prefix)
    if (parsed === undefined || used.has(parsed)) {
      if (parsed !== undefined) {
        diag(
          sink,
          UNREPRESENTABLE_MSPDI_VALUE,
          'warning',
          `${label} '${id}' parses to MSPDI uid ${parsed}, which is already taken by another ${label}; a deterministic replacement uid was synthesized`,
          id,
        )
      }
      unmapped.push(id)
      continue
    }
    uids.set(id, parsed)
    used.add(parsed)
  }
  let next = 0
  for (const id of unmapped) {
    while (used.has(next)) next += 1
    uids.set(id, next)
    used.add(next)
    diag(
      sink,
      UNREPRESENTABLE_MSPDI_VALUE,
      'warning',
      `${label} '${id}' does not follow the deterministic import mapping ('${prefix}<uid>'); it is exported as MSPDI uid ${next} and re-imports as '${prefix}${next}' — all references are remapped consistently, but the id string does not round-trip`,
      id,
    )
    next += 1
  }
  return { uids }
}

/** Task/resource uids: verbatim when they are non-negative integers and
 * unique, otherwise synthesized (smallest unused, in canonical order) with
 * an `INVALID_MSPDI_EXPORT` error. */
function assignEntityUids(
  entities: Array<{ id: string; uid: number }>,
  label: string,
  sink: Sink,
): UidTable {
  const uids = new Map<string, number>()
  const used = new Set<number>()
  const unmapped: Array<{ id: string; uid: number }> = []
  for (const entity of entities) {
    if (Number.isInteger(entity.uid) && entity.uid >= 0 && !used.has(entity.uid)) {
      uids.set(entity.id, entity.uid)
      used.add(entity.uid)
      continue
    }
    if (Number.isInteger(entity.uid) && entity.uid >= 0) {
      diag(
        sink,
        INVALID_MSPDI_EXPORT,
        'error',
        `${label} '${entity.id}' duplicates uid ${entity.uid}; a deterministic replacement uid was synthesized`,
        entity.id,
      )
    } else {
      diag(
        sink,
        INVALID_MSPDI_EXPORT,
        'error',
        `${label} '${entity.id}' has uid ${String(entity.uid)}, which is not a non-negative integer MSPDI uid; a deterministic replacement uid was synthesized`,
        entity.id,
      )
    }
    unmapped.push(entity)
  }
  let next = 0
  for (const entity of unmapped) {
    while (used.has(next)) next += 1
    uids.set(entity.id, next)
    used.add(next)
    next += 1
  }
  return { uids }
}

// ---- task ordering + outline projection ---------------------------------------

/**
 * Project the canonical hierarchy (`parentTaskId` + canonical task array
 * order) into MSPDI outline numbers.
 *
 * The exported task order is the deterministic DFS order of the parent
 * forest: parents before children, siblings in canonical array order. The
 * outline number of a task is `parentOutline + '.' + siblingIndex` (top level
 * `'1'`, `'2'`, …) — exactly the form the accepted importer reconstructs
 * `parentTaskId` from, so the exported hierarchy imports back into the same
 * canonical parent relationships. WBS is a deterministic projection, never
 * canonical identity (architecture-lock §4). When the derived outline differs
 * from the canonical `task.wbs` string, the exporter emits the derived value
 * and warns (the hierarchy itself is preserved exactly).
 */
function projectOutline(tasks: Task[]): {
  order: Task[]
  outlines: Map<string, string>
  reordered: boolean
} {
  const idToTask = new Map<string, Task>(tasks.map((t) => [t.id as string, t]))
  // Children grouped per parent in canonical array order; roots keep the
  // array order among themselves.
  const childrenOf = new Map<string, Task[]>()
  const roots: Task[] = []
  for (const task of tasks) {
    const parent = task.parentTaskId !== undefined ? idToTask.get(task.parentTaskId) : undefined
    if (parent === undefined) {
      roots.push(task)
      continue
    }
    const list = childrenOf.get(parent.id as string)
    if (list === undefined) childrenOf.set(parent.id as string, [task])
    else list.push(task)
  }
  const order: Task[] = []
  const outlines = new Map<string, string>()
  let reordered = false
  let emitted = 0
  const visit = (task: Task, number: string): void => {
    emitted += 1
    if (tasks[emitted - 1] !== task) reordered = true
    outlines.set(task.id as string, number)
    order.push(task)
    const children = childrenOf.get(task.id as string) ?? []
    let childIndex = 1
    for (const child of children) {
      visit(child, `${number}.${childIndex}`)
      childIndex += 1
    }
  }
  let rootIndex = 1
  for (const root of roots) {
    visit(root, String(rootIndex))
    rootIndex += 1
  }
  return { order, outlines, reordered }
}

// ---- calendar weekday materialization ------------------------------------------

/**
 * Materialize a calendar's weekday periods for MSPDI emission.
 *
 * The accepted importer materializes ALL SEVEN weekday keys on import, so a
 * canonical derived calendar with a partial `workingWeek` (absent keys
 * inherit the base chain's periods under `resolveCalendar`) cannot be
 * exported raw without changing its resolved semantics. Instead the exporter
 * emits, for every weekday key, the child's own periods when present and
 * otherwise the nearest ancestor's periods for that key (never-working when
 * no ancestor defines it) — a pure record walk over `baseCalendarId` links,
 * NOT a scheduling computation (no dates or durations are derived). The
 * `BaseCalendarUID` link is preserved, and the round-tripped calendar
 * re-resolves to exactly the same working times. This is the sanctioned
 * normalization ("MSPDI requires a normalized representation and the
 * semantics remain exactly recoverable by the accepted importer"); it is
 * surfaced with an `MSPDI_EXPORT_NORMALIZED` info note whenever at least one
 * inherited key with working periods was materialized.
 */
function materializeWeek(
  calendar: Calendar,
  byId: Map<string, Calendar>,
): { week: Record<number, CalendarPeriod[]>; materialized: boolean } {
  const chain: Calendar[] = [calendar]
  const seen = new Set<string>([calendar.id as string])
  let current = calendar
  while (current.baseCalendarId !== undefined && !seen.has(current.baseCalendarId)) {
    const base = byId.get(current.baseCalendarId as string)
    if (base === undefined) break
    seen.add(base.id as string)
    chain.push(base)
    current = base
  }
  const week: Record<number, CalendarPeriod[]> = {}
  let materialized = false
  const hasKey = (cal: Calendar, key: number): boolean =>
    Object.prototype.hasOwnProperty.call(cal.workingWeek, key)
  for (let key = 0; key <= 6; key += 1) {
    if (hasKey(calendar, key)) {
      week[key] = calendar.workingWeek[key]
      continue
    }
    // Inherit from the nearest ancestor that defines this weekday.
    let inherited: CalendarPeriod[] | undefined
    for (const ancestor of chain) {
      if (ancestor !== calendar && hasKey(ancestor, key)) {
        inherited = ancestor.workingWeek[key]
        break
      }
    }
    week[key] = inherited ?? []
    if (inherited !== undefined && inherited.length > 0) materialized = true
  }
  return { week, materialized }
}

// ---- duration emission -----------------------------------------------------------

/** Emit an MSPDI ISO-8601 duration leaf, warning on non-representable
 * (non-integer or negative) working-minute values — never rounding. */
function emitDuration(
  writer: XmlWriter,
  sink: Sink,
  tag: string,
  minutes: number,
  entity: string,
): void {
  if (!Number.isFinite(minutes) || !Number.isInteger(minutes) || minutes < 0) {
    diag(
      sink,
      UNREPRESENTABLE_MSPDI_VALUE,
      'warning',
      `${entity} ${tag.toLowerCase()} ${numberText(minutes)} is not a non-negative integer working-minute value; the exact value is emitted but the accepted importer will reject it (INVALID_MSPDI_DURATION) and default it to 0`,
    )
  }
  writer.leaf(tag, minutesToIsoDuration(minutes))
}

// ---- main entry point --------------------------------------------------------------

/**
 * Export a canonical `ProjectDocument` to deterministic MSPDI XML bytes.
 *
 * Refusal: a document that fails `validateProjectDocument` is NOT exported —
 * the result carries zero bytes, one `INVALID_MSPDI_EXPORT` error, and the
 * engine's diagnostics surfaced verbatim as error-level entries (the
 * exporter never serializes semantically invalid canonical state).
 */
export function exportMspdi(document: ProjectDocument): MspdiExportResult {
  const sink: ImportDiagnostic[] = []

  // ---- refusal gate -------------------------------------------------------
  const validation = validateProjectDocument(document)
  if (!validation.accepted) {
    diag(
      sink,
      INVALID_MSPDI_EXPORT,
      'error',
      'canonical ProjectDocument failed validateProjectDocument; refusing to export semantically invalid state',
    )
    for (const d of validation.diagnostics) {
      sink.push({ code: d.code, severity: 'error', message: d.message })
    }
    return { bytes: new Uint8Array(0), diagnostics: sink }
  }

  const writer = new XmlWriter()
  writer.declaration()

  // ---- identity tables ----------------------------------------------------
  const taskUids = assignEntityUids(
    document.tasks.map((t) => ({ id: t.id as string, uid: t.uid })),
    'task',
    sink,
  )
  const resourceUids = assignEntityUids(
    document.resources.map((r) => ({ id: r.id as string, uid: r.uid })),
    'resource',
    sink,
  )
  // Calendars: parse-or-synthesize in canonical id order (deterministic;
  // distinct ids that parse to the same uid resolve to the code-point-first
  // id, the loser is synthesized with a warning).
  const calendarIdsSorted = document.calendars.map((c) => c.id as string).sort(compareStrings)
  const calendarUids = synthesizeUids(calendarIdsSorted, 'c', 'calendar id', sink)
  // Assignments: canonical iteration order = (task uid, resource uid) — the
  // semantic key of an assignment (the pair is unique per the engine).
  const assignmentsSorted = stableSort(document.assignments, (a, b) => {
    const ta = taskUids.uids.get(a.taskId as string) ?? 0
    const tb = taskUids.uids.get(b.taskId as string) ?? 0
    if (ta !== tb) return ta - tb
    const ra = resourceUids.uids.get(a.resourceId as string) ?? 0
    const rb = resourceUids.uids.get(b.resourceId as string) ?? 0
    if (ra !== rb) return ra - rb
    return compareStrings(a.id as string, b.id as string)
  })
  const assignmentUids = synthesizeUids(
    assignmentsSorted.map((a) => a.id as string),
    'a',
    'assignment id',
    sink,
  )
  // Baselines: parse-or-synthesize in canonical id order.
  const baselineIdsSorted = document.baselines.map((b) => b.id as string).sort(compareStrings)
  const baselineSlots = synthesizeUids(baselineIdsSorted, 'b', 'baseline id', sink)
  const taskUid = (taskId: string): number => taskUids.uids.get(taskId) ?? 0
  const calendarUid = (calendarId: string): number => calendarUids.uids.get(calendarId) ?? 0

  const calendarById = new Map<string, Calendar>(document.calendars.map((c) => [c.id as string, c]))

  // ---- project properties -------------------------------------------------
  const props = document.properties
  writer.open('Project', { xmlns: MSPDI_NS })
  writer.leaf('SaveVersion', '16')
  if (props.id.length > 0) {
    writer.leaf('UID', props.id)
  } else {
    diag(
      sink,
      UNREPRESENTABLE_MSPDI_VALUE,
      'warning',
      'project properties id is empty; the importer will fall back to the project name as the canonical id',
    )
  }
  if (props.name.length === 0) {
    diag(
      sink,
      UNREPRESENTABLE_MSPDI_VALUE,
      'warning',
      'project name is empty; the importer will substitute the placeholder name "Imported MSPDI Project"',
    )
  }
  writer.leaf('Name', props.name)
  writer.leaf('StartDate', props.startDate)
  if (props.finishDate !== undefined) writer.leaf('FinishDate', props.finishDate)
  if (props.statusDate !== undefined) writer.leaf('StatusDate', props.statusDate)

  // ---- baseline capturedAt carrier -----------------------------------------
  // MSPDI carries no per-baseline captured date; the accepted importer falls
  // back `<LastSaved>` → `<CreationDate>` → `<StartDate>`. A uniform
  // capturedAt round-trips exactly through `<LastSaved>`; divergent values
  // are carried by the first (lowest-slot) baseline's capturedAt with
  // per-baseline warnings. No baselines → no carrier (timestamps are never
  // invented merely to make the XML look complete).
  if (document.baselines.length > 0) {
    const slotsSorted = stableSort(document.baselines, (a, b) => {
      const sa = baselineSlots.uids.get(a.id as string) ?? 0
      const sb = baselineSlots.uids.get(b.id as string) ?? 0
      if (sa !== sb) return sa - sb
      return compareStrings(a.id as string, b.id as string)
    })
    const carrier = slotsSorted[0].capturedAt
    writer.leaf('LastSaved', carrier)
    for (const baseline of slotsSorted.slice(1)) {
      if (baseline.capturedAt !== carrier) {
        diag(
          sink,
          UNREPRESENTABLE_MSPDI_VALUE,
          'warning',
          `baseline '${baseline.id}' capturedAt ${baseline.capturedAt} cannot be represented separately (MSPDI carries a single <LastSaved>); it re-imports as ${carrier}`,
          baseline.id as string,
        )
      }
    }
  }

  // ---- calendars (default first, then ascending uid) ------------------------
  if (document.calendars.length > 0) {
    const defaultCalendarId = props.defaultCalendarId as string
    const calendarsSorted = stableSort(document.calendars, (a, b) => {
      const ad = a.id === defaultCalendarId
      const bd = b.id === defaultCalendarId
      if (ad !== bd) return ad ? -1 : 1 // the default calendar is emitted first
      const ua = calendarUid(a.id as string)
      const ub = calendarUid(b.id as string)
      if (ua !== ub) return ua - ub
      return compareStrings(a.id as string, b.id as string)
    })
    writer.open('Calendars')
    for (const calendar of calendarsSorted) {
      const calendarId = calendar.id as string
      writer.open('Calendar')
      writer.leaf('UID', String(calendarUid(calendarId)))
      writer.leaf('Name', calendar.name)
      writer.leaf('IsBaseCalendar', calendar.baseCalendarId === undefined ? 'true' : 'false')
      if (calendar.baseCalendarId !== undefined) {
        writer.leaf('BaseCalendarUID', String(calendarUid(calendar.baseCalendarId as string)))
      }
      writer.leaf('IsBaseCalendarDefault', calendarId === defaultCalendarId ? 'true' : 'false')
      const { week, materialized } = materializeWeek(calendar, calendarById)
      if (materialized) {
        diag(
          sink,
          MSPDI_EXPORT_NORMALIZED,
          'info',
          `calendar '${calendarId}' has a partial workingWeek; the inherited weekday periods were materialized from the base-calendar chain (the accepted importer materializes all seven weekday keys — the resolved semantics are exactly recoverable)`,
          calendarId,
        )
      }
      writer.open('WeekDays')
      for (let key = 0; key <= 6; key += 1) {
        const periods = week[key] ?? []
        writer.open('WeekDay')
        writer.leaf('DayType', String(key + 1)) // MSPDI DayType 1=Sunday..7=Saturday
        writer.leaf('DayWorking', periods.length > 0 ? 'true' : 'false')
        if (periods.length > 0) {
          writer.open('WorkingTimes')
          for (const period of stableSort(periods, (a, b) => {
            if (a.startMinute !== b.startMinute) return a.startMinute - b.startMinute
            return a.endMinute - b.endMinute
          })) {
            if (period.endMinute >= 1440) {
              diag(
                sink,
                UNREPRESENTABLE_MSPDI_VALUE,
                'warning',
                `calendar '${calendarId}' has a working period ending at ${minutesToTime(period.endMinute)}; the accepted importer's HH:MM:SS rule (hours ≤ 23) drops this period on re-import`,
                calendarId,
              )
            }
            writer.open('WorkingTime')
            writer.leaf('FromTime', minutesToTime(period.startMinute))
            writer.leaf('ToTime', minutesToTime(period.endMinute))
            writer.close('WorkingTime')
          }
          writer.close('WorkingTimes')
        }
        writer.close('WeekDay')
      }
      writer.close('WeekDays')
      if (calendar.exceptions.length > 0) {
        writer.open('Exceptions')
        for (const exception of stableSort(calendar.exceptions, (a, b) =>
          compareStrings(a.date, b.date),
        )) {
          writer.open('Exception')
          writer.leaf('Start', `${exception.date}T00:00:00`)
          writer.leaf('Finish', `${exception.date}T00:00:00`)
          if (exception.periods.length > 0) {
            writer.open('WorkingTimes')
            for (const period of stableSort(exception.periods, (a, b) => {
              if (a.startMinute !== b.startMinute) return a.startMinute - b.startMinute
              return a.endMinute - b.endMinute
            })) {
              if (period.endMinute >= 1440) {
                diag(
                  sink,
                  UNREPRESENTABLE_MSPDI_VALUE,
                  'warning',
                  `calendar '${calendarId}' has an exception period on ${exception.date} ending at ${minutesToTime(period.endMinute)}; the accepted importer's HH:MM:SS rule (hours ≤ 23) drops this period on re-import`,
                  calendarId,
                )
              }
              writer.open('WorkingTime')
              writer.leaf('FromTime', minutesToTime(period.startMinute))
              writer.leaf('ToTime', minutesToTime(period.endMinute))
              writer.close('WorkingTime')
            }
            writer.close('WorkingTimes')
          }
          writer.close('Exception')
        }
        writer.close('Exceptions')
      }
      writer.close('Calendar')
    }
    writer.close('Calendars')
  }

  // ---- custom-field definitions (sorted by FieldID) --------------------------
  if (document.customFields.length > 0) {
    writer.open('ExtendedAttributes')
    for (const field of stableSort(document.customFields, (a, b) =>
      compareStrings(a.id as string, b.id as string),
    )) {
      writer.open('ExtendedAttribute')
      writer.leaf('FieldID', field.id as string)
      writer.leaf('Alias', field.name)
      writer.leaf('Type', field.type)
      writer.close('ExtendedAttribute')
    }
    writer.close('ExtendedAttributes')
  }

  // ---- tasks (hierarchical DFS order) ----------------------------------------
  const { order: tasksInOrder, outlines, reordered } = projectOutline(document.tasks)
  if (reordered) {
    diag(
      sink,
      UNSUPPORTED_MSPDI_EXPORT_FEATURE,
      'warning',
      'canonical task array order is not hierarchical-DFS (a child precedes its parent); tasks are exported in deterministic DFS order with sibling order preserved',
    )
  }
  // Dependencies grouped per successor, in canonical (successor uid,
  // predecessor uid, link type) order.
  const dependenciesBySuccessor = new Map<string, Dependency[]>()
  for (const dep of stableSort(document.dependencies, (a, b) => {
    const sa = taskUid(a.successorId as string)
    const sb = taskUid(b.successorId as string)
    if (sa !== sb) return sa - sb
    const pa = taskUid(a.predecessorId as string)
    const pb = taskUid(b.predecessorId as string)
    if (pa !== pb) return pa - pb
    return LINK_TYPE_CODES[a.type] - LINK_TYPE_CODES[b.type]
  })) {
    const key = dep.successorId as string
    const list = dependenciesBySuccessor.get(key)
    if (list === undefined) dependenciesBySuccessor.set(key, [dep])
    else list.push(dep)
  }
  // Baseline snapshots keyed by task id → slot entries.
  const snapshotsByTask = new Map<string, Array<{ slot: number; baseline: Baseline }>>()
  for (const baseline of document.baselines) {
    const slot = baselineSlots.uids.get(baseline.id as string) ?? 0
    for (const taskKey of Object.keys(baseline.taskSnapshots)) {
      const list = snapshotsByTask.get(taskKey)
      const entry = { slot, baseline }
      if (list === undefined) snapshotsByTask.set(taskKey, [entry])
      else list.push(entry)
    }
  }
  // task.baseline reverse-index transparency (derived state; the top-level
  // snapshots are the authoritative export source). The accepted importer
  // reconstructs the index EMPTY — the imported shape — so a non-empty
  // canonical index does not round-trip byte-wise; a listing with no matching
  // snapshot is genuinely lost. Both are surfaced, never silent.
  for (const task of document.tasks) {
    const snapshottedBy = new Set(
      (snapshotsByTask.get(task.id as string) ?? []).map((e) => e.baseline.id as string),
    )
    const listed = (task.baseline ?? []).map((id) => id as string)
    const dangling = listed.filter((id) => !snapshottedBy.has(id))
    if (dangling.length > 0) {
      diag(
        sink,
        UNSUPPORTED_MSPDI_EXPORT_FEATURE,
        'warning',
        `task '${task.id}' baseline reverse-index lists baseline(s) [${dangling.join(', ')}] with no snapshot for this task; those listings are not representable in MSPDI and are dropped`,
        task.id as string,
      )
    } else if (listed.length > 0) {
      diag(
        sink,
        UNSUPPORTED_MSPDI_EXPORT_FEATURE,
        'warning',
        `task '${task.id}' baseline reverse-index is derived state; the accepted importer reconstructs it empty (the top-level snapshots it derives from are preserved exactly)`,
        task.id as string,
      )
    }
  }

  if (document.tasks.length > 0) {
    writer.open('Tasks')
    let displayId = 0
    for (const task of tasksInOrder) {
      displayId += 1
      const taskId = task.id as string
      const uid = taskUid(taskId)
      const outline = outlines.get(taskId) ?? ''
      if (outline !== task.wbs) {
        diag(
          sink,
          UNREPRESENTABLE_MSPDI_VALUE,
          'warning',
          `task '${taskId}' canonical wbs ${JSON.stringify(task.wbs)} differs from the hierarchy-derived outline ${JSON.stringify(outline)}; the derived outline is exported (the parent/child relationships are preserved exactly)`,
          taskId,
        )
      }
      writer.open('Task')
      writer.leaf('UID', String(uid))
      writer.leaf('ID', String(displayId))
      if (task.name.length === 0) {
        diag(
          sink,
          UNREPRESENTABLE_MSPDI_VALUE,
          'warning',
          `task '${taskId}' has an empty name; the importer will substitute the placeholder "Task ${uid}"`,
          taskId,
        )
      }
      writer.leaf('Name', task.name)
      writer.leaf('WBS', outline)
      writer.leaf('OutlineNumber', outline)
      writer.leaf('OutlineLevel', String(task.outlineLevel))
      writer.leaf('Summary', task.summary ? 'true' : 'false')
      writer.leaf('Milestone', task.milestone ? 'true' : 'false')
      if (task.autoScheduled !== !task.manualScheduled) {
        diag(
          sink,
          UNSUPPORTED_MSPDI_EXPORT_FEATURE,
          'warning',
          `task '${taskId}' has inconsistent manualScheduled/autoScheduled flags; the pair is normalized from manualScheduled on re-import`,
          taskId,
        )
      }
      writer.leaf('Manual', task.manualScheduled ? 'true' : 'false')
      writer.leaf('Type', String(TASK_TYPE_CODES[task.taskType]))
      if (task.constraintType !== undefined) {
        writer.leaf('ConstraintType', String(CONSTRAINT_TYPE_CODES[task.constraintType]))
      }
      if (task.constraintDate !== undefined) writer.leaf('ConstraintDate', task.constraintDate)
      if (task.deadline !== undefined) writer.leaf('Deadline', task.deadline)
      writer.leaf('Priority', numberText(task.priority))
      if (task.calendarId !== undefined) {
        writer.leaf('CalendarUID', String(calendarUid(task.calendarId as string)))
      }
      if (task.start !== undefined) writer.leaf('Start', task.start)
      if (task.finish !== undefined) writer.leaf('Finish', task.finish)
      emitDuration(writer, sink, 'Duration', task.duration, `task '${taskId}'`)
      emitDuration(writer, sink, 'Work', task.work, `task '${taskId}'`)
      emitDuration(writer, sink, 'RemainingWork', task.remainingWork, `task '${taskId}'`)
      emitDuration(writer, sink, 'ActualWork', task.actualWork, `task '${taskId}'`)
      writer.leaf('Cost', numberText(task.cost))
      writer.leaf('ActualCost', numberText(task.actualCost))
      writer.leaf('RemainingCost', numberText(task.remainingCost))
      writer.leaf('PercentComplete', numberText(task.percentComplete))
      if (task.physicalPercentComplete !== undefined) {
        writer.leaf('PhysicalPercentComplete', numberText(task.physicalPercentComplete))
        diag(
          sink,
          UNSUPPORTED_MSPDI_EXPORT_FEATURE,
          'warning',
          `task '${taskId}' physicalPercentComplete is emitted for MSPDI fidelity but the accepted PROJECT-015 importer does not read it back (round-trip limitation)`,
          taskId,
        )
      }
      if (task.notes.length > 1) {
        diag(
          sink,
          UNSUPPORTED_MSPDI_EXPORT_FEATURE,
          'warning',
          `task '${taskId}' carries ${task.notes.length} notes; MSPDI has a single <Notes> field, so they are joined with newlines and re-import as one note`,
          taskId,
        )
      }
      if (task.notes.length > 0) writer.leaf('Notes', task.notes.join('\n'))
      // Dependencies where this task is the successor.
      for (const dep of dependenciesBySuccessor.get(taskId) ?? []) {
        writer.open('PredecessorLink')
        writer.leaf('PredecessorUID', String(taskUid(dep.predecessorId as string)))
        writer.leaf('Type', String(LINK_TYPE_CODES[dep.type]))
        const linkLag = dep.lagMinutes * 10
        if (!Number.isSafeInteger(linkLag)) {
          diag(
            sink,
            INVALID_MSPDI_EXPORT_LAG,
            'error',
            `dependency lagMinutes ${numberText(dep.lagMinutes)} cannot be represented as LinkLag = lagMinutes × 10 (outside the safe-integer range); exported as lag 0 — lag semantics are never silently changed`,
            taskId,
          )
          writer.leaf('LinkLag', '0')
        } else {
          writer.leaf('LinkLag', String(linkLag))
        }
        // The deterministic canonical lag format: working minutes
        // (`LinkLagFormat` = 1), `LinkLag` = lagMinutes × 10.
        writer.leaf('LinkLagFormat', String(LAG_FORMAT_MINUTE))
        writer.close('PredecessorLink')
      }
      // Baseline slots (ascending slot order).
      for (const entry of (snapshotsByTask.get(taskId) ?? []).sort((a, b) => a.slot - b.slot)) {
        const snapshot = entry.baseline.taskSnapshots[taskId]
        const slotName = BASELINE_SLOT_NAMES[entry.slot] ?? `Baseline${entry.slot}`
        writer.open(slotName)
        if (snapshot.start !== undefined) writer.leaf('Start', snapshot.start)
        if (snapshot.finish !== undefined) writer.leaf('Finish', snapshot.finish)
        emitDuration(
          writer,
          sink,
          'Duration',
          snapshot.duration,
          `baseline '${entry.baseline.id}' task '${taskId}'`,
        )
        emitDuration(
          writer,
          sink,
          'Work',
          snapshot.work,
          `baseline '${entry.baseline.id}' task '${taskId}'`,
        )
        writer.leaf('Cost', numberText(snapshot.cost))
        writer.close(slotName)
      }
      // Custom-field values (sorted by field id — never insertion order).
      for (const fieldId of Object.keys(task.customFields).sort(compareStrings)) {
        writer.open('ExtendedAttribute')
        writer.leaf('FieldID', fieldId)
        const value: string | number | boolean | null | undefined = (
          task.customFields as Record<string, string | number | boolean | null>
        )[fieldId]
        if (value === null || value === undefined) {
          // No <Value> child → the importer records null.
        } else if (typeof value === 'number') {
          if (!Number.isFinite(value)) {
            diag(
              sink,
              UNREPRESENTABLE_MSPDI_VALUE,
              'warning',
              `custom field '${fieldId}' on task '${taskId}' has non-finite value ${numberText(value)}; it re-imports as the string form`,
              taskId,
            )
          }
          writer.leaf('Value', numberText(value))
        } else if (typeof value === 'boolean') {
          writer.leaf('Value', value ? 'true' : 'false')
        } else {
          const trimmed = value.trim()
          if (/^-?\d+(\.\d+)?$/.test(trimmed) || trimmed === 'true' || trimmed === 'false') {
            diag(
              sink,
              UNREPRESENTABLE_MSPDI_VALUE,
              'warning',
              `string custom-field value ${JSON.stringify(value)} on task '${taskId}' re-imports as a number/boolean (the importer re-parses values independent of the field type)`,
              taskId,
            )
          }
          writer.leaf('Value', value)
        }
        writer.close('ExtendedAttribute')
      }
      writer.close('Task')
    }
    writer.close('Tasks')
  }

  // ---- resources (ascending uid) ----------------------------------------------
  if (document.resources.length > 0) {
    writer.open('Resources')
    for (const resource of stableSort(document.resources, (a, b) => {
      if (a.uid !== b.uid) return a.uid - b.uid
      return compareStrings(a.id as string, b.id as string)
    })) {
      writer.open('Resource')
      writer.leaf('UID', String(resourceUids.uids.get(resource.id as string) ?? 0))
      writer.leaf('Name', resource.name)
      writer.leaf('Type', String(RESOURCE_KIND_CODES[resource.kind]))
      writer.leaf('MaxUnits', numberText(resource.maxUnits))
      writer.leaf('StandardRate', numberText(resource.standardRate))
      writer.leaf('OvertimeRate', numberText(resource.overtimeRate))
      writer.leaf('CostPerUse', numberText(resource.costPerUse))
      if (resource.calendarId !== undefined) {
        writer.leaf('CalendarUID', String(calendarUid(resource.calendarId as string)))
      }
      if (resource.availability.length > 0) {
        writer.open('AvailabilityPeriods')
        for (const period of stableSort(resource.availability, (a, b) =>
          compareStrings(a.start, b.start),
        )) {
          writer.open('AvailabilityPeriod')
          writer.leaf('AvailableFrom', period.start)
          if (period.finish !== undefined) writer.leaf('AvailableTo', period.finish)
          writer.leaf('AvailableUnits', numberText(period.units))
          writer.close('AvailabilityPeriod')
        }
        writer.close('AvailabilityPeriods')
      }
      writer.close('Resource')
    }
    writer.close('Resources')
  }

  // ---- assignments (sorted by task uid, then resource uid) ---------------------
  if (document.assignments.length > 0) {
    writer.open('Assignments')
    for (const assignment of assignmentsSorted) {
      writer.open('Assignment')
      writer.leaf('UID', String(assignmentUids.uids.get(assignment.id as string) ?? 0))
      writer.leaf('TaskUID', String(taskUid(assignment.taskId as string)))
      writer.leaf(
        'ResourceUID',
        String(resourceUids.uids.get(assignment.resourceId as string) ?? 0),
      )
      writer.leaf('Units', numberText(assignment.units))
      emitDuration(writer, sink, 'Work', assignment.work, `assignment '${assignment.id}'`)
      emitDuration(
        writer,
        sink,
        'ActualWork',
        assignment.actualWork,
        `assignment '${assignment.id}'`,
      )
      emitDuration(
        writer,
        sink,
        'RemainingWork',
        assignment.remainingWork,
        `assignment '${assignment.id}'`,
      )
      writer.leaf('Cost', numberText(assignment.cost))
      writer.leaf('ActualCost', numberText(assignment.actualCost))
      writer.leaf('RemainingCost', numberText(assignment.remainingCost))
      writer.close('Assignment')
    }
    writer.close('Assignments')
  }

  writer.close('Project')

  // ---- view-layer collections with no MSPDI representation ---------------------
  const viewCount =
    document.views.length +
    document.tables.length +
    document.filters.length +
    document.groups.length
  if (viewCount > 0) {
    diag(
      sink,
      UNSUPPORTED_MSPDI_EXPORT_FEATURE,
      'warning',
      `document carries ${viewCount} view/table/filter/group definition(s) with no MSPDI representation; they are not exported (the accepted importer reconstructs them empty)`,
    )
  }

  diag(
    sink,
    MSPDI_WRITTEN,
    'info',
    `Exported canonical ProjectDocument as MSPDI XML (SaveVersion 16, adapter schema ${MSPDI_FORMAT_VERSION}): ${document.tasks.length} task(s), ${document.resources.length} resource(s), ${document.assignments.length} assignment(s), ${document.dependencies.length} dependency(ies), ${document.calendars.length} calendar(s), ${document.baselines.length} baseline(s), ${document.customFields.length} custom field(s)`,
  )

  return { bytes: encodeUtf8(writer.toString()), diagnostics: sink }
}
