/**
 * PROJECT-015 — MSPDI XML → canonical `ProjectDocument` importer.
 *
 * Architectural flow (PROJECT-015 brief):
 *
 *   MSPDI XML → MSPDIAdapter → ProjectDocument → validateProjectDocument →
 *   canonical scheduling engine
 *
 * The adapter does NOT become the canonical model. MSPDI-specific identity is
 * NOT preserved as GenOffice identity (see `./identity.ts`). No MSPDI XML is
 * written into React/browser code (this package has no React/browser imports —
 * `project-foundation.yml` greps it). No MSPDI export is implemented
 * (PROJECT-016 is explicitly unauthorized); the `MspdiFileAdapter` exposes
 * only `inspect` + `import`, the minimal shared-adapter typing.
 *
 * Pipeline (mirrors the accepted PROJECT-014 `.gproj` adapter conventions):
 *
 *   1. byte-size guard
 *   2. parseXml (host-neutral pure-TS tokenizer; throws XmlParseError on
 *      malformed/oversized/DoS input)
 *   3. root-element + SaveVersion gate
 *   4. build cross-entity UID→identity maps (identity.ts)
 *   5. per-entity schema validation + canonical construction
 *      (calendars → resources → tasks → dependencies → assignments →
 *       baselines → custom fields → properties)
 *   6. partial recovery: drop bad entity + emit error diagnostic; file-level
 *      failure → empty document + single error
 *   7. delegate to `validateProjectDocument` → surface its diagnostics
 *      verbatim (severity 'error', no entityId)
 *   8. emit info `MSPDI_READ` with the save version
 *
 * Determinism: the same MSPDI bytes always produce the same `ProjectDocument`
 * bytes (verifiable via `serializeGproj`). Element child-order variations do
 * not affect the canonical output (every field is extracted by name).
 */
import type {
  Assignment,
  Baseline,
  Calendar,
  CalendarException,
  CalendarId,
  CalendarPeriod,
  ConstraintType,
  CustomField,
  Dependency,
  DependencyType,
  ImportDiagnostic,
  ISODateTime,
  ProjectDocument,
  ProjectFileMetadata,
  ProjectProperties,
  Resource,
  ResourceId,
  Task,
  TaskId,
  WorkingMinutes,
} from '@genoffice/project-contracts'
import {
  asCalendarId,
  asCustomFieldId,
  asISODateTime,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import { emptyProjectDocument } from '../serialize.js'
import { validateProjectDocument } from '@genoffice/project-engine'
import {
  childText,
  childrenNamed,
  firstChild,
  parseXml,
  XmlParseError,
  type XmlNode,
} from './xml-parser.js'
import {
  isSupportedMspdiSaveVersion,
  MSPDI_FORMAT,
  MSPDI_FORMAT_VERSION,
  MSPDI_ROOT_ELEMENT,
} from './envelope.js'
import {
  baselineIndexToId,
  dependencyId,
  uidToAssignmentId,
  uidToCalendarId,
  uidToResourceId,
  uidToTaskId,
} from './identity.js'
import {
  isoDurationToMinutes,
  isValidExceptionDate,
  lagToMinutes,
  mspdiTimeToMinutes,
  normalizeMspdiDate,
} from './conversions.js'
import {
  INVALID_MSPDI,
  INVALID_MSPDI_CALENDAR,
  INVALID_MSPDI_CONSTRAINT,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_REFERENCE,
  INVALID_MSPDI_RESOURCE,
  MISSING_MSPDI_FIELD,
  MSPDI_READ,
  UNSUPPORTED_MSPDI_FEATURE,
  UNSUPPORTED_MSPDI_VERSION,
} from './diagnostics.js'

export interface MspdiImportResult {
  document: ProjectDocument
  diagnostics: ImportDiagnostic[]
}

/** Minimal shared-adapter typing: import + inspect only (NO export —
 * PROJECT-016 MSPDI export is explicitly unauthorized). */
export interface MspdiFileAdapter {
  readonly format: typeof MSPDI_FORMAT
  inspect(input: Uint8Array, metadata?: ProjectFileMetadata): ProjectFileMetadata
  import(input: Uint8Array, metadata?: ProjectFileMetadata): MspdiImportResult
}

type Diag = ImportDiagnostic
interface Sink {
  push(d: Diag): void
}

function diag(
  sink: Sink,
  code: string,
  severity: Diag['severity'],
  message: string,
  entityId?: string,
): void {
  const entry: Diag = { code, severity, message }
  if (entityId !== undefined) entry.entityId = entityId
  sink.push(entry)
}

// ---- typed extraction helpers (operate on parsed XML element children) ----

/** True if an element carries `xsi:nil="true"` (MSPDI null marker). */
function isNil(node: XmlNode): boolean {
  return node.attributes['nil'] === 'true' || node.attributes['xsi:nil'] === 'true'
}

function optionalNumber(node: XmlNode, name: string): number | undefined {
  const text = childText(node, name)
  if (text === undefined) return undefined
  const n = Number(text)
  return Number.isFinite(n) ? n : undefined
}

function optionalInteger(node: XmlNode, name: string): number | undefined {
  const text = childText(node, name)
  if (text === undefined) return undefined
  const n = Number(text)
  return Number.isInteger(n) ? n : undefined
}

function optionalBoolean(node: XmlNode, name: string): boolean | undefined {
  const text = childText(node, name)
  if (text === undefined) return undefined
  const t = text.toLowerCase()
  if (t === 'true' || t === '1') return true
  if (t === 'false' || t === '0') return false
  return undefined
}

function optionalDate(
  node: XmlNode,
  name: string,
  sink: Sink,
  entityId?: string,
): ISODateTime | undefined {
  const child = firstChild(node, name)
  if (child === undefined) return undefined
  if (isNil(child)) return undefined
  const raw = child.text.trim()
  if (raw.length === 0) return undefined
  const r = normalizeMspdiDate(raw)
  if ('invalid' in r) {
    diag(
      sink,
      INVALID_MSPDI_DATE,
      'error',
      `<${name}> is not a valid ISO-8601 date: ${JSON.stringify(raw)}`,
      entityId,
    )
    return undefined
  }
  return r.iso
}

function optionalDuration(
  node: XmlNode,
  name: string,
  sink: Sink,
  entityId?: string,
): WorkingMinutes {
  const text = childText(node, name)
  if (text === undefined) return asWorkingMinutes(0)
  const r = isoDurationToMinutes(text)
  if (r.ok) return r.minutes
  diag(
    sink,
    r.reason === 'unsupported' ? UNSUPPORTED_MSPDI_FEATURE : INVALID_MSPDI_DURATION,
    r.reason === 'unsupported' ? 'warning' : 'error',
    `<${name}> cannot be converted to integer working minutes: ${JSON.stringify(text)} (${r.reason})`,
    entityId,
  )
  return asWorkingMinutes(0)
}

// ---- constraint / link / resource-kind / task-type maps ----------------

const CONSTRAINT_TYPE_MAP: Partial<Record<number, ConstraintType>> = {
  0: 'asSoonAsPossible',
  1: 'asLateAsPossible',
  2: 'mustStartOn',
  3: 'mustFinishOn',
  4: 'startNoEarlierThan',
  5: 'finishNoEarlierThan',
  6: 'startNoLaterThan',
  7: 'finishNoLaterThan',
}

function mapConstraintType(
  value: number | undefined,
): { type: ConstraintType } | { unsupported: true } | { absent: true } {
  if (value === undefined) return { absent: true }
  const mapped = CONSTRAINT_TYPE_MAP[value]
  if (mapped === undefined) return { unsupported: true }
  return { type: mapped }
}

const LINK_TYPE_MAP: Partial<Record<number, DependencyType>> = {
  0: 'FS',
  1: 'FF',
  2: 'SS',
  3: 'SF',
}

function mapLinkType(value: number | undefined): DependencyType | undefined {
  if (value === undefined) return undefined
  return LINK_TYPE_MAP[value]
}

const RESOURCE_KIND_MAP: Partial<Record<number, Resource['kind']>> = {
  1: 'work',
  2: 'material',
  3: 'cost',
}

const TASK_TYPE_MAP: Partial<Record<number, Task['taskType']>> = {
  0: 'fixedUnits',
  1: 'fixedDuration',
  2: 'fixedWork',
}

// ---- calendar parsing ----------------------------------------------------

function parseWorkingTimes(parent: XmlNode): CalendarPeriod[] {
  const wt = firstChild(parent, 'WorkingTimes')
  if (wt === undefined) return []
  const periods: CalendarPeriod[] = []
  for (const period of childrenNamed(wt, 'WorkingTime')) {
    const from = childText(period, 'FromTime')
    const to = childText(period, 'ToTime')
    if (from === undefined || to === undefined) continue
    const start = mspdiTimeToMinutes(from)
    const end = mspdiTimeToMinutes(to)
    if (start === null || end === null || start >= end) continue
    periods.push({ startMinute: start, endMinute: end })
  }
  return periods
}

function parseCalendar(node: XmlNode, sink: Sink): Calendar | undefined {
  const uidRaw = optionalInteger(node, 'UID')
  if (uidRaw === undefined) {
    diag(sink, MISSING_MSPDI_FIELD, 'error', '<Calendar> is missing <UID>')
    return undefined
  }
  const id = uidToCalendarId(uidRaw)
  const name = childText(node, 'Name') ?? `Calendar ${uidRaw}`
  const baseUid = optionalInteger(node, 'BaseCalendarUID')
  const baseCalendarId = baseUid === undefined ? undefined : uidToCalendarId(baseUid)
  // workingWeek: keys 0..6 (0=Sunday). MSPDI DayType 1=Sunday..7=Saturday.
  const workingWeek: Record<number, CalendarPeriod[]> = {
    0: [],
    1: [],
    2: [],
    3: [],
    4: [],
    5: [],
    6: [],
  }
  const weekDays = firstChild(node, 'WeekDays')
  if (weekDays !== undefined) {
    for (const day of childrenNamed(weekDays, 'WeekDay')) {
      const dayType = optionalInteger(day, 'DayType')
      if (dayType === undefined || dayType < 1 || dayType > 7) {
        diag(
          sink,
          INVALID_MSPDI_CALENDAR,
          'error',
          `<Calendar uid=${uidRaw}> has a <WeekDay> with missing/invalid <DayType>`,
          String(id),
        )
        continue
      }
      const key = dayType - 1
      const working = optionalBoolean(day, 'DayWorking') ?? false
      workingWeek[key] = working ? parseWorkingTimes(day) : []
    }
  }
  // exceptions
  const exceptions: CalendarException[] = []
  const excNode = firstChild(node, 'Exceptions')
  if (excNode !== undefined) {
    for (const exc of childrenNamed(excNode, 'Exception')) {
      const startRaw = childText(exc, 'Start')
      if (startRaw === undefined) {
        diag(
          sink,
          INVALID_MSPDI_CALENDAR,
          'error',
          `<Calendar uid=${uidRaw}> has an <Exception> without <Start>`,
          String(id),
        )
        continue
      }
      const startDate = startRaw.slice(0, 10)
      if (!isValidExceptionDate(startDate)) {
        diag(
          sink,
          INVALID_MSPDI_DATE,
          'error',
          `<Calendar uid=${uidRaw}> exception <Start> is not a valid date: ${JSON.stringify(startRaw)}`,
          String(id),
        )
        continue
      }
      const finishRaw = childText(exc, 'Finish')
      const typeRaw = optionalInteger(exc, 'Type')
      const recurring = typeRaw !== undefined && typeRaw !== 1
      const multiDay = finishRaw !== undefined && finishRaw.slice(0, 10) !== startDate
      if (recurring || multiDay) {
        diag(
          sink,
          UNSUPPORTED_MSPDI_FEATURE,
          'warning',
          `<Calendar uid=${uidRaw}> has a ${recurring ? 'recurring' : 'multi-day'} exception on ${startDate}; the canonical model only supports single-date exceptions — mapped to a single ${startDate} exception`,
          String(id),
        )
      }
      const periods = parseWorkingTimes(exc)
      exceptions.push({ date: startDate, periods })
    }
  }
  return {
    id,
    name,
    ...(baseCalendarId !== undefined ? { baseCalendarId } : {}),
    workingWeek,
    exceptions,
  }
}

// ---- resource parsing ----------------------------------------------------

function parseAvailabilityPeriods(node: XmlNode, sink: Sink, entityId: string) {
  const periodsNode = firstChild(node, 'AvailabilityPeriods')
  if (periodsNode === undefined) return []
  const out: Resource['availability'] = []
  for (const period of childrenNamed(periodsNode, 'AvailabilityPeriod')) {
    const fromRaw = childText(period, 'AvailableFrom')
    const toRaw = childText(period, 'AvailableTo')
    const unitsRaw = optionalNumber(period, 'AvailableUnits')
    if (fromRaw === undefined) {
      diag(
        sink,
        INVALID_MSPDI_DATE,
        'error',
        '<AvailabilityPeriod> is missing <AvailableFrom>',
        entityId,
      )
      continue
    }
    const from = normalizeMspdiDate(fromRaw)
    if ('invalid' in from) {
      diag(
        sink,
        INVALID_MSPDI_DATE,
        'error',
        `<AvailableFrom> is not a valid date: ${JSON.stringify(fromRaw)}`,
        entityId,
      )
      continue
    }
    let finish: ISODateTime | undefined
    if (toRaw !== undefined) {
      const to = normalizeMspdiDate(toRaw)
      if ('invalid' in to) {
        diag(
          sink,
          INVALID_MSPDI_DATE,
          'error',
          `<AvailableTo> is not a valid date: ${JSON.stringify(toRaw)}`,
          entityId,
        )
      } else {
        finish = to.iso
      }
    }
    const units = unitsRaw ?? 1
    out.push({ start: from.iso, ...(finish !== undefined ? { finish } : {}), units })
  }
  return out
}

function parseResource(
  node: XmlNode,
  calendarUidMap: ReadonlyMap<number, string>,
  sink: Sink,
): Resource | undefined {
  const uidRaw = optionalInteger(node, 'UID')
  if (uidRaw === undefined) {
    diag(sink, MISSING_MSPDI_FIELD, 'error', '<Resource> is missing <UID>')
    return undefined
  }
  const id = uidToResourceId(uidRaw)
  const name = childText(node, 'Name') ?? `Resource ${uidRaw}`
  const typeRaw = optionalInteger(node, 'Type')
  let kind: Resource['kind']
  if (typeRaw === undefined) {
    kind = 'work'
  } else {
    const mapped = RESOURCE_KIND_MAP[typeRaw]
    if (mapped === undefined) {
      diag(
        sink,
        INVALID_MSPDI_RESOURCE,
        'error',
        `<Resource uid=${uidRaw}> has unsupported <Type> ${typeRaw}; defaulted to 'work'`,
        String(id),
      )
      kind = 'work'
    } else {
      kind = mapped
    }
  }
  const maxUnits = optionalNumber(node, 'MaxUnits') ?? (kind === 'work' ? 1 : 0)
  const standardRate = optionalNumber(node, 'StandardRate') ?? 0
  const overtimeRate = optionalNumber(node, 'OvertimeRate') ?? 0
  const costPerUse = optionalNumber(node, 'CostPerUse') ?? 0
  let calendarId: CalendarId | undefined
  const calUid = optionalInteger(node, 'CalendarUID')
  if (calUid !== undefined) {
    const mapped = calendarUidMap.get(calUid)
    if (mapped === undefined) {
      diag(
        sink,
        INVALID_MSPDI_REFERENCE,
        'error',
        `<Resource uid=${uidRaw}> references <CalendarUID> ${calUid} which is not a declared calendar; resource will use the project default calendar`,
        String(id),
      )
    } else {
      calendarId = asCalendarId(mapped)
    }
  }
  const availability = parseAvailabilityPeriods(node, sink, String(id))
  return {
    id,
    uid: uidRaw,
    name,
    kind,
    maxUnits,
    standardRate,
    overtimeRate,
    costPerUse,
    ...(calendarId !== undefined ? { calendarId } : {}),
    availability,
  }
}

// ---- task parsing --------------------------------------------------------

function parseCustomFieldValues(
  node: XmlNode,
  definedFields: ReadonlySet<string>,
  sink: Sink,
  entityId: string,
): Task['customFields'] {
  const out: Task['customFields'] = {}
  for (const ext of childrenNamed(node, 'ExtendedAttribute')) {
    const fieldId = childText(ext, 'FieldID')
    const valueText = childText(ext, 'Value')
    if (fieldId === undefined) continue
    if (!definedFields.has(fieldId)) {
      diag(
        sink,
        INVALID_MSPDI_REFERENCE,
        'error',
        `task references ExtendedAttribute FieldID ${fieldId} which is not a declared custom field; value dropped`,
        entityId,
      )
      continue
    }
    if (valueText === undefined) {
      out[asCustomFieldId(fieldId)] = null
      continue
    }
    // Parse deterministically: number, boolean, else string.
    const asNum = Number(valueText)
    if (
      valueText.trim() !== '' &&
      Number.isFinite(asNum) &&
      /^-?\d+(\.\d+)?$/.test(valueText.trim())
    ) {
      out[asCustomFieldId(fieldId)] = asNum
    } else if (valueText === 'true') {
      out[asCustomFieldId(fieldId)] = true
    } else if (valueText === 'false') {
      out[asCustomFieldId(fieldId)] = false
    } else {
      out[asCustomFieldId(fieldId)] = valueText
    }
  }
  return out
}

function parseTask(
  node: XmlNode,
  calendarUidMap: ReadonlyMap<number, string>,
  definedCustomFields: ReadonlySet<string>,
  sink: Sink,
): Task | undefined {
  const uidRaw = optionalInteger(node, 'UID')
  if (uidRaw === undefined) {
    diag(sink, MISSING_MSPDI_FIELD, 'error', '<Task> is missing <UID>')
    return undefined
  }
  const id = uidToTaskId(uidRaw)
  const name = childText(node, 'Name') ?? `Task ${uidRaw}`
  const wbs = childText(node, 'OutlineNumber') ?? childText(node, 'WBS') ?? ''
  const outlineLevel = optionalInteger(node, 'OutlineLevel') ?? 1
  const summary = optionalBoolean(node, 'Summary') ?? false
  const milestone = optionalBoolean(node, 'Milestone') ?? false
  const manual = optionalBoolean(node, 'Manual') ?? false
  const manualScheduled = manual
  const autoScheduled = !manual
  // Task type.
  const typeRaw = optionalInteger(node, 'Type')
  let taskType: Task['taskType']
  if (typeRaw === undefined) {
    taskType = 'fixedUnits'
  } else {
    const mapped = TASK_TYPE_MAP[typeRaw]
    if (mapped === undefined) {
      diag(
        sink,
        UNSUPPORTED_MSPDI_FEATURE,
        'warning',
        `<Task uid=${uidRaw}> has unsupported <Type> ${typeRaw}; defaulted to 'fixedUnits'`,
        String(id),
      )
      taskType = 'fixedUnits'
    } else {
      taskType = mapped
    }
  }
  const duration = optionalDuration(node, 'Duration', sink, String(id))
  const work = optionalDuration(node, 'Work', sink, String(id))
  const actualWork = optionalDuration(node, 'ActualWork', sink, String(id))
  const remainingWork = optionalDuration(node, 'RemainingWork', sink, String(id))
  const cost = optionalNumber(node, 'Cost') ?? 0
  const actualCost = optionalNumber(node, 'ActualCost') ?? 0
  const remainingCost = optionalNumber(node, 'RemainingCost') ?? 0
  const percentComplete = optionalNumber(node, 'PercentComplete') ?? 0
  const priority = optionalNumber(node, 'Priority') ?? 500
  const start = optionalDate(node, 'Start', sink, String(id))
  const finish = optionalDate(node, 'Finish', sink, String(id))
  const deadline = optionalDate(node, 'Deadline', sink, String(id))
  // Constraint.
  let constraintType: ConstraintType | undefined
  let constraintDate: ISODateTime | undefined
  const constraintRaw = optionalInteger(node, 'ConstraintType')
  const mappedConstraint = mapConstraintType(constraintRaw)
  if ('type' in mappedConstraint) {
    constraintType = mappedConstraint.type
    constraintDate = optionalDate(node, 'ConstraintDate', sink, String(id))
    if (
      constraintDate === undefined &&
      constraintType !== 'asSoonAsPossible' &&
      constraintType !== 'asLateAsPossible'
    ) {
      diag(
        sink,
        INVALID_MSPDI_CONSTRAINT,
        'error',
        `<Task uid=${uidRaw}> has date-bounded constraint ${constraintType} but no valid <ConstraintDate>`,
        String(id),
      )
    }
  } else if ('unsupported' in mappedConstraint) {
    diag(
      sink,
      UNSUPPORTED_MSPDI_FEATURE,
      'warning',
      `<Task uid=${uidRaw}> has unsupported <ConstraintType> ${constraintRaw}; constraint dropped`,
      String(id),
    )
  }
  // Calendar.
  let calendarId: CalendarId | undefined
  const calUid = optionalInteger(node, 'CalendarUID')
  if (calUid !== undefined) {
    const mapped = calendarUidMap.get(calUid)
    if (mapped === undefined) {
      diag(
        sink,
        INVALID_MSPDI_REFERENCE,
        'error',
        `<Task uid=${uidRaw}> references <CalendarUID> ${calUid} which is not a declared calendar; task will use the project default calendar`,
        String(id),
      )
    } else {
      calendarId = asCalendarId(mapped)
    }
  }
  // Notes.
  const notesRaw = childText(node, 'Notes')
  const notes: string[] = notesRaw === undefined ? [] : [notesRaw]
  // Custom field values.
  const customFields = parseCustomFieldValues(node, definedCustomFields, sink, String(id))
  return {
    id,
    uid: uidRaw,
    wbs,
    outlineLevel,
    name,
    taskType,
    summary,
    milestone,
    manualScheduled,
    autoScheduled,
    duration,
    priority,
    percentComplete,
    work,
    remainingWork,
    actualWork,
    cost,
    actualCost,
    remainingCost,
    baseline: [],
    customFields,
    notes,
    ...(start !== undefined ? { start } : {}),
    ...(finish !== undefined ? { finish } : {}),
    ...(constraintType !== undefined ? { constraintType } : {}),
    ...(constraintDate !== undefined ? { constraintDate } : {}),
    ...(deadline !== undefined ? { deadline } : {}),
    ...(calendarId !== undefined ? { calendarId } : {}),
  }
}

// ---- dependency parsing --------------------------------------------------

function parseDependencies(
  root: XmlNode,
  taskUidMap: ReadonlyMap<number, TaskId>,
  sink: Sink,
): Dependency[] {
  const tasksNode = firstChild(root, 'Tasks')
  if (tasksNode === undefined) return []
  const deps: Dependency[] = []
  for (const taskNode of childrenNamed(tasksNode, 'Task')) {
    const successorUid = optionalInteger(taskNode, 'UID')
    if (successorUid === undefined) continue
    const successorId = taskUidMap.get(successorUid)
    if (successorId === undefined) continue
    for (const link of childrenNamed(taskNode, 'PredecessorLink')) {
      const predUid = optionalInteger(link, 'PredecessorUID')
      if (predUid === undefined) {
        diag(
          sink,
          MISSING_MSPDI_FIELD,
          'error',
          '<PredecessorLink> is missing <PredecessorUID>; dependency dropped',
        )
        continue
      }
      const predecessorId = taskUidMap.get(predUid)
      if (predecessorId === undefined) {
        diag(
          sink,
          INVALID_MSPDI_REFERENCE,
          'error',
          `<PredecessorLink> references <PredecessorUID> ${predUid} which is not a declared task; dependency dropped`,
        )
        continue
      }
      const typeRaw = optionalInteger(link, 'Type')
      const type = mapLinkType(typeRaw)
      if (type === undefined) {
        diag(
          sink,
          UNSUPPORTED_MSPDI_FEATURE,
          'warning',
          `<PredecessorLink> from ${predUid} has unsupported link type ${typeRaw}; dependency dropped`,
        )
        continue
      }
      const linkLag = optionalInteger(link, 'LinkLag') ?? 0
      const linkLagFormat = optionalInteger(link, 'LinkLagFormat')
      const lagResult = lagToMinutes(linkLag, linkLagFormat)
      let lagMinutes: number
      if (lagResult.ok) {
        lagMinutes = lagResult.minutes
      } else if (lagResult.reason === 'unsupported') {
        diag(
          sink,
          UNSUPPORTED_MSPDI_FEATURE,
          'warning',
          `<PredecessorLink> from ${predUid} has ${linkLagFormat === 35 ? 'percentage' : 'elapsed'} lag (format ${linkLagFormat}); not faithfully representable as integer working-minutes, defaulted to 0`,
        )
        lagMinutes = 0
      } else {
        diag(
          sink,
          INVALID_MSPDI_DURATION,
          'error',
          `<PredecessorLink> from ${predUid} has malformed lag (LinkLag=${linkLag}, Format=${linkLagFormat}); defaulted to 0`,
        )
        lagMinutes = 0
      }
      deps.push({
        id: dependencyId(String(successorId), String(predecessorId), type),
        predecessorId,
        successorId,
        type,
        lagMinutes,
      })
    }
  }
  return deps
}

// ---- assignment parsing --------------------------------------------------

function parseAssignments(
  root: XmlNode,
  taskUidMap: ReadonlyMap<number, TaskId>,
  resourceUidMap: ReadonlyMap<number, ResourceId>,
  sink: Sink,
): Assignment[] {
  const node = firstChild(root, 'Assignments')
  if (node === undefined) return []
  const out: Assignment[] = []
  for (const aNode of childrenNamed(node, 'Assignment')) {
    const uidRaw = optionalInteger(aNode, 'UID')
    if (uidRaw === undefined) {
      diag(sink, MISSING_MSPDI_FIELD, 'error', '<Assignment> is missing <UID>')
      continue
    }
    const id = uidToAssignmentId(uidRaw)
    const taskUid = optionalInteger(aNode, 'TaskUID')
    const taskId = taskUid === undefined ? undefined : taskUidMap.get(taskUid)
    if (taskId === undefined) {
      diag(
        sink,
        INVALID_MSPDI_REFERENCE,
        'error',
        `<Assignment uid=${uidRaw}> references <TaskUID> ${taskUid} which is not a declared task; assignment dropped`,
        String(id),
      )
      continue
    }
    const resourceUid = optionalInteger(aNode, 'ResourceUID')
    const resourceId = resourceUid === undefined ? undefined : resourceUidMap.get(resourceUid)
    if (resourceId === undefined) {
      diag(
        sink,
        INVALID_MSPDI_REFERENCE,
        'error',
        `<Assignment uid=${uidRaw}> references <ResourceUID> ${resourceUid} which is not a declared resource; assignment dropped`,
        String(id),
      )
      continue
    }
    const units = optionalNumber(aNode, 'Units') ?? 1
    const work = optionalDuration(aNode, 'Work', sink, String(id))
    const actualWork = optionalDuration(aNode, 'ActualWork', sink, String(id))
    const remainingWork = optionalDuration(aNode, 'RemainingWork', sink, String(id))
    const cost = optionalNumber(aNode, 'Cost') ?? 0
    const actualCost = optionalNumber(aNode, 'ActualCost') ?? 0
    const remainingCost = optionalNumber(aNode, 'RemainingCost') ?? 0
    out.push({
      id,
      taskId,
      resourceId,
      units,
      work,
      actualWork,
      remainingWork,
      cost,
      actualCost,
      remainingCost,
    })
  }
  return out
}

// ---- baseline parsing ----------------------------------------------------

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

function parseBaselines(
  root: XmlNode,
  taskUidMap: ReadonlyMap<number, TaskId>,
  capturedAtFallback: ISODateTime,
  sink: Sink,
): Baseline[] {
  const tasksNode = firstChild(root, 'Tasks')
  if (tasksNode === undefined) return []
  // slotIndex -> { name, snapshots: Record<taskIdStr, snapshot> }
  const slots = new Map<
    number,
    { name: string; snapshots: Record<string, Baseline['taskSnapshots'][string]> }
  >()
  for (const taskNode of childrenNamed(tasksNode, 'Task')) {
    const uidRaw = optionalInteger(taskNode, 'UID')
    if (uidRaw === undefined) continue
    const taskId = taskUidMap.get(uidRaw)
    if (taskId === undefined) continue
    for (let slot = 0; slot < BASELINE_SLOT_NAMES.length; slot++) {
      const slotNode = firstChild(taskNode, BASELINE_SLOT_NAMES[slot])
      if (slotNode === undefined) continue
      const start = optionalDate(slotNode, 'Start', sink, `b${slot}`)
      const finish = optionalDate(slotNode, 'Finish', sink, `b${slot}`)
      const duration = optionalDuration(slotNode, 'Duration', sink, `b${slot}`)
      const work = optionalDuration(slotNode, 'Work', sink, `b${slot}`)
      const cost = optionalNumber(slotNode, 'Cost') ?? 0
      let entry = slots.get(slot)
      if (entry === undefined) {
        entry = {
          name: slot === 0 ? 'Baseline' : `Baseline ${slot}`,
          snapshots: {},
        }
        slots.set(slot, entry)
      }
      entry.snapshots[String(taskId)] = {
        duration,
        work,
        cost,
        ...(start !== undefined ? { start } : {}),
        ...(finish !== undefined ? { finish } : {}),
      }
    }
  }
  const baselines: Baseline[] = []
  for (const [slot, entry] of [...slots.entries()].sort((a, b) => a[0] - b[0])) {
    baselines.push({
      id: baselineIndexToId(slot),
      name: entry.name,
      capturedAt: capturedAtFallback,
      taskSnapshots: entry.snapshots,
    })
  }
  return baselines
}

// ---- custom-field definitions -------------------------------------------

function parseCustomFieldDefinitions(root: XmlNode, sink: Sink): CustomField[] {
  const eaNode = firstChild(root, 'ExtendedAttributes')
  if (eaNode === undefined) return []
  const out: CustomField[] = []
  for (const def of childrenNamed(eaNode, 'ExtendedAttribute')) {
    const fieldId = childText(def, 'FieldID')
    if (fieldId === undefined) {
      diag(sink, MISSING_MSPDI_FIELD, 'error', 'ExtendedAttribute definition missing <FieldID>')
      continue
    }
    const alias = childText(def, 'Alias') ?? fieldId
    const typeRaw = childText(def, 'Type')
    let type: CustomField['type'] = 'text'
    if (typeRaw !== undefined) {
      if (
        typeRaw === 'text' ||
        typeRaw === 'number' ||
        typeRaw === 'boolean' ||
        typeRaw === 'date'
      ) {
        type = typeRaw
      } else {
        diag(
          sink,
          UNSUPPORTED_MSPDI_FEATURE,
          'warning',
          `ExtendedAttribute ${fieldId} has unsupported <Type> ${typeRaw}; defaulted to 'text'`,
          fieldId,
        )
      }
    }
    out.push({ id: asCustomFieldId(fieldId), name: alias, type })
  }
  return out
}

// ---- properties ----------------------------------------------------------

function parseProperties(
  root: XmlNode,
  defaultCalendarId: string,
  sink: Sink,
): ProjectProperties | undefined {
  const nameRaw = childText(root, 'Name') ?? childText(root, 'Title')
  const idRaw = childText(root, 'UID') ?? nameRaw ?? 'project'
  const startDateRaw = childText(root, 'StartDate')
  if (startDateRaw === undefined) {
    diag(sink, MISSING_MSPDI_FIELD, 'error', '<Project> is missing <StartDate>')
    return undefined
  }
  const startDateNorm = normalizeMspdiDate(startDateRaw)
  if ('invalid' in startDateNorm) {
    diag(
      sink,
      INVALID_MSPDI_DATE,
      'error',
      `<Project><StartDate> is not a valid ISO-8601 date: ${JSON.stringify(startDateRaw)}`,
    )
    return undefined
  }
  const finishDate = optionalDate(root, 'FinishDate', sink)
  const statusDate = optionalDate(root, 'StatusDate', sink)
  return {
    id: idRaw,
    name: nameRaw ?? 'Imported MSPDI Project',
    startDate: startDateNorm.iso,
    ...(finishDate !== undefined ? { finishDate } : {}),
    ...(statusDate !== undefined ? { statusDate } : {}),
    defaultCalendarId: asCalendarId(defaultCalendarId),
  }
}

/**
 * Reconstruct the canonical task hierarchy from MSPDI outline information only.
 *
 * Canonical identity is NOT WBS (architecture-lock §4), but MSPDI's
 * `<OutlineNumber>` (e.g. `1`, `1.1`, `1.1.2`) deterministically implies the
 * parent: the parent's outline number is the current one with the last `.…`
 * suffix removed. `parentTaskId` is resolved by looking that prefix up in the
 * WBS→TaskId map.
 *
 * This never invents identity — `parentTaskId` is set only when the parent
 * outline number actually exists among the imported tasks. A dangling parent
 * (outline `1.1` present but `1` absent) emits `INVALID_MSPDI_REFERENCE` and
 * leaves `parentTaskId` unset (the engine's hierarchy validator then reports
 * `MISSING_PARENT`). The canonical hierarchy engine remains authoritative for
 * `outlineLevel`/`summary` consistency (`INCONSISTENT_OUTLINE_LEVEL`,
 * `INCONSISTENT_SUMMARY_FLAG`).
 */
function reconstructHierarchy(tasks: Task[], sink: Sink): Task[] {
  const wbsToId = new Map<string, TaskId>()
  for (const t of tasks) {
    if (t.wbs.length > 0) wbsToId.set(t.wbs, t.id)
  }
  const out: Task[] = []
  for (const t of tasks) {
    if (t.wbs.length === 0) {
      out.push(t)
      continue
    }
    const lastDot = t.wbs.lastIndexOf('.')
    if (lastDot === -1) {
      // Top-level outline (e.g. "1") — no parent.
      out.push(t)
      continue
    }
    const parentWbs = t.wbs.slice(0, lastDot)
    const parentId = wbsToId.get(parentWbs)
    if (parentId === undefined) {
      diag(
        sink,
        INVALID_MSPDI_REFERENCE,
        'error',
        `Task wbs ${t.wbs} references parent outline ${parentWbs} which is not present among imported tasks; parentTaskId left unset`,
        String(t.id),
      )
      out.push(t)
      continue
    }
    out.push({ ...t, parentTaskId: parentId })
  }
  return out
}

// ---- main entry points ---------------------------------------------------

/**
 * Import MSPDI XML bytes into a canonical `ProjectDocument`.
 *
 * File-level failures (bad XML, wrong root, unsupported save version, missing
 * `<StartDate>`) return `{ document: emptyProjectDocument(), diagnostics }`
 * with a single error diagnostic. Entity-level failures drop the entity and
 * emit an error diagnostic (partial recovery — never silent discard). After
 * construction the document is delegated to `validateProjectDocument` and its
 * diagnostics are surfaced verbatim (severity 'error').
 */
export function importMspdi(input: Uint8Array, metadata?: ProjectFileMetadata): MspdiImportResult {
  const sink: Diag[] = []
  let root: XmlNode
  try {
    root = parseXml(input)
  } catch (error) {
    const message = error instanceof XmlParseError ? error.message : 'XML parse failed'
    diag(sink, INVALID_MSPDI, 'error', `MSPDI XML is not well-formed: ${message}`)
    return { document: emptyProjectDocument(), diagnostics: sink }
  }
  if (root.name !== MSPDI_ROOT_ELEMENT) {
    diag(
      sink,
      INVALID_MSPDI,
      'error',
      `MSPDI root element must be <${MSPDI_ROOT_ELEMENT}>; found <${root.name}>`,
    )
    return { document: emptyProjectDocument(), diagnostics: sink }
  }
  // SaveVersion gate.
  const saveVersion = optionalInteger(root, 'SaveVersion')
  if (saveVersion !== undefined && !isSupportedMspdiSaveVersion(saveVersion)) {
    diag(
      sink,
      UNSUPPORTED_MSPDI_VERSION,
      'error',
      `MSPDI <SaveVersion> ${saveVersion} is not in the supported set; refusing to forward-read an unknown schema`,
    )
    return { document: emptyProjectDocument(), diagnostics: sink }
  }

  // ---- build UID→identity maps (so references can be resolved) -----------
  // Calendars first (tasks/resources reference calendars).
  const calendarsNode = firstChild(root, 'Calendars')
  const calendars: Calendar[] = []
  const calendarUidMap = new Map<number, string>()
  let defaultCalendarIdFallback = 'standard'
  if (calendarsNode !== undefined) {
    let foundDefault = false
    for (const calNode of childrenNamed(calendarsNode, 'Calendar')) {
      const cal = parseCalendar(calNode, sink)
      if (cal === undefined) continue
      calendars.push(cal)
      const calUid = optionalInteger(calNode, 'UID')
      if (calUid !== undefined) {
        calendarUidMap.set(calUid, String(cal.id))
        const isDefault = optionalBoolean(calNode, 'IsBaseCalendarDefault') ?? false
        const isBase = optionalBoolean(calNode, 'IsBaseCalendar') ?? false
        if (!foundDefault && (isDefault || isBase)) {
          defaultCalendarIdFallback = String(cal.id)
          foundDefault = true
        }
      }
    }
    if (!foundDefault && calendars.length > 0) {
      defaultCalendarIdFallback = String(calendars[0].id)
    }
  }

  // Resources.
  const resourcesNode = firstChild(root, 'Resources')
  const resources: Resource[] = []
  const resourceUidMap = new Map<number, ResourceId>()
  if (resourcesNode !== undefined) {
    for (const resNode of childrenNamed(resourcesNode, 'Resource')) {
      const res = parseResource(resNode, calendarUidMap, sink)
      if (res === undefined) continue
      resources.push(res)
      const resUid = optionalInteger(resNode, 'UID')
      if (resUid !== undefined) resourceUidMap.set(resUid, res.id)
    }
  }

  // Custom-field definitions (needed before tasks for value validation).
  const customFields = parseCustomFieldDefinitions(root, sink)
  const definedCustomFieldIds = new Set(customFields.map((c) => String(c.id)))

  // Tasks.
  const tasksNode = firstChild(root, 'Tasks')
  const tasksRaw: Task[] = []
  const taskUidMap = new Map<number, TaskId>()
  if (tasksNode !== undefined) {
    for (const taskNode of childrenNamed(tasksNode, 'Task')) {
      const task = parseTask(taskNode, calendarUidMap, definedCustomFieldIds, sink)
      if (task === undefined) continue
      tasksRaw.push(task)
      taskUidMap.set(task.uid, task.id)
    }
  }
  // Reconstruct canonical hierarchy from MSPDI outline info only (PROJECT-015:
  // canonical identity is NOT WBS, but WBS reconstructs parentTaskId). The
  // engine's hierarchy validator then checks parentTaskId/outlineLevel/
  // summary consistency — malformed MSPDI hierarchy produces diagnostics,
  // never an invalid ProjectDocument.
  const tasks = reconstructHierarchy(tasksRaw, sink)

  // Dependencies, assignments, baselines.
  const dependencies = parseDependencies(root, taskUidMap, sink)
  const assignments = parseAssignments(root, taskUidMap, resourceUidMap, sink)
  // Baselines need a capturedAt fallback (MSPDI carries no per-baseline
  // captured date). Use <LastSaved>, then <CreationDate>, then the project
  // <StartDate> as a deterministic, documented fallback.
  const lastSaved = optionalDate(root, 'LastSaved', sink)
  const creationDate = optionalDate(root, 'CreationDate', sink)
  const startDateText = childText(root, 'StartDate')
  let capturedAtFallback: ISODateTime
  if (lastSaved !== undefined) {
    capturedAtFallback = lastSaved
  } else if (creationDate !== undefined) {
    capturedAtFallback = creationDate
  } else if (startDateText !== undefined) {
    const r = normalizeMspdiDate(startDateText)
    capturedAtFallback = 'iso' in r ? r.iso : asISODateTime('1970-01-01T00:00:00.000Z')
  } else {
    capturedAtFallback = asISODateTime('1970-01-01T00:00:00.000Z')
  }
  const baselines = parseBaselines(root, taskUidMap, capturedAtFallback, sink)

  // Properties.
  const properties = parseProperties(root, defaultCalendarIdFallback, sink)
  if (properties === undefined) {
    return { document: emptyProjectDocument(), diagnostics: sink }
  }

  const document: ProjectDocument = {
    schemaVersion: 1,
    properties,
    tasks,
    resources,
    assignments,
    dependencies,
    calendars,
    baselines,
    customFields,
    views: [],
    tables: [],
    filters: [],
    groups: [],
  }

  // Delegate to the engine's canonical validator; surface its diagnostics
  // verbatim (severity 'error', no entityId) — the MSPDI adapter does NOT
  // invent a parallel semantic validator.
  const validation = validateProjectDocument(document)
  for (const d of validation.diagnostics) {
    sink.push({ code: d.code, severity: 'error', message: d.message })
  }

  diag(
    sink,
    MSPDI_READ,
    'info',
    `Imported MSPDI (SaveVersion ${saveVersion ?? 'unknown'}${metadata?.sourceName ? ', source ' + metadata.sourceName : ''}) as canonical ProjectDocument schemaVersion ${MSPDI_FORMAT_VERSION}`,
  )

  return { document, diagnostics: sink }
}

/** Fast metadata inspection without constructing the document. */
export function inspectMspdi(
  input: Uint8Array,
  metadata?: ProjectFileMetadata,
): ProjectFileMetadata {
  const fallback: ProjectFileMetadata = {
    format: MSPDI_FORMAT,
    version: String(MSPDI_FORMAT_VERSION),
    ...(metadata?.sourceName !== undefined ? { sourceName: metadata.sourceName } : {}),
  }
  let root: XmlNode
  try {
    root = parseXml(input)
  } catch {
    return fallback
  }
  if (root.name !== MSPDI_ROOT_ELEMENT) return fallback
  const saveVersion = optionalInteger(root, 'SaveVersion')
  return {
    format: MSPDI_FORMAT,
    version: String(saveVersion ?? MSPDI_FORMAT_VERSION),
    ...(metadata?.sourceName !== undefined ? { sourceName: metadata.sourceName } : {}),
  }
}

/** The canonical MSPDI file adapter (import + inspect only; NO export). */
export const mspdiFileAdapter: MspdiFileAdapter = {
  format: MSPDI_FORMAT,
  inspect: (input, metadata) => inspectMspdi(input, metadata),
  import: (input, metadata) => importMspdi(input, metadata),
}
