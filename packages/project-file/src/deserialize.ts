/**
 * PROJECT-014 — Canonical `.gproj` deserializer.
 *
 * Pipeline (per the PROJECT-014 brief):
 *   1. parse safely        — UTF-8 decode + `JSON.parse` (no `eval`, no
 *                            `Function`, no reviver, no prototype
 *                            deserialization, no arbitrary constructors).
 *   2. validate envelope   — magic `format`, `formatVersion` ∈ supported set.
 *   3. validate schema     — every required field present + correctly typed;
 *                            branded identity fields are strings; enum fields
 *                            are in their union; record-map keys are filtered
 *                            against prototype-pollution hazards.
 *   4. construct canonical — promote raw strings/numbers to branded values
 *                            through the single canonical promotion point
 *                            (`asTaskId`/`asISODateTime`/`asWorkingMinutes`/
 *                            … from `@genoffice/project-contracts`).
 *   5. validate document   — delegate to the engine's canonical
 *                            `validateProjectDocument` (no duplicated
 *                            diagnostic system — the engine's codes are
 *                            already valid `ImportDiagnostic.code` strings).
 *   6. return diagnostics  — every dropped entity or invalid reference is
 *                            surfaced as an error-level diagnostic; nothing is
 *                            silently discarded.
 *   7. malformed → fail    — file-level errors return an empty document + a
 *                            single error diagnostic; entity-level errors drop
 *                            the entity + emit an error diagnostic (partial
 *                            recovery, explicitly represented).
 *
 * The parser is host-independent (no Node/Electron/browser globals) and is
 * safe against prototype-pollution payloads (`__proto__`, `constructor`,
 * `prototype` keys are filtered), deeply nested structures (depth limit), and
 * oversized inputs (byte limit).
 */
import {
  asAssignmentId,
  asBaselineId,
  asCalendarId,
  asCustomFieldId,
  asDependencyId,
  asISODateTime,
  asProjectFilterId,
  asProjectGroupId,
  asProjectTableId,
  asProjectViewId,
  asResourceId,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type {
  Assignment,
  Baseline,
  Calendar,
  CalendarPeriod,
  CustomField,
  Dependency,
  DependencyType,
  ProjectDocument,
  ProjectFileMetadata,
  ProjectFilter,
  ProjectGroup,
  ProjectProperties,
  ProjectTable,
  ProjectView,
  Resource,
  Task,
  TaskType,
} from '@genoffice/project-contracts'
import { validateProjectDocument } from '@genoffice/project-engine'
import { decodeUtf8 } from './utf8.js'
import {
  GPROJ_FORMAT,
  GPROJ_FORMAT_VERSION,
  GPROJ_FORBIDDEN_KEYS,
  GPROJ_MAX_INPUT_BYTES,
  GPROJ_MAX_PARSE_DEPTH,
  isSupportedGprojVersion,
} from './envelope.js'
import {
  INVALID_BASELINE,
  INVALID_CALENDAR,
  INVALID_ASSIGNMENT,
  INVALID_GPROJ,
  INVALID_IDENTITY,
  INVALID_REFERENCE,
  INVALID_RESOURCE,
  INVALID_TASK,
  MISSING_REQUIRED_FIELD,
  SCHEMA_INVALID,
  UNSUPPORTED_GPROJ_VERSION,
} from './diagnostics.js'
import { emptyProjectDocument } from './serialize.js'
import { buildGprojEnvelope } from './serialize.js'

export interface GprojImportResult {
  document: ProjectDocument
  diagnostics: import('@genoffice/project-contracts').ImportDiagnostic[]
}

// ---- primitive type guards ----------------------------------------------

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}
function isString(v: unknown): v is string {
  return typeof v === 'string'
}
function isNumber(v: unknown): v is number {
  return typeof v === 'number' && Number.isFinite(v)
}
function isInteger(v: unknown): v is number {
  return isNumber(v) && Number.isInteger(v)
}
function isBoolean(v: unknown): v is boolean {
  return typeof v === 'boolean'
}
function isArray(v: unknown): v is unknown[] {
  return Array.isArray(v)
}

/** ISO-8601 UTC validator (strict). Allows optional milliseconds. */
const ISO_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,9})?Z$/
function isISODateTime(v: unknown): v is string {
  return isString(v) && ISO_RE.test(v) && !Number.isNaN(new Date(v).getTime())
}

/** Exception date (YYYY-MM-DD) for calendar exceptions. */
const EXCEPTION_DATE_RE = /^\d{4}-\d{2}-\d{2}$/
function isExceptionDate(v: unknown): v is string {
  return (
    isString(v) &&
    EXCEPTION_DATE_RE.test(v) &&
    !Number.isNaN(new Date(v + 'T00:00:00.000Z').getTime())
  )
}

/** Safe object-key list — filters `__proto__`/`constructor`/`prototype`. */
function safeKeys(obj: Record<string, unknown>): string[] {
  return Object.keys(obj).filter((k) => !GPROJ_FORBIDDEN_KEYS.includes(k))
}

// ---- diagnostic helper ---------------------------------------------------

type Diag = import('@genoffice/project-contracts').ImportDiagnostic
function diag(
  code: string,
  severity: 'info' | 'warning' | 'error',
  message: string,
  entityId?: string,
): Diag {
  return entityId !== undefined
    ? { code, severity, message, entityId }
    : { code, severity, message }
}

// ---- depth-safe JSON parse ----------------------------------------------

function safeJsonParse(text: string, diagnostics: Diag[]): unknown | undefined {
  let value: unknown
  try {
    value = JSON.parse(text)
  } catch {
    diagnostics.push(diag(INVALID_GPROJ, 'error', 'Input is not valid JSON'))
    return undefined
  }
  if (!isObject(value) && !isArray(value)) {
    diagnostics.push(diag(INVALID_GPROJ, 'error', 'JSON root is not an object'))
    return undefined
  }
  // Walk and validate depth + key safety. Throws on violation.
  try {
    walkSafe(value, 0)
  } catch (message) {
    if (typeof message === 'string') {
      diagnostics.push(diag(SCHEMA_INVALID, 'error', message))
    } else {
      diagnostics.push(diag(SCHEMA_INVALID, 'error', 'Structure too deep or malformed'))
    }
    return undefined
  }
  return value
}

function walkSafe(value: unknown, depth: number): void {
  if (depth > GPROJ_MAX_PARSE_DEPTH) {
    throw 'Exceeded maximum parse depth (' + GPROJ_MAX_PARSE_DEPTH + ')'
  }
  if (isArray(value)) {
    for (const item of value) walkSafe(item, depth + 1)
    return
  }
  if (isObject(value)) {
    for (const key of Object.keys(value)) {
      if (GPROJ_FORBIDDEN_KEYS.includes(key)) {
        throw 'Forbidden key "' + key + '" in parsed object'
      }
      walkSafe(value[key], depth + 1)
    }
  }
}

// ---- schema-level entity validators --------------------------------------
//
// Each `expect*` helper takes the raw parsed value and a diagnostics sink.
// On success it returns the canonical (brand-promoted) entity; on failure it
// pushes an error diagnostic and returns `undefined` (the caller drops the
// entity — partial recovery, explicitly represented).

// Each `expect*` helper emits `MISSING_REQUIRED_FIELD` when the value is
// absent (`undefined`) and the caller-supplied entity code when the value is
// present-but-wrong-type. This aligns with the PROJECT-014 brief, which
// distinguishes "a required field is missing" from "a field has the wrong
// primitive type / is not in the allowed enum".
function expectString(
  v: unknown,
  field: string,
  entityId: string,
  code: string,
  diagnostics: Diag[],
): string | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isString(v)) {
    diagnostics.push(
      diag(code, 'error', `Field "${field}" on ${entityId} must be a string`, entityId),
    )
    return undefined
  }
  return v
}
function expectNonEmptyString(
  v: unknown,
  field: string,
  entityId: string,
  code: string,
  diagnostics: Diag[],
): string | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isString(v) || v.length === 0) {
    diagnostics.push(
      diag(code, 'error', `Field "${field}" on ${entityId} must be a non-empty string`, entityId),
    )
    return undefined
  }
  return v
}
function expectNumber(
  v: unknown,
  field: string,
  entityId: string,
  code: string,
  diagnostics: Diag[],
): number | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isNumber(v)) {
    diagnostics.push(
      diag(code, 'error', `Field "${field}" on ${entityId} must be a finite number`, entityId),
    )
    return undefined
  }
  return v
}
function expectInteger(
  v: unknown,
  field: string,
  entityId: string,
  code: string,
  diagnostics: Diag[],
): number | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isInteger(v)) {
    diagnostics.push(
      diag(code, 'error', `Field "${field}" on ${entityId} must be an integer`, entityId),
    )
    return undefined
  }
  return v
}
function expectBoolean(
  v: unknown,
  field: string,
  entityId: string,
  code: string,
  diagnostics: Diag[],
): boolean | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isBoolean(v)) {
    diagnostics.push(
      diag(code, 'error', `Field "${field}" on ${entityId} must be a boolean`, entityId),
    )
    return undefined
  }
  return v
}
function expectISODateTime(
  v: unknown,
  field: string,
  entityId: string,
  code: string,
  diagnostics: Diag[],
): string | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isISODateTime(v)) {
    diagnostics.push(
      diag(
        code,
        'error',
        `Field "${field}" on ${entityId} must be an ISO-8601 UTC timestamp`,
        entityId,
      ),
    )
    return undefined
  }
  return v
}
function expectEnum<T extends string>(
  v: unknown,
  field: string,
  entityId: string,
  allowed: readonly T[],
  code: string,
  diagnostics: Diag[],
): T | undefined {
  if (v === undefined) {
    diagnostics.push(
      diag(MISSING_REQUIRED_FIELD, 'error', `Field "${field}" on ${entityId} is missing`, entityId),
    )
    return undefined
  }
  if (!isString(v) || !allowed.includes(v as T)) {
    diagnostics.push(
      diag(
        code,
        'error',
        `Field "${field}" on ${entityId} must be one of [${allowed.join(', ')}]`,
        entityId,
      ),
    )
    return undefined
  }
  return v as T
}

// ---- entity constructors -------------------------------------------------

const TASK_TYPES: readonly TaskType[] = ['fixedUnits', 'fixedWork', 'fixedDuration']
const DEPENDENCY_TYPES: readonly DependencyType[] = ['FS', 'SS', 'FF', 'SF']
const CONSTRAINT_TYPES = [
  'asSoonAsPossible',
  'asLateAsPossible',
  'startNoEarlierThan',
  'startNoLaterThan',
  'mustStartOn',
  'finishNoEarlierThan',
  'finishNoLaterThan',
  'mustFinishOn',
] as const
const RESOURCE_KINDS = ['work', 'material', 'cost'] as const
const CUSTOM_FIELD_TYPES = ['text', 'number', 'boolean', 'date'] as const

function parseTask(raw: unknown, diagnostics: Diag[]): Task | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(INVALID_TASK, 'error', 'Task entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(task)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const uid = expectInteger(raw.uid, 'uid', id, INVALID_TASK, diagnostics)
  const wbs = expectString(raw.wbs, 'wbs', id, INVALID_TASK, diagnostics)
  const outlineLevel = expectInteger(
    raw.outlineLevel,
    'outlineLevel',
    id,
    INVALID_TASK,
    diagnostics,
  )
  const name = expectString(raw.name, 'name', id, INVALID_TASK, diagnostics)
  const taskType = expectEnum(raw.taskType, 'taskType', id, TASK_TYPES, INVALID_TASK, diagnostics)
  const summary = expectBoolean(raw.summary, 'summary', id, INVALID_TASK, diagnostics)
  const milestone = expectBoolean(raw.milestone, 'milestone', id, INVALID_TASK, diagnostics)
  const manualScheduled = expectBoolean(
    raw.manualScheduled,
    'manualScheduled',
    id,
    INVALID_TASK,
    diagnostics,
  )
  const autoScheduled = expectBoolean(
    raw.autoScheduled,
    'autoScheduled',
    id,
    INVALID_TASK,
    diagnostics,
  )
  const duration = expectInteger(raw.duration, 'duration', id, INVALID_TASK, diagnostics)
  const priority = expectInteger(raw.priority, 'priority', id, INVALID_TASK, diagnostics)
  const percentComplete = expectNumber(
    raw.percentComplete,
    'percentComplete',
    id,
    INVALID_TASK,
    diagnostics,
  )
  const work = expectInteger(raw.work, 'work', id, INVALID_TASK, diagnostics)
  const remainingWork = expectInteger(
    raw.remainingWork,
    'remainingWork',
    id,
    INVALID_TASK,
    diagnostics,
  )
  const actualWork = expectInteger(raw.actualWork, 'actualWork', id, INVALID_TASK, diagnostics)
  const cost = expectNumber(raw.cost, 'cost', id, INVALID_TASK, diagnostics)
  const actualCost = expectNumber(raw.actualCost, 'actualCost', id, INVALID_TASK, diagnostics)
  const remainingCost = expectNumber(
    raw.remainingCost,
    'remainingCost',
    id,
    INVALID_TASK,
    diagnostics,
  )
  if (
    uid === undefined ||
    wbs === undefined ||
    outlineLevel === undefined ||
    name === undefined ||
    taskType === undefined ||
    summary === undefined ||
    milestone === undefined ||
    manualScheduled === undefined ||
    autoScheduled === undefined ||
    duration === undefined ||
    priority === undefined ||
    percentComplete === undefined ||
    work === undefined ||
    remainingWork === undefined ||
    actualWork === undefined ||
    cost === undefined ||
    actualCost === undefined ||
    remainingCost === undefined
  ) {
    return undefined
  }
  // Optional fields
  const start =
    raw.start === undefined
      ? undefined
      : expectISODateTime(raw.start, 'start', id, INVALID_TASK, diagnostics)
  const finish =
    raw.finish === undefined
      ? undefined
      : expectISODateTime(raw.finish, 'finish', id, INVALID_TASK, diagnostics)
  const constraintType =
    raw.constraintType === undefined
      ? undefined
      : expectEnum(
          raw.constraintType,
          'constraintType',
          id,
          CONSTRAINT_TYPES,
          INVALID_TASK,
          diagnostics,
        )
  const constraintDate =
    raw.constraintDate === undefined
      ? undefined
      : expectISODateTime(raw.constraintDate, 'constraintDate', id, INVALID_TASK, diagnostics)
  const deadline =
    raw.deadline === undefined
      ? undefined
      : expectISODateTime(raw.deadline, 'deadline', id, INVALID_TASK, diagnostics)
  const calendarId =
    raw.calendarId === undefined
      ? undefined
      : expectNonEmptyString(raw.calendarId, 'calendarId', id, INVALID_REFERENCE, diagnostics)
  const parentTaskId =
    raw.parentTaskId === undefined
      ? undefined
      : expectNonEmptyString(raw.parentTaskId, 'parentTaskId', id, INVALID_REFERENCE, diagnostics)
  const physicalPercentComplete =
    raw.physicalPercentComplete === undefined
      ? undefined
      : expectNumber(
          raw.physicalPercentComplete,
          'physicalPercentComplete',
          id,
          INVALID_TASK,
          diagnostics,
        )
  // baseline: BaselineId[]
  const baselineRaw = raw.baseline
  let baseline: string[] = []
  if (baselineRaw !== undefined) {
    if (!isArray(baselineRaw)) {
      diagnostics.push(
        diag(
          INVALID_TASK,
          'error',
          `Field "baseline" on ${id} must be an array of baseline ids`,
          id,
        ),
      )
      return undefined
    }
    const baselineIds: string[] = []
    for (const b of baselineRaw) {
      const bId = expectNonEmptyString(b, 'baseline[]', id, INVALID_REFERENCE, diagnostics)
      if (bId !== undefined) baselineIds.push(bId)
    }
    baseline = baselineIds
  }
  // customFields: Record<CustomFieldId, string|number|boolean|null>
  const customFieldsRaw = raw.customFields
  const customFields: Record<string, string | number | boolean | null> = {}
  if (customFieldsRaw !== undefined) {
    if (!isObject(customFieldsRaw)) {
      diagnostics.push(
        diag(INVALID_TASK, 'error', `Field "customFields" on ${id} must be an object`, id),
      )
      return undefined
    }
    for (const key of safeKeys(customFieldsRaw)) {
      const val = customFieldsRaw[key]
      if (val === null || isString(val) || isNumber(val) || isBoolean(val)) {
        customFields[key] = val
      } else {
        diagnostics.push(
          diag(
            INVALID_TASK,
            'error',
            `customField "${key}" on ${id} must be string|number|boolean|null`,
            id,
          ),
        )
      }
    }
  }
  // notes: string[]
  const notesRaw = raw.notes
  const notes: string[] = []
  if (notesRaw !== undefined) {
    if (!isArray(notesRaw)) {
      diagnostics.push(
        diag(INVALID_TASK, 'error', `Field "notes" on ${id} must be an array of strings`, id),
      )
      return undefined
    }
    for (const n of notesRaw) {
      const ns = expectString(n, 'notes[]', id, INVALID_TASK, diagnostics)
      if (ns !== undefined) notes.push(ns)
    }
  }
  if (
    start === null ||
    finish === null ||
    constraintType === null ||
    constraintDate === null ||
    deadline === null ||
    calendarId === null ||
    parentTaskId === null ||
    physicalPercentComplete === null
  ) {
    return undefined
  }
  const task: Task = {
    id: asTaskId(id),
    uid,
    wbs,
    outlineLevel,
    name,
    taskType,
    summary,
    milestone,
    manualScheduled,
    autoScheduled,
    duration: asWorkingMinutes(duration),
    priority,
    percentComplete,
    work: asWorkingMinutes(work),
    remainingWork: asWorkingMinutes(remainingWork),
    actualWork: asWorkingMinutes(actualWork),
    cost,
    actualCost,
    remainingCost,
    baseline: baseline.map(asBaselineId),
    customFields: customFields as Task['customFields'],
    notes,
  }
  if (start !== undefined) task.start = asISODateTime(start)
  if (finish !== undefined) task.finish = asISODateTime(finish)
  if (constraintType !== undefined) task.constraintType = constraintType
  if (constraintDate !== undefined) task.constraintDate = asISODateTime(constraintDate)
  if (deadline !== undefined) task.deadline = asISODateTime(deadline)
  if (calendarId !== undefined) task.calendarId = asCalendarId(calendarId)
  if (parentTaskId !== undefined) task.parentTaskId = asTaskId(parentTaskId)
  if (physicalPercentComplete !== undefined) task.physicalPercentComplete = physicalPercentComplete
  return task
}

function parseResource(raw: unknown, diagnostics: Diag[]): Resource | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(INVALID_RESOURCE, 'error', 'Resource entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(resource)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const uid = expectInteger(raw.uid, 'uid', id, INVALID_RESOURCE, diagnostics)
  const name = expectString(raw.name, 'name', id, INVALID_RESOURCE, diagnostics)
  const kind = expectEnum(raw.kind, 'kind', id, RESOURCE_KINDS, INVALID_RESOURCE, diagnostics)
  const maxUnits = expectNumber(raw.maxUnits, 'maxUnits', id, INVALID_RESOURCE, diagnostics)
  const standardRate = expectNumber(
    raw.standardRate,
    'standardRate',
    id,
    INVALID_RESOURCE,
    diagnostics,
  )
  const overtimeRate = expectNumber(
    raw.overtimeRate,
    'overtimeRate',
    id,
    INVALID_RESOURCE,
    diagnostics,
  )
  const costPerUse = expectNumber(raw.costPerUse, 'costPerUse', id, INVALID_RESOURCE, diagnostics)
  if (
    uid === undefined ||
    name === undefined ||
    kind === undefined ||
    maxUnits === undefined ||
    standardRate === undefined ||
    overtimeRate === undefined ||
    costPerUse === undefined
  ) {
    return undefined
  }
  const calendarId =
    raw.calendarId === undefined
      ? undefined
      : expectNonEmptyString(raw.calendarId, 'calendarId', id, INVALID_REFERENCE, diagnostics)
  // availability: Array<{ start, finish?, units }>
  const availabilityRaw = raw.availability
  const availability: Resource['availability'] = []
  if (availabilityRaw !== undefined) {
    if (!isArray(availabilityRaw)) {
      diagnostics.push(
        diag(INVALID_RESOURCE, 'error', `Field "availability" on ${id} must be an array`, id),
      )
      return undefined
    }
    for (const a of availabilityRaw) {
      if (!isObject(a)) {
        diagnostics.push(
          diag(INVALID_RESOURCE, 'error', `availability entry on ${id} is not an object`, id),
        )
        continue
      }
      const aStart = expectISODateTime(
        a.start,
        'availability[].start',
        id,
        INVALID_RESOURCE,
        diagnostics,
      )
      const aUnits = expectNumber(
        a.units,
        'availability[].units',
        id,
        INVALID_RESOURCE,
        diagnostics,
      )
      const aFinish =
        a.finish === undefined
          ? undefined
          : expectISODateTime(a.finish, 'availability[].finish', id, INVALID_RESOURCE, diagnostics)
      if (aStart === undefined || aUnits === undefined || aFinish === null) continue
      availability.push({
        start: asISODateTime(aStart),
        units: aUnits,
        ...(aFinish !== undefined ? { finish: asISODateTime(aFinish) } : {}),
      })
    }
  }
  if (calendarId === null) return undefined
  const resource: Resource = {
    id: asResourceId(id),
    uid,
    name,
    kind,
    maxUnits,
    standardRate,
    overtimeRate,
    costPerUse,
    availability,
  }
  if (calendarId !== undefined) resource.calendarId = asCalendarId(calendarId)
  return resource
}

function parseAssignment(raw: unknown, diagnostics: Diag[]): Assignment | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(INVALID_ASSIGNMENT, 'error', 'Assignment entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(assignment)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const taskId = expectNonEmptyString(raw.taskId, 'taskId', id, INVALID_REFERENCE, diagnostics)
  const resourceId = expectNonEmptyString(
    raw.resourceId,
    'resourceId',
    id,
    INVALID_REFERENCE,
    diagnostics,
  )
  const units = expectNumber(raw.units, 'units', id, INVALID_ASSIGNMENT, diagnostics)
  const work = expectInteger(raw.work, 'work', id, INVALID_ASSIGNMENT, diagnostics)
  const actualWork = expectInteger(
    raw.actualWork,
    'actualWork',
    id,
    INVALID_ASSIGNMENT,
    diagnostics,
  )
  const remainingWork = expectInteger(
    raw.remainingWork,
    'remainingWork',
    id,
    INVALID_ASSIGNMENT,
    diagnostics,
  )
  const cost = expectNumber(raw.cost, 'cost', id, INVALID_ASSIGNMENT, diagnostics)
  const actualCost = expectNumber(raw.actualCost, 'actualCost', id, INVALID_ASSIGNMENT, diagnostics)
  const remainingCost = expectNumber(
    raw.remainingCost,
    'remainingCost',
    id,
    INVALID_ASSIGNMENT,
    diagnostics,
  )
  if (
    taskId === undefined ||
    resourceId === undefined ||
    units === undefined ||
    work === undefined ||
    actualWork === undefined ||
    remainingWork === undefined ||
    cost === undefined ||
    actualCost === undefined ||
    remainingCost === undefined
  ) {
    return undefined
  }
  return {
    id: asAssignmentId(id),
    taskId: asTaskId(taskId),
    resourceId: asResourceId(resourceId),
    units,
    work: asWorkingMinutes(work),
    actualWork: asWorkingMinutes(actualWork),
    remainingWork: asWorkingMinutes(remainingWork),
    cost,
    actualCost,
    remainingCost,
  }
}

function parseDependency(raw: unknown, diagnostics: Diag[]): Dependency | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(INVALID_REFERENCE, 'error', 'Dependency entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(dependency)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const predecessorId = expectNonEmptyString(
    raw.predecessorId,
    'predecessorId',
    id,
    INVALID_REFERENCE,
    diagnostics,
  )
  const successorId = expectNonEmptyString(
    raw.successorId,
    'successorId',
    id,
    INVALID_REFERENCE,
    diagnostics,
  )
  const type = expectEnum(raw.type, 'type', id, DEPENDENCY_TYPES, INVALID_REFERENCE, diagnostics)
  const lagMinutes = expectNumber(raw.lagMinutes, 'lagMinutes', id, INVALID_REFERENCE, diagnostics)
  if (
    predecessorId === undefined ||
    successorId === undefined ||
    type === undefined ||
    lagMinutes === undefined
  ) {
    return undefined
  }
  return {
    id: asDependencyId(id),
    predecessorId: asTaskId(predecessorId),
    successorId: asTaskId(successorId),
    type,
    lagMinutes,
  }
}

function parseCalendarPeriod(
  raw: unknown,
  entityId: string,
  field: string,
  diagnostics: Diag[],
): CalendarPeriod | undefined {
  if (!isObject(raw)) {
    diagnostics.push(
      diag(INVALID_CALENDAR, 'error', `${field} on ${entityId} is not an object`, entityId),
    )
    return undefined
  }
  const startMinute = expectInteger(
    raw.startMinute,
    'startMinute',
    entityId,
    INVALID_CALENDAR,
    diagnostics,
  )
  const endMinute = expectInteger(
    raw.endMinute,
    'endMinute',
    entityId,
    INVALID_CALENDAR,
    diagnostics,
  )
  if (startMinute === undefined || endMinute === undefined) return undefined
  if (startMinute < 0 || endMinute > 1440 || startMinute >= endMinute) {
    diagnostics.push(
      diag(
        INVALID_CALENDAR,
        'error',
        `${field} on ${entityId} has a period outside 00:00-24:00 or with an empty interval`,
        entityId,
      ),
    )
    return undefined
  }
  return { startMinute, endMinute }
}

function parseCalendar(raw: unknown, diagnostics: Diag[]): Calendar | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(INVALID_CALENDAR, 'error', 'Calendar entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(calendar)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, INVALID_CALENDAR, diagnostics)
  const baseCalendarId =
    raw.baseCalendarId === undefined
      ? undefined
      : expectNonEmptyString(
          raw.baseCalendarId,
          'baseCalendarId',
          id,
          INVALID_REFERENCE,
          diagnostics,
        )
  if (name === undefined || baseCalendarId === null) return undefined
  // workingWeek: Record<number, CalendarPeriod[]>
  const workingWeekRaw = raw.workingWeek
  const workingWeek: Record<number, CalendarPeriod[]> = {}
  if (workingWeekRaw === undefined) {
    // treated as empty — a calendar with no workingWeek is a non-working calendar
  } else if (!isObject(workingWeekRaw)) {
    diagnostics.push(
      diag(INVALID_CALENDAR, 'error', `Field "workingWeek" on ${id} must be an object`, id),
    )
    return undefined
  } else {
    for (const key of safeKeys(workingWeekRaw)) {
      const dayNum = Number(key)
      if (!Number.isInteger(dayNum) || dayNum < 0 || dayNum > 6) {
        diagnostics.push(
          diag(
            INVALID_CALENDAR,
            'error',
            `workingWeek key "${key}" on ${id} must be an integer 0-6`,
            id,
          ),
        )
        continue
      }
      const periodsRaw = workingWeekRaw[key]
      if (!isArray(periodsRaw)) {
        diagnostics.push(
          diag(
            INVALID_CALENDAR,
            'error',
            `workingWeek[${key}] on ${id} must be an array of periods`,
            id,
          ),
        )
        continue
      }
      const periods: CalendarPeriod[] = []
      for (const p of periodsRaw) {
        const period = parseCalendarPeriod(p, id, `workingWeek[${key}]`, diagnostics)
        if (period !== undefined) periods.push(period)
      }
      workingWeek[dayNum] = periods
    }
  }
  // exceptions: CalendarException[]
  const exceptionsRaw = raw.exceptions
  const exceptions: Calendar['exceptions'] = []
  if (exceptionsRaw !== undefined) {
    if (!isArray(exceptionsRaw)) {
      diagnostics.push(
        diag(INVALID_CALENDAR, 'error', `Field "exceptions" on ${id} must be an array`, id),
      )
      return undefined
    }
    for (const e of exceptionsRaw) {
      if (!isObject(e)) {
        diagnostics.push(
          diag(INVALID_CALENDAR, 'error', `exception entry on ${id} is not an object`, id),
        )
        continue
      }
      if (!isExceptionDate(e.date)) {
        diagnostics.push(
          diag(INVALID_CALENDAR, 'error', `exception date on ${id} must be YYYY-MM-DD`, id),
        )
        continue
      }
      if (!isArray(e.periods)) {
        diagnostics.push(
          diag(INVALID_CALENDAR, 'error', `exception.periods on ${id} must be an array`, id),
        )
        continue
      }
      const periods: CalendarPeriod[] = []
      for (const p of e.periods) {
        const period = parseCalendarPeriod(p, id, `exception[${e.date}].periods`, diagnostics)
        if (period !== undefined) periods.push(period)
      }
      exceptions.push({ date: e.date, periods })
    }
  }
  const calendar: Calendar = {
    id: asCalendarId(id),
    name,
    workingWeek,
    exceptions,
  }
  if (baseCalendarId !== undefined) calendar.baseCalendarId = asCalendarId(baseCalendarId)
  return calendar
}

function parseBaseline(raw: unknown, diagnostics: Diag[]): Baseline | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(INVALID_BASELINE, 'error', 'Baseline entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(baseline)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, INVALID_BASELINE, diagnostics)
  const capturedAt = expectISODateTime(
    raw.capturedAt,
    'capturedAt',
    id,
    INVALID_BASELINE,
    diagnostics,
  )
  if (name === undefined || capturedAt === undefined) return undefined
  const snapshotsRaw = raw.taskSnapshots
  const taskSnapshots: Baseline['taskSnapshots'] = {}
  if (snapshotsRaw !== undefined) {
    if (!isObject(snapshotsRaw)) {
      diagnostics.push(
        diag(INVALID_BASELINE, 'error', `Field "taskSnapshots" on ${id} must be an object`, id),
      )
      return undefined
    }
    for (const key of safeKeys(snapshotsRaw)) {
      const snap = snapshotsRaw[key]
      if (!isObject(snap)) {
        diagnostics.push(
          diag(INVALID_BASELINE, 'error', `taskSnapshots["${key}"] on ${id} is not an object`, id),
        )
        continue
      }
      const duration = expectInteger(snap.duration, 'duration', id, INVALID_BASELINE, diagnostics)
      const work = expectInteger(snap.work, 'work', id, INVALID_BASELINE, diagnostics)
      const cost = expectNumber(snap.cost, 'cost', id, INVALID_BASELINE, diagnostics)
      const start =
        snap.start === undefined
          ? undefined
          : expectISODateTime(snap.start, 'start', id, INVALID_BASELINE, diagnostics)
      const finish =
        snap.finish === undefined
          ? undefined
          : expectISODateTime(snap.finish, 'finish', id, INVALID_BASELINE, diagnostics)
      if (
        duration === undefined ||
        work === undefined ||
        cost === undefined ||
        start === null ||
        finish === null
      )
        continue
      taskSnapshots[key] = {
        duration: asWorkingMinutes(duration),
        work: asWorkingMinutes(work),
        cost,
        ...(start !== undefined ? { start: asISODateTime(start) } : {}),
        ...(finish !== undefined ? { finish: asISODateTime(finish) } : {}),
      }
    }
  }
  return {
    id: asBaselineId(id),
    name,
    capturedAt: asISODateTime(capturedAt),
    taskSnapshots,
  }
}

function parseCustomField(raw: unknown, diagnostics: Diag[]): CustomField | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', 'CustomField entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(custom field)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, SCHEMA_INVALID, diagnostics)
  const type = expectEnum(raw.type, 'type', id, CUSTOM_FIELD_TYPES, SCHEMA_INVALID, diagnostics)
  if (name === undefined || type === undefined) return undefined
  return { id: asCustomFieldId(id), name, type }
}

function parseView(raw: unknown, diagnostics: Diag[]): ProjectView | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', 'ProjectView entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(view)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, SCHEMA_INVALID, diagnostics)
  const type = expectString(raw.type, 'type', id, SCHEMA_INVALID, diagnostics)
  if (name === undefined || type === undefined) return undefined
  const tableId =
    raw.tableId === undefined
      ? undefined
      : expectNonEmptyString(raw.tableId, 'tableId', id, INVALID_REFERENCE, diagnostics)
  const filterId =
    raw.filterId === undefined
      ? undefined
      : expectNonEmptyString(raw.filterId, 'filterId', id, INVALID_REFERENCE, diagnostics)
  const groupId =
    raw.groupId === undefined
      ? undefined
      : expectNonEmptyString(raw.groupId, 'groupId', id, INVALID_REFERENCE, diagnostics)
  if (tableId === null || filterId === null || groupId === null) return undefined
  const view: ProjectView = { id: asProjectViewId(id), name, type }
  if (tableId !== undefined) view.tableId = asProjectTableId(tableId)
  if (filterId !== undefined) view.filterId = asProjectFilterId(filterId)
  if (groupId !== undefined) view.groupId = asProjectGroupId(groupId)
  return view
}

function parseTable(raw: unknown, diagnostics: Diag[]): ProjectTable | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', 'ProjectTable entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(table)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, SCHEMA_INVALID, diagnostics)
  if (name === undefined) return undefined
  if (!isArray(raw.columns)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', `Field "columns" on ${id} must be an array`, id))
    return undefined
  }
  const columns: string[] = []
  for (const c of raw.columns) {
    const cs = expectString(c, 'columns[]', id, SCHEMA_INVALID, diagnostics)
    if (cs !== undefined) columns.push(cs)
  }
  return { id: asProjectTableId(id), name, columns }
}

function parseFilter(raw: unknown, diagnostics: Diag[]): ProjectFilter | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', 'ProjectFilter entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(filter)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, SCHEMA_INVALID, diagnostics)
  const expression = expectString(raw.expression, 'expression', id, SCHEMA_INVALID, diagnostics)
  if (name === undefined || expression === undefined) return undefined
  return { id: asProjectFilterId(id), name, expression }
}

function parseGroup(raw: unknown, diagnostics: Diag[]): ProjectGroup | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', 'ProjectGroup entry is not an object'))
    return undefined
  }
  const id = expectNonEmptyString(raw.id, 'id', '(group)', INVALID_IDENTITY, diagnostics)
  if (id === undefined) return undefined
  const name = expectString(raw.name, 'name', id, SCHEMA_INVALID, diagnostics)
  const expression = expectString(raw.expression, 'expression', id, SCHEMA_INVALID, diagnostics)
  if (name === undefined || expression === undefined) return undefined
  return { id: asProjectGroupId(id), name, expression }
}

function parseProperties(raw: unknown, diagnostics: Diag[]): ProjectProperties | undefined {
  if (!isObject(raw)) {
    diagnostics.push(diag(MISSING_REQUIRED_FIELD, 'error', 'properties is not an object'))
    return undefined
  }
  const id = expectString(raw.id, 'id', '(properties)', MISSING_REQUIRED_FIELD, diagnostics)
  const name = expectString(raw.name, 'name', '(properties)', MISSING_REQUIRED_FIELD, diagnostics)
  const startDate = expectISODateTime(
    raw.startDate,
    'startDate',
    '(properties)',
    MISSING_REQUIRED_FIELD,
    diagnostics,
  )
  const defaultCalendarId = expectNonEmptyString(
    raw.defaultCalendarId,
    'defaultCalendarId',
    '(properties)',
    MISSING_REQUIRED_FIELD,
    diagnostics,
  )
  if (
    id === undefined ||
    name === undefined ||
    startDate === undefined ||
    defaultCalendarId === undefined
  )
    return undefined
  const finishDate =
    raw.finishDate === undefined
      ? undefined
      : expectISODateTime(raw.finishDate, 'finishDate', '(properties)', SCHEMA_INVALID, diagnostics)
  const statusDate =
    raw.statusDate === undefined
      ? undefined
      : expectISODateTime(raw.statusDate, 'statusDate', '(properties)', SCHEMA_INVALID, diagnostics)
  if (finishDate === null || statusDate === null) return undefined
  const props: ProjectProperties = {
    id,
    name,
    startDate: asISODateTime(startDate),
    defaultCalendarId: asCalendarId(defaultCalendarId),
  }
  if (finishDate !== undefined) props.finishDate = asISODateTime(finishDate)
  if (statusDate !== undefined) props.statusDate = asISODateTime(statusDate)
  return props
}

function parseArray<T>(
  raw: unknown,
  field: string,
  parse: (item: unknown, d: Diag[]) => T | undefined,
  diagnostics: Diag[],
): T[] {
  if (raw === undefined) return []
  if (!isArray(raw)) {
    diagnostics.push(diag(SCHEMA_INVALID, 'error', `Field "${field}" must be an array`))
    return []
  }
  const out: T[] = []
  for (const item of raw) {
    const entity = parse(item, diagnostics)
    if (entity !== undefined) out.push(entity)
  }
  return out
}

// ---- top-level deserializer ----------------------------------------------

/**
 * Deserialize `.gproj` bytes into a canonical `ProjectDocument`.
 *
 * Returns `{ document, diagnostics }`. On file-level errors (bad JSON, wrong
 * magic, unsupported version) the document is the canonical empty document
 * and the diagnostics contain exactly one error. On entity-level errors the
 * document omits the malformed entities and the diagnostics enumerate each
 * dropped entity. The canonical engine validator runs after construction so
 * any remaining semantic violation (duplicate id, dangling reference, …) is
 * surfaced as an additional error diagnostic.
 */
export function deserializeGproj(
  input: Uint8Array,
  metadata?: ProjectFileMetadata,
): GprojImportResult {
  const diagnostics: Diag[] = []

  // (1) Input size guard.
  if (input.byteLength > GPROJ_MAX_INPUT_BYTES) {
    diagnostics.push(
      diag(
        INVALID_GPROJ,
        'error',
        'Input exceeds maximum size (' + GPROJ_MAX_INPUT_BYTES + ' bytes)',
      ),
    )
    return { document: emptyProjectDocument(), diagnostics }
  }

  // (2) UTF-8 decode.
  let text: string
  try {
    text = decodeUtf8(input)
  } catch (e) {
    diagnostics.push(
      diag(INVALID_GPROJ, 'error', 'Input is not valid UTF-8: ' + (e as Error).message),
    )
    return { document: emptyProjectDocument(), diagnostics }
  }

  // (3) JSON parse + depth/key safety walk.
  const root = safeJsonParse(text, diagnostics)
  if (root === undefined) {
    return { document: emptyProjectDocument(), diagnostics }
  }
  if (!isObject(root)) {
    diagnostics.push(diag(INVALID_GPROJ, 'error', 'Root is not a JSON object'))
    return { document: emptyProjectDocument(), diagnostics }
  }

  // (4) Envelope: format magic.
  const format = root.format
  if (!isString(format) || format !== GPROJ_FORMAT) {
    diagnostics.push(
      diag(
        INVALID_GPROJ,
        'error',
        `Envelope "format" must be "${GPROJ_FORMAT}", got ${JSON.stringify(format)}`,
      ),
    )
    return { document: emptyProjectDocument(), diagnostics }
  }

  // (5) Envelope: format version.
  const formatVersion = root.formatVersion
  if (!isInteger(formatVersion) || !isSupportedGprojVersion(formatVersion)) {
    diagnostics.push(
      diag(
        UNSUPPORTED_GPROJ_VERSION,
        'error',
        `Unsupported .gproj format version: ${JSON.stringify(formatVersion)}`,
      ),
    )
    return { document: emptyProjectDocument(), diagnostics }
  }

  // (6) Envelope: document payload.
  const documentRaw = root.document
  if (!isObject(documentRaw)) {
    diagnostics.push(
      diag(SCHEMA_INVALID, 'error', 'Envelope "document" payload is missing or not an object'),
    )
    return { document: emptyProjectDocument(), diagnostics }
  }

  // (7) Schema-validate + construct the canonical ProjectDocument.
  const schemaVersion = documentRaw.schemaVersion
  if (!isInteger(schemaVersion) || schemaVersion !== 1) {
    diagnostics.push(
      diag(
        SCHEMA_INVALID,
        'error',
        `Unsupported document schemaVersion: ${JSON.stringify(schemaVersion)} (expected 1)`,
      ),
    )
    return { document: emptyProjectDocument(), diagnostics }
  }

  const properties = parseProperties(documentRaw.properties, diagnostics)
  if (properties === undefined) {
    return { document: emptyProjectDocument(), diagnostics }
  }

  const tasks = parseArray<Task>(documentRaw.tasks, 'tasks', parseTask, diagnostics)
  const resources = parseArray<Resource>(
    documentRaw.resources,
    'resources',
    parseResource,
    diagnostics,
  )
  const assignments = parseArray<Assignment>(
    documentRaw.assignments,
    'assignments',
    parseAssignment,
    diagnostics,
  )
  const dependencies = parseArray<Dependency>(
    documentRaw.dependencies,
    'dependencies',
    parseDependency,
    diagnostics,
  )
  const calendars = parseArray<Calendar>(
    documentRaw.calendars,
    'calendars',
    parseCalendar,
    diagnostics,
  )
  const baselines = parseArray<Baseline>(
    documentRaw.baselines,
    'baselines',
    parseBaseline,
    diagnostics,
  )
  const customFields = parseArray<CustomField>(
    documentRaw.customFields,
    'customFields',
    parseCustomField,
    diagnostics,
  )
  const views = parseArray<ProjectView>(documentRaw.views, 'views', parseView, diagnostics)
  const tables = parseArray<ProjectTable>(documentRaw.tables, 'tables', parseTable, diagnostics)
  const filters = parseArray<ProjectFilter>(
    documentRaw.filters,
    'filters',
    parseFilter,
    diagnostics,
  )
  const groups = parseArray<ProjectGroup>(documentRaw.groups, 'groups', parseGroup, diagnostics)

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
    views,
    tables,
    filters,
    groups,
  }

  // (8) Delegate semantic validation to the canonical engine validator.
  // The engine's diagnostics have shape { code: string; message: string }
  // (no severity, no entityId). We map them to ImportDiagnostic with
  // severity 'error' so the caller treats them uniformly. The engine's codes
  // (DUPLICATE_TASK_ID, MISSING_TASK_REFERENCE, CALENDAR_PERIOD_MALFORMED,
  // CALENDAR_CYCLE, MISSING_BASE_CALENDAR, …) are already valid
  // ImportDiagnostic.code strings — no duplication.
  const engineResult = validateProjectDocument(document)
  for (const d of engineResult.diagnostics) {
    diagnostics.push({ code: d.code, severity: 'error', message: d.message })
  }

  // (9) Canonicalize: emit an info diagnostic noting the file was read at the
  // canonical version. (Hosts may suppress info-level diagnostics.)
  diagnostics.push({
    code: 'GPROJ_READ',
    severity: 'info',
    message: `Read .gproj formatVersion ${formatVersion}${metadata?.sourceName ? ' from ' + metadata.sourceName : ''}`,
  })

  return { document, diagnostics }
}

/**
 * Inspect `.gproj` bytes (fast path): decode + parse the envelope only and
 * return the `ProjectFileMetadata`. Does NOT validate the document payload.
 * `sourceName` comes from the host-supplied `metadata` parameter (the file
 * does not store its own filename). If the bytes cannot be parsed at all, the
 * returned metadata falls back to the current format version so callers
 * always receive a typed `ProjectFileMetadata` (the failure is surfaced through
 * the `import` path, which returns error diagnostics).
 */
export function inspectGproj(
  input: Uint8Array,
  metadata?: ProjectFileMetadata,
): ProjectFileMetadata {
  let version = String(GPROJ_FORMAT_VERSION)
  try {
    const text = decodeUtf8(input)
    const root = JSON.parse(text)
    if (isObject(root) && isInteger(root.formatVersion)) {
      version = String(root.formatVersion)
    }
  } catch {
    // Unparseable — fall through with the default version.
  }
  return {
    format: GPROJ_FORMAT,
    version,
    sourceName: metadata?.sourceName,
  }
}

// Re-export so callers can reach the canonical envelope builder + empty doc.
export { buildGprojEnvelope, emptyProjectDocument }
// Re-export the canonical serializer for the byte-identity invariant.
export { serializeGproj } from './serialize.js'
