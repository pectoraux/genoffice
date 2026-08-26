/**
 * PROJECT-016 — MSPDI export golden fixtures (E01–E15).
 *
 * Canonical `ProjectDocument` builders exercising every required export
 * surface. E01–E14 are valid documents whose identities follow the accepted
 * PROJECT-015 deterministic import mapping (`'t'+uid`, `'r'+uid`, `'a'+uid`,
 * `'c'+uid`, `'b'+slot`, `'d-'+succ+'-'+pred+'-'+type`) and whose array
 * orders match the documented canonical export ordering — so
 * `exportMspdi → importMspdi` reproduces byte-identical canonical documents
 * (provable via `serializeGproj`). E15 is a family of documents carrying
 * unsupported/unrepresentable canonical state, each asserting explicit export
 * diagnostics instead of silent loss.
 *
 * Every builder is a pure function returning fresh objects (no module-level
 * counters, no shared mutable state) so repeated calls are byte-deterministic.
 */
import {
  asAssignmentId,
  asBaselineId,
  asCalendarId,
  asCustomFieldId,
  asDependencyId,
  asISODateTime,
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
  Resource,
  Task,
  TaskType,
} from '@genoffice/project-contracts'

// ---- shared constants ------------------------------------------------------

export const MONDAY = '2026-08-03T09:00:00.000Z'
export const TUESDAY = '2026-08-04T09:00:00.000Z'
export const WEDNESDAY = '2026-08-05T09:00:00.000Z'
export const THURSDAY = '2026-08-06T09:00:00.000Z'
export const FRIDAY = '2026-08-07T09:00:00.000Z'
export const CAPTURED_AT = '2026-08-01T12:00:00.000Z'

const standardDay = (): CalendarPeriod[] => [{ startMinute: 540, endMinute: 1020 }]

/** Full seven-key Mon–Fri 09:00–17:00 week (the import materialized shape). */
export const standardWeek = (): Record<number, CalendarPeriod[]> => ({
  0: [],
  1: standardDay(),
  2: standardDay(),
  3: standardDay(),
  4: standardDay(),
  5: standardDay(),
  6: [],
})

// ---- entity builders ---------------------------------------------------------

export interface TaskOverrides {
  uid?: number
  wbs?: string
  outlineLevel?: number
  name?: string
  taskType?: TaskType
  summary?: boolean
  milestone?: boolean
  manualScheduled?: boolean
  start?: string
  finish?: string
  duration?: number
  constraintType?: Task['constraintType']
  constraintDate?: string
  deadline?: string
  priority?: number
  calendarId?: string
  percentComplete?: number
  physicalPercentComplete?: number
  work?: number
  remainingWork?: number
  actualWork?: number
  cost?: number
  actualCost?: number
  remainingCost?: number
  baseline?: string[]
  customFields?: Record<string, string | number | boolean | null>
  notes?: string[]
  parentTaskId?: string
}

/** Build a canonical task with the import-convention id `'t'+uid`. */
export function makeTask(uid: number, overrides: TaskOverrides = {}): Task {
  const id = asTaskId(`t${uid}`)
  return {
    id,
    uid,
    wbs: overrides.wbs ?? '',
    outlineLevel: overrides.outlineLevel ?? 1,
    name: overrides.name ?? `Task ${uid}`,
    taskType: overrides.taskType ?? 'fixedDuration',
    summary: overrides.summary ?? false,
    milestone: overrides.milestone ?? false,
    manualScheduled: overrides.manualScheduled ?? false,
    autoScheduled: !(overrides.manualScheduled ?? false),
    duration: asWorkingMinutes(overrides.duration ?? 480),
    priority: overrides.priority ?? 500,
    percentComplete: overrides.percentComplete ?? 0,
    work: asWorkingMinutes(overrides.work ?? 0),
    remainingWork: asWorkingMinutes(overrides.remainingWork ?? 0),
    actualWork: asWorkingMinutes(overrides.actualWork ?? 0),
    cost: overrides.cost ?? 0,
    actualCost: overrides.actualCost ?? 0,
    remainingCost: overrides.remainingCost ?? 0,
    baseline: (overrides.baseline ?? []).map(asBaselineId),
    customFields: (overrides.customFields ?? {}) as Task['customFields'],
    notes: overrides.notes ?? [],
    ...(overrides.start !== undefined ? { start: asISODateTime(overrides.start) } : {}),
    ...(overrides.finish !== undefined ? { finish: asISODateTime(overrides.finish) } : {}),
    ...(overrides.constraintType !== undefined ? { constraintType: overrides.constraintType } : {}),
    ...(overrides.constraintDate !== undefined
      ? { constraintDate: asISODateTime(overrides.constraintDate) }
      : {}),
    ...(overrides.deadline !== undefined ? { deadline: asISODateTime(overrides.deadline) } : {}),
    ...(overrides.calendarId !== undefined
      ? { calendarId: asCalendarId(overrides.calendarId) }
      : {}),
    ...(overrides.physicalPercentComplete !== undefined
      ? { physicalPercentComplete: overrides.physicalPercentComplete }
      : {}),
    ...(overrides.parentTaskId !== undefined
      ? { parentTaskId: asTaskId(overrides.parentTaskId) }
      : {}),
  }
}

export interface ResourceOverrides {
  uid?: number
  name?: string
  kind?: Resource['kind']
  maxUnits?: number
  standardRate?: number
  overtimeRate?: number
  costPerUse?: number
  calendarId?: string
  availability?: Array<{ start: string; finish?: string; units: number }>
}

/** Build a canonical resource with the import-convention id `'r'+uid`. */
export function makeResource(uid: number, overrides: ResourceOverrides = {}): Resource {
  const kind = overrides.kind ?? 'work'
  return {
    id: asResourceId(`r${uid}`),
    uid,
    name: overrides.name ?? `Resource ${uid}`,
    kind,
    maxUnits: overrides.maxUnits ?? (kind === 'work' ? 1 : 0),
    standardRate: overrides.standardRate ?? 0,
    overtimeRate: overrides.overtimeRate ?? 0,
    costPerUse: overrides.costPerUse ?? 0,
    availability: (overrides.availability ?? []).map((p) => ({
      start: asISODateTime(p.start),
      ...(p.finish !== undefined ? { finish: asISODateTime(p.finish) } : {}),
      units: p.units,
    })),
    ...(overrides.calendarId !== undefined
      ? { calendarId: asCalendarId(overrides.calendarId) }
      : {}),
  }
}

export interface AssignmentOverrides {
  units?: number
  work?: number
  actualWork?: number
  remainingWork?: number
  cost?: number
  actualCost?: number
  remainingCost?: number
}

/** Build a canonical assignment with the import-convention id `'a'+uid`. */
export function makeAssignment(
  uid: number,
  taskId: number,
  resourceId: number,
  overrides: AssignmentOverrides = {},
): Assignment {
  return {
    id: asAssignmentId(`a${uid}`),
    taskId: asTaskId(`t${taskId}`),
    resourceId: asResourceId(`r${resourceId}`),
    units: overrides.units ?? 1,
    work: asWorkingMinutes(overrides.work ?? 0),
    actualWork: asWorkingMinutes(overrides.actualWork ?? 0),
    remainingWork: asWorkingMinutes(overrides.remainingWork ?? 0),
    cost: overrides.cost ?? 0,
    actualCost: overrides.actualCost ?? 0,
    remainingCost: overrides.remainingCost ?? 0,
  }
}

/** Deterministic dependency id (the accepted import mapping). */
export function depId(successorUid: number, predecessorUid: number, type: DependencyType): string {
  return `d-t${successorUid}-t${predecessorUid}-${type}`
}

export function makeDependency(
  successorUid: number,
  predecessorUid: number,
  type: DependencyType,
  lagMinutes = 0,
): Dependency {
  return {
    id: asDependencyId(depId(successorUid, predecessorUid, type)),
    predecessorId: asTaskId(`t${predecessorUid}`),
    successorId: asTaskId(`t${successorUid}`),
    type,
    lagMinutes: asWorkingMinutes(lagMinutes),
  }
}

export interface CalendarOverrides {
  name?: string
  baseCalendarId?: string
  workingWeek?: Record<number, CalendarPeriod[]>
  exceptions?: Array<{ date: string; periods: CalendarPeriod[] }>
}

/** Build a canonical calendar with the import-convention id `'c'+uid`. */
export function makeCalendar(uid: number, overrides: CalendarOverrides = {}): Calendar {
  return {
    id: asCalendarId(`c${uid}`),
    name: overrides.name ?? `Calendar ${uid}`,
    workingWeek: overrides.workingWeek ?? standardWeek(),
    exceptions: (overrides.exceptions ?? []).map((e) => ({ date: e.date, periods: e.periods })),
    ...(overrides.baseCalendarId !== undefined
      ? { baseCalendarId: asCalendarId(overrides.baseCalendarId) }
      : {}),
  }
}

export function makeBaseline(
  slot: number,
  snapshots: Record<
    number,
    { start?: string; finish?: string; duration: number; work: number; cost: number }
  >,
  capturedAt = CAPTURED_AT,
): Baseline {
  const taskSnapshots: Baseline['taskSnapshots'] = {}
  for (const taskUid of Object.keys(snapshots)
    .map(Number)
    .sort((a, b) => a - b)) {
    const s = snapshots[taskUid]
    taskSnapshots[`t${taskUid}`] = {
      duration: asWorkingMinutes(s.duration),
      work: asWorkingMinutes(s.work),
      cost: s.cost,
      ...(s.start !== undefined ? { start: asISODateTime(s.start) } : {}),
      ...(s.finish !== undefined ? { finish: asISODateTime(s.finish) } : {}),
    }
  }
  return {
    id: asBaselineId(`b${slot}`),
    name: slot === 0 ? 'Baseline' : `Baseline ${slot}`,
    capturedAt: asISODateTime(capturedAt),
    taskSnapshots,
  }
}

export function makeCustomField(id: string, type: CustomField['type'], name = id): CustomField {
  return { id: asCustomFieldId(id), name, type }
}

// ---- document assembly ---------------------------------------------------------

export interface DocumentOptions {
  id?: string
  name?: string
  startDate?: string
  finishDate?: string
  statusDate?: string
  tasks: Task[]
  calendars: Calendar[]
  resources?: Resource[]
  assignments?: Assignment[]
  dependencies?: Dependency[]
  baselines?: Baseline[]
  customFields?: CustomField[]
}

export function makeDocument(options: DocumentOptions): ProjectDocument {
  return {
    schemaVersion: 1,
    properties: {
      id: options.id ?? 'p1',
      name: options.name ?? 'Export Fixture',
      startDate: asISODateTime(options.startDate ?? MONDAY),
      ...(options.finishDate !== undefined
        ? { finishDate: asISODateTime(options.finishDate) }
        : {}),
      ...(options.statusDate !== undefined
        ? { statusDate: asISODateTime(options.statusDate) }
        : {}),
      defaultCalendarId: asCalendarId('c1'),
    },
    tasks: options.tasks,
    calendars: options.calendars,
    resources: options.resources ?? [],
    assignments: options.assignments ?? [],
    dependencies: options.dependencies ?? [],
    baselines: options.baselines ?? [],
    customFields: options.customFields ?? [],
    views: [],
    tables: [],
    filters: [],
    groups: [],
  }
}

// ---- E01–E14 valid goldens -------------------------------------------------------

/** E01 — minimal project: one task, one calendar. */
export function e01Minimal(): ProjectDocument {
  return makeDocument({
    name: 'Minimal',
    tasks: [makeTask(1, { name: 'Only Task', wbs: '1' })],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E02 — WBS hierarchy: one summary with two children + a second top level. */
export function e02Hierarchy(): ProjectDocument {
  return makeDocument({
    name: 'Hierarchy',
    tasks: [
      makeTask(1, {
        name: 'Phase A',
        wbs: '1',
        outlineLevel: 1,
        summary: true,
        duration: 960,
      }),
      makeTask(2, { name: 'Work A1', wbs: '1.1', outlineLevel: 2, parentTaskId: 't1' }),
      makeTask(3, { name: 'Work A2', wbs: '1.2', outlineLevel: 2, parentTaskId: 't1' }),
      makeTask(4, { name: 'Phase B', wbs: '2', outlineLevel: 1 }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E03 — dependencies: one of each FS/SS/FF/SF type. */
export function e03Dependencies(): ProjectDocument {
  return makeDocument({
    name: 'Dependencies',
    tasks: [
      makeTask(1, { name: 'Start', wbs: '1' }),
      makeTask(2, { name: 'FS Successor', wbs: '2' }),
      makeTask(3, { name: 'SS Successor', wbs: '3' }),
      makeTask(4, { name: 'FF+SF Successor', wbs: '4' }),
    ],
    dependencies: [
      makeDependency(2, 1, 'FS'),
      makeDependency(3, 1, 'SS'),
      makeDependency(4, 2, 'FF'),
      makeDependency(4, 3, 'SF'),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E04 — lag/lead: positive lag, negative lead, zero lag. */
export function e04LagLead(): ProjectDocument {
  return makeDocument({
    name: 'Lag Lead',
    tasks: [
      makeTask(1, { name: 'Anchor', wbs: '1' }),
      makeTask(2, { name: 'Lagged', wbs: '2' }),
      makeTask(3, { name: 'Led', wbs: '3' }),
      makeTask(4, { name: 'Zero Lag', wbs: '4' }),
    ],
    dependencies: [
      makeDependency(2, 1, 'FS', 150),
      makeDependency(3, 1, 'FS', -240),
      makeDependency(4, 1, 'FS', 0),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E05 — constraints: all eight canonical constraint types. */
export function e05Constraints(): ProjectDocument {
  return makeDocument({
    name: 'Constraints',
    tasks: [
      makeTask(1, { name: 'ASAP', wbs: '1', constraintType: 'asSoonAsPossible' }),
      makeTask(2, {
        name: 'ALAP',
        wbs: '2',
        constraintType: 'asLateAsPossible',
      }),
      makeTask(3, {
        name: 'SNET',
        wbs: '3',
        constraintType: 'startNoEarlierThan',
        constraintDate: TUESDAY,
      }),
      makeTask(4, {
        name: 'SNLT',
        wbs: '4',
        constraintType: 'startNoLaterThan',
        constraintDate: WEDNESDAY,
      }),
      makeTask(5, {
        name: 'MSO',
        wbs: '5',
        constraintType: 'mustStartOn',
        constraintDate: THURSDAY,
      }),
      makeTask(6, {
        name: 'FNET',
        wbs: '6',
        constraintType: 'finishNoEarlierThan',
        constraintDate: FRIDAY,
      }),
      makeTask(7, {
        name: 'FNLT',
        wbs: '7',
        constraintType: 'finishNoLaterThan',
        constraintDate: MONDAY,
      }),
      makeTask(8, {
        name: 'MFO',
        wbs: '8',
        constraintType: 'mustFinishOn',
        constraintDate: TUESDAY,
      }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E06 — deadlines and progress (includes physicalPercentComplete, which the
 * accepted importer does not read back — the export warns). */
export function e06DeadlinesProgress(): ProjectDocument {
  return makeDocument({
    name: 'Deadlines Progress',
    statusDate: TUESDAY,
    tasks: [
      makeTask(1, {
        name: 'Tracked',
        wbs: '1',
        deadline: FRIDAY,
        percentComplete: 50,
        physicalPercentComplete: 25,
        work: 480,
        actualWork: 240,
        remainingWork: 240,
        cost: 1000,
        actualCost: 500,
        remainingCost: 500,
      }),
      makeTask(2, { name: 'Milestone', wbs: '2', milestone: true, duration: 0, deadline: FRIDAY }),
      makeTask(3, {
        name: 'Complete',
        wbs: '3',
        percentComplete: 100,
        work: 480,
        actualWork: 480,
      }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E07 — calendars: base (default) + derived override + working/non-working
 * exceptions. */
export function e07Calendars(): ProjectDocument {
  const fourToTen = (): CalendarPeriod[] => [{ startMinute: 960, endMinute: 1320 }]
  const lateWeek = (): Record<number, CalendarPeriod[]> => ({
    0: [],
    1: fourToTen(),
    2: fourToTen(),
    3: fourToTen(),
    4: fourToTen(),
    5: fourToTen(),
    6: [],
  })
  return makeDocument({
    name: 'Calendars',
    tasks: [
      makeTask(1, { name: 'Standard Task', wbs: '1' }),
      makeTask(2, { name: 'Late Task', wbs: '2', calendarId: 'c2' }),
    ],
    calendars: [
      makeCalendar(1, {
        name: 'Standard',
        exceptions: [{ date: '2026-12-25', periods: [] }],
      }),
      makeCalendar(2, {
        name: 'Late Shift',
        baseCalendarId: 'c1',
        workingWeek: lateWeek(),
        exceptions: [
          { date: '2026-08-08', periods: [{ startMinute: 600, endMinute: 840 }] },
          { date: '2026-12-25', periods: [] },
        ],
      }),
    ],
    dependencies: [],
  })
}

/** E08 — resources: work + material + cost kinds, availability windows. */
export function e08Resources(): ProjectDocument {
  return makeDocument({
    name: 'Resources',
    tasks: [makeTask(1, { name: 'Work Item', wbs: '1' })],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    resources: [
      makeResource(1, {
        name: 'Engineer',
        kind: 'work',
        maxUnits: 1,
        standardRate: 50,
        overtimeRate: 75,
        calendarId: 'c1',
        availability: [{ start: MONDAY, finish: '2026-12-31T17:00:00.000Z', units: 1 }],
      }),
      makeResource(2, {
        name: 'Concrete',
        kind: 'material',
        maxUnits: 0,
        standardRate: 120,
        costPerUse: 10,
      }),
      makeResource(3, { name: 'Legal Fee', kind: 'cost', costPerUse: 500 }),
    ],
  })
}

/** E09 — assignments across the three resource kinds. */
export function e09Assignments(): ProjectDocument {
  return makeDocument({
    name: 'Assignments',
    tasks: [
      makeTask(1, { name: 'Build', wbs: '1' }),
      makeTask(2, { name: 'Pour', wbs: '2' }),
      makeTask(3, { name: 'Sign Off', wbs: '3' }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    resources: [
      makeResource(1, { name: 'Engineer', kind: 'work', standardRate: 50 }),
      makeResource(2, { name: 'Concrete', kind: 'material', standardRate: 120 }),
      makeResource(3, { name: 'Legal Fee', kind: 'cost', costPerUse: 500 }),
    ],
    assignments: [
      makeAssignment(1, 1, 1, { units: 1, work: 480, remainingWork: 480 }),
      makeAssignment(2, 2, 2, { units: 10, cost: 1210 }),
      makeAssignment(3, 3, 3, { cost: 500 }),
    ],
  })
}

/** E10 — a single baseline snapshot set (uniform capturedAt). */
export function e10Baseline(): ProjectDocument {
  return makeDocument({
    name: 'Baseline',
    tasks: [
      makeTask(1, { name: 'Planned', wbs: '1' }),
      makeTask(2, { name: 'Also Planned', wbs: '2' }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    baselines: [
      makeBaseline(0, {
        1: { start: MONDAY, finish: TUESDAY, duration: 480, work: 480, cost: 100 },
        2: { duration: 960, work: 0, cost: 0 },
      }),
    ],
  })
}

/** E11 — multiple baselines (slots 0 and 1, uniform capturedAt). */
export function e11MultipleBaselines(): ProjectDocument {
  return makeDocument({
    name: 'Multiple Baselines',
    tasks: [makeTask(1, { name: 'Planned', wbs: '1' })],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    baselines: [
      makeBaseline(0, {
        1: { start: MONDAY, finish: TUESDAY, duration: 480, work: 480, cost: 100 },
      }),
      makeBaseline(1, {
        1: { start: TUESDAY, finish: WEDNESDAY, duration: 480, work: 480, cost: 120 },
      }),
    ],
  })
}

/** E12 — custom fields: text/number/boolean/date definitions + values
 * (including null and undefined-free records). */
export function e12CustomFields(): ProjectDocument {
  return makeDocument({
    name: 'Custom Fields',
    tasks: [
      makeTask(1, {
        name: 'Rich',
        wbs: '1',
        customFields: {
          cf1: 'hello world',
          cf2: 42,
          cf3: true,
          cf4: '2026-08-03T00:00:00.000Z',
          cf5: null,
        },
      }),
      makeTask(2, {
        name: 'Sparse',
        wbs: '2',
        customFields: { cf1: 'second', cf2: -7.5, cf3: false },
      }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    customFields: [
      makeCustomField('cf1', 'text', 'Comment'),
      makeCustomField('cf2', 'number', 'Score'),
      makeCustomField('cf3', 'boolean', 'Flag'),
      makeCustomField('cf4', 'date', 'Reviewed On'),
      makeCustomField('cf5', 'text', 'Optional'),
    ],
  })
}

/** E13 — comprehensive: hierarchy + dependencies + lags + constraints +
 * deadlines + progress + calendars + resources + assignments + baselines +
 * custom fields + notes. */
export function e13Comprehensive(): ProjectDocument {
  return makeDocument({
    id: 'comprehensive-1',
    name: 'Comprehensive & Co <Alpha>',
    finishDate: FRIDAY,
    statusDate: TUESDAY,
    tasks: [
      makeTask(1, {
        name: 'Phase A',
        wbs: '1',
        outlineLevel: 1,
        summary: true,
        duration: 1920,
        deadline: FRIDAY,
      }),
      makeTask(2, {
        name: 'Design',
        wbs: '1.1',
        outlineLevel: 2,
        parentTaskId: 't1',
        constraintType: 'startNoEarlierThan',
        constraintDate: MONDAY,
        percentComplete: 100,
        work: 480,
        actualWork: 480,
        notes: ['Design note'],
      }),
      makeTask(3, {
        name: 'Build',
        wbs: '1.2',
        outlineLevel: 2,
        parentTaskId: 't1',
        constraintType: 'finishNoLaterThan',
        constraintDate: FRIDAY,
        calendarId: 'c2',
        percentComplete: 50,
        work: 960,
        actualWork: 480,
        remainingWork: 480,
        cost: 2000,
        actualCost: 1000,
        remainingCost: 1000,
        customFields: { cf1: 'in progress', cf2: 7, cf3: true },
      }),
      makeTask(4, {
        name: 'Milestone: Ready',
        wbs: '2',
        outlineLevel: 1,
        milestone: true,
        duration: 0,
        deadline: FRIDAY,
      }),
    ],
    calendars: [
      makeCalendar(1, {
        name: 'Standard',
        exceptions: [{ date: '2026-12-25', periods: [] }],
      }),
      makeCalendar(2, {
        name: 'Late Shift',
        baseCalendarId: 'c1',
        workingWeek: {
          0: [],
          1: [{ startMinute: 960, endMinute: 1320 }],
          2: [{ startMinute: 960, endMinute: 1320 }],
          3: [{ startMinute: 960, endMinute: 1320 }],
          4: [{ startMinute: 960, endMinute: 1320 }],
          5: [{ startMinute: 960, endMinute: 1320 }],
          6: [],
        },
      }),
    ],
    resources: [
      makeResource(1, { name: 'Engineer', kind: 'work', standardRate: 50, overtimeRate: 75 }),
      makeResource(2, { name: 'Concrete', kind: 'material', standardRate: 120 }),
      makeResource(3, { name: 'Legal Fee', kind: 'cost', costPerUse: 500 }),
    ],
    assignments: [
      makeAssignment(1, 2, 1, { units: 1, work: 480, remainingWork: 480 }),
      makeAssignment(2, 3, 1, { units: 2, work: 960, remainingWork: 960 }),
      makeAssignment(3, 3, 3, { cost: 500 }),
    ],
    dependencies: [makeDependency(3, 2, 'FS', 150), makeDependency(4, 3, 'FS', -240)],
    baselines: [
      makeBaseline(0, {
        2: { start: MONDAY, finish: TUESDAY, duration: 480, work: 480, cost: 100 },
        3: { duration: 960, work: 960, cost: 2000 },
      }),
      makeBaseline(1, {
        2: { start: TUESDAY, finish: WEDNESDAY, duration: 480, work: 480, cost: 110 },
      }),
    ],
    customFields: [
      makeCustomField('cf1', 'text', 'Comment'),
      makeCustomField('cf2', 'number', 'Score'),
      makeCustomField('cf3', 'boolean', 'Flag'),
    ],
  })
}

/** E14 — large project: 60 tasks, 59 dependencies, 10 resources,
 * 60 assignments (deterministic generation, no counters). The task array is
 * in hierarchical-DFS order (parents before children, sibling order = array
 * order) — the order the exporter preserves and the importer reproduces — so
 * the fixture round-trips byte-exactly. */
export function e14Large(): ProjectDocument {
  // DFS task order: roots 1–2 (leaves), then each summary (3–6) immediately
  // followed by its whole subtree.
  const order: number[] = [1, 2, 3, 7, 8, 9, 4]
  for (let uid = 10; uid <= 12; uid += 1) order.push(uid)
  for (let uid = 19; uid <= 24; uid += 1) order.push(uid)
  for (let uid = 37; uid <= 48; uid += 1) order.push(uid)
  order.push(5)
  for (let uid = 13; uid <= 15; uid += 1) order.push(uid)
  for (let uid = 25; uid <= 30; uid += 1) order.push(uid)
  for (let uid = 49; uid <= 60; uid += 1) order.push(uid)
  order.push(6)
  for (let uid = 16; uid <= 18; uid += 1) order.push(uid)
  for (let uid = 31; uid <= 36; uid += 1) order.push(uid)
  const parentOf = (uid: number): number | undefined => {
    if (uid >= 7 && uid <= 9) return 3
    if (uid >= 10 && uid <= 12) return 4
    if (uid >= 19 && uid <= 24) return 4
    if (uid >= 37 && uid <= 48) return 4
    if (uid >= 13 && uid <= 15) return 5
    if (uid >= 25 && uid <= 30) return 5
    if (uid >= 49 && uid <= 60) return 5
    if (uid >= 16 && uid <= 18) return 6
    if (uid >= 31 && uid <= 36) return 6
    return undefined
  }
  const tasks: Task[] = order.map((uid) => {
    const parent = parentOf(uid)
    const isSummary = uid >= 3 && uid <= 6
    return makeTask(uid, {
      name: `Task ${uid}`,
      wbs: '',
      outlineLevel: parent === undefined ? 1 : 2,
      summary: isSummary,
      duration: isSummary ? 4800 : 480,
      ...(parent !== undefined ? { parentTaskId: `t${parent}` } : {}),
    })
  })
  // Outline numbers derived exactly as the exporter derives them (DFS with
  // sibling order = array order), so the fixture wbs round-trips byte-exactly.
  const childrenOf = new Map<string, Task[]>()
  for (const task of tasks) {
    if (task.parentTaskId === undefined) continue
    const list = childrenOf.get(task.parentTaskId as string)
    if (list === undefined) childrenOf.set(task.parentTaskId as string, [task])
    else list.push(task)
  }
  const wbsOf = new Map<string, string>()
  const assign = (task: Task, number: string): void => {
    wbsOf.set(task.id as string, number)
    let childIndex = 1
    for (const child of childrenOf.get(task.id as string) ?? []) {
      assign(child, `${number}.${childIndex}`)
      childIndex += 1
    }
  }
  let rootIndex = 1
  for (const task of tasks) {
    if (task.parentTaskId !== undefined) continue
    assign(task, String(rootIndex))
    rootIndex += 1
  }
  for (const task of tasks) {
    ;(task as { wbs: string }).wbs = wbsOf.get(task.id as string) ?? String(task.uid)
  }
  // 59 dependencies: chain 8..60 plus cross links (deterministic, acyclic),
  // ordered exactly as the importer reconstructs them (successor in DFS task
  // order, then predecessor uid, then link type).
  const dfsPos = new Map<number, number>()
  order.forEach((uid, index) => dfsPos.set(uid, index))
  const dependencies: Dependency[] = []
  for (let uid = 8; uid <= 60; uid += 1) {
    dependencies.push(makeDependency(uid, uid - 1, 'FS', (uid % 5) * 30))
  }
  for (let uid = 10; uid <= 30; uid += 10) {
    dependencies.push(makeDependency(uid, 1, 'SS', 0))
    dependencies.push(makeDependency(uid + 1, 2, 'FF', -60))
  }
  const TYPE_ORDER: Record<DependencyType, number> = { FS: 0, FF: 1, SS: 2, SF: 3 }
  dependencies.sort(
    (a, b) =>
      (dfsPos.get(Number((a.successorId as string).slice(1))) ?? 0) -
        (dfsPos.get(Number((b.successorId as string).slice(1))) ?? 0) ||
      Number((a.predecessorId as string).slice(1)) - Number((b.predecessorId as string).slice(1)) ||
      TYPE_ORDER[a.type] - TYPE_ORDER[b.type],
  )
  const resources: Resource[] = []
  for (let uid = 1; uid <= 10; uid += 1) {
    resources.push(
      makeResource(uid, {
        name: `Resource ${uid}`,
        kind: uid % 3 === 0 ? 'material' : uid % 3 === 1 ? 'work' : 'cost',
        standardRate: uid * 10,
      }),
    )
  }
  const assignments: Assignment[] = []
  for (let uid = 1; uid <= 60; uid += 1) {
    const resourceId = ((uid - 1) % 10) + 1
    assignments.push(
      makeAssignment(uid, uid, resourceId, { units: 1, work: 480, remainingWork: 480 }),
    )
  }
  return makeDocument({
    name: 'Large Project',
    tasks,
    calendars: [makeCalendar(1, { name: 'Standard' })],
    resources,
    assignments,
    dependencies,
  })
}

// ---- E15 — unsupported / unrepresentable canonical state ---------------------------

/** E15a — derived calendar with a PARTIAL workingWeek (weekday inheritance
 * materialization; resolved semantics exactly recoverable). */
export function e15PartialInheritance(): ProjectDocument {
  return makeDocument({
    name: 'Partial Inheritance',
    tasks: [makeTask(1, { name: 'Task', wbs: '1', calendarId: 'c2' })],
    calendars: [
      makeCalendar(1, { name: 'Standard' }),
      makeCalendar(2, {
        name: 'Tuesdays Only Override',
        baseCalendarId: 'c1',
        workingWeek: { 2: [{ startMinute: 600, endMinute: 840 }] },
      }),
    ],
  })
}

/** E15b — working period ending at 24:00 (1440) — unrepresentable through the
 * accepted importer's HH:MM:SS rule. */
export function e15MidnightEnd(): ProjectDocument {
  return makeDocument({
    name: 'Midnight End',
    tasks: [makeTask(1, { name: 'Night Owl', wbs: '1' })],
    calendars: [
      makeCalendar(1, {
        name: 'Standard',
        workingWeek: {
          0: [],
          1: [{ startMinute: 1020, endMinute: 1440 }],
          2: [],
          3: [],
          4: [],
          5: [],
          6: [],
        },
      }),
    ],
  })
}

/** E15c — baselines with divergent capturedAt values (single <LastSaved>
 * carrier). */
export function e15DivergentCapturedAt(): ProjectDocument {
  return makeDocument({
    name: 'Divergent Captured At',
    tasks: [makeTask(1, { name: 'Planned', wbs: '1' })],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    baselines: [
      makeBaseline(0, { 1: { duration: 480, work: 0, cost: 0 } }, MONDAY),
      makeBaseline(1, { 1: { duration: 960, work: 0, cost: 0 } }, FRIDAY),
    ],
  })
}

/** E15d — non-convention calendar identity ('standard') — deterministic uid
 * synthesis + consistent reference remap. */
export function e15NativeCalendarId(): ProjectDocument {
  const doc = e01Minimal()
  const standard = asCalendarId('standard')
  return {
    ...doc,
    properties: { ...doc.properties, defaultCalendarId: standard },
    calendars: [{ ...doc.calendars[0], id: standard }],
    tasks: [makeTask(1, { name: 'Only Task', wbs: '1', calendarId: 'standard' })],
  }
}

/** E15e — multiple notes (collapsed into the single MSPDI <Notes> field). */
export function e15MultipleNotes(): ProjectDocument {
  return makeDocument({
    name: 'Multiple Notes',
    tasks: [makeTask(1, { name: 'Chatty', wbs: '1', notes: ['first note', 'second note'] })],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E15f — string custom-field values that re-parse as number/boolean. */
export function e15NumericStringCustomField(): ProjectDocument {
  return makeDocument({
    name: 'Numeric Strings',
    tasks: [
      makeTask(1, {
        name: 'Typed',
        wbs: '1',
        customFields: { cf1: '123', cf2: 'true' },
      }),
    ],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    customFields: [
      makeCustomField('cf1', 'text', 'Comment'),
      makeCustomField('cf2', 'text', 'Flag Text'),
    ],
  })
}

/** E15g — lag magnitude beyond the safe-integer LinkLag range. */
export function e15HugeLag(): ProjectDocument {
  return makeDocument({
    name: 'Huge Lag',
    tasks: [makeTask(1, { name: 'A', wbs: '1' }), makeTask(2, { name: 'B', wbs: '2' })],
    dependencies: [makeDependency(2, 1, 'FS', 1_000_000_000_000_000)],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** E15h — task/resource uids that are not non-negative integers. */
export function e15InvalidUid(): ProjectDocument {
  const doc = makeDocument({
    name: 'Invalid Uid',
    tasks: [makeTask(1, { name: 'Good', wbs: '1' })],
    calendars: [makeCalendar(1, { name: 'Standard' })],
    resources: [makeResource(2, { name: 'Good Resource' })],
  })
  const badTask = makeTask(-5, { name: 'Negative Uid', wbs: '2' })
  const badResource = makeResource(2.5, { name: 'Fractional Uid' })
  return {
    ...doc,
    tasks: [...doc.tasks, badTask],
    resources: [...doc.resources, badResource],
  }
}

// ---- E15z — malformed canonical documents (export REFUSAL) ------------------------

/** Duplicate task id → engine-invalid → export refused. */
export function e15zDuplicateTaskId(): ProjectDocument {
  const doc = e01Minimal()
  return { ...doc, tasks: [...doc.tasks, { ...doc.tasks[0], name: 'Duplicate' }] }
}

/** Dependency cycle → engine-invalid → export refused. */
export function e15zDependencyCycle(): ProjectDocument {
  return makeDocument({
    name: 'Cycle',
    tasks: [makeTask(1, { name: 'A', wbs: '1' }), makeTask(2, { name: 'B', wbs: '2' })],
    dependencies: [makeDependency(2, 1, 'FS'), makeDependency(1, 2, 'FS')],
    calendars: [makeCalendar(1, { name: 'Standard' })],
  })
}

/** Missing referenced calendar → engine-invalid → export refused. */
export function e15zMissingCalendar(): ProjectDocument {
  const doc = e01Minimal()
  return {
    ...doc,
    tasks: [makeTask(1, { name: 'Only Task', wbs: '1', calendarId: 'c99' })],
  }
}

/** Negative duration → engine-invalid → export refused. */
export function e15zNegativeDuration(): ProjectDocument {
  const doc = e01Minimal()
  return {
    ...doc,
    tasks: [makeTask(1, { name: 'Broken', wbs: '1', duration: -480 })],
  }
}

/** Date-bounded constraint without a constraintDate → engine-invalid →
 * export refused. */
export function e15zMissingConstraintDate(): ProjectDocument {
  const doc = e01Minimal()
  return {
    ...doc,
    tasks: [makeTask(1, { name: 'Undated', wbs: '1', constraintType: 'mustStartOn' })],
  }
}
