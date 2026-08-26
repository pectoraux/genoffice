/**
 * PROJECT-014 — `.gproj` test fixtures.
 *
 * Local, standalone builders for the 15 golden fixtures (G01–G12 valid,
 * G13–G15 malformed). These mirror the scheduling/engine test-fixture
 * conventions (standard Mon–Fri 09:00–17:00 week) but live in this package's
 * own tests so the file-adapter tests do not import another package's test
 * internals (which are not part of any package's public `exports`).
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
  ProjectFilter,
  ProjectGroup,
  ProjectTable,
  ProjectView,
  Resource,
  Task,
  TaskType,
} from '@genoffice/project-contracts'
import { encodeUtf8 } from '../src/utf8.js'

// ---- shared constants ----------------------------------------------------

export const MONDAY = '2026-08-03T09:00:00.000Z'
export const TUESDAY = '2026-08-04T09:00:00.000Z'
export const WEDNESDAY = '2026-08-05T09:00:00.000Z'
export const THURSDAY = '2026-08-06T09:00:00.000Z'
export const FRIDAY = '2026-08-07T09:00:00.000Z'

const standardDay = (): CalendarPeriod[] => [{ startMinute: 540, endMinute: 1020 }]

export const standardWeek = (): Record<number, CalendarPeriod[]> => ({
  0: [],
  1: standardDay(),
  2: standardDay(),
  3: standardDay(),
  4: standardDay(),
  5: standardDay(),
  6: [],
})

// ---- entity builders -----------------------------------------------------

export function makeCalendar(id: string, overrides: Partial<Calendar> = {}): Calendar {
  const { id: _ignored, ...rest } = overrides
  void _ignored
  return {
    name: id,
    workingWeek: standardWeek(),
    exceptions: [],
    ...rest,
    id: asCalendarId(id),
  }
}

let taskUid = 0
let resourceUid = 0

/**
 * Reset the auto-incrementing uid counters. Called at the start of every
 * golden builder so that each `build()` is byte-deterministic (calling
 * `build()` twice produces identical bytes). This is required by the
 * PROJECT-014 byte-identity invariant.
 */
export function resetFixtureCounters(): void {
  taskUid = 0
  resourceUid = 0
}

export function makeTask(overrides: Omit<Partial<Task>, 'id'> & { id: string }): Task {
  const { id, ...rest } = overrides
  taskUid += 1
  return {
    uid: taskUid,
    wbs: '',
    outlineLevel: 1,
    name: id,
    taskType: 'fixedDuration' as TaskType,
    summary: false,
    milestone: false,
    manualScheduled: false,
    autoScheduled: true,
    duration: asWorkingMinutes(480),
    priority: 500,
    percentComplete: 0,
    work: asWorkingMinutes(0),
    remainingWork: asWorkingMinutes(0),
    actualWork: asWorkingMinutes(0),
    cost: 0,
    actualCost: 0,
    remainingCost: 0,
    baseline: [],
    customFields: {},
    notes: [],
    ...rest,
    id: asTaskId(id),
  }
}

export function makeResource(overrides: Omit<Partial<Resource>, 'id'> & { id: string }): Resource {
  const { id, ...rest } = overrides
  resourceUid += 1
  return {
    uid: resourceUid,
    name: id,
    kind: 'work',
    maxUnits: 1,
    standardRate: 0,
    overtimeRate: 0,
    costPerUse: 0,
    availability: [],
    ...rest,
    id: asResourceId(id),
  }
}

export function makeAssignment(
  id: string,
  taskId: string,
  resourceId: string,
  overrides: Partial<Omit<Assignment, 'id' | 'taskId' | 'resourceId'>> = {},
): Assignment {
  return {
    units: 1,
    work: asWorkingMinutes(0),
    actualWork: asWorkingMinutes(0),
    remainingWork: asWorkingMinutes(0),
    cost: 0,
    actualCost: 0,
    remainingCost: 0,
    ...overrides,
    id: asAssignmentId(id),
    taskId: asTaskId(taskId),
    resourceId: asResourceId(resourceId),
  }
}

export function makeDependency(
  id: string,
  predecessorId: string,
  successorId: string,
  type: DependencyType,
  lagMinutes = 0,
): Dependency {
  return {
    id: asDependencyId(id),
    predecessorId: asTaskId(predecessorId),
    successorId: asTaskId(successorId),
    type,
    lagMinutes,
  }
}

export function makeBaseline(
  id: string,
  capturedAt: string,
  snapshots: Record<
    string,
    { start?: string; finish?: string; duration?: number; work?: number; cost?: number }
  > = {},
  overrides: Partial<Omit<Baseline, 'id' | 'taskSnapshots'>> = {},
): Baseline {
  const taskSnapshots: Baseline['taskSnapshots'] = {}
  for (const [taskKey, snap] of Object.entries(snapshots)) {
    taskSnapshots[taskKey] = {
      duration: asWorkingMinutes(snap.duration ?? 0),
      work: asWorkingMinutes(snap.work ?? 0),
      cost: snap.cost ?? 0,
      ...(snap.start !== undefined ? { start: asISODateTime(snap.start) } : {}),
      ...(snap.finish !== undefined ? { finish: asISODateTime(snap.finish) } : {}),
    }
  }
  return {
    name: id,
    ...overrides,
    id: asBaselineId(id),
    capturedAt: asISODateTime(capturedAt),
    taskSnapshots,
  }
}

export function makeCustomField(
  id: string,
  name: string,
  type: 'text' | 'number' | 'boolean' | 'date',
): CustomField {
  return { id: asCustomFieldId(id), name, type }
}

export interface DocumentParts {
  startDate?: string
  calendars?: Calendar[]
  tasks?: Task[]
  resources?: Resource[]
  assignments?: Assignment[]
  dependencies?: Dependency[]
  baselines?: Baseline[]
  customFields?: CustomField[]
  views?: ProjectView[]
  tables?: ProjectTable[]
  filters?: ProjectFilter[]
  groups?: ProjectGroup[]
  propertiesId?: string
  propertiesName?: string
}

export function makeDocument(parts: DocumentParts = {}): ProjectDocument {
  const calendar = parts.calendars?.[0] ?? makeCalendar('standard')
  return {
    schemaVersion: 1,
    properties: {
      id: parts.propertiesId ?? 'project',
      name: parts.propertiesName ?? 'Validation Project',
      startDate: asISODateTime(parts.startDate ?? MONDAY),
      defaultCalendarId: calendar.id,
    },
    tasks: parts.tasks ?? [],
    resources: parts.resources ?? [],
    assignments: parts.assignments ?? [],
    dependencies: parts.dependencies ?? [],
    calendars: parts.calendars ?? [calendar],
    baselines: parts.baselines ?? [],
    customFields: parts.customFields ?? [],
    views: parts.views ?? [],
    tables: parts.tables ?? [],
    filters: parts.filters ?? [],
    groups: parts.groups ?? [],
  }
}

// ---- valid golden builders (G01–G12) ------------------------------------

/** G01 — minimal project: properties + one calendar, no tasks. */
export function g01Minimal(): ProjectDocument {
  resetFixtureCounters()
  return makeDocument({ propertiesId: 'g01', propertiesName: 'Minimal' })
}

/** G02 — WBS hierarchy: summary + two children. */
export function g02Wbs(): ProjectDocument {
  resetFixtureCounters()
  const parent = makeTask({
    id: 't1',
    name: 'Phase A',
    summary: true,
    outlineLevel: 1,
    wbs: '1',
    duration: asWorkingMinutes(960),
  })
  const child1 = makeTask({
    id: 't2',
    name: 'Design',
    outlineLevel: 2,
    wbs: '1.1',
    parentTaskId: asTaskId('t1'),
    duration: asWorkingMinutes(480),
  })
  const child2 = makeTask({
    id: 't3',
    name: 'Build',
    outlineLevel: 2,
    wbs: '1.2',
    parentTaskId: asTaskId('t1'),
    duration: asWorkingMinutes(480),
  })
  return makeDocument({
    propertiesId: 'g02',
    propertiesName: 'WBS',
    tasks: [parent, child1, child2],
  })
}

/** G03 — dependency graph: 3 tasks + FS + SS dependencies. */
export function g03Dependencies(): ProjectDocument {
  resetFixtureCounters()
  const a = makeTask({ id: 'a', wbs: '1', duration: asWorkingMinutes(480) })
  const b = makeTask({ id: 'b', wbs: '2', duration: asWorkingMinutes(480) })
  const c = makeTask({ id: 'c', wbs: '3', duration: asWorkingMinutes(480) })
  const deps = [makeDependency('d1', 'a', 'b', 'FS', 0), makeDependency('d2', 'a', 'c', 'SS', 240)]
  return makeDocument({
    propertiesId: 'g03',
    propertiesName: 'Dependencies',
    tasks: [a, b, c],
    dependencies: deps,
  })
}

/** G04 — calendar-rich: base calendar + derived calendar with exception. */
export function g04CalendarRich(): ProjectDocument {
  resetFixtureCounters()
  const base = makeCalendar('base-cal', { name: 'Base Mon-Fri' })
  const derived = makeCalendar('derived-cal', {
    name: 'Derived',
    baseCalendarId: asCalendarId('base-cal'),
    workingWeek: standardWeek(),
    exceptions: [{ date: '2026-08-03', periods: [] }],
  })
  return makeDocument({
    propertiesId: 'g04',
    propertiesName: 'Calendar Rich',
    calendars: [base, derived],
  })
}

/** G05 — resources + assignments. */
export function g05Resources(): ProjectDocument {
  resetFixtureCounters()
  const task = makeTask({ id: 'task-r', wbs: '1', duration: asWorkingMinutes(480) })
  const r1 = makeResource({ id: 'r1', name: 'Alice', kind: 'work', maxUnits: 1, standardRate: 50 })
  const r2 = makeResource({
    id: 'r2',
    name: 'Concrete',
    kind: 'material',
    maxUnits: 0,
    standardRate: 100,
  })
  const a1 = makeAssignment('a1', 'task-r', 'r1', {
    units: 1,
    work: asWorkingMinutes(480),
    remainingWork: asWorkingMinutes(480),
  })
  const a2 = makeAssignment('a2', 'task-r', 'r2', { units: 5, cost: 500, remainingCost: 500 })
  return makeDocument({
    propertiesId: 'g05',
    propertiesName: 'Resources',
    tasks: [task],
    resources: [r1, r2],
    assignments: [a1, a2],
  })
}

/** G06 — baseline-rich: one baseline with snapshots. */
export function g06BaselineRich(): ProjectDocument {
  resetFixtureCounters()
  const t1 = makeTask({ id: 'bt1', wbs: '1', duration: asWorkingMinutes(480) })
  const t2 = makeTask({ id: 'bt2', wbs: '2', duration: asWorkingMinutes(480) })
  const baseline = makeBaseline(
    'b1',
    MONDAY,
    {
      bt1: { start: MONDAY, finish: MONDAY, duration: 480, work: 480, cost: 100 },
      bt2: { duration: 480, work: 0, cost: 0 },
    },
    { name: 'Baseline 1' },
  )
  return makeDocument({
    propertiesId: 'g06',
    propertiesName: 'Baseline Rich',
    tasks: [t1, t2],
    baselines: [baseline],
  })
}

/** G07 — constraints / deadlines / progress. */
export function g07Constraints(): ProjectDocument {
  resetFixtureCounters()
  const t1 = makeTask({
    id: 'ct1',
    wbs: '1',
    duration: asWorkingMinutes(480),
    constraintType: 'startNoEarlierThan',
    constraintDate: asISODateTime(MONDAY),
    deadline: asISODateTime(FRIDAY),
    percentComplete: 50,
  })
  const t2 = makeTask({
    id: 'ct2',
    wbs: '2',
    duration: asWorkingMinutes(0),
    milestone: true,
    constraintType: 'mustFinishOn',
    constraintDate: asISODateTime(FRIDAY),
    percentComplete: 0,
  })
  return makeDocument({ propertiesId: 'g07', propertiesName: 'Constraints', tasks: [t1, t2] })
}

/** G08 — work / cost project: assignments carry work + cost. */
export function g08WorkCost(): ProjectDocument {
  resetFixtureCounters()
  const t = makeTask({ id: 'wt1', wbs: '1', duration: asWorkingMinutes(480), percentComplete: 25 })
  const r = makeResource({
    id: 'wr1',
    name: 'Worker',
    kind: 'work',
    maxUnits: 1,
    standardRate: 60,
    costPerUse: 10,
  })
  const a = makeAssignment('wa1', 'wt1', 'wr1', {
    units: 1,
    work: asWorkingMinutes(480),
    actualWork: asWorkingMinutes(120),
    remainingWork: asWorkingMinutes(360),
    cost: 490,
    actualCost: 122.5,
    remainingCost: 367.5,
  })
  return makeDocument({
    propertiesId: 'g08',
    propertiesName: 'Work Cost',
    tasks: [t],
    resources: [r],
    assignments: [a],
  })
}

/** G09 — custom fields: definitions + per-task values. */
export function g09CustomFields(): ProjectDocument {
  resetFixtureCounters()
  const cf1 = makeCustomField('cf1', 'Sponsor', 'text')
  const cf2 = makeCustomField('cf2', 'Budget', 'number')
  const cf3 = makeCustomField('cf3', 'Flagged', 'boolean')
  const t = makeTask({
    id: 'cft1',
    wbs: '1',
    duration: asWorkingMinutes(480),
    customFields: {
      cf1: 'Acme Corp',
      cf2: 5000,
      cf3: true,
    } as Task['customFields'],
  })
  return makeDocument({
    propertiesId: 'g09',
    propertiesName: 'Custom Fields',
    tasks: [t],
    customFields: [cf1, cf2, cf3],
  })
}

/** G10 — views / tables / filters / groups. */
export function g10Views(): ProjectDocument {
  resetFixtureCounters()
  const table = {
    id: asProjectTableId('tbl1'),
    name: 'Entry',
    columns: ['name', 'duration', 'start', 'finish'],
  }
  const filter = {
    id: asProjectFilterId('flt1'),
    name: 'Critical',
    expression: 'critical === true',
  }
  const group = { id: asProjectGroupId('grp1'), name: 'By Phase', expression: 'phase' }
  const view: ProjectView = {
    id: asProjectViewId('vw1'),
    name: 'Gantt',
    type: 'gantt',
    tableId: asProjectTableId('tbl1'),
    filterId: asProjectFilterId('flt1'),
    groupId: asProjectGroupId('grp1'),
  }
  return makeDocument({
    propertiesId: 'g10',
    propertiesName: 'Views',
    views: [view],
    tables: [table],
    filters: [filter],
    groups: [group],
  })
}

/** G11 — multiple baselines. */
export function g11MultiBaseline(): ProjectDocument {
  resetFixtureCounters()
  const t = makeTask({ id: 'mt1', wbs: '1', duration: asWorkingMinutes(480) })
  const b1 = makeBaseline(
    'mb1',
    MONDAY,
    { mt1: { start: MONDAY, duration: 480, work: 480, cost: 100 } },
    { name: 'Baseline 1' },
  )
  const b2 = makeBaseline(
    'mb2',
    TUESDAY,
    { mt1: { start: TUESDAY, duration: 480, work: 480, cost: 110 } },
    { name: 'Baseline 2' },
  )
  const b3 = makeBaseline(
    'mb3',
    WEDNESDAY,
    { mt1: { start: WEDNESDAY, duration: 480, work: 480, cost: 120 } },
    { name: 'Baseline 3' },
  )
  return makeDocument({
    propertiesId: 'g11',
    propertiesName: 'Multi Baseline',
    tasks: [t],
    baselines: [b1, b2, b3],
  })
}

/** G12 — large project: 60 tasks, 50 dependencies, 10 resources, 60 assignments. */
export function g12Large(): ProjectDocument {
  resetFixtureCounters()
  const tasks: Task[] = []
  for (let i = 1; i <= 60; i++) {
    tasks.push(
      makeTask({
        id: 'lt' + i,
        wbs: String(i),
        duration: asWorkingMinutes(480),
        name: 'Task ' + i,
      }),
    )
  }
  const deps: Dependency[] = []
  for (let i = 1; i <= 50; i++) {
    const type: DependencyType = (['FS', 'SS', 'FF', 'SF'] as const)[i % 4]
    deps.push(makeDependency('ld' + i, 'lt' + i, 'lt' + (i + 1), type, (i % 3) * 60))
  }
  const resources: Resource[] = []
  for (let i = 1; i <= 10; i++) {
    resources.push(
      makeResource({
        id: 'lr' + i,
        name: 'Resource ' + i,
        kind: 'work',
        maxUnits: 1,
        standardRate: 50 + i,
      }),
    )
  }
  const assignments: Assignment[] = []
  for (let i = 1; i <= 60; i++) {
    assignments.push(
      makeAssignment('la' + i, 'lt' + i, 'lr' + (1 + (i % 10)), {
        units: 1,
        work: asWorkingMinutes(480),
        remainingWork: asWorkingMinutes(480),
      }),
    )
  }
  return makeDocument({
    propertiesId: 'g12',
    propertiesName: 'Large',
    tasks,
    resources,
    assignments,
    dependencies: deps,
  })
}

export const VALID_GOLDEN_BUILDERS: ReadonlyArray<{
  id: string
  build: () => ProjectDocument
  note: string
}> = [
  { id: 'G01', build: g01Minimal, note: 'minimal project' },
  { id: 'G02', build: g02Wbs, note: 'WBS hierarchy' },
  { id: 'G03', build: g03Dependencies, note: 'dependency graph' },
  { id: 'G04', build: g04CalendarRich, note: 'calendar-rich' },
  { id: 'G05', build: g05Resources, note: 'resources + assignments' },
  { id: 'G06', build: g06BaselineRich, note: 'baseline-rich' },
  { id: 'G07', build: g07Constraints, note: 'constraints/deadlines/progress' },
  { id: 'G08', build: g08WorkCost, note: 'work/cost' },
  { id: 'G09', build: g09CustomFields, note: 'custom fields' },
  { id: 'G10', build: g10Views, note: 'views/tables/filters/groups' },
  { id: 'G11', build: g11MultiBaseline, note: 'multi-baseline' },
  { id: 'G12', build: g12Large, note: 'large project' },
]

// ---- invalid golden bytes (G13–G15) -------------------------------------
//
// These are hand-crafted malformed `.gproj` byte payloads (not generated by
// the serializer) so they exercise the deserializer's rejection paths.

/** G13 — invalid schema: a task whose `duration` is a string, not a number. */
export function g13InvalidSchemaBytes(): Uint8Array {
  const text = [
    '{',
    '  "document": {',
    '    "schemaVersion": 1,',
    '    "properties": {',
    '      "id": "g13",',
    '      "name": "Invalid Schema",',
    '      "startDate": "2026-08-03T09:00:00.000Z",',
    '      "defaultCalendarId": "standard"',
    '    },',
    '    "calendars": [',
    '      {',
    '        "id": "standard",',
    '        "name": "Standard",',
    '        "workingWeek": { "1": [{ "startMinute": 540, "endMinute": 1020 }] },',
    '        "exceptions": []',
    '      }',
    '    ],',
    '    "tasks": [',
    '      {',
    '        "id": "x1", "uid": 1, "wbs": "1", "outlineLevel": 1, "name": "Bad",',
    '        "taskType": "fixedDuration", "summary": false, "milestone": false,',
    '        "manualScheduled": false, "autoScheduled": true,',
    '        "duration": "not-a-number",',
    '        "priority": 500, "percentComplete": 0, "work": 0, "remainingWork": 0,',
    '        "actualWork": 0, "cost": 0, "actualCost": 0, "remainingCost": 0,',
    '        "baseline": [], "customFields": {}, "notes": []',
    '      }',
    '    ],',
    '    "resources": [], "assignments": [], "dependencies": [], "baselines": [],',
    '    "customFields": [], "views": [], "tables": [], "filters": [], "groups": []',
    '  },',
    '  "format": "gproj",',
    '  "formatVersion": 1,',
    '  "metadata": { "format": "gproj", "version": "1" }',
    '}',
    '',
  ].join('\n')
  return encodeUtf8(text)
}

/** G14 — unsupported version: formatVersion 999. */
export function g14UnsupportedVersionBytes(): Uint8Array {
  const text = [
    '{',
    '  "document": {',
    '    "schemaVersion": 1,',
    '    "properties": {',
    '      "id": "g14", "name": "Unsupported", "startDate": "2026-08-03T09:00:00.000Z",',
    '      "defaultCalendarId": "standard"',
    '    },',
    '    "calendars": [], "tasks": [], "resources": [], "assignments": [],',
    '    "dependencies": [], "baselines": [], "customFields": [], "views": [],',
    '    "tables": [], "filters": [], "groups": []',
    '  },',
    '  "format": "gproj",',
    '  "formatVersion": 999,',
    '  "metadata": { "format": "gproj", "version": "999" }',
    '}',
    '',
  ].join('\n')
  return encodeUtf8(text)
}

/** G15 — malformed references: a dependency whose predecessor does not exist. */
export function g15MalformedReferenceBytes(): Uint8Array {
  const text = [
    '{',
    '  "document": {',
    '    "schemaVersion": 1,',
    '    "properties": {',
    '      "id": "g15", "name": "Bad Refs", "startDate": "2026-08-03T09:00:00.000Z",',
    '      "defaultCalendarId": "standard"',
    '    },',
    '    "calendars": [',
    '      {',
    '        "id": "standard", "name": "Standard",',
    '        "workingWeek": { "1": [{ "startMinute": 540, "endMinute": 1020 }] },',
    '        "exceptions": []',
    '      }',
    '    ],',
    '    "tasks": [',
    '      { "id": "ok1", "uid": 1, "wbs": "1", "outlineLevel": 1, "name": "OK1",',
    '        "taskType": "fixedDuration", "summary": false, "milestone": false,',
    '        "manualScheduled": false, "autoScheduled": true, "duration": 480,',
    '        "priority": 500, "percentComplete": 0, "work": 0, "remainingWork": 0,',
    '        "actualWork": 0, "cost": 0, "actualCost": 0, "remainingCost": 0,',
    '        "baseline": [], "customFields": {}, "notes": [] }',
    '    ],',
    '    "dependencies": [',
    '      { "id": "dangling", "predecessorId": "does-not-exist", "successorId": "ok1",',
    '        "type": "FS", "lagMinutes": 0 }',
    '    ],',
    '    "resources": [], "assignments": [], "baselines": [], "customFields": [],',
    '    "views": [], "tables": [], "filters": [], "groups": []',
    '  },',
    '  "format": "gproj",',
    '  "formatVersion": 1,',
    '  "metadata": { "format": "gproj", "version": "1" }',
    '}',
    '',
  ].join('\n')
  return encodeUtf8(text)
}

export const INVALID_GOLDEN_BYTES: ReadonlyArray<{
  id: string
  bytes: () => Uint8Array
  expectedCode: string
  note: string
}> = [
  { id: 'G13', bytes: g13InvalidSchemaBytes, expectedCode: 'INVALID_TASK', note: 'invalid schema' },
  {
    id: 'G14',
    bytes: g14UnsupportedVersionBytes,
    expectedCode: 'UNSUPPORTED_GPROJ_VERSION',
    note: 'unsupported version',
  },
  {
    id: 'G15',
    bytes: g15MalformedReferenceBytes,
    expectedCode: 'MISSING_TASK_REFERENCE',
    note: 'malformed references',
  },
]

/** Helper: every TaskId string set in a document (for identity-preservation assertions). */
export function taskIdSet(doc: ProjectDocument): Set<string> {
  return new Set(doc.tasks.map((t) => t.id as string))
}
