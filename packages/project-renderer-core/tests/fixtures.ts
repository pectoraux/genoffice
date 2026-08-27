/**
 * PROJECT-021 renderer-core test fixtures.
 *
 * Self-contained canonical document builders (mirroring the engine test
 * fixture semantics) so the renderer-core suite never depends on another
 * package's test internals. The scheduling package is imported at the TEST
 * layer only (the accepted project-file precedent): the package under test
 * stays scheduling-free with the scheduler injected.
 */
import {
  asCalendarId,
  asDependencyId,
  asISODateTime,
  asResourceId,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type {
  Assignment,
  Calendar,
  CalendarPeriod,
  Dependency,
  DependencyType,
  DerivedSchedule,
  ProjectDocument,
  Resource,
  Task,
  TaskSchedule,
} from '@genoffice/project-contracts'

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

export function makeCalendar(id: string, overrides: Partial<Calendar> = {}): Calendar {
  const { id: _ignored, ...rest } = overrides
  return {
    name: id,
    workingWeek: standardWeek(),
    exceptions: [],
    ...rest,
    id: asCalendarId(id),
  }
}

let fixtureUid = 0

export function makeTask(overrides: Omit<Partial<Task>, 'id'> & { id: string }): Task {
  const { id, ...rest } = overrides
  fixtureUid += 1
  return {
    uid: fixtureUid,
    wbs: '',
    outlineLevel: 1,
    name: id,
    taskType: 'fixedDuration',
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

export function makeDependency(
  id: string,
  predecessorId: string,
  successorId: string,
  type: DependencyType = 'FS',
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

let fixtureResourceUid = 0

export function makeResource(overrides: Omit<Partial<Resource>, 'id'> & { id: string }): Resource {
  const { id, ...rest } = overrides
  fixtureResourceUid += 1
  return {
    uid: fixtureResourceUid,
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
    id: asTaskId(id) as unknown as Assignment['id'],
    taskId: asTaskId(taskId),
    resourceId: asResourceId(resourceId),
  }
}

export interface DocumentParts {
  tasks?: Task[]
  dependencies?: Dependency[]
  calendars?: Calendar[]
  resources?: Resource[]
  assignments?: Assignment[]
  startDate?: string
  finishDate?: string
  views?: ProjectDocument['views']
  tables?: ProjectDocument['tables']
  filters?: ProjectDocument['filters']
  groups?: ProjectDocument['groups']
}

export function makeDocument(parts: DocumentParts = {}): ProjectDocument {
  const calendar = parts.calendars?.[0] ?? makeCalendar('standard')
  return {
    schemaVersion: 1,
    properties: {
      id: 'project',
      name: 'Renderer Fixture',
      startDate: asISODateTime(parts.startDate ?? '2026-08-03T09:00:00.000Z'),
      ...(parts.finishDate !== undefined ? { finishDate: asISODateTime(parts.finishDate) } : {}),
      defaultCalendarId: calendar.id,
    },
    tasks: parts.tasks ?? [],
    resources: parts.resources ?? [],
    assignments: parts.assignments ?? [],
    dependencies: parts.dependencies ?? [],
    calendars: parts.calendars ?? [calendar],
    baselines: [],
    customFields: [],
    views: parts.views ?? [],
    tables: parts.tables ?? [],
    filters: parts.filters ?? [],
    groups: parts.groups ?? [],
  }
}

/** A canonical three-level outline: root > (a, b), a > a1. Derived hierarchy
 * fields (outlineLevel, summary, wbs) are consistent with PROJECT-007 so the
 * document passes canonical validation (commands only execute against valid
 * documents). */
export function outlineDocument(): ProjectDocument {
  return makeDocument({
    tasks: [
      makeTask({ id: 'root', outlineLevel: 1, summary: true, wbs: '1' }),
      makeTask({
        id: 'a',
        parentTaskId: asTaskId('root'),
        outlineLevel: 2,
        summary: true,
        wbs: '1.1',
      }),
      makeTask({ id: 'a1', parentTaskId: asTaskId('a'), outlineLevel: 3, wbs: '1.1.1' }),
      makeTask({ id: 'b', parentTaskId: asTaskId('root'), outlineLevel: 2, wbs: '1.2' }),
    ],
  })
}

/** A canonical outline with a THREE-sibling group (p > a1, a2, a3) and two
 * roots — the non-last-anchor counterexample fixture for the CreateTask
 * insert-position semantics: `a2` and `root1` are anchors that are NOT the
 * last member of their sibling group, so an append lands after a3/root2
 * respectively (PROJECT-021 review round 1). */
export function multiSiblingDocument(): ProjectDocument {
  return makeDocument({
    tasks: [
      makeTask({ id: 'root1', outlineLevel: 1, summary: true, wbs: '1' }),
      makeTask({
        id: 'p',
        parentTaskId: asTaskId('root1'),
        outlineLevel: 2,
        summary: true,
        wbs: '1.1',
      }),
      makeTask({ id: 'a1', parentTaskId: asTaskId('p'), outlineLevel: 3, wbs: '1.1.1' }),
      makeTask({ id: 'a2', parentTaskId: asTaskId('p'), outlineLevel: 3, wbs: '1.1.2' }),
      makeTask({ id: 'a3', parentTaskId: asTaskId('p'), outlineLevel: 3, wbs: '1.1.3' }),
      makeTask({ id: 'root2', outlineLevel: 1, wbs: '2' }),
    ],
  })
}

// ---------------------------------------------------------------------------
// PROJECT-022 — Gantt view fixtures.
// ---------------------------------------------------------------------------

/** A hand-authored schedule entry (exact dates under the test's control —
 * the scheduling authority itself is exercised for real in the session and
 * gantt-view integration tests). */
export function makeScheduleEntry(
  taskId: string,
  start: string,
  finish: string,
  overrides: Partial<TaskSchedule> = {},
): TaskSchedule {
  const { taskId: _ignored, ...rest } = overrides
  return {
    duration: asWorkingMinutes(480),
    totalSlack: 0,
    freeSlack: 0,
    critical: false,
    ...rest,
    taskId: asTaskId(taskId),
    scheduledStart: asISODateTime(start),
    scheduledFinish: asISODateTime(finish),
  }
}

/** The PROJECT-022 Gantt document: root > (a > a1,a2, b, m), one flagged
 * milestone, two dependencies (FS a1→a2, SS a2→b with lag). Document order
 * is the canonical outline order; uids are explicit so grid cells are
 * stable and readable. */
export function ganttDocument(): ProjectDocument {
  return makeDocument({
    startDate: '2026-08-01T00:00:00.000Z',
    finishDate: '2026-08-31T00:00:00.000Z',
    tasks: [
      makeTask({ id: 'root', uid: 1, outlineLevel: 1, summary: true, wbs: '1' }),
      makeTask({
        id: 'a',
        uid: 2,
        parentTaskId: asTaskId('root'),
        outlineLevel: 2,
        summary: true,
        wbs: '1.1',
      }),
      makeTask({
        id: 'a1',
        uid: 3,
        parentTaskId: asTaskId('a'),
        outlineLevel: 3,
        duration: asWorkingMinutes(480),
        percentComplete: 50,
        wbs: '1.1.1',
      }),
      makeTask({
        id: 'a2',
        uid: 4,
        parentTaskId: asTaskId('a'),
        outlineLevel: 3,
        duration: asWorkingMinutes(960),
        wbs: '1.1.2',
      }),
      makeTask({
        id: 'b',
        uid: 5,
        parentTaskId: asTaskId('root'),
        outlineLevel: 2,
        duration: asWorkingMinutes(1440),
        wbs: '1.2',
      }),
      makeTask({
        id: 'm',
        uid: 6,
        parentTaskId: asTaskId('root'),
        outlineLevel: 2,
        milestone: true,
        duration: asWorkingMinutes(0),
        wbs: '1.3',
      }),
    ],
    dependencies: [makeDependency('d1', 'a1', 'a2'), makeDependency('d2', 'a2', 'b', 'SS', 60)],
  })
}

/** A hand-authored derived schedule for `ganttDocument()` (working-day
 * shaped instants; roll-up windows on the summaries; 50% progress on a1).
 * The milestone `m` carries a zero-span window (start === finish). */
export function ganttSchedule(): DerivedSchedule {
  return {
    taskSchedules: {
      [asTaskId('root')]: makeScheduleEntry(
        'root',
        '2026-08-03T09:00:00.000Z',
        '2026-08-12T17:00:00.000Z',
        {
          duration: asWorkingMinutes(4800),
          percentComplete: 25,
        },
      ),
      [asTaskId('a')]: makeScheduleEntry(
        'a',
        '2026-08-03T09:00:00.000Z',
        '2026-08-05T17:00:00.000Z',
        {
          duration: asWorkingMinutes(2400),
          percentComplete: 25,
        },
      ),
      [asTaskId('a1')]: makeScheduleEntry(
        'a1',
        '2026-08-03T09:00:00.000Z',
        '2026-08-03T17:00:00.000Z',
        {
          percentComplete: 50,
        },
      ),
      [asTaskId('a2')]: makeScheduleEntry(
        'a2',
        '2026-08-04T09:00:00.000Z',
        '2026-08-05T17:00:00.000Z',
        {
          duration: asWorkingMinutes(960),
        },
      ),
      [asTaskId('b')]: makeScheduleEntry(
        'b',
        '2026-08-10T09:00:00.000Z',
        '2026-08-12T17:00:00.000Z',
        {
          duration: asWorkingMinutes(1440),
        },
      ),
      [asTaskId('m')]: makeScheduleEntry(
        'm',
        '2026-08-07T09:00:00.000Z',
        '2026-08-07T09:00:00.000Z',
        {
          duration: asWorkingMinutes(0),
        },
      ),
    },
    projectStart: asISODateTime('2026-08-03T09:00:00.000Z'),
    projectFinish: asISODateTime('2026-08-12T17:00:00.000Z'),
    diagnostics: [],
  }
}
