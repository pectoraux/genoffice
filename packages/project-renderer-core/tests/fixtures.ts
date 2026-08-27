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
  ProjectDocument,
  Resource,
  Task,
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
