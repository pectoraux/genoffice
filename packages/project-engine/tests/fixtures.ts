import {
  asAssignmentId,
  asBaselineId,
  asCalendarId,
  asDependencyId,
  asISODateTime,
  asResourceId,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type {
  Assignment,
  Baseline,
  BaselineId,
  Calendar,
  CalendarPeriod,
  CustomField,
  Dependency,
  DependencyType,
  ProjectDocument,
  Resource,
  Task,
  TaskId,
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

// Task UIDs must be unique within a document (interoperability identity).
// Fixture tasks receive an auto-incrementing UID so multi-task documents stay
// valid; tests may always pass an explicit uid override.
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

// PROJECT-010: resource UIDs must be unique within a document (interoperability
// identity), mirroring the Task uid rule. Fixture resources receive an
// auto-incrementing UID so multi-resource documents stay valid; tests may
// always pass an explicit uid override.
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
    id: asAssignmentId(id),
    taskId: asTaskId(taskId),
    resourceId: asResourceId(resourceId),
  }
}

// PROJECT-009: a baseline is an immutable snapshot of task start/finish/
// duration/work/cost captured at a point in time. The helper accepts a
// task-id → snapshot map (or an empty snapshot when the caller only needs
// the baseline shell for a rejection test).
export interface BaselineSnapshotInput {
  start?: string
  finish?: string
  duration?: number
  work?: number
  cost?: number
}

export function makeBaseline(
  id: string,
  capturedAt: string,
  snapshots: Record<string, BaselineSnapshotInput> = {},
  overrides: Partial<Omit<Baseline, 'id' | 'taskSnapshots'>> = {},
): Baseline {
  const taskSnapshots: Baseline['taskSnapshots'] = {}
  for (const [taskKey, snap] of Object.entries(snapshots)) {
    taskSnapshots[taskKey] = {
      start: snap.start !== undefined ? asISODateTime(snap.start) : undefined,
      finish: snap.finish !== undefined ? asISODateTime(snap.finish) : undefined,
      duration: asWorkingMinutes(snap.duration ?? 0),
      work: asWorkingMinutes(snap.work ?? 0),
      cost: snap.cost ?? 0,
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

export { asBaselineId }
export type { BaselineId, TaskId }

export interface DocumentParts {
  tasks?: Task[]
  dependencies?: Dependency[]
  calendars?: Calendar[]
  resources?: Resource[]
  assignments?: Assignment[]
  baselines?: Baseline[]
  customFields?: CustomField[]
  startDate?: string
}

export function makeDocument(parts: DocumentParts = {}): ProjectDocument {
  const calendar = parts.calendars?.[0] ?? makeCalendar('standard')
  return {
    schemaVersion: 1,
    properties: {
      id: 'project',
      name: 'Validation Project',
      startDate: asISODateTime(parts.startDate ?? '2026-08-03T09:00:00.000Z'),
      defaultCalendarId: calendar.id,
    },
    tasks: parts.tasks ?? [],
    resources: parts.resources ?? [],
    assignments: parts.assignments ?? [],
    dependencies: parts.dependencies ?? [],
    calendars: parts.calendars ?? [calendar],
    baselines: parts.baselines ?? [],
    customFields: parts.customFields ?? [],
    views: [],
    tables: [],
    filters: [],
    groups: [],
  }
}
