import {
  asCalendarId,
  asDependencyId,
  asISODateTime,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type {
  Assignment,
  Calendar,
  CalendarException,
  CalendarPeriod,
  Dependency,
  DependencyType,
  ISODateTime,
  ProjectDocument,
  Resource,
  Task,
  TaskId,
  WorkingMinutes,
} from '@genoffice/project-contracts'

export const iso = asISODateTime
export const wm = asWorkingMinutes
export const taskId = asTaskId

export const MONDAY = '2026-08-03T09:00:00.000Z'
export const MONDAY_FINISH = '2026-08-03T17:00:00.000Z'
export const TUESDAY = '2026-08-04T09:00:00.000Z'
export const TUESDAY_FINISH = '2026-08-04T17:00:00.000Z'
export const WEDNESDAY = '2026-08-05T09:00:00.000Z'
export const WEDNESDAY_FINISH = '2026-08-05T17:00:00.000Z'
export const THURSDAY = '2026-08-06T09:00:00.000Z'
export const THURSDAY_FINISH = '2026-08-06T17:00:00.000Z'
export const FRIDAY = '2026-08-07T09:00:00.000Z'
export const FRIDAY_FINISH = '2026-08-07T17:00:00.000Z'
export const NEXT_MONDAY = '2026-08-10T09:00:00.000Z'
export const PREV_FRIDAY = '2026-07-31T17:00:00.000Z'

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
  const { id: _ignoredId, ...rest } = overrides
  return {
    name: id,
    workingWeek: standardWeek(),
    exceptions: [],
    ...rest,
    id: asCalendarId(id),
  }
}

export function holiday(date: string): CalendarException {
  return { date, periods: [] }
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
    duration: wm(480),
    priority: 500,
    percentComplete: 0,
    work: wm(0),
    remainingWork: wm(0),
    actualWork: wm(0),
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

export interface DocumentParts {
  tasks?: Task[]
  dependencies?: Dependency[]
  calendars?: Calendar[]
  resources?: Resource[]
  assignments?: Assignment[]
  startDate?: ISODateTime
}

export function makeDocument(parts: DocumentParts = {}): ProjectDocument {
  const calendar = parts.calendars?.[0] ?? makeCalendar('standard')
  return {
    schemaVersion: 1,
    properties: {
      id: 'project',
      name: 'Golden Project',
      startDate: parts.startDate ?? iso(MONDAY),
      defaultCalendarId: calendar.id,
    },
    tasks: parts.tasks ?? [],
    resources: parts.resources ?? [],
    assignments: parts.assignments ?? [],
    dependencies: parts.dependencies ?? [],
    calendars: parts.calendars ?? [calendar],
    baselines: [],
    customFields: [],
    views: [],
    tables: [],
    filters: [],
    groups: [],
  }
}

export function parseDocument(json: string): ProjectDocument {
  return JSON.parse(json) as ProjectDocument
}

export type { Task, TaskId, WorkingMinutes, Calendar, Dependency }
