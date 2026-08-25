import type {
  AssignmentId,
  BaselineId,
  CalendarId,
  CustomFieldId,
  DependencyId,
  ISODateTime,
  ProjectFilterId,
  ProjectGroupId,
  ProjectTableId,
  ProjectViewId,
  ResourceId,
  TaskId,
  WorkingMinutes,
} from './types.js'

// Explicit brand-boundary conversions. Branded types cannot be constructed from
// raw values without an assertion, so these helpers are the single canonical
// place where plain strings/numbers are promoted to branded domain values.
// Using them keeps `as`-casts out of engines and tests.

export const asTaskId = (value: string): TaskId => value as TaskId
export const asResourceId = (value: string): ResourceId => value as ResourceId
export const asAssignmentId = (value: string): AssignmentId => value as AssignmentId
export const asDependencyId = (value: string): DependencyId => value as DependencyId
export const asCalendarId = (value: string): CalendarId => value as CalendarId
export const asBaselineId = (value: string): BaselineId => value as BaselineId
export const asCustomFieldId = (value: string): CustomFieldId => value as CustomFieldId
export const asProjectViewId = (value: string): ProjectViewId => value as ProjectViewId
export const asProjectTableId = (value: string): ProjectTableId => value as ProjectTableId
export const asProjectFilterId = (value: string): ProjectFilterId => value as ProjectFilterId
export const asProjectGroupId = (value: string): ProjectGroupId => value as ProjectGroupId
export const asISODateTime = (value: string): ISODateTime => value as ISODateTime
export const asWorkingMinutes = (value: number): WorkingMinutes => value as WorkingMinutes

/** Recover the plain numeric value of a WorkingMinutes brand. */
export const plainMinutes = (value: WorkingMinutes): number => value as number
