export type Brand<T, Name extends string> = T & { readonly __brand: Name }

export type TaskId = Brand<string, 'TaskId'>
export type ResourceId = Brand<string, 'ResourceId'>
export type AssignmentId = Brand<string, 'AssignmentId'>
export type DependencyId = Brand<string, 'DependencyId'>
export type CalendarId = Brand<string, 'CalendarId'>
export type BaselineId = Brand<string, 'BaselineId'>
export type CustomFieldId = Brand<string, 'CustomFieldId'>
export type ProjectViewId = Brand<string, 'ProjectViewId'>
export type ProjectTableId = Brand<string, 'ProjectTableId'>
export type ProjectFilterId = Brand<string, 'ProjectFilterId'>
export type ProjectGroupId = Brand<string, 'ProjectGroupId'>

export type ISODateTime = string & { readonly __brand: 'ISODateTime' }
export type WorkingMinutes = number & { readonly __brand: 'WorkingMinutes' }

export type TaskType = 'fixedUnits' | 'fixedWork' | 'fixedDuration'
export type DependencyType = 'FS' | 'SS' | 'FF' | 'SF'
export type ConstraintType =
  | 'asSoonAsPossible'
  | 'asLateAsPossible'
  | 'startNoEarlierThan'
  | 'startNoLaterThan'
  | 'mustStartOn'
  | 'finishNoEarlierThan'
  | 'finishNoLaterThan'
  | 'mustFinishOn'

/**
 * PROJECT-008 derived progress status. The engine derives this deterministically
 * from `percentComplete`, the scheduled window, and `ProjectProperties.statusDate`.
 * It is never authoritative renderer state and never depends on wall-clock time.
 *
 * Canonical precedence (documented in spec/project/requirements.md):
 *  - `complete`: percentComplete >= 100.
 *  - `inProgress`: 0 < percentComplete < 100, OR (percentComplete == 0 AND a
 *     statusDate is set that has reached/passed the scheduled start).
 *  - `notStarted`: percentComplete == 0 AND (no statusDate, OR statusDate is
 *     still before the scheduled start).
 *
 * Milestones are zero-duration binary events: `complete` at 100%, otherwise
 * `notStarted` (the "in progress" window is empty for a zero-duration task).
 */
export type TaskProgressStatus = 'notStarted' | 'inProgress' | 'complete'

export interface ProjectProperties {
  id: string
  name: string
  startDate: ISODateTime
  finishDate?: ISODateTime
  statusDate?: ISODateTime
  defaultCalendarId: CalendarId
}

export interface Task {
  id: TaskId
  uid: number
  wbs: string
  outlineLevel: number
  name: string
  taskType: TaskType
  summary: boolean
  milestone: boolean
  manualScheduled: boolean
  autoScheduled: boolean
  start?: ISODateTime
  finish?: ISODateTime
  duration: WorkingMinutes
  constraintType?: ConstraintType
  constraintDate?: ISODateTime
  deadline?: ISODateTime
  priority: number
  calendarId?: CalendarId
  percentComplete: number
  physicalPercentComplete?: number
  work: WorkingMinutes
  remainingWork: WorkingMinutes
  actualWork: WorkingMinutes
  cost: number
  actualCost: number
  remainingCost: number
  baseline: BaselineId[]
  customFields: Record<CustomFieldId, string | number | boolean | null>
  notes: string[]
  parentTaskId?: TaskId
}

export interface Resource {
  id: ResourceId
  uid: number
  name: string
  kind: 'work' | 'material' | 'cost'
  maxUnits: number
  standardRate: number
  overtimeRate: number
  costPerUse: number
  calendarId?: CalendarId
  availability: Array<{ start: ISODateTime; finish?: ISODateTime; units: number }>
}

export interface Assignment {
  id: AssignmentId
  taskId: TaskId
  resourceId: ResourceId
  units: number
  work: WorkingMinutes
  actualWork: WorkingMinutes
  remainingWork: WorkingMinutes
  cost: number
  actualCost: number
  remainingCost: number
}

export interface Dependency {
  id: DependencyId
  predecessorId: TaskId
  successorId: TaskId
  type: DependencyType
  lagMinutes: number
}

export interface CalendarPeriod {
  startMinute: number
  endMinute: number
}
export interface CalendarException {
  date: string
  periods: CalendarPeriod[]
}
export interface Calendar {
  id: CalendarId
  name: string
  baseCalendarId?: CalendarId
  workingWeek: Record<number, CalendarPeriod[]>
  exceptions: CalendarException[]
}

export interface Baseline {
  id: BaselineId
  name: string
  capturedAt: ISODateTime
  taskSnapshots: Record<
    string,
    {
      start?: ISODateTime
      finish?: ISODateTime
      duration: WorkingMinutes
      work: WorkingMinutes
      cost: number
    }
  >
}

export interface CustomField {
  id: CustomFieldId
  name: string
  type: 'text' | 'number' | 'boolean' | 'date'
}
export interface ProjectView {
  id: ProjectViewId
  name: string
  type: string
  tableId?: ProjectTableId
  filterId?: ProjectFilterId
  groupId?: ProjectGroupId
}
export interface ProjectTable {
  id: ProjectTableId
  name: string
  columns: string[]
}
export interface ProjectFilter {
  id: ProjectFilterId
  name: string
  expression: string
}
export interface ProjectGroup {
  id: ProjectGroupId
  name: string
  expression: string
}

export interface ProjectDocument {
  schemaVersion: 1
  properties: ProjectProperties
  tasks: Task[]
  resources: Resource[]
  assignments: Assignment[]
  dependencies: Dependency[]
  calendars: Calendar[]
  baselines: Baseline[]
  customFields: CustomField[]
  views: ProjectView[]
  tables: ProjectTable[]
  filters: ProjectFilter[]
  groups: ProjectGroup[]
}

export interface TaskSchedule {
  taskId: TaskId
  earlyStart?: ISODateTime
  earlyFinish?: ISODateTime
  lateStart?: ISODateTime
  lateFinish?: ISODateTime
  totalSlack: number
  freeSlack: number
  critical: boolean
  scheduledStart?: ISODateTime
  scheduledFinish?: ISODateTime
  duration: WorkingMinutes
  // ---- PROJECT-008 derived deadline state ----
  /**
   * Deadline echoed from the task. A deadline is NOT a scheduling constraint:
   * it never moves the task. It is only used to derive variance/missed state
   * for downstream reporting layers.
   */
  deadline?: ISODateTime
  /**
   * Signed working-minute variance from scheduledFinish to the deadline,
   * computed in the task's resolved calendar. Positive when the task finishes
   * before the deadline (ahead/on time); negative when the task finishes after
   * the deadline (missed); zero when the finish equals the deadline.
   */
  deadlineVariance?: number
  /** True when the task finishes strictly after its deadline. */
  deadlineMissed?: boolean
  // ---- PROJECT-008 derived progress state ----
  /**
   * Derived progress status. For leaf/milestone tasks this is derived from
   * `percentComplete` + the status date; for summary tasks it is derived from
   * the rolled-up progress of the subtree.
   */
  status?: TaskProgressStatus
  /**
   * Derived percent-complete echo. For leaf/milestone tasks equals the stored
   * `Task.percentComplete`; for summary tasks equals the duration-weighted
   * roll-up of the subtree. Always in [0, 100].
   */
  percentComplete?: number
  /** Echo of the task's physical percent-complete (leaf only; undefined on summaries). */
  physicalPercentComplete?: number
  /** Working minutes already accomplished (rounded). Equals duration*percent/100 on leaves. */
  actualDuration?: WorkingMinutes
  /** Working minutes remaining (= duration - actualDuration on leaves). */
  remainingDuration?: WorkingMinutes
}

export interface DerivedSchedule {
  taskSchedules: Record<TaskId, TaskSchedule>
  projectStart?: ISODateTime
  projectFinish?: ISODateTime
  diagnostics: ImportDiagnostic[]
}

// ---- PROJECT-009 derived baseline-variance state ----
/**
 * PROJECT-009 per-task baseline variance. A baseline is an immutable snapshot
 * of task start/finish/duration/work/cost captured at a point in time. The
 * comparison projects the CURRENT derived schedule against that snapshot.
 *
 * Sign convention (explicit — mirrors the Microsoft Project "Variance" table):
 *  - `startVariance` / `finishVariance` are SIGNED WORKING-MINUTE spans,
 *    computed in the task's resolved calendar, measured as
 *    `signedWorkingDuration(baseline, current)`:
 *      • positive when the current date is LATER than the baseline (the task
 *        has slipped past its planned date);
 *      • negative when the current date is EARLIER (the task is ahead of plan);
 *      • zero when the dates coincide.
 *  - `durationVariance` is a plain signed working-minute span
 *    (`currentDuration - baselineDuration`): positive when the current task
 *    is longer than planned, negative when shorter, zero when equal.
 *
 * Both `startVariance` and `finishVariance` are `undefined` when either the
 * baseline snapshot or the current schedule lacks the corresponding date
 * (for example a baseline captured before a task was scheduled, or a summary
 * whose baseline has no finish). `durationVariance` is always defined because
 * duration is always present in both the snapshot and the derived schedule.
 */
export interface BaselineVariance {
  taskId: TaskId
  baselineId: BaselineId
  baselineStart?: ISODateTime
  baselineFinish?: ISODateTime
  baselineDuration: WorkingMinutes
  baselineWork: WorkingMinutes
  baselineCost: number
  /** Signed working-minutes; + when current starts after baseline (slipped). */
  startVariance?: number
  /** Signed working-minutes; + when current finishes after baseline (slipped). */
  finishVariance?: number
  /** Signed minutes; + when current duration exceeds baseline. */
  durationVariance: number
}

/**
 * PROJECT-009 baseline comparison result. Projects the current `DerivedSchedule`
 * against a single baseline's immutable snapshots. Tasks without a snapshot in
 * the baseline are omitted (a baseline only reports variance for tasks it
 * captured). Pure and deterministic: the same document + schedule + baseline
 * always produces byte-identical variance bytes.
 */
export interface BaselineComparison {
  baselineId: BaselineId
  variances: Record<TaskId, BaselineVariance>
}

export interface ProjectSavePlan {
  format: 'gproj' | 'mspdi' | 'mpp'
  path?: string
  document: ProjectDocument
}
export interface ProjectFileMetadata {
  format: ProjectSavePlan['format']
  version: string
  sourceName?: string
}
export interface ImportDiagnostic {
  code: string
  severity: 'info' | 'warning' | 'error'
  message: string
  entityId?: string
}
