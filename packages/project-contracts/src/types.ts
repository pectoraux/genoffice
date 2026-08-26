export type Brand<T, Name extends string> = T & { readonly __brand: Name }

// `ProjectCommand` is defined in `./commands.js` and references several types
// from this file. The reference is type-only (erased at runtime), so the
// mutual type-only dependency between `types.ts` and `commands.ts` does not
// create a runtime circular import.
import type { ProjectCommand } from './commands.js'

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
  // ---- PROJECT-010 resolved calendar + resource scheduling inputs ----
  /**
   * The resolved calendar id used to schedule this task. Equals
   * `task.calendarId ?? properties.defaultCalendarId` (the PROJECT-006 task-
   * calendar precedence, unchanged). Exposed so downstream layers can read the
   * deterministic calendar choice without re-deriving it. Never renderer state.
   */
  resolvedCalendarId?: CalendarId
  // ---- PROJECT-011 derived work/cost state ----
  /**
   * PROJECT-011 derived task work (WorkingMinutes). For leaf tasks: the sum
   * of derived assignment work across all assignments on this task. For
   * summary tasks: the sum of direct children's derived work (rolled up
   * recursively). A task with no assignments has work 0 (no resources
   * assigned means no resource work). This is a derived schedule value, never
   * authoritative document state.
   */
  work?: WorkingMinutes
  /**
   * PROJECT-011 derived actual work accomplished. For leaf tasks: the sum of
   * derived assignment actualWork. For summary tasks: the rolled-up sum of
   * children's actualWork. Derived from `task.percentComplete` (the
   * authoritative progress input) — the status date does NOT override
   * percentComplete for work calculations (PROJECT-008 precedence preserved).
   */
  actualWork?: WorkingMinutes
  /**
   * PROJECT-011 derived remaining work. For leaf tasks: `work − actualWork`.
   * For summary tasks: the rolled-up sum of children's remainingWork.
   */
  remainingWork?: WorkingMinutes
  /**
   * PROJECT-011 derived task cost. For leaf tasks: the sum of derived
   * assignment cost. For summary tasks: the rolled-up sum of children's cost.
   */
  cost?: number
  /** PROJECT-011 derived actual cost. Sum of assignment actualCost, rolled up for summaries. */
  actualCost?: number
  /** PROJECT-011 derived remaining cost. `cost − actualCost`, rolled up for summaries. */
  remainingCost?: number
}

/**
 * PROJECT-010 / PROJECT-011 derived assignment scheduling inputs.
 *
 * An `AssignmentSchedule` is a deterministic projection of an `Assignment`
 * paired with its resolved `Resource` scheduling inputs. PROJECT-010
 * established the resolved calendar/resource-type/units/maxUnits inputs.
 * PROJECT-011 extends it with derived work and cost values computed from the
 * accepted schedule (task duration, assignment units, resource rates, and
 * task progress). All derived values are deterministic: the same serialized
 * `ProjectDocument` + options always produce byte-identical schedule bytes.
 *
 * `resolvedCalendarId` is the resource's resolved calendar id
 * (`resource.calendarId ?? properties.defaultCalendarId`). It is independent
 * of the task's resolved calendar: per the documented PROJECT-010 calendar
 * precedence, task scheduling stays task-calendar-based (PROJECT-006 frozen),
 * and resource calendars are scheduling INPUTS only — they do not move task
 * dates in this increment.
 *
 * `maxUnits` echoes the resource's max units. For non-work resources (material
 * and cost) the engine treats this as a non-capacity value: a cost resource
 * never carries work capacity, so `maxUnits` is echoed but never used as a
 * capacity bound.
 *
 * PROJECT-011 canonical work/cost semantics:
 *
 * - **Work resources**: `work = task.duration × units` (WorkingMinutes). The
 *   task duration is the already-scheduled duration in the task's resolved
 *   calendar (the PROJECT-006 accepted schedule, unchanged). `units` is a
 *   capacity fraction (1.0 = 100%, 0.5 = 50%, 2.0 = 200%). The resource
 *   calendar is an INPUT but does not move task dates or change the work
 *   formula in this increment (leveling is PROJECT-013). `actualWork =
 *   round(work × task.percentComplete / 100)`; `remainingWork = work −
 *   actualWork`. `cost = standardRateCost + overtimeCost + costPerUse` where
 *   `standardRateCost = (work / 60) × standardRate` (rate is per hour; work
 *   is in minutes) and `overtimeCost = 0` (overtime work input is not present
 *   in the frozen Assignment contract, so overtime cost is deferred — see
 *   the PROJECT-011 spec clarifications).
 * - **Material resources**: no work capacity. `work = actualWork =
 *   remainingWork = 0`. `units` is a material quantity. `cost = units ×
 *   standardRate + costPerUse`. `actualCost = round(cost × percentComplete /
 *   100)`; `remainingCost = cost − actualCost`.
 * - **Cost resources**: pure cost, no work. `work = actualWork =
 *   remainingWork = 0`. The `Assignment.cost` field is the authoritative
 *   cost input. `actualCost = round(cost × percentComplete / 100)`;
 *   `remainingCost = cost − actualCost`.
 *
 * The resource rates (`standardRate`, `overtimeRate`, `costPerUse`) are echoed
 * so downstream layers never re-derive them. Branded `WorkingMinutes` is used
 * for all work fields; `number` is used for cost fields (cost is not a
 * working-minute concept).
 */
export interface AssignmentSchedule {
  assignmentId: AssignmentId
  taskId: TaskId
  resourceId: ResourceId
  /** Echo of `Resource.kind`. Cost/material resources are never work-capacity. */
  resourceType: 'work' | 'material' | 'cost'
  /**
   * The resource's resolved calendar id
   * (`resource.calendarId ?? properties.defaultCalendarId`). Deterministic and
   * independent of the task's resolved calendar.
   */
  resolvedCalendarId: CalendarId
  /** Echo of `Resource.maxUnits` (work-capacity input; non-capacity for cost). */
  maxUnits: number
  /** Echo of `Assignment.units`. */
  units: number
  /**
   * PROJECT-011 derived assignment work (WorkingMinutes). For work resources:
   * `task.duration × units`. For material/cost: 0 (no work capacity).
   */
  work?: WorkingMinutes
  /**
   * PROJECT-011 derived actual work accomplished. For work resources:
   * `round(work × task.percentComplete / 100)`. For material/cost: 0.
   */
  actualWork?: WorkingMinutes
  /**
   * PROJECT-011 derived remaining work. For work resources: `work −
   * actualWork`. For material/cost: 0.
   */
  remainingWork?: WorkingMinutes
  /**
   * PROJECT-011 derived assignment cost. Work resources: `standardRateCost +
   * overtimeCost + costPerUse`. Material: `units × standardRate + costPerUse`.
   * Cost: the authoritative `Assignment.cost` field.
   */
  cost?: number
  /** PROJECT-011 derived actual cost: `round(cost × percentComplete / 100)`. */
  actualCost?: number
  /** PROJECT-011 derived remaining cost: `cost − actualCost`. */
  remainingCost?: number
  /** Echo of `Resource.standardRate` (per-hour rate for work; per-unit for material). */
  standardRate?: number
  /** Echo of `Resource.overtimeRate` (per-hour rate; overtime cost deferred in PROJECT-011). */
  overtimeRate?: number
  /** Echo of `Resource.costPerUse` (flat per-assignment cost). */
  costPerUse?: number
}

export interface DerivedSchedule {
  taskSchedules: Record<TaskId, TaskSchedule>
  projectStart?: ISODateTime
  projectFinish?: ISODateTime
  diagnostics: ImportDiagnostic[]
  /**
   * PROJECT-010 derived per-assignment scheduling inputs. Keyed by `AssignmentId`
   * and built deterministically (sorted by AssignmentId) so the same serialized
   * `ProjectDocument` + options always produce byte-identical schedule bytes.
   * Optional and absent when the document carries no assignments, so existing
   * PROJECT-006..009 consumers that do not read it are unaffected.
   */
  assignmentSchedules?: Record<AssignmentId, AssignmentSchedule>
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

// ---- PROJECT-013 resource leveling ----
/**
 * PROJECT-013 canonical leveling policy options.
 *
 * The leveler is a pure deterministic function
 * `levelResources(document, options): LevelingResult`. It detects work-resource
 * over-allocation across the current derived schedule, deterministically
 * selects which eligible task to delay to resolve each conflict, and emits
 * semantic `SetTaskStart` commands that — when applied through the canonical
 * `applyProjectCommand` and re-scheduled — move whole tasks later in time so
 * no resource is asked to work above its `maxUnits` at any working instant.
 *
 * The leveler MUST NOT mutate the input document. The returned
 * `LevelingResult.proposedCommands` are the semantic output; the host (or the
 * `LevelResources` engine command) applies them and then re-schedules. The
 * scheduler remains the sole scheduling authority; leveling only proposes
 * task-date commands that flow back through the canonical pipeline.
 *
 * Policy decisions (documented in `spec/project/requirements.md`):
 *
 * - **Scope**: `taskIds` restricts leveling to the named subset (by `TaskId`).
 *   When undefined or empty, every auto-scheduled leaf task in the document is
 *   in scope. Scope is consulted at conflict resolution: a conflict between an
 *   in-scope and an out-of-scope task delays the in-scope task (the out-of-
 *   scope task is treated as immovable for this leveling pass).
 * - **Date window**: `levelingDateWindow` ({ start?, finish? }) restricts
 *   over-allocation detection to assignments whose scheduled window overlaps
 *   the window. When undefined, the entire project span is considered.
 * - **Manual tasks**: ALWAYS protected. `task.manualScheduled === true` is
 *   never delayed; the leveler treats it as immovable and, if it is the only
 *   resolvable side of a conflict, emits `LEVELING_PROTECTED_MANUAL`.
 * - **Critical tasks**: `respectCritical` (default `false`) protects critical
 *   tasks. When true, a critical task is never delayed; the leveler picks the
 *   non-critical side. When both sides are critical and protection is on, the
 *   leveler emits `LEVELING_PROTECTED_CRITICAL`.
 * - **Priority**: `respectPriority` (default `true`) uses `task.priority` to
 *   order conflict resolution. Higher priority is kept in place; lower
 *   priority is delayed first (mirrors the MS Project "higher priority =
 *   harder to move" convention).
 * - **Tie-breaking** (deterministic, locale-free): when priority is equal (or
 *   `respectPriority` is off), the task with the earlier `scheduledStart` is
 *   kept; when starts are equal, the task with the lexicographically smaller
 *   `TaskId` is kept. The leveler NEVER uses `Date.now()`, `Math.random()`,
 *   `localeCompare`, or array position as ordering identity.
 * - **Constraints**: hard constraints (`mustStartOn`, `mustFinishOn`) make a
 *   task immovable. Soft constraints are respected as floors/ceilings on the
 *   delayed start/finish:
 *     • `startNoEarlierThan`: the new start is clamped up to the SNET date.
 *     • `startNoLaterThan`: the new start must be ≤ SNLT; otherwise
 *       `LEVELING_CONSTRAINT_CONFLICT`.
 *     • `finishNoEarlierThan`: always satisfied when delaying (later finish).
 *     • `finishNoLaterThan`: the new finish must be ≤ FNLT; otherwise
 *       `LEVELING_CONSTRAINT_CONFLICT`.
 *     • `asLateAsPossible`: delaying is allowed (already at late dates);
 *       negative slack may result, observable through the re-scheduled
 *       `DerivedSchedule`.
 * - **Deadlines**: a deadline is NOT a constraint. Leveling may produce a
 *   deadline miss; the re-scheduled `DerivedSchedule` exposes
 *   `deadlineVariance`/`deadlineMissed` faithfully. `respectDeadlines`
 *   (default `false`) — when true, the leveler refuses to delay a task past
 *   its deadline and emits `LEVELING_DEADLINE_CONFLICT` instead.
 * - **Milestones**: zero-duration milestones have no work demand and are
 *   never levelable for capacity. They are skipped.
 * - **Summaries**: summary tasks are never directly delayed. Conflicts are
 *   always attributed to leaf tasks.
 * - **Splitting**: NOT supported by the frozen `Task` model (a task has a
 *   single contiguous `[start, finish]` window). Leveling moves whole tasks
 *   only. A single assignment whose `units` exceed `resource.maxUnits` cannot
 *   be resolved by moving the task and emits `LEVELING_INCOMPLETE` (splitting
 *   is deferred to PROJECT-045).
 * - **Negative slack**: leveling may produce negative slack (delaying a
 *   critical task extends the project). This is observable in the re-scheduled
 *   `DerivedSchedule`; the leveler does not clamp slack.
 * - **Identity preservation**: the leveler NEVER changes `TaskId`,
 *   `DependencyId`, `ResourceId`, `AssignmentId`, or any baseline snapshot.
 *   Baselines are immutable; only the current schedule moves.
 *
 * Determinism contract: given the same serialized `ProjectDocument` and
 * `LevelingOptions`, the leveler produces byte-identical
 * `LevelingResult.proposedCommands` and `diagnostics`. Reversed task arrays,
 * reversed assignment arrays, reversed resource arrays, and serialized
 * round-trips all produce the same output.
 */
export interface LevelingOptions {
  /** Restricts leveling to the named subset of tasks (by `TaskId`). */
  taskIds?: TaskId[]
  /** Restricts over-allocation detection to assignments overlapping this window. */
  levelingDateWindow?: { start?: ISODateTime; finish?: ISODateTime }
  /**
   * When true, critical tasks (totalSlack ≤ 0) are protected from delay.
   * Default false — critical tasks are levelable and may extend the project.
   */
  respectCritical?: boolean
  /**
   * When true, `task.priority` orders conflict resolution (higher priority is
   * kept in place). Default true.
   */
  respectPriority?: boolean
  /**
   * When true, leveling refuses to delay a task past its `deadline` and emits
   * `LEVELING_DEADLINE_CONFLICT`. Default false — deadlines are not
   * constraints and may be missed (the re-scheduled DerivedSchedule reports
   * `deadlineMissed` faithfully).
   */
  respectDeadlines?: boolean
}

/**
 * A single over-allocation detected by the leveler.
 *
 * `resourceId` is the over-allocated work resource. `assignmentIds` are the
 * conflicting assignments (all on `resourceId`). `taskIds` are the
 * corresponding tasks. `peakDemand` is the maximum combined `units` observed
 * at any working instant in the conflict window; `maxUnits` is the resource's
 * capacity. `window` is the { start, finish } interval where the
 * over-allocation was detected (assignment-scheduled endpoints, not working
 * minutes). `resolved` is true when the leveler proposed a delay that
 * eliminates this over-allocation in the working copy.
 */
export interface LevelingOverallocation {
  resourceId: ResourceId
  assignmentIds: AssignmentId[]
  taskIds: TaskId[]
  peakDemand: number
  maxUnits: number
  window: { start: ISODateTime; finish: ISODateTime }
  resolved: boolean
}

/**
 * A leveling action proposed by the leveler. Each action is a semantic
 * `SetTaskStart` command plus the reason it was proposed and the conflict it
 * resolves. The `proposedCommand` is the canonical semantic output: the host
 * (or the `LevelResources` engine command) applies it through
 * `applyProjectCommand` and then re-schedules.
 */
export interface LevelingAction {
  taskId: TaskId
  resourceId: ResourceId
  /** The original scheduled start (before this action's delay). */
  originalStart: ISODateTime
  /** The proposed new start (the kept task's scheduledFinish, advanced to the
   * next working instant in the resource's resolved calendar, clamped to any
   * SNET floor). */
  newStart: ISODateTime
  /** The semantic command the host applies to realize this delay. */
  proposedCommand: ProjectCommand
  /** The conflict that triggered this action (over-allocation signature). */
  reason: 'over-allocation'
  /** The assignment on the delayed task that participates in the conflict. */
  assignmentId: AssignmentId
}

/**
 * PROJECT-013 canonical leveling diagnostic codes. The leveler emits these
 * in `LevelingResult.diagnostics` (and the engine surfaces them through
 * `ProjectCommandResult.diagnostics`) so the host can report exactly why a
 * leveling pass did or did not fully resolve every over-allocation.
 *
 * - `LEVELING_NO_OVERALLOCATION`: no work-resource over-allocation was
 *   detected; the document is already level (info).
 * - `LEVELING_INCOMPLETE`: at least one over-allocation could not be fully
 *   resolved without splitting or reducing units (e.g. a single 200%
 *   assignment on a 100% resource). The leveler applied every resolvable
 *   delay and reports the remaining conflict (warning).
 * - `LEVELING_CONSTRAINT_CONFLICT`: delaying the chosen task would violate a
 *   hard or ceiling soft constraint (SNLT/FNLT/MFO/MSO). The leveler did not
 *   move that task and reports the conflict (warning).
 * - `LEVELING_NO_ELIGIBLE_TASK`: a conflict exists but no eligible task can be
 *   delayed (all sides are manual, summary, milestone, out-of-scope, or
 *   protected by policy). The conflict remains (warning).
 * - `LEVELING_PROTECTED_CRITICAL`: `respectCritical` is on and every side of
 *   the conflict is critical. The conflict remains (warning).
 * - `LEVELING_PROTECTED_MANUAL`: a conflict involves a manual task on the only
 *   resolvable side. Manual tasks are never delayed; the conflict remains
 *   (warning).
 * - `LEVELING_DEADLINE_CONFLICT`: `respectDeadlines` is on and delaying the
 *   chosen task would push it past its deadline. The conflict remains
 *   (warning).
 * - `LEVELING_SCOPE_EMPTY`: the `taskIds` scope filter matched no tasks
 *   (info). No leveling was performed.
 */
export type LevelingDiagnosticCode =
  | 'LEVELING_NO_OVERALLOCATION'
  | 'LEVELING_INCOMPLETE'
  | 'LEVELING_CONSTRAINT_CONFLICT'
  | 'LEVELING_NO_ELIGIBLE_TASK'
  | 'LEVELING_PROTECTED_CRITICAL'
  | 'LEVELING_PROTECTED_MANUAL'
  | 'LEVELING_DEADLINE_CONFLICT'
  | 'LEVELING_SCOPE_EMPTY'

export interface LevelingDiagnostic {
  code: LevelingDiagnosticCode
  severity: 'info' | 'warning' | 'error'
  message: string
  taskId?: TaskId
  resourceId?: ResourceId
  assignmentId?: AssignmentId
}

/**
 * PROJECT-013 leveling result. Pure and deterministic: the same serialized
 * `ProjectDocument` + `LevelingOptions` always produces byte-identical
 * `proposedCommands`, `actions`, `overallocations`, and `diagnostics`.
 *
 * The leveler does NOT mutate the input document. The host applies
 * `proposedCommands` (typically via the `LevelResources` engine command, which
 * applies them as a batch and returns a new document) and then calls
 * `schedule()` on the result. The scheduler remains the sole scheduling
 * authority.
 *
 * `overallocations` lists every conflict detected (in deterministic order:
 * sorted by `resourceId`, then by conflict window start, then by the sorted
 * `TaskId` set of the conflicting assignments). `resolved` flags whether the
 * leveler's proposed delays eliminate that conflict in the working copy.
 * `actions` lists every `SetTaskStart` proposed, in deterministic
 * application order. `affectedTaskIds` is the sorted unique set of tasks the
 * actions move.
 */
export interface LevelingResult {
  /** Semantic commands the host applies (in order) to realize the leveling. */
  proposedCommands: ProjectCommand[]
  /** Per-action detail (one entry per proposed delay). */
  actions: LevelingAction[]
  /** Every over-allocation detected, with `resolved` flags. */
  overallocations: LevelingOverallocation[]
  /** Sorted unique set of tasks the proposed actions move. */
  affectedTaskIds: TaskId[]
  /** Diagnostics surfacing incomplete / impossible / no-op leveling. */
  diagnostics: LevelingDiagnostic[]
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
