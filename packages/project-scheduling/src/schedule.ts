import type {
  AssignmentId,
  AssignmentSchedule,
  Calendar,
  CalendarId,
  Dependency,
  DerivedSchedule,
  ISODateTime,
  ImportDiagnostic,
  ProjectDocument,
  ResourceId,
  Task,
  TaskId,
  TaskProgressStatus,
  TaskSchedule,
  WorkingMinutes,
} from '@genoffice/project-contracts'
import { validateProjectDocument } from '@genoffice/project-engine'
import {
  CalendarBook,
  CalendarError,
  addWorkingTime,
  nextWorkingInstant,
  resolveCalendar,
  signedWorkingDuration,
  subtractWorkingTime,
  workingDuration,
  workingMinutesOf,
} from './calendar.js'
import { DependencyGraphError, buildDependencyGraph } from './graph.js'

export interface SchedulingOptions {
  projectStart?: ISODateTime
}

/**
 * PROJECT-010 deterministic resource calendar id resolution.
 *
 * Returns the resolved calendar id for a resource
 * (`resource.calendarId ?? properties.defaultCalendarId`). This is the same
 * resolution the scheduling engine uses to build `AssignmentSchedule`
 * values; exposing it keeps a single canonical boundary so tests and
 * downstream layers never re-derive (and potentially diverge from) the
 * engine's calendar choice. Pure and deterministic: the same document +
 * resourceId always resolve to the same CalendarId.
 */
export function resolveResourceCalendarId(
  document: ProjectDocument,
  resourceId: ResourceId,
): CalendarId | undefined {
  const resource = document.resources.find((item) => item.id === resourceId)
  if (!resource) return undefined
  return (resource.calendarId ?? document.properties.defaultCalendarId) as CalendarId
}

const clampPercent = (value: number): number =>
  Number.isFinite(value) ? Math.max(0, Math.min(100, value)) : 0

/**
 * PROJECT-008 leaf/milestone progress status derivation.
 *
 * Canonical precedence (no wall-clock; only `statusDate` + the scheduled
 * window + percentComplete):
 *  - percentComplete >= 100 → complete (even when finish is in the future)
 *  - 0 < percentComplete < 100 → inProgress
 *  - percentComplete == 0:
 *      • statusDate set AND not before scheduledStart → inProgress
 *        (work should have begun but is not yet reported complete)
 *      • otherwise → notStarted
 *
 * Milestones are zero-duration binary events; their [start, finish) window is
 * empty, so a 0% milestone is `notStarted` and a 100% milestone is `complete`.
 */
function deriveLeafStatus(
  percent: number,
  milestone: boolean,
  statusDate: ISODateTime | undefined,
  scheduledStart: ISODateTime | undefined,
): TaskProgressStatus {
  if (percent >= 100) return 'complete'
  if (milestone) return percent > 0 ? 'inProgress' : 'notStarted'
  if (percent > 0) return 'inProgress'
  if (statusDate && scheduledStart && !isBefore(statusDate, scheduledStart)) return 'inProgress'
  return 'notStarted'
}

interface ComputedTask {
  task: Task
  calendar: Calendar
  /** Rolled-up duration; equals task.duration for leaf tasks. */
  duration: WorkingMinutes
  earlyStart?: ISODateTime
  earlyFinish?: ISODateTime
  lateStart?: ISODateTime
  lateFinish?: ISODateTime
}

const isBefore = (a: ISODateTime, b: ISODateTime): boolean =>
  new Date(a).getTime() < new Date(b).getTime()
const earlier = (a: ISODateTime, b: ISODateTime): ISODateTime => (isBefore(a, b) ? a : b)
const later = (a: ISODateTime, b: ISODateTime): ISODateTime => (isBefore(a, b) ? b : a)
const compareIds = (a: TaskId, b: TaskId): number => (a < b ? -1 : a > b ? 1 : 0)

/**
 * Constraint semantics (canonical foundation behavior):
 * - asSoonAsPossible: default; early dates come from dependencies.
 * - startNoEarlierThan / finishNoEarlierThan: push early dates later only.
 * - mustStartOn / mustFinishOn: pin early AND late dates exactly; total slack is
 *   measured against the pinned dates, so an otherwise slack-rich task becomes
 *   critical. This is what separates MSO/MFO from SNET/FNET.
 * - startNoLaterThan / finishNoLaterThan / asLateAsPossible: scheduled dates are
 *   pulled to the late dates (bounded by the constraint for SNLT/FNLT).
 */
const LATE_SCHEDULED_CONSTRAINTS = new Set([
  'asLateAsPossible',
  'startNoLaterThan',
  'finishNoLaterThan',
])

function requiredStartFrom(
  dependency: Dependency,
  predecessor: ComputedTask,
  successorDuration: WorkingMinutes,
  calendar: Calendar,
): ISODateTime | undefined {
  const lag = workingMinutesOf(dependency.lagMinutes)
  switch (dependency.type) {
    case 'FS':
      return predecessor.earlyFinish
        ? addWorkingTime(calendar, predecessor.earlyFinish, lag)
        : undefined
    case 'SS':
      return predecessor.earlyStart
        ? addWorkingTime(calendar, predecessor.earlyStart, lag)
        : undefined
    case 'FF':
      return predecessor.earlyFinish
        ? subtractWorkingTime(
            calendar,
            addWorkingTime(calendar, predecessor.earlyFinish, lag),
            successorDuration,
          )
        : undefined
    case 'SF':
      return predecessor.earlyStart
        ? subtractWorkingTime(
            calendar,
            addWorkingTime(calendar, predecessor.earlyStart, lag),
            successorDuration,
          )
        : undefined
  }
}

function applyForwardConstraints(
  task: Task,
  calendar: Calendar,
  candidate: ISODateTime,
): ISODateTime {
  if (!task.constraintDate || !task.constraintType) return candidate
  switch (task.constraintType) {
    case 'startNoEarlierThan':
      return later(candidate, task.constraintDate)
    case 'mustStartOn':
      return task.constraintDate
    case 'finishNoEarlierThan': {
      const nominalFinish = addWorkingTime(calendar, candidate, task.duration)
      return isBefore(nominalFinish, task.constraintDate)
        ? subtractWorkingTime(calendar, task.constraintDate, task.duration)
        : candidate
    }
    case 'mustFinishOn':
      return subtractWorkingTime(calendar, task.constraintDate, task.duration)
    default:
      // asSoonAsPossible, asLateAsPossible, startNoLaterThan, finishNoLaterThan
      // have no forward-pass effect.
      return candidate
  }
}

function applyBackwardConstraints(
  task: Task,
  calendar: Calendar,
  duration: WorkingMinutes,
  lateFinish: ISODateTime,
): { lateStart: ISODateTime; lateFinish: ISODateTime } {
  let start = subtractWorkingTime(calendar, lateFinish, duration)
  let finish = lateFinish
  if (!task.constraintDate || !task.constraintType) return { lateStart: start, lateFinish: finish }
  switch (task.constraintType) {
    case 'mustStartOn':
      start = task.constraintDate
      finish = addWorkingTime(calendar, start, duration)
      break
    case 'startNoLaterThan':
      start = earlier(start, task.constraintDate)
      finish = addWorkingTime(calendar, start, duration)
      break
    case 'mustFinishOn':
      finish = task.constraintDate
      start = subtractWorkingTime(calendar, finish, duration)
      break
    case 'finishNoLaterThan':
      finish = earlier(finish, task.constraintDate)
      start = subtractWorkingTime(calendar, finish, duration)
      break
    default:
      break
  }
  return { lateStart: start, lateFinish: finish }
}

/** Latest finish bound (finish domain) that a dependency imposes on its predecessor. */
function lateFinishBoundFrom(
  dependency: Dependency,
  successor: ComputedTask,
  predecessorDuration: WorkingMinutes,
  calendar: Calendar,
): ISODateTime | undefined {
  const lag = workingMinutesOf(dependency.lagMinutes)
  switch (dependency.type) {
    case 'FS':
      return successor.lateStart
        ? subtractWorkingTime(calendar, successor.lateStart, lag)
        : undefined
    case 'FF':
      return successor.lateFinish
        ? subtractWorkingTime(calendar, successor.lateFinish, lag)
        : undefined
    case 'SS':
      return successor.lateStart
        ? addWorkingTime(
            calendar,
            subtractWorkingTime(calendar, successor.lateStart, lag),
            predecessorDuration,
          )
        : undefined
    case 'SF':
      return successor.lateFinish
        ? addWorkingTime(
            calendar,
            subtractWorkingTime(calendar, successor.lateFinish, lag),
            predecessorDuration,
          )
        : undefined
  }
}

/** Free-slack bound: how far the task can slip before this successor's early dates move. */
function freeSlackBoundFrom(
  dependency: Dependency,
  successor: ComputedTask,
  calendar: Calendar,
): { bound: ISODateTime; anchor: 'start' | 'finish' } | undefined {
  const lag = workingMinutesOf(dependency.lagMinutes)
  switch (dependency.type) {
    case 'FS':
      return successor.earlyStart
        ? { bound: subtractWorkingTime(calendar, successor.earlyStart, lag), anchor: 'finish' }
        : undefined
    case 'SS':
      return successor.earlyStart
        ? { bound: subtractWorkingTime(calendar, successor.earlyStart, lag), anchor: 'start' }
        : undefined
    case 'FF':
      return successor.earlyFinish
        ? { bound: subtractWorkingTime(calendar, successor.earlyFinish, lag), anchor: 'finish' }
        : undefined
    case 'SF':
      return successor.earlyFinish
        ? { bound: subtractWorkingTime(calendar, successor.earlyFinish, lag), anchor: 'start' }
        : undefined
  }
}

function computeSchedule(document: ProjectDocument, options: SchedulingOptions): DerivedSchedule {
  const graph = buildDependencyGraph(document)
  const book: CalendarBook = { calendars: document.calendars }
  const calendarCache = new Map<string, Calendar>()
  const calendarFor = (task: Task): Calendar => {
    const calendarId = (task.calendarId ?? document.properties.defaultCalendarId) as CalendarId
    const key = calendarId as string
    let resolved = calendarCache.get(key)
    if (!resolved) {
      resolved = resolveCalendar(book, calendarId)
      calendarCache.set(key, resolved)
    }
    return resolved
  }

  // PROJECT-010 resource scheduling inputs. Resource calendars are scheduling
  // INPUTS only — they do not move task dates in this increment (the PROJECT-006
  // task-calendar precedence is unchanged). The resolved resource calendar id
  // (`resource.calendarId ?? properties.defaultCalendarId`) is exposed on the
  // derived `AssignmentSchedule` so downstream layers can reason about resource
  // capacity without re-deriving calendar/resource resolution. Calendar
  // precedence (documented, deterministic, never renderer-driven):
  //   task calendar > resource calendar > project default
  // Task scheduling stays task-calendar-based; resource calendars are
  // independent derived inputs reserved for PROJECT-011 work/cost calculations.
  const resourceById = new Map<string, (typeof document.resources)[number]>()
  for (const resource of document.resources) {
    resourceById.set(resource.id as string, resource)
  }
  const taskResolvedCalendarId = (task: Task): CalendarId =>
    (task.calendarId ?? document.properties.defaultCalendarId) as CalendarId
  const resourceResolvedCalendarId = (resourceId: ResourceId): CalendarId => {
    const resource = resourceById.get(resourceId as string)
    return (resource?.calendarId ?? document.properties.defaultCalendarId) as CalendarId
  }

  const childrenOf = new Map<TaskId, Task[]>()
  for (const task of document.tasks) {
    if (task.parentTaskId) {
      const bucket = childrenOf.get(task.parentTaskId) ?? []
      bucket.push(task)
      childrenOf.set(task.parentTaskId, bucket)
    }
  }

  const computed = new Map<TaskId, ComputedTask>()
  for (const task of document.tasks) {
    computed.set(task.id, { task, calendar: calendarFor(task), duration: task.duration })
  }

  const projectStart = options.projectStart ?? document.properties.startDate
  // PROJECT-008: status evaluation is deterministic and uses the project
  // status date only. Wall-clock "today" never enters the scheduling engine.
  const statusDate = document.properties.statusDate

  // ---- Forward pass: early dates in combined dependency + hierarchy order ----
  for (const taskId of graph.topologicalOrder) {
    const entry = computed.get(taskId)!
    const { task, calendar } = entry
    const children = childrenOf.get(taskId) ?? []
    const isSummary = task.summary && children.length > 0

    if (isSummary) {
      const childEntries = children
        .map((child) => computed.get(child.id)!)
        .filter((child) => child.earlyStart && child.earlyFinish)
      const start = childEntries.map((child) => child.earlyStart!).reduce(earlier)
      const finish = childEntries.map((child) => child.earlyFinish!).reduce(later)
      entry.earlyStart = start
      entry.earlyFinish = finish
      entry.duration = workingDuration(calendar, start, finish)
      continue
    }

    let candidate: ISODateTime = task.start ?? projectStart
    for (const dependency of graph.predecessors.get(taskId) ?? []) {
      const predecessor = computed.get(dependency.predecessorId)!
      const required = requiredStartFrom(dependency, predecessor, task.duration, calendar)
      if (required) candidate = later(candidate, required)
    }
    candidate = applyForwardConstraints(task, calendar, candidate)
    const earlyStart = nextWorkingInstant(calendar, candidate)
    entry.earlyStart = earlyStart
    entry.earlyFinish = addWorkingTime(calendar, earlyStart, task.duration)
  }

  const leafEntries = [...computed.values()].filter(
    (entry) => !(entry.task.summary && (childrenOf.get(entry.task.id) ?? []).length > 0),
  )
  const projectFinish = leafEntries
    .map((entry) => entry.earlyFinish)
    .filter((value): value is ISODateTime => Boolean(value))
    .reduce<ISODateTime | undefined>((acc, value) => (acc ? later(acc, value) : value), undefined)

  // ---- Backward pass: late dates in reverse dependency order ----
  if (projectFinish) {
    for (const taskId of [...graph.topologicalOrder].reverse()) {
      const entry = computed.get(taskId)!
      const { task, calendar, duration } = entry
      const bounds: ISODateTime[] = []
      for (const dependency of graph.successors.get(taskId) ?? []) {
        const successor = computed.get(dependency.successorId)!
        const bound = lateFinishBoundFrom(dependency, successor, duration, calendar)
        if (bound) bounds.push(bound)
      }
      // Late dates are bounded by the project finish so that slack always
      // measures "slip without extending the project". Start-domain bounds
      // (SS/SF) alone can be looser than the project finish.
      const baseFinish = bounds.reduce(earlier, projectFinish)
      const constrained = applyBackwardConstraints(task, calendar, duration, baseFinish)
      entry.lateStart = constrained.lateStart
      entry.lateFinish = constrained.lateFinish
    }

    // Summary late dates roll up from children (deepest first), capped by any
    // dependency-imposed bounds already computed above.
    const summaries = document.tasks
      .filter((task) => task.summary && (childrenOf.get(task.id) ?? []).length > 0)
      .sort((a, b) => b.outlineLevel - a.outlineLevel || compareIds(a.id, b.id))
    for (const task of summaries) {
      const entry = computed.get(task.id)!
      const childEntries = childrenOf
        .get(task.id)!
        .map((child) => computed.get(child.id)!)
        .filter((child) => child.lateStart && child.lateFinish)
      if (!childEntries.length) continue
      const childLateFinish = childEntries.map((child) => child.lateFinish!).reduce(later)
      const childLateStart = childEntries.map((child) => child.lateStart!).reduce(earlier)
      entry.lateFinish = entry.lateFinish
        ? earlier(entry.lateFinish, childLateFinish)
        : childLateFinish
      entry.lateStart = entry.lateStart ? earlier(entry.lateStart, childLateStart) : childLateStart
    }
  }

  // ---- Assemble per-task schedules ----
  const taskSchedules: Record<string, TaskSchedule> = {}
  for (const taskId of document.tasks.map((task) => task.id).sort(compareIds)) {
    const entry = computed.get(taskId)!
    const { task, calendar, duration } = entry
    const earlyStart = entry.earlyStart
    const earlyFinish = entry.earlyFinish
    const lateStart = entry.lateStart
    const lateFinish = entry.lateFinish

    const totalSlack =
      earlyFinish && lateFinish ? signedWorkingDuration(calendar, earlyFinish, lateFinish) : 0

    let freeSlack = totalSlack
    const successors = graph.successors.get(taskId) ?? []
    if (successors.length) {
      const linkSlacks: number[] = []
      for (const dependency of successors) {
        const successor = computed.get(dependency.successorId)!
        const bound = freeSlackBoundFrom(dependency, successor, calendar)
        const anchor = bound?.anchor === 'start' ? earlyStart : earlyFinish
        if (bound && anchor) linkSlacks.push(signedWorkingDuration(calendar, anchor, bound.bound))
      }
      if (linkSlacks.length) freeSlack = Math.min(...linkSlacks)
    }

    const lateScheduled =
      task.constraintType !== undefined && LATE_SCHEDULED_CONSTRAINTS.has(task.constraintType)
    const scheduledStart = lateScheduled && lateStart ? lateStart : earlyStart
    const scheduledFinish = lateScheduled && lateFinish ? lateFinish : earlyFinish

    // ---- PROJECT-008 deadline derivation ----
    // A deadline is NOT a constraint and never moved the task above. It only
    // yields variance/missed state for downstream reporting. Variance is the
    // signed working-minute span from scheduledFinish to the deadline in the
    // task's calendar: positive when the task finishes ahead of the deadline,
    // negative when it finishes after (missed), zero when they coincide.
    let deadline: ISODateTime | undefined
    let deadlineVariance: number | undefined
    let deadlineMissed: boolean | undefined
    if (task.deadline !== undefined && scheduledFinish !== undefined) {
      deadline = task.deadline
      deadlineVariance = signedWorkingDuration(calendar, scheduledFinish, task.deadline)
      deadlineMissed = isBefore(task.deadline, scheduledFinish)
    }

    // ---- PROJECT-008 leaf/milestone progress derivation ----
    // Summary progress is rolled up in a separate pass below (it needs every
    // child's derived state). Leaves and milestones derive directly from the
    // stored percentComplete, the scheduled window, and the status date.
    const isSummaryTask = task.summary && (childrenOf.get(taskId) ?? []).length > 0
    let status: TaskProgressStatus | undefined
    let percent: number | undefined
    let physicalPercent: number | undefined
    let actualDuration: WorkingMinutes | undefined
    let remainingDuration: WorkingMinutes | undefined
    if (!isSummaryTask) {
      const stored = clampPercent(task.percentComplete)
      percent = stored
      physicalPercent = task.physicalPercentComplete
      // Milestones have zero duration: actual/remaining are both 0 regardless
      // of percent (a 100% milestone is complete but consumed no work-minutes).
      const actual =
        task.milestone || duration <= 0 ? 0 : Math.round(((duration as number) * stored) / 100)
      actualDuration = workingMinutesOf(actual)
      remainingDuration = workingMinutesOf((duration as number) - actual)
      status = deriveLeafStatus(stored, task.milestone, statusDate, scheduledStart)
    }

    taskSchedules[taskId as string] = {
      taskId,
      earlyStart,
      earlyFinish,
      lateStart,
      lateFinish,
      totalSlack,
      freeSlack,
      critical: totalSlack <= 0,
      scheduledStart,
      scheduledFinish,
      duration,
      deadline,
      deadlineVariance,
      deadlineMissed,
      status,
      percentComplete: percent,
      physicalPercentComplete: physicalPercent,
      actualDuration,
      remainingDuration,
      // PROJECT-010: the resolved calendar id used to schedule this task
      // (task.calendarId ?? properties.defaultCalendarId). Exposed so downstream
      // layers read the deterministic calendar choice without re-deriving it.
      resolvedCalendarId: taskResolvedCalendarId(task),
    }
  }

  // ---- PROJECT-008 summary progress roll-up (deepest first) ----
  // Summary progress is a duration-weighted roll-up of the subtree: it sums
  // each child's derived actual/remaining work-minutes. This is deliberately
  // NOT resource-weighted (PROJECT-011 concern) and does not copy a single
  // child's percent. A summary's stored percentComplete is never authoritative.
  const summaryTasks = document.tasks
    .filter((task) => task.summary && (childrenOf.get(task.id) ?? []).length > 0)
    .sort((a, b) => b.outlineLevel - a.outlineLevel || compareIds(a.id, b.id))
  for (const summary of summaryTasks) {
    const id = summary.id as string
    const entry = taskSchedules[id]
    if (!entry) continue
    const childEntries = (childrenOf.get(summary.id) ?? [])
      .map((child) => taskSchedules[child.id as string])
      .filter((child): child is TaskSchedule => Boolean(child))
    let actual = 0
    let remaining = 0
    let allComplete = childEntries.length > 0
    let anyInProgress = false
    let allNotStarted = childEntries.length > 0
    for (const child of childEntries) {
      actual += (child.actualDuration as number | undefined) ?? 0
      remaining += (child.remainingDuration as number | undefined) ?? 0
      if (child.status !== 'complete') allComplete = false
      if (child.status === 'inProgress') anyInProgress = true
      if (child.status !== 'notStarted') allNotStarted = false
    }
    const total = actual + remaining
    const summaryPercent = total > 0 ? Math.round((actual / total) * 100) : allComplete ? 100 : 0
    const summaryStatus: TaskProgressStatus =
      summaryPercent >= 100
        ? 'complete'
        : summaryPercent <= 0
          ? allNotStarted
            ? 'notStarted'
            : 'inProgress'
          : 'inProgress'
    void anyInProgress // retained for clarity of the precedence intent
    entry.percentComplete = summaryPercent
    entry.actualDuration = workingMinutesOf(actual)
    entry.remainingDuration = workingMinutesOf(remaining)
    entry.status = summaryStatus
    // physicalPercentComplete is intentionally undefined on summaries.
  }

  // ---- PROJECT-010 / PROJECT-011 derived assignment scheduling inputs ----
  // Build per-assignment derived schedules deterministically: keyed by
  // AssignmentId and assembled from assignments sorted by AssignmentId so the
  // same serialized ProjectDocument + options always produce byte-identical
  // schedule bytes (independent of input array order). PROJECT-010 established
  // the resolved resource scheduling inputs (calendar id, max units, type,
  // units); PROJECT-011 extends each assignment schedule with derived work and
  // cost computed from the accepted schedule.
  //
  // Canonical work/cost semantics (PROJECT-011):
  //  - Work resources: `work = task.duration × units` (WorkingMinutes, using
  //    the task's already-scheduled duration in the task's resolved calendar).
  //    `actualWork = round(work × percentComplete / 100)`; `remainingWork =
  //    work − actualWork`. Cost = `standardRateCost + overtimeCost +
  //    costPerUse` where `standardRateCost = (work / 60) × standardRate` (rate
  //    is per hour; work is in minutes). Overtime cost is 0 because the frozen
  //    Assignment contract has no overtimeWork input — this is a documented
  //    deferred limitation, not a guess.
  //  - Material resources: no work capacity. `units` is a material quantity.
  //    `cost = units × standardRate + costPerUse`. Work = 0.
  //  - Cost resources: pure cost. `cost = assignment.cost` (authoritative
  //    document input). Work = 0. `actualCost = round(cost × percentComplete /
  //    100)`; `remainingCost = cost − actualCost`.
  //
  // The status date does NOT override percentComplete for work/cost
  // calculations (PROJECT-008 precedence preserved). The resource calendar is
  // an INPUT but does not move task dates or change the work formula in this
  // increment (leveling is PROJECT-013).
  const assignmentCompare = (a: AssignmentId, b: AssignmentId): number =>
    (a as string) < (b as string) ? -1 : (a as string) > (b as string) ? 1 : 0
  const assignmentSchedules: Record<string, AssignmentSchedule> = {}
  // Per-task work/cost accumulators (leaf-level aggregation from assignments).
  // Summary tasks roll these up in a separate pass below.
  const taskWork = new Map<TaskId, number>()
  const taskActualWork = new Map<TaskId, number>()
  const taskRemainingWork = new Map<TaskId, number>()
  const taskCost = new Map<TaskId, number>()
  const taskActualCost = new Map<TaskId, number>()
  const taskRemainingCost = new Map<TaskId, number>()
  const sortedAssignments = [...document.assignments].sort((a, b) => assignmentCompare(a.id, b.id))
  for (const assignment of sortedAssignments) {
    const resource = resourceById.get(assignment.resourceId as string)
    // A valid document guarantees the resource exists (validation rejects
    // dangling references), but the schedule is still built defensively so a
    // partial/repairing document cannot crash the engine. An assignment whose
    // resource is missing is omitted from the derived map; the document's own
    // diagnostics (from validation) already report the broken reference.
    if (!resource) continue
    // The task's scheduled duration (leaf = task.duration; summary = rolled-up
    // duration). Work is computed against this accepted duration; the resource
    // calendar does not move task dates in PROJECT-011.
    const taskEntry = computed.get(assignment.taskId)
    const taskDuration = (taskEntry?.duration as number) ?? 0
    const percent = clampPercent(taskEntry?.task.percentComplete ?? 0)
    // Canonical work/cost derivation by resource kind. Work resources derive
    // work from duration × units and cost from the standard rate + cost-per-
    // use. Material resources carry a quantity in `units` and derive cost
    // from quantity × rate. Cost resources use the authoritative
    // `Assignment.cost` field. Overtime cost is 0 (deferred — the frozen
    // Assignment contract has no overtimeWork input).
    const work = resource.kind === 'work' ? Math.round(taskDuration * assignment.units) : 0
    const cost =
      resource.kind === 'work'
        ? (work / 60) * resource.standardRate + resource.costPerUse
        : resource.kind === 'material'
          ? assignment.units * resource.standardRate + resource.costPerUse
          : assignment.cost
    const actualWork = Math.round((work * percent) / 100)
    const remainingWork = work - actualWork
    const actualCost = Math.round((cost * percent) / 100)
    const remainingCost = cost - actualCost
    // Accumulate per-task (leaf) work/cost for task schedule aggregation.
    const tid = assignment.taskId
    taskWork.set(tid, (taskWork.get(tid) ?? 0) + work)
    taskActualWork.set(tid, (taskActualWork.get(tid) ?? 0) + actualWork)
    taskRemainingWork.set(tid, (taskRemainingWork.get(tid) ?? 0) + remainingWork)
    taskCost.set(tid, (taskCost.get(tid) ?? 0) + cost)
    taskActualCost.set(tid, (taskActualCost.get(tid) ?? 0) + actualCost)
    taskRemainingCost.set(tid, (taskRemainingCost.get(tid) ?? 0) + remainingCost)
    assignmentSchedules[assignment.id as string] = {
      assignmentId: assignment.id,
      taskId: assignment.taskId,
      resourceId: assignment.resourceId,
      resourceType: resource.kind,
      resolvedCalendarId: resourceResolvedCalendarId(assignment.resourceId),
      maxUnits: resource.maxUnits,
      units: assignment.units,
      work: workingMinutesOf(work),
      actualWork: workingMinutesOf(actualWork),
      remainingWork: workingMinutesOf(remainingWork),
      cost,
      actualCost,
      remainingCost,
      standardRate: resource.standardRate,
      overtimeRate: resource.overtimeRate,
      costPerUse: resource.costPerUse,
    }
  }

  // ---- PROJECT-011 leaf task work/cost aggregation ----
  // For each leaf task, set the derived work/cost from the per-task
  // accumulators (sum of assignment work/cost). A task with no assignments has
  // work 0 (no resources means no resource work) and cost 0. Summary task
  // work/cost is rolled up in the next pass.
  for (const task of document.tasks) {
    const isSummaryTask = task.summary && (childrenOf.get(task.id) ?? []).length > 0
    if (isSummaryTask) continue
    const entry = taskSchedules[task.id as string]
    if (!entry) continue
    const w = taskWork.get(task.id) ?? 0
    const aw = taskActualWork.get(task.id) ?? 0
    const rw = taskRemainingWork.get(task.id) ?? 0
    const c = taskCost.get(task.id) ?? 0
    const ac = taskActualCost.get(task.id) ?? 0
    const rc = taskRemainingCost.get(task.id) ?? 0
    entry.work = workingMinutesOf(w)
    entry.actualWork = workingMinutesOf(aw)
    entry.remainingWork = workingMinutesOf(rw)
    entry.cost = c
    entry.actualCost = ac
    entry.remainingCost = rc
  }

  // ---- PROJECT-011 summary work/cost roll-up (deepest first) ----
  // Summary work/cost is the deterministic sum of direct children's derived
  // work/cost, rolled up recursively (a child summary contributes its own
  // rolled-up value). This mirrors the PROJECT-008 summary-progress roll-up
  // pattern and is NOT resource-weighted. Nested summaries roll up correctly
  // because we process deepest-first (children before parents). A summary's
  // own assignments (if any) are NOT included — assignments belong on leaf
  // tasks, and the roll-up is from children only.
  const summaryWorkTasks = document.tasks
    .filter((task) => task.summary && (childrenOf.get(task.id) ?? []).length > 0)
    .sort((a, b) => b.outlineLevel - a.outlineLevel || compareIds(a.id, b.id))
  for (const summary of summaryWorkTasks) {
    const entry = taskSchedules[summary.id as string]
    if (!entry) continue
    const childEntries = (childrenOf.get(summary.id) ?? [])
      .map((child) => taskSchedules[child.id as string])
      .filter((child): child is TaskSchedule => Boolean(child))
    if (childEntries.length === 0) continue
    let work = 0
    let actualWork = 0
    let remainingWork = 0
    let cost = 0
    let actualCost = 0
    let remainingCost = 0
    for (const child of childEntries) {
      work += (child.work as number | undefined) ?? 0
      actualWork += (child.actualWork as number | undefined) ?? 0
      remainingWork += (child.remainingWork as number | undefined) ?? 0
      cost += child.cost ?? 0
      actualCost += child.actualCost ?? 0
      remainingCost += child.remainingCost ?? 0
    }
    entry.work = workingMinutesOf(work)
    entry.actualWork = workingMinutesOf(actualWork)
    entry.remainingWork = workingMinutesOf(remainingWork)
    entry.cost = cost
    entry.actualCost = actualCost
    entry.remainingCost = remainingCost
  }

  return {
    taskSchedules: taskSchedules as Record<TaskId, TaskSchedule>,
    projectStart,
    projectFinish,
    diagnostics: [],
    ...(Object.keys(assignmentSchedules).length > 0
      ? { assignmentSchedules: assignmentSchedules as Record<AssignmentId, AssignmentSchedule> }
      : {}),
  }
}

export function schedule(
  document: ProjectDocument,
  options: SchedulingOptions = {},
): DerivedSchedule {
  const validation = validateProjectDocument(document)
  if (!validation.accepted) {
    const diagnostics: ImportDiagnostic[] = validation.diagnostics.map((diagnostic) => ({
      code: diagnostic.code,
      severity: 'error' as const,
      message: diagnostic.message,
    }))
    return { taskSchedules: {}, diagnostics }
  }
  try {
    return computeSchedule(document, options)
  } catch (error) {
    if (error instanceof DependencyGraphError || error instanceof CalendarError) {
      return {
        taskSchedules: {},
        diagnostics: [{ code: error.code, severity: 'error', message: error.message }],
      }
    }
    throw error
  }
}
