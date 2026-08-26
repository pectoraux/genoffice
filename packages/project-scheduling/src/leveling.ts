import type {
  Assignment,
  Calendar,
  CalendarId,
  DerivedSchedule,
  ISODateTime,
  LevelingAction,
  LevelingDiagnostic,
  LevelingOptions,
  LevelingOverallocation,
  LevelingResult,
  ProjectCommand,
  ProjectDocument,
  Resource,
  ResourceId,
  Task,
  TaskId,
} from '@genoffice/project-contracts'
import { asTaskId } from '@genoffice/project-contracts'
import { CalendarBook, addWorkingTime, nextWorkingInstant, resolveCalendar } from './calendar.js'
import { resolveResourceCalendarId, schedule, type SchedulingOptions } from './schedule.js'

// ===========================================================================
// PROJECT-013 deterministic resource leveling.
//
// The leveler is a pure function `levelResources(document, options)` that
// detects work-resource over-allocation across the current derived schedule,
// deterministically selects which eligible task to delay to resolve each
// conflict, and emits semantic `SetTaskStart` commands that — when applied
// through `applyProjectCommand` and re-scheduled — move whole tasks later so
// no resource is over-allocated at any working instant.
//
// Architecture:
//   levelResources(document, options) → LevelingResult
//     → proposedCommands (SetTaskStart[])
//     → applyProjectCommand (LevelResources applies the batch)
//     → schedule(documentAfterLeveling)
//     → DerivedSchedule
//
// The leveler does NOT mutate the input document. The scheduler remains the
// sole scheduling authority. See `spec/project/requirements.md` for the full
// canonical policy.
// ===========================================================================

const compareIds = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0)
const isBefore = (a: ISODateTime, b: ISODateTime): boolean =>
  new Date(a).getTime() < new Date(b).getTime()
const later = (a: ISODateTime, b: ISODateTime): ISODateTime => (isBefore(a, b) ? b : a)

/**
 * Default leveling policy. Mirrors the documented PROJECT-013 canonical
 * defaults: critical tasks ARE levelable (may extend the project), priority
 * IS respected (higher priority stays), deadlines are NOT constraints
 * (may be missed). Manual tasks are ALWAYS protected (hardcoded, not a
 * policy toggle).
 */
const DEFAULT_OPTIONS: Required<Omit<LevelingOptions, 'taskIds' | 'levelingDateWindow'>> & {
  taskIds: undefined
  levelingDateWindow: undefined
} = {
  respectCritical: false,
  respectPriority: true,
  respectDeadlines: false,
  taskIds: undefined,
  levelingDateWindow: undefined,
}

const normalizeOptions = (
  options?: LevelingOptions,
): {
  taskIds: Set<string> | undefined
  levelingDateWindow: { start?: ISODateTime; finish?: ISODateTime } | undefined
  respectCritical: boolean
  respectPriority: boolean
  respectDeadlines: boolean
} => {
  const merged = { ...DEFAULT_OPTIONS, ...options }
  const taskIds =
    merged.taskIds && merged.taskIds.length > 0
      ? new Set(merged.taskIds.map((id) => id as string))
      : undefined
  return {
    taskIds,
    levelingDateWindow: merged.levelingDateWindow,
    respectCritical: merged.respectCritical,
    respectPriority: merged.respectPriority,
    respectDeadlines: merged.respectDeadlines,
  }
}

interface AssignmentInterval {
  assignment: Assignment
  task: Task
  start: ISODateTime
  finish: ISODateTime
  units: number
}

interface Conflict {
  resourceId: ResourceId
  resource: Resource
  sides: AssignmentInterval[]
  window: { start: ISODateTime; finish: ISODateTime }
  peakDemand: number
  /** Effective (tightest) max-units that was exceeded during the conflict window. */
  maxUnits: number
}

interface DelayDecision {
  ok: boolean
  diagnostic?: LevelingDiagnostic
  action?: LevelingAction
  proposedCommand?: ProjectCommand
}

/** True if the task is a leaf (not a summary with children). */
const isLeaf = (document: ProjectDocument, task: Task): boolean => {
  if (!task.summary) return true
  return !document.tasks.some((child) => child.parentTaskId === task.id)
}

/**
 * True if the task CONTRIBUTES to work-resource demand. A task contributes
 * demand when it has a non-zero duration and is not a summary-with-children
 * (summaries roll up from children; their own assignments, if any, are not
 * double-counted). Manual tasks, auto tasks, and leaf summaries all
 * contribute demand — the resource is consumed when the task runs regardless
 * of whether the task is manually scheduled. Whether a task can be DELAYED is
 * a separate question answered by `isProtected`.
 *
 * Milestones and zero-duration tasks have an empty [start, finish) window,
 * so they contribute no work demand and are skipped.
 */
const contributesToDemand = (document: ProjectDocument, task: Task): boolean => {
  if (task.summary && !isLeaf(document, task)) return false
  if (task.milestone) return false
  if ((task.duration as number) <= 0) return false
  return true
}

/**
 * Resolves the calendar for a task (task.calendarId ?? defaultCalendarId). This
 * mirrors the scheduler's calendarFor() so the leveler reasons about the same
 * working-time windows the scheduler will apply.
 */
const taskCalendarFor = (
  document: ProjectDocument,
  book: CalendarBook,
  cache: Map<string, Calendar>,
  task: Task,
): Calendar => {
  const calendarId = (task.calendarId ?? document.properties.defaultCalendarId) as CalendarId
  const key = calendarId as string
  let resolved = cache.get(key)
  if (!resolved) {
    resolved = resolveCalendar(book, calendarId)
    cache.set(key, resolved)
  }
  return resolved
}

/**
 * Effective max-units capacity for a work resource at a given instant. When
 * an availability window covers the instant, its `units` override
 * `resource.maxUnits` (MS Project semantics: availability windows define the
 * resource's max units over time). When no window covers the instant, the
 * resource's `maxUnits` is the capacity. When `maxUnits` is 0 the resource has
 * no capacity at any time.
 *
 * Windows with `finish === undefined` are open-ended (active from `start`
 * onwards). Overlapping windows are resolved by taking the MINIMUM units
 * (the tightest capacity) so an over-allocation is never masked by a wider
 * window.
 */
const effectiveMaxUnits = (resource: Resource, timestamp: ISODateTime): number => {
  if (resource.maxUnits <= 0) return 0
  const ts = new Date(timestamp).getTime()
  let effective = resource.maxUnits
  let covered = false
  for (const slot of resource.availability) {
    const startMs = new Date(slot.start).getTime()
    const finishMs =
      slot.finish !== undefined ? new Date(slot.finish).getTime() : Number.POSITIVE_INFINITY
    if (ts >= startMs && ts < finishMs) {
      covered = true
      // Take the tightest capacity across all covering windows.
      if (slot.units < effective) effective = slot.units
    }
  }
  return covered ? effective : resource.maxUnits
}

/**
 * Resolves the resource's calendar (resource.calendarId ?? defaultCalendarId).
 * Used to compute the proposed new start for a delayed task: the task starts
 * at the next instant the RESOURCE is working (after the kept tasks finish),
 * so the resource is never asked to work outside its own calendar. The task's
 * own calendar is then re-applied by `schedule()` when it computes the
 * scheduled start from the pinned `task.start` candidate.
 */
const resourceCalendarFor = (
  document: ProjectDocument,
  book: CalendarBook,
  cache: Map<string, Calendar>,
  resource: Resource,
): Calendar => {
  const calendarId =
    resolveResourceCalendarId(document, resource.id) ?? document.properties.defaultCalendarId
  const key = calendarId as string
  let resolved = cache.get(key)
  if (!resolved) {
    resolved = resolveCalendar(book, calendarId)
    cache.set(key, resolved)
  }
  return resolved
}

/**
 * Builds the list of assignment intervals for a given work resource. Only
 * assignments on levelable tasks with valid scheduled windows are included.
 * Material and cost resources are never work-capacity and are skipped.
 */
const assignmentIntervalsForResource = (
  document: ProjectDocument,
  schedule: DerivedSchedule,
  resource: Resource,
  options: ReturnType<typeof normalizeOptions>,
): AssignmentInterval[] => {
  if (resource.kind !== 'work') return []
  const intervals: AssignmentInterval[] = []
  for (const assignment of document.assignments) {
    if (assignment.resourceId !== resource.id) continue
    const task = document.tasks.find((t) => t.id === assignment.taskId)
    if (!task) continue
    if (!contributesToDemand(document, task)) continue
    // Scope filter: if a taskIds scope is set, only assignments on those tasks
    // are DELAYABLE. Assignments on out-of-scope tasks still contribute to
    // demand (they are the immovable side of a conflict).
    const taskSchedule = schedule.taskSchedules[task.id]
    if (!taskSchedule) continue
    const start = taskSchedule.scheduledStart
    const finish = taskSchedule.scheduledFinish
    if (!start || !finish) continue
    // Date window filter: skip assignments whose scheduled window does not
    // overlap the leveling window.
    if (options.levelingDateWindow) {
      const w = options.levelingDateWindow
      if (w.start && isBefore(finish, w.start)) continue
      if (w.finish && isBefore(w.finish, start)) continue
    }
    intervals.push({ assignment, task, start, finish, units: assignment.units })
  }
  return intervals
}

/**
 * Detects over-allocation conflicts for a single work resource using a sweep
 * over assignment-interval endpoints. Returns the maximal windows where
 * combined `units` exceed the resource's effective max-units (considering
 * availability windows), with the conflicting assignments active during each
 * window.
 */
const detectConflictsForResource = (
  resource: Resource,
  intervals: AssignmentInterval[],
): Conflict[] => {
  if (intervals.length === 0 || resource.maxUnits <= 0) return []
  type Event = { time: ISODateTime; kind: 'end' | 'start'; interval: AssignmentInterval }
  const events: Event[] = []
  for (const interval of intervals) {
    events.push({ time: interval.start, kind: 'start', interval })
    events.push({ time: interval.finish, kind: 'end', interval })
  }
  // Sort: at equal times, process 'end' before 'start' so half-open intervals
  // [start, finish) that touch at a point do not falsely overlap.
  events.sort((a, b) => {
    const ta = new Date(a.time).getTime()
    const tb = new Date(b.time).getTime()
    if (ta !== tb) return ta - tb
    // 'end' < 'start' at the same instant.
    return a.kind === b.kind ? 0 : a.kind === 'end' ? -1 : 1
  })
  const conflicts: Conflict[] = []
  let active: AssignmentInterval[] = []
  let inConflict = false
  let conflictStart: ISODateTime | undefined
  let conflictSides: AssignmentInterval[] = []
  let peakDemand = 0
  let conflictMaxUnits = resource.maxUnits
  let lastTime: ISODateTime | undefined
  for (const event of events) {
    if (event.kind === 'end') {
      active = active.filter((i) => i !== event.interval)
    } else {
      active.push(event.interval)
    }
    const demand = active.reduce((sum, i) => sum + i.units, 0)
    // Effective capacity at this instant (considers availability windows).
    const capacity = effectiveMaxUnits(resource, event.time)
    if (demand > capacity) {
      if (!inConflict) {
        inConflict = true
        conflictStart = event.time
        conflictSides = [...active]
        peakDemand = demand
        conflictMaxUnits = capacity
      } else {
        if (demand > peakDemand) peakDemand = demand
        if (capacity < conflictMaxUnits) conflictMaxUnits = capacity
        for (const i of active) {
          if (!conflictSides.includes(i)) conflictSides.push(i)
        }
      }
    } else if (inConflict) {
      const window = {
        start: conflictStart!,
        finish: event.time,
      }
      conflicts.push({
        resourceId: resource.id,
        resource,
        sides: conflictSides,
        window,
        peakDemand,
        // Report the effective (tightest) capacity that was exceeded. The
        // nominal `resource.maxUnits` is still echoed on the `resource` field
        // for downstream layers.
        maxUnits: conflictMaxUnits,
      })
      inConflict = false
      conflictStart = undefined
      conflictSides = []
      peakDemand = 0
      conflictMaxUnits = resource.maxUnits
    }
    lastTime = event.time
  }
  void lastTime
  return conflicts
}

/**
 * Detects all over-allocation conflicts across the document, in deterministic
 * order: sorted by resourceId, then by conflict window start, then by the
 * sorted TaskId set of the conflicting sides.
 */
const detectAllConflicts = (
  document: ProjectDocument,
  schedule: DerivedSchedule,
  options: ReturnType<typeof normalizeOptions>,
): Conflict[] => {
  const all: Conflict[] = []
  for (const resource of document.resources) {
    if (resource.kind !== 'work') continue
    const intervals = assignmentIntervalsForResource(document, schedule, resource, options)
    const conflicts = detectConflictsForResource(resource, intervals)
    all.push(...conflicts)
  }
  all.sort((a, b) => {
    const rid = compareIds(a.resourceId as string, b.resourceId as string)
    if (rid !== 0) return rid
    const ws = new Date(a.window.start).getTime() - new Date(b.window.start).getTime()
    if (ws !== 0) return ws
    const aTasks = a.sides
      .map((s) => s.task.id as string)
      .sort(compareIds)
      .join('|')
    const bTasks = b.sides
      .map((s) => s.task.id as string)
      .sort(compareIds)
      .join('|')
    return compareIds(aTasks, bTasks)
  })
  return all
}

/**
 * Determines whether a task is "protected immovable" for this leveling pass.
 * Protected tasks contribute to demand but cannot be the delayed side.
 *
 * - Manual tasks: always protected.
 * - Summary tasks with children: always protected (never directly levelable).
 * - Milestones / zero-duration: not levelable (but no demand either).
 * - Hard-constrained tasks (MSO/MFO): always protected (cannot move).
 * - Out-of-scope tasks (when taskIds filter is set): protected.
 * - Critical tasks when respectCritical: protected.
 */
const isProtected = (
  document: ProjectDocument,
  task: Task,
  schedule: DerivedSchedule,
  options: ReturnType<typeof normalizeOptions>,
): { protected: boolean; reason?: LevelingDiagnostic['code'] } => {
  if (task.manualScheduled) return { protected: true, reason: 'LEVELING_PROTECTED_MANUAL' }
  if (task.summary && !isLeaf(document, task))
    return { protected: true, reason: 'LEVELING_NO_ELIGIBLE_TASK' }
  if (task.milestone || (task.duration as number) <= 0)
    return { protected: true, reason: 'LEVELING_NO_ELIGIBLE_TASK' }
  if (options.taskIds && !options.taskIds.has(task.id as string))
    return { protected: true, reason: 'LEVELING_NO_ELIGIBLE_TASK' }
  if (task.constraintType === 'mustStartOn' || task.constraintType === 'mustFinishOn')
    return { protected: true, reason: 'LEVELING_CONSTRAINT_CONFLICT' }
  if (options.respectCritical) {
    const ts = schedule.taskSchedules[task.id]
    if (ts?.critical) return { protected: true, reason: 'LEVELING_PROTECTED_CRITICAL' }
  }
  return { protected: false }
}

/**
 * "Keep score" comparator for "which side to keep vs delay". The KEPT side is
 * the one that is:
 *  - HIGHER priority (more important stays), when `respectPriority` is on;
 *  - EARLIER scheduledStart (the task that started first stays), as a
 *    deterministic tie-break that matches the MS Project convention;
 *  - SMALLER TaskId (lexicographic, locale-free), as the final deterministic
 *    tie-break.
 *
 * The side with the SMALLEST keep-score is KEPT; the side with the LARGEST
 * keep-score is DELAYED. So we sort ascending and pick the LAST element as
 * the side to delay.
 *
 * The leveler NEVER uses `localeCompare`, `Date.now()`, `Math.random()`, or
 * array position as identity — only `priority`, `scheduledStart` (from the
 * derived schedule), and `TaskId` (locale-free code-unit comparison).
 */
interface KeepScore {
  // Negative priority so higher priority → more negative → smaller → kept.
  priorityNeg: number
  // Positive epoch so earlier start → smaller epoch → smaller → kept.
  startEpoch: number
  // TaskId string, compared locale-free (code-unit).
  taskId: string
}

const keepScoreOf = (
  side: AssignmentInterval,
  options: ReturnType<typeof normalizeOptions>,
): KeepScore => ({
  priorityNeg: options.respectPriority ? -side.task.priority : 0,
  startEpoch: new Date(side.start).getTime(),
  taskId: side.task.id as string,
})

const compareKeepScore = (a: KeepScore, b: KeepScore): number => {
  if (a.priorityNeg !== b.priorityNeg) return a.priorityNeg < b.priorityNeg ? -1 : 1
  if (a.startEpoch !== b.startEpoch) return a.startEpoch < b.startEpoch ? -1 : 1
  return compareIds(a.taskId, b.taskId)
}

/**
 * Computes the proposed new start for the delayed side: the latest
 * scheduledFinish of the OTHER sides, advanced to the next working instant in
 * the RESOURCE's resolved calendar. This ensures the delayed task starts at
 * the next instant the resource is working (after every kept task finishes),
 * removing the resource overlap. The task's own calendar is then re-applied
 * by `schedule()` when it computes the scheduled start from the pinned
 * `task.start` candidate (the scheduler's `nextWorkingInstant(taskCalendar, …)`
 * may advance further if the task calendar has no working time at the
 * resource's next-working instant).
 */
const computeNewStart = (
  document: ProjectDocument,
  book: CalendarBook,
  cache: Map<string, Calendar>,
  conflict: Conflict,
  delayedSide: AssignmentInterval,
): ISODateTime => {
  const otherFinishes = conflict.sides.filter((s) => s !== delayedSide).map((s) => s.finish)
  const latestFinish = otherFinishes.reduce<ISODateTime | undefined>(
    (acc, f) => (acc ? later(acc, f) : f),
    undefined,
  )
  const resourceCal = resourceCalendarFor(document, book, cache, conflict.resource)
  const base = latestFinish ?? delayedSide.start
  return nextWorkingInstant(resourceCal, base)
}

/**
 * Validates the proposed new start against the delayed task's constraints.
 * Returns the (possibly clamped) new start, or a diagnostic if the delay is
 * impossible.
 */
const validateConstraintsForDelay = (
  document: ProjectDocument,
  book: CalendarBook,
  cache: Map<string, Calendar>,
  task: Task,
  newStart: ISODateTime,
  options: ReturnType<typeof normalizeOptions>,
): { ok: true; newStart: ISODateTime } | { ok: false; diagnostic: LevelingDiagnostic } => {
  const taskCal = taskCalendarFor(document, book, cache, task)
  const diag = (
    code: LevelingDiagnostic['code'],
    message: string,
    extra: Partial<LevelingDiagnostic> = {},
  ): LevelingDiagnostic => ({
    code,
    severity: 'warning',
    message,
    taskId: task.id,
    ...extra,
  })
  if (task.constraintType && task.constraintDate) {
    switch (task.constraintType) {
      case 'mustStartOn':
      case 'mustFinishOn':
        return {
          ok: false,
          diagnostic: diag(
            'LEVELING_CONSTRAINT_CONFLICT',
            `Task ${task.id} has hard constraint ${task.constraintType}; cannot delay`,
          ),
        }
      case 'startNoEarlierThan': {
        // Clamp the new start up to the SNET date.
        if (isBefore(newStart, task.constraintDate)) {
          return { ok: true, newStart: nextWorkingInstant(taskCal, task.constraintDate) }
        }
        break
      }
      case 'startNoLaterThan': {
        if (isBefore(task.constraintDate, newStart)) {
          return {
            ok: false,
            diagnostic: diag(
              'LEVELING_CONSTRAINT_CONFLICT',
              `Task ${task.id} SNLT ${task.constraintDate} would be violated by start ${newStart}`,
            ),
          }
        }
        break
      }
      case 'finishNoLaterThan': {
        const newFinish = addWorkingTime(taskCal, newStart, task.duration)
        if (isBefore(task.constraintDate, newFinish)) {
          return {
            ok: false,
            diagnostic: diag(
              'LEVELING_CONSTRAINT_CONFLICT',
              `Task ${task.id} FNLT ${task.constraintDate} would be violated by finish ${newFinish}`,
            ),
          }
        }
        break
      }
      case 'finishNoEarlierThan':
        // Delaying always makes finish later, so FNET is always satisfied.
        break
      default:
        // asSoonAsPossible, asLateAsPossible: no ceiling.
        break
    }
  }
  // Deadline protection (only when respectDeadlines is on).
  if (options.respectDeadlines && task.deadline) {
    const newFinish = addWorkingTime(taskCal, newStart, task.duration)
    if (isBefore(task.deadline, newFinish)) {
      return {
        ok: false,
        diagnostic: diag(
          'LEVELING_DEADLINE_CONFLICT',
          `Task ${task.id} deadline ${task.deadline} would be missed by finish ${newFinish}`,
        ),
      }
    }
  }
  return { ok: true, newStart }
}

/**
 * Decides which side of a conflict to delay and computes the proposed new
 * start. Returns a DelayDecision with either an action + proposedCommand or a
 * diagnostic explaining why the conflict cannot be resolved.
 */
const decideDelay = (
  document: ProjectDocument,
  book: CalendarBook,
  cache: Map<string, Calendar>,
  conflict: Conflict,
  schedule: DerivedSchedule,
  options: ReturnType<typeof normalizeOptions>,
): DelayDecision => {
  // Single-side conflict: the assignment's units alone exceed the effective
  // capacity. Delaying the task does not help (it still over-allocates when it
  // runs). This is the "200% assignment on a 100% resource" case — splitting
  // or reducing units is required (splitting is deferred to PROJECT-045).
  if (conflict.sides.length === 1) {
    const side = conflict.sides[0]
    return {
      ok: false,
      diagnostic: {
        code: 'LEVELING_INCOMPLETE',
        severity: 'warning',
        message: `Assignment ${side.assignment.id} on resource ${conflict.resourceId} demands ${side.units} units (max ${conflict.maxUnits}); cannot resolve without splitting or reducing units`,
        taskId: side.task.id,
        resourceId: conflict.resourceId,
        assignmentId: side.assignment.id,
      },
    }
  }
  // Partition sides into delayable vs protected.
  const candidates: { side: AssignmentInterval; score: KeepScore }[] = []
  const protectedReasons: LevelingDiagnostic['code'][] = []
  for (const side of conflict.sides) {
    const prot = isProtected(document, side.task, schedule, options)
    if (prot.protected) {
      if (prot.reason) protectedReasons.push(prot.reason)
      continue
    }
    candidates.push({ side, score: keepScoreOf(side, options) })
  }
  if (candidates.length === 0) {
    // No eligible side. Pick the most specific reason.
    const reason =
      protectedReasons.find((r) => r === 'LEVELING_PROTECTED_CRITICAL') ??
      protectedReasons.find((r) => r === 'LEVELING_PROTECTED_MANUAL') ??
      protectedReasons.find((r) => r === 'LEVELING_CONSTRAINT_CONFLICT') ??
      'LEVELING_NO_ELIGIBLE_TASK'
    const sampleTask = conflict.sides[0].task
    return {
      ok: false,
      diagnostic: {
        code: reason,
        severity: 'warning',
        message: `Conflict on ${conflict.resourceId} has no eligible task to delay (${conflict.sides
          .map((s) => s.task.id as string)
          .sort(compareIds)
          .join(', ')})`,
        taskId: sampleTask.id,
        resourceId: conflict.resourceId,
      },
    }
  }
  // The side to DELAY has the LARGEST keep-score (lowest priority, latest
  // start, largest TaskId). Sort ascending and pick the LAST.
  candidates.sort((a, b) => compareKeepScore(a.score, b.score))
  const delayed = candidates[candidates.length - 1]
  const delayedSide = delayed.side
  // Compute the new start.
  let newStart = computeNewStart(document, book, cache, conflict, delayedSide)
  // Validate against the delayed task's constraints.
  const validated = validateConstraintsForDelay(
    document,
    book,
    cache,
    delayedSide.task,
    newStart,
    options,
  )
  if (!validated.ok) {
    return { ok: false, diagnostic: validated.diagnostic }
  }
  newStart = validated.newStart
  const proposedCommand: ProjectCommand = {
    type: 'SetTaskStart',
    taskId: delayedSide.task.id,
    start: newStart,
  }
  const action: LevelingAction = {
    taskId: delayedSide.task.id,
    resourceId: conflict.resourceId,
    originalStart: delayedSide.start,
    newStart,
    proposedCommand,
    reason: 'over-allocation',
    assignmentId: delayedSide.assignment.id,
  }
  return { ok: true, action, proposedCommand }
}

/**
 * Applies a SetTaskStart to the working copy (immutable). The working copy is
 * internal to the leveler; the proposed command is what the host applies via
 * the canonical applyProjectCommand path.
 */
const applyDelayToWorkingCopy = (
  document: ProjectDocument,
  taskId: TaskId,
  newStart: ISODateTime,
): ProjectDocument => ({
  ...document,
  tasks: document.tasks.map((t) => (t.id === taskId ? { ...t, start: newStart } : t)),
})

/** Signature for deduplication of over-allocations across iterations. */
const conflictSignature = (conflict: Conflict): string => {
  const tasks = conflict.sides
    .map((s) => s.task.id as string)
    .sort(compareIds)
    .join(',')
  const assignments = conflict.sides
    .map((s) => s.assignment.id as string)
    .sort(compareIds)
    .join(',')
  return `${conflict.resourceId as string}|${tasks}|${assignments}`
}

const MAX_ITERATIONS = 256

/**
 * PROJECT-013 canonical leveling entry point. Pure and deterministic: the
 * same serialized `ProjectDocument` + `LevelingOptions` always produces
 * byte-identical `LevelingResult`.
 *
 * The leveler:
 *  1. Validates the document (returns empty result + diagnostic on failure).
 *  2. Schedules the current state to get the derived schedule.
 *  3. Detects over-allocation conflicts per work resource.
 *  4. Iteratively: pick the first conflict (deterministic order), decide
 *     which eligible side to delay, apply the delay to the working copy,
 *     re-schedule, re-detect. Repeat until no conflicts or no eligible sides.
 *  5. Returns the accumulated proposedCommands, actions, overallocations
 *     (with resolved flags), affectedTaskIds, and diagnostics.
 */
export function levelResources(
  document: ProjectDocument,
  options?: LevelingOptions,
  schedulingOptions?: SchedulingOptions,
): LevelingResult {
  const opts = normalizeOptions(options)
  const diagnostics: LevelingDiagnostic[] = []
  const proposedCommands: ProjectCommand[] = []
  const actions: LevelingAction[] = []
  const affectedTaskIds = new Set<string>()
  // Track every conflict signature we have already resolved (so we can flag
  // them resolved:true) and every signature we have already reported (so the
  // final over-allocations list has no duplicates).
  const resolvedSignatures = new Set<string>()
  const reportedSignatures = new Map<string, LevelingOverallocation>()
  // The working copy carries task.start overrides for proposed delays. It
  // starts as the input document (the leveler does NOT mutate the input).
  let workingDocument: ProjectDocument = document
  // Build the calendar book + cache once (calendars don't change between
  // iterations; only task.start values do).
  const book: CalendarBook = { calendars: document.calendars }
  const calendarCache = new Map<string, Calendar>()

  // Scope-empty fast path: if a taskIds filter is set and matches NO task
  // in the document, leveling is a no-op. (When the filter matches at least
  // one task but none can be delayed, that's LEVELING_NO_ELIGIBLE_TASK below.)
  if (opts.taskIds) {
    const documentTaskIds = new Set(document.tasks.map((t) => t.id as string))
    let anyInScope = false
    for (const id of opts.taskIds) {
      if (documentTaskIds.has(id)) {
        anyInScope = true
        break
      }
    }
    if (!anyInScope) {
      return {
        proposedCommands: [],
        actions: [],
        overallocations: [],
        affectedTaskIds: [],
        diagnostics: [
          {
            code: 'LEVELING_SCOPE_EMPTY',
            severity: 'info',
            message: 'Leveling scope matches no task in the document; no tasks to level',
          },
        ],
      }
    }
  }

  let unresolvedRemaining = false
  let lastDiagnostic: LevelingDiagnostic | undefined
  for (let iteration = 0; iteration < MAX_ITERATIONS; iteration += 1) {
    const result = schedule(workingDocument, schedulingOptions ?? {})
    if (result.diagnostics.length > 0) {
      // The working copy failed to schedule (cycle, calendar error, etc.).
      // Stop leveling and report.
      diagnostics.push({
        code: 'LEVELING_INCOMPLETE',
        severity: 'warning',
        message: `Leveling stopped: schedule produced ${result.diagnostics.length} diagnostic(s)`,
      })
      unresolvedRemaining = true
      break
    }
    const conflicts = detectAllConflicts(workingDocument, result, opts)
    if (conflicts.length === 0) break
    // Pick the first conflict (already sorted deterministically).
    const conflict = conflicts[0]
    const sig = conflictSignature(conflict)
    const decision = decideDelay(workingDocument, book, calendarCache, conflict, result, opts)
    if (!decision.ok) {
      // Record this conflict as unresolved and stop trying to resolve it
      // (further iterations would pick the same conflict). Record the
      // diagnostic. Other conflicts may also remain; we report them as
      // unresolved in the final pass.
      if (!reportedSignatures.has(sig)) {
        reportedSignatures.set(sig, {
          resourceId: conflict.resourceId,
          assignmentIds: conflict.sides
            .map((s) => s.assignment.id)
            .sort((a, b) => compareIds(a as string, b as string)),
          taskIds: conflict.sides
            .map((s) => s.task.id)
            .sort((a, b) => compareIds(a as string, b as string)),
          peakDemand: conflict.peakDemand,
          maxUnits: conflict.maxUnits,
          window: conflict.window,
          resolved: false,
        })
      }
      if (decision.diagnostic) {
        // Deduplicate diagnostics by code+taskId+resourceId so the same
        // impossible conflict does not spam the diagnostic list across
        // iterations (it can't, since we break here, but be defensive).
        const d = decision.diagnostic
        const dupsig = `${d.code}|${d.taskId ?? ''}|${d.resourceId ?? ''}`
        if (
          !diagnostics.some(
            (existing) =>
              `${existing.code}|${existing.taskId ?? ''}|${existing.resourceId ?? ''}` === dupsig,
          )
        ) {
          diagnostics.push(d)
        }
        lastDiagnostic = d
      }
      unresolvedRemaining = true
      break
    }
    // Apply the delay to the working copy.
    const { action, proposedCommand } = decision
    proposedCommands.push(proposedCommand!)
    actions.push(action!)
    affectedTaskIds.add(action!.taskId as string)
    // Mark this conflict signature as resolved (it will be eliminated in the
    // next iteration's schedule).
    resolvedSignatures.add(sig)
    // Record the over-allocation as resolved:true (if not already reported).
    if (!reportedSignatures.has(sig)) {
      reportedSignatures.set(sig, {
        resourceId: conflict.resourceId,
        assignmentIds: conflict.sides
          .map((s) => s.assignment.id)
          .sort((a, b) => compareIds(a as string, b as string)),
        taskIds: conflict.sides
          .map((s) => s.task.id)
          .sort((a, b) => compareIds(a as string, b as string)),
        peakDemand: conflict.peakDemand,
        maxUnits: conflict.maxUnits,
        window: conflict.window,
        resolved: true,
      })
    }
    workingDocument = applyDelayToWorkingCopy(workingDocument, action!.taskId, action!.newStart)
    // Continue to the next iteration: re-schedule and re-detect.
  }

  // Final pass: detect any remaining conflicts in the final working copy.
  const finalResult = schedule(workingDocument, schedulingOptions ?? {})
  if (finalResult.diagnostics.length === 0) {
    const remaining = detectAllConflicts(workingDocument, finalResult, opts)
    for (const conflict of remaining) {
      const sig = conflictSignature(conflict)
      if (reportedSignatures.has(sig)) {
        // Already reported; if it's still present, it's unresolved.
        const existing = reportedSignatures.get(sig)!
        existing.resolved = false
        continue
      }
      reportedSignatures.set(sig, {
        resourceId: conflict.resourceId,
        assignmentIds: conflict.sides
          .map((s) => s.assignment.id)
          .sort((a, b) => compareIds(a as string, b as string)),
        taskIds: conflict.sides
          .map((s) => s.task.id)
          .sort((a, b) => compareIds(a as string, b as string)),
        peakDemand: conflict.peakDemand,
        maxUnits: conflict.maxUnits,
        window: conflict.window,
        resolved: false,
      })
      unresolvedRemaining = true
    }
  }

  // Assemble the final over-allocations list in deterministic order.
  const overallocations: LevelingOverallocation[] = [...reportedSignatures.values()].sort(
    (a, b) => {
      const rid = compareIds(a.resourceId as string, b.resourceId as string)
      if (rid !== 0) return rid
      const ws = new Date(a.window.start).getTime() - new Date(b.window.start).getTime()
      if (ws !== 0) return ws
      return compareIds(a.taskIds.join('|'), b.taskIds.join('|'))
    },
  )

  // Diagnostics summary.
  if (proposedCommands.length === 0 && overallocations.length === 0) {
    diagnostics.unshift({
      code: 'LEVELING_NO_OVERALLOCATION',
      severity: 'info',
      message: 'No work-resource over-allocation detected',
    })
  } else if (unresolvedRemaining) {
    // Already pushed specific diagnostics above; ensure a summary exists.
    if (!diagnostics.some((d) => d.code === 'LEVELING_INCOMPLETE')) {
      diagnostics.push({
        code: 'LEVELING_INCOMPLETE',
        severity: 'warning',
        message: `Leveling incomplete: ${overallocations.filter((o) => !o.resolved).length} over-allocation(s) remain`,
      })
    }
  }
  void lastDiagnostic

  const affectedSorted = [...affectedTaskIds].sort(compareIds).map((id) => asTaskId(id))
  // Deduplicate proposedCommands: a task may be delayed multiple times across
  // iterations (each delay supersedes the previous). Keep only the LAST
  // SetTaskStart per task — this is the final proposed start the host applies.
  // The full audit trail (every delay decision) remains in `actions`.
  // The deduped list is sorted by TaskId (locale-free, deterministic) so the
  // output is independent of the iteration order in which tasks were first
  // delayed — a reversed input array produces the same proposedCommands bytes.
  const lastCommandPerTask = new Map<string, ProjectCommand>()
  for (const cmd of proposedCommands) {
    if (cmd.type === 'SetTaskStart') {
      lastCommandPerTask.set(cmd.taskId as string, cmd)
    }
  }
  const dedupedCommands: ProjectCommand[] = [...lastCommandPerTask.values()].sort((a, b) => {
    // Both come from lastCommandPerTask (only SetTaskStart entries).
    const aId = (a as { taskId: TaskId }).taskId as string
    const bId = (b as { taskId: TaskId }).taskId as string
    return compareIds(aId, bId)
  })
  return {
    proposedCommands: dedupedCommands,
    actions,
    overallocations,
    affectedTaskIds: affectedSorted,
    diagnostics: diagnostics.sort((a, b) => {
      const c = compareIds(a.code, b.code)
      if (c !== 0) return c
      const at = (a.taskId as string) ?? ''
      const bt = (b.taskId as string) ?? ''
      return compareIds(at, bt)
    }),
  }
}
