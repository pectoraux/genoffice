import type {
  Baseline,
  BaselineComparison,
  BaselineId,
  BaselineVariance,
  Calendar,
  CalendarId,
  DerivedSchedule,
  ISODateTime,
  ProjectDocument,
  Task,
  TaskId,
  WorkingMinutes,
} from '@genoffice/project-contracts'
import { asBaselineId, asISODateTime, asTaskId } from '@genoffice/project-contracts'
import {
  CalendarBook,
  resolveCalendar,
  signedWorkingDuration,
  workingMinutesOf,
} from './calendar.js'

/**
 * PROJECT-009 baseline snapshot selection.
 *
 * `captureBaseline` projects the current derived schedule into an immutable
 * `Baseline`. The selection determines WHICH tasks are snapshotted:
 *  - `all`: every task in the document (the Microsoft Project default: the
 *    whole plan is captured).
 *  - `leaves`: only leaf + milestone tasks (summary snapshots are derived by
 *    roll-up downstream, matching how MS Project stores summary baselines).
 *  - `tasks`: an explicit TaskId list (for partial baselines).
 *
 * Defaults to `all` so a caller that omits the selection captures the entire
 * plan, which is the canonical "set a baseline" behavior.
 */
export type BaselineSelection =
  { kind: 'all' } | { kind: 'leaves' } | { kind: 'tasks'; taskIds: TaskId[] }

export interface CaptureBaselineOptions {
  selection?: BaselineSelection
  /**
   * The `capturedAt` instant. When omitted the engine defaults to
   * `ProjectProperties.statusDate` (the deterministic project status date),
   * NEVER to wall-clock `Date.now()`. This keeps baseline capture byte-
   * identical across repeated runs of the same serialized document, which is
   * the PROJECT-008/009 determinism invariant. When neither a `capturedAt`
   * override nor a project status date is set the capture is rejected: a
   * baseline cannot be captured without a deterministic instant.
   */
  capturedAt?: ISODateTime
  name?: string
}

/**
 * Builds an immutable `Baseline` from the current `DerivedSchedule`.
 *
 * Pure and deterministic: the same document + schedule + options always
 * produce the same baseline bytes. The snapshot records each selected task's
 * scheduled start, scheduled finish, duration, work, and cost. No scheduling
 * is recomputed here — the supplied `DerivedSchedule` is the authoritative
 * current state. The `capturedAt` instant comes from `options.capturedAt`,
 * falling back to `document.properties.statusDate`; it never falls back to
 * `Date.now()` (that would break determinism).
 *
 * Returns `undefined` when no deterministic `capturedAt` instant is available
 * (neither an explicit override nor a project status date) so the caller can
 * surface a clean diagnostic instead of silently inventing a timestamp.
 */
export function captureBaseline(
  document: ProjectDocument,
  schedule: DerivedSchedule,
  baselineId: BaselineId,
  options: CaptureBaselineOptions = {},
): Baseline | undefined {
  const capturedAt = options.capturedAt ?? document.properties.statusDate
  if (capturedAt === undefined) return undefined

  const selection = options.selection ?? { kind: 'all' }
  const selectedTasks = selectTasks(document, selection)
  const taskSnapshots: Baseline['taskSnapshots'] = {}
  for (const task of selectedTasks) {
    const entry = schedule.taskSchedules[task.id]
    // A task may be unschedulable (the schedule produced no entry). Omitting
    // such a task keeps the baseline free of empty snapshots; the task is
    // simply not tracked by this baseline.
    if (!entry) continue
    taskSnapshots[task.id as string] = {
      start: entry.scheduledStart,
      finish: entry.scheduledFinish,
      duration: entry.duration,
      work: task.work,
      cost: task.cost,
    }
  }
  return {
    id: baselineId,
    name: options.name ?? (baselineId as string),
    capturedAt: asISODateTime(capturedAt),
    taskSnapshots,
  }
}

function selectTasks(document: ProjectDocument, selection: BaselineSelection): Task[] {
  if (selection.kind === 'all') return document.tasks
  if (selection.kind === 'leaves') {
    const parentsWithChildren = new Set(
      document.tasks.filter((task) => task.parentTaskId).map((task) => task.parentTaskId!),
    )
    return document.tasks.filter((task) => !parentsWithChildren.has(task.id))
  }
  const wanted = new Set(selection.taskIds)
  return document.tasks.filter((task) => wanted.has(task.id))
}

/**
 * PROJECT-009 baseline comparison.
 *
 * Projects the current `DerivedSchedule` against a single baseline's immutable
 * snapshots and emits per-task `BaselineVariance`. Pure and deterministic.
 *
 * Sign convention (mirrors the Microsoft Project "Variance" table):
 *  - `startVariance` / `finishVariance` are signed working-minute spans,
 *    computed in the task's resolved calendar as
 *    `signedWorkingDuration(baseline, current)`:
 *      • positive when current is LATER than baseline (the task slipped);
 *      • negative when current is EARLIER (the task is ahead of plan);
 *      • zero when the dates coincide.
 *  - `durationVariance` is `currentDuration - baselineDuration` (plain signed
 *    working-minutes): positive when the current task is longer than planned.
 *
 * Tasks that the baseline did not snapshot are omitted from the result (a
 * baseline only reports variance for tasks it captured). Tasks that the
 * baseline snapshotted but which no longer exist in the document are also
 * omitted (the snapshot is dangling and the document is being repaired).
 *
 * `startVariance`/`finishVariance` are `undefined` when either the baseline
 * snapshot or the current schedule lacks the corresponding date. This lets a
 * caller distinguish "no variance computable" from "zero variance".
 */
export function compareBaseline(
  document: ProjectDocument,
  schedule: DerivedSchedule,
  baseline: Baseline,
): BaselineComparison {
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

  const taskById = new Map<TaskId, Task>(document.tasks.map((task) => [task.id, task]))
  const variances: Record<string, BaselineVariance> = {}
  // Deterministic iteration: snapshot keys are sorted by TaskId so repeated
  // runs over the same baseline produce byte-identical comparison bytes.
  for (const taskKey of Object.keys(baseline.taskSnapshots).sort()) {
    const snapshot = baseline.taskSnapshots[taskKey]
    const task = taskById.get(asTaskId(taskKey))
    const current = task ? schedule.taskSchedules[task.id] : undefined
    if (!task || !current) continue
    const calendar = calendarFor(task)

    const baselineStart = snapshot.start
    const baselineFinish = snapshot.finish
    const currentStart = current.scheduledStart
    const currentFinish = current.scheduledFinish

    // startVariance: signed working-minutes from baseline.start to current.start.
    // Undefined when either date is absent (no variance computable, not zero).
    const startVariance =
      baselineStart !== undefined && currentStart !== undefined
        ? signedWorkingDuration(calendar, baselineStart, currentStart)
        : undefined
    const finishVariance =
      baselineFinish !== undefined && currentFinish !== undefined
        ? signedWorkingDuration(calendar, baselineFinish, currentFinish)
        : undefined

    const baselineDuration = snapshot.duration
    const currentDuration = current.duration
    const durationVariance = (currentDuration as number) - (baselineDuration as number)

    variances[taskKey] = {
      taskId: task.id,
      baselineId: baseline.id,
      baselineStart,
      baselineFinish,
      baselineDuration,
      baselineWork: snapshot.work,
      baselineCost: snapshot.cost,
      startVariance,
      finishVariance,
      durationVariance: workingMinutesOf(durationVariance) as WorkingMinutes,
    }
  }
  return {
    baselineId: baseline.id,
    variances: variances as Record<TaskId, BaselineVariance>,
  }
}

/** Convenience: build a baseline id from a plain string at the brand boundary. */
export const baselineIdOf = (value: string): BaselineId => asBaselineId(value)
