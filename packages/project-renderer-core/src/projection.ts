/**
 * PROJECT-021 — the shared view projection.
 *
 * `projectDocumentView` joins the canonical `ProjectDocument` with the
 * authoritative `DerivedSchedule` (when the host wired a scheduler) and the
 * view state into the immutable `ProjectViewProjection` that both hosts
 * render: the visible task rows in canonical outline order (collapsed
 * subtrees hidden), the project window, and the resolved active canonical
 * view definitions.
 *
 * Architecture-lock §11 is enforced structurally: every scheduling value on
 * a row is the VERBATIM `TaskSchedule` object produced by the scheduling
 * authority (`row.schedule` is the exact object from
 * `DerivedSchedule.taskSchedules` — reference-equal, never copied,
 * never recomputed). The projection layer owns no scheduling semantics: when
 * a task has no schedule entry (no scheduler wired, or the schedule failed),
 * the row simply carries `schedule: undefined` — dates are never invented,
 * defaulted, or degraded. Task-identity/labeling fields are verbatim echoes
 * of the canonical `Task`; the only derived values are pure projections of
 * canonical structure (visibility from the collapse set, `resourceNames`
 * from the document's assignment array order).
 *
 * PROJECT-023 — selection/edit reflection: each row also carries the
 * interaction-state projection `selected`/`focused`/`editingField` — a
 * by-value echo of the view state's selection/focus/editing target joined
 * onto the row the same way `collapsed` is (pure state projection, lock
 * §11-clean: no scheduling value is involved). The edit DRAFT is live user
 * input and stays on `state.editing.draft` — the projection reflects only
 * the edit TARGET (which row, which field). A task selected but hidden by
 * collapse is NOT projected as a row (visibility rule) — its selection is
 * retained in the view state (the documented policy: hiding is not
 * deselecting) and re-projects when expanded.
 */
import type {
  DerivedSchedule,
  ImportDiagnostic,
  ProjectDocument,
  ProjectFilter,
  ProjectGroup,
  ProjectTable,
  ProjectView,
  Task,
  TaskSchedule,
  WorkingMinutes,
  ConstraintType,
  ISODateTime,
  TaskId,
  TaskType,
} from '@genoffice/project-contracts'
import type { ProjectViewState } from './state.js'
import type { EditableTaskField } from './editing.js'

/**
 * One visible task row: the projection of a canonical `Task` (identity,
 * structure, labeling) joined with its authoritative `TaskSchedule` (when
 * one exists) and the view state (collapse flag, PROJECT-023 selection /
 * focus / edit-target reflection). Rows appear in canonical outline order;
 * a task is hidden iff one of its ancestors is collapsed.
 */
export interface ProjectTaskRow {
  readonly taskId: TaskId
  readonly uid: number
  readonly wbs: string
  readonly outlineLevel: number
  readonly name: string
  readonly taskType: TaskType
  readonly summary: boolean
  readonly milestone: boolean
  readonly manualScheduled: boolean
  readonly priority: number
  /** Canonical document duration (the input `Task.duration`). */
  readonly duration: WorkingMinutes
  /** Canonical task percent-complete (the input `Task.percentComplete`). */
  readonly percentComplete: number
  readonly constraintType?: ConstraintType
  readonly constraintDate?: ISODateTime
  readonly deadline?: ISODateTime
  /**
   * Assigned resource names, projected from the document's assignment array
   * in document order (deduplicated, first occurrence kept). A pure
   * structural projection — no semantic re-derivation.
   */
  readonly resourceNames: readonly string[]
  /** Whether this row's subtree is collapsed in the view state. */
  readonly collapsed: boolean
  /** Whether this row's task is in the view state's task selection
   * (PROJECT-023 reflection — a pure echo of `state.tasks.taskIds`
   * membership; hiding by collapse does NOT deselect, so a selected-but-
   * hidden task keeps its selection in state and re-projects as a selected
   * row when expanded). */
  readonly selected: boolean
  /** Whether this row's task is the keyboard focus
   * (`state.tasks.focusId` — the moveTaskFocus target). */
  readonly focused: boolean
  /** The field of the ACTIVE cell edit targeting this row
   * (`state.editing.field`), present iff the editing target is this row's
   * task. The live draft text stays on `state.editing.draft` — the
   * projection reflects the edit TARGET only. */
  readonly editingField?: EditableTaskField
  /**
   * The VERBATIM authoritative schedule for this task — the exact
   * `TaskSchedule` object from `DerivedSchedule.taskSchedules[taskId]`
   * (reference-equal, asserted by test). `undefined` when no schedule
   * exists for the task (no scheduler wired, schedule failure, or the task
   * was not schedulable). Never recomputed, never partially copied.
   */
  readonly schedule?: TaskSchedule
}

/**
 * The projected view both hosts render. `projectStart`/`projectFinish`
 * prefer the derived schedule (the scheduling authority) and fall back to
 * the canonical properties window — never an invented value.
 */
export interface ProjectViewProjection {
  /** Visible rows in canonical outline order (collapsed subtrees hidden). */
  readonly rows: readonly ProjectTaskRow[]
  readonly projectStart?: ISODateTime
  readonly projectFinish?: ISODateTime
  /** Whether an authoritative derived schedule was joined into the rows. */
  readonly hasSchedule: boolean
  /** The schedule's diagnostics, echoed verbatim (read-only reference). */
  readonly scheduleDiagnostics: readonly ImportDiagnostic[]
  /** The resolved active canonical view definitions (when set and present). */
  readonly activeView?: ProjectView
  readonly activeTable?: ProjectTable
  readonly activeFilter?: ProjectFilter
  readonly activeGroup?: ProjectGroup
}

function parentChainOf(tasks: readonly Task[]): Map<TaskId, TaskId | undefined> {
  const parents = new Map<TaskId, TaskId | undefined>()
  for (const task of tasks) {
    if (task.parentTaskId !== undefined) parents.set(task.id, task.parentTaskId)
  }
  return parents
}

/**
 * Projects the document view. Pure: the same `(document, schedule, state)`
 * triple always produces the same projection (deterministic, no allocation
 * of new scheduling values — schedule objects are joined by reference).
 * Neither the document, the schedule, nor the state is mutated.
 */
export function projectDocumentView(
  document: ProjectDocument,
  schedule: DerivedSchedule | undefined,
  state: ProjectViewState,
): ProjectViewProjection {
  const collapsed = new Set(state.collapsed)
  const selectedTaskIds = new Set(state.tasks.taskIds)
  const focusedTaskId = state.tasks.focusId
  const editingTaskId = state.editing?.taskId
  const editingField = state.editing?.field
  const parents = parentChainOf(document.tasks)
  const resourceNames = new Map<string, Map<string, string>>()
  for (const assignment of document.assignments) {
    const byTask = resourceNames.get(assignment.taskId) ?? new Map<string, string>()
    const resource = document.resources.find((candidate) => candidate.id === assignment.resourceId)
    const name = resource?.name
    if (name !== undefined && !byTask.has(assignment.resourceId)) {
      byTask.set(assignment.resourceId, name)
    }
    resourceNames.set(assignment.taskId, byTask)
  }

  const rows: ProjectTaskRow[] = []
  for (const task of document.tasks) {
    if (isHiddenByCollapse(parents, collapsed, task)) continue
    const names = resourceNames.get(task.id)
    const taskSchedule = schedule?.taskSchedules[task.id]
    rows.push({
      taskId: task.id,
      uid: task.uid,
      wbs: task.wbs,
      outlineLevel: task.outlineLevel,
      name: task.name,
      taskType: task.taskType,
      summary: task.summary,
      milestone: task.milestone,
      manualScheduled: task.manualScheduled,
      priority: task.priority,
      duration: task.duration,
      percentComplete: task.percentComplete,
      ...(task.constraintType !== undefined ? { constraintType: task.constraintType } : {}),
      ...(task.constraintDate !== undefined ? { constraintDate: task.constraintDate } : {}),
      ...(task.deadline !== undefined ? { deadline: task.deadline } : {}),
      resourceNames: names !== undefined ? [...names.values()] : [],
      collapsed: collapsed.has(task.id),
      selected: selectedTaskIds.has(task.id),
      focused: focusedTaskId === task.id,
      ...(editingTaskId === task.id && editingField !== undefined ? { editingField } : {}),
      ...(taskSchedule !== undefined ? { schedule: taskSchedule } : {}),
    })
  }

  const activeView =
    state.activeViewId !== undefined
      ? document.views.find((view) => view.id === state.activeViewId)
      : undefined
  const activeTable =
    state.activeTableId !== undefined
      ? document.tables.find((table) => table.id === state.activeTableId)
      : undefined
  const activeFilter =
    state.activeFilterId !== undefined
      ? document.filters.find((filter) => filter.id === state.activeFilterId)
      : undefined
  const activeGroup =
    state.activeGroupId !== undefined
      ? document.groups.find((group) => group.id === state.activeGroupId)
      : undefined

  return {
    rows,
    ...(schedule?.projectStart !== undefined
      ? { projectStart: schedule.projectStart }
      : document.properties.startDate !== undefined
        ? { projectStart: document.properties.startDate }
        : {}),
    ...(schedule?.projectFinish !== undefined
      ? { projectFinish: schedule.projectFinish }
      : document.properties.finishDate !== undefined
        ? { projectFinish: document.properties.finishDate }
        : {}),
    hasSchedule: schedule !== undefined,
    scheduleDiagnostics: schedule?.diagnostics ?? [],
    ...(activeView !== undefined ? { activeView } : {}),
    ...(activeTable !== undefined ? { activeTable } : {}),
    ...(activeFilter !== undefined ? { activeFilter } : {}),
    ...(activeGroup !== undefined ? { activeGroup } : {}),
  }
}

/** A task is hidden iff any ancestor is collapsed. */
function isHiddenByCollapse(
  parents: ReadonlyMap<TaskId, TaskId | undefined>,
  collapsed: ReadonlySet<TaskId>,
  task: Task,
): boolean {
  let current: TaskId | undefined = task.parentTaskId
  let guard = 0
  while (current !== undefined && guard < 10_000) {
    guard += 1
    if (collapsed.has(current)) return true
    current = parents.get(current)
  }
  return false
}
