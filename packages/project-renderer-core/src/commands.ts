/**
 * PROJECT-021 — structural command builders.
 *
 * The shared translation layer for the outline gestures whose canonical
 * mapping must NOT be re-implemented per host: new-task identity allocation,
 * the preceding-sibling indent rule (PROJECT-007), subtree deletion, and
 * outdent eligibility. Every builder is a pure function of the canonical
 * document; every mutation leaves here as a semantic `ProjectCommand`
 * (architecture-lock §9) that the host applies through the session
 * controller (`./session.js`) — the renderer core never mutates a document
 * itself.
 *
 * Field-level edits (rename, duration, constraints, percent complete, …)
 * are 1:1 `ProjectCommand` mappings that hosts construct directly from the
 * contracts; their interaction surfaces are PROJECT-023/024 scope. The
 * builders below are the ones where two independent host implementations
 * would otherwise risk divergent semantics.
 */
import { asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type { ProjectCommand, ProjectDocument, Task, TaskId } from '@genoffice/project-contracts'
import { HierarchyError, buildTaskHierarchy, previousSiblingOf } from '@genoffice/project-engine'

/** Canonical local-id pattern for renderer-allocated tasks: `t` + a positive
 * decimal integer. Allocation scans the document's existing task ids and
 * takes `max + 1` — deterministic for the same document regardless of array
 * order. */
const RENDERER_TASK_ID_PATTERN = /^t(\d+)$/

/** Canonical creation defaults (presentation-layer choices only; every
 * derived field — wbs, outlineLevel, summary — is recomputed by the engine
 * on acceptance). Mirrors the engine's accepted fixture semantics. */
const DEFAULT_TASK_DURATION_MINUTES = 480
const DEFAULT_TASK_PRIORITY = 500
const DEFAULT_TASK_NAME = 'New Task'

/**
 * The deterministic identity for the next renderer-created task: the smallest
 * `t{n}` greater than every existing `t{n}` id (or `t1` when none), and the
 * smallest uid greater than every existing uid (uids are unique
 * interoperability identifiers — PROJECT-007). Pure: the same document
 * always yields the same identity.
 */
export function nextTaskIdentity(document: ProjectDocument): { id: TaskId; uid: number } {
  let maxId = 0
  for (const task of document.tasks) {
    const match = RENDERER_TASK_ID_PATTERN.exec(task.id)
    if (match !== null) {
      const numeric = Number.parseInt(match[1]!, 10)
      if (Number.isFinite(numeric) && numeric > maxId) maxId = numeric
    }
  }
  let maxUid = 0
  for (const task of document.tasks) {
    if (Number.isFinite(task.uid) && task.uid > maxUid) maxUid = task.uid
  }
  return { id: asTaskId(`t${maxId + 1}`), uid: maxUid + 1 }
}

/** The canonical creation defaults for a renderer-allocated task. */
export function defaultNewTask(identity: { id: TaskId; uid: number }): Task {
  return {
    id: identity.id,
    uid: identity.uid,
    wbs: '',
    outlineLevel: 1,
    name: DEFAULT_TASK_NAME,
    taskType: 'fixedDuration',
    summary: false,
    milestone: false,
    manualScheduled: false,
    autoScheduled: true,
    duration: asWorkingMinutes(DEFAULT_TASK_DURATION_MINUTES),
    priority: DEFAULT_TASK_PRIORITY,
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
  }
}

/** Where a new task is inserted: as the last child of `parentId`, as the
 * sibling after `afterTaskId`, or as the last root task when neither is
 * given. `afterTaskId` wins when both are provided (they are alternative
 * gestures); `undefined` fields are omitted so the canonical shape stays
 * clean. */
export interface TaskInsertPosition {
  readonly parentId?: TaskId
  readonly afterTaskId?: TaskId
}

/**
 * Builds the `CreateTask` command for a new task at the given position:
 * identity allocated deterministically from the document, canonical
 * creation defaults, and the parent wiring derived from the position
 * (afterTaskId → the same parent as the anchor task; parentId → the named
 * parent). The engine recomputes all derived fields and enforces every
 * semantic rule on acceptance.
 */
export function buildCreateTaskCommand(
  document: ProjectDocument,
  position: TaskInsertPosition = {},
): ProjectCommand {
  const identity = nextTaskIdentity(document)
  const task = defaultNewTask(identity)
  if (position.afterTaskId !== undefined) {
    const anchor = document.tasks.find((candidate) => candidate.id === position.afterTaskId)
    if (anchor !== undefined && anchor.parentTaskId !== undefined) {
      return { type: 'CreateTask', task: { ...task, parentTaskId: anchor.parentTaskId } }
    }
    return { type: 'CreateTask', task }
  }
  if (position.parentId !== undefined) {
    return { type: 'CreateTask', task: { ...task, parentTaskId: position.parentId } }
  }
  return { type: 'CreateTask', task }
}

/**
 * Builds the `IndentTask` command for a task: the target parent is the
 * task's immediately preceding sibling in canonical outline order — computed
 * with the ENGINE's own `previousSiblingOf` so the builder and the
 * acceptance rule can never disagree (PROJECT-007). Returns `undefined`
 * when the task does not exist, has no preceding sibling, or the document
 * hierarchy is structurally invalid (the gesture is unavailable — the host
 * surfaces a disabled state, it never invents a parent).
 */
export function buildIndentCommand(
  document: ProjectDocument,
  taskId: TaskId,
): ProjectCommand | undefined {
  try {
    const hierarchy = buildTaskHierarchy(document.tasks)
    const node = hierarchy.byId.get(taskId)
    if (node === undefined) return undefined
    const previous = previousSiblingOf(hierarchy, node)
    if (previous === undefined) return undefined
    return { type: 'IndentTask', taskId, parentTaskId: previous.task.id }
  } catch (error) {
    if (error instanceof HierarchyError) return undefined
    throw error
  }
}

/**
 * Builds the `OutdentTask` command for a task. Returns `undefined` when the
 * task does not exist or is a root task (outdenting a root is invalid).
 */
export function buildOutdentCommand(
  document: ProjectDocument,
  taskId: TaskId,
): ProjectCommand | undefined {
  const task = document.tasks.find((candidate) => candidate.id === taskId)
  if (task === undefined || task.parentTaskId === undefined) return undefined
  return { type: 'OutdentTask', taskId }
}

/**
 * Builds the `DeleteTask` commands for a task selection: each id whose
 * ANCESTOR is also selected is dropped (the ancestor's subtree deletion
 * already covers it — the Microsoft Project outline-deletion behavior), and
 * the remaining deletions are emitted in REVERSE canonical outline order
 * (deterministic; each `DeleteTask` removes a whole subtree and looks tasks
 * up by identity, so the order is a stability guarantee, not a correctness
 * dependency).
 */
export function buildDeleteSelectionCommands(
  document: ProjectDocument,
  selectedTaskIds: readonly TaskId[],
): ProjectCommand[] {
  const selected = new Set(selectedTaskIds)
  if (selected.size === 0) return []
  const parentOf = new Map<TaskId, TaskId | undefined>()
  for (const task of document.tasks) {
    if (task.parentTaskId !== undefined) parentOf.set(task.id, task.parentTaskId)
  }
  const topMost: TaskId[] = []
  for (const task of document.tasks) {
    if (!selected.has(task.id)) continue
    let ancestor = parentOf.get(task.id)
    let covered = false
    let guard = 0
    while (ancestor !== undefined && guard < 10_000) {
      guard += 1
      if (selected.has(ancestor)) {
        covered = true
        break
      }
      ancestor = parentOf.get(ancestor)
    }
    if (!covered) topMost.push(task.id)
  }
  return topMost
    .map((id) => document.tasks.findIndex((task) => task.id === id))
    .filter((index) => index >= 0)
    .sort((a, b) => b - a)
    .map((index) => ({ type: 'DeleteTask', taskId: document.tasks[index]!.id }) as ProjectCommand)
}
