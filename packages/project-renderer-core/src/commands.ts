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
 *
 * PROJECT-024 adds the dependency builders: deterministic dependency-id
 * allocation, the creation defaults (FS, zero lag — the Microsoft Project
 * link defaults), and the gesture guards for link creation (unknown or
 * self-referencing endpoints are disabled gestures, never invented
 * commands). The engine remains the single validation authority — a builder
 * only refuses to build what the engine would deterministically reject.
 */
import { asDependencyId, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type {
  DependencyId,
  DependencyType,
  ProjectCommand,
  ProjectDocument,
  Task,
  TaskId,
} from '@genoffice/project-contracts'
import { HierarchyError, buildTaskHierarchy, previousSiblingOf } from '@genoffice/project-engine'

/** Canonical local-id pattern for renderer-allocated tasks: `t` + a positive
 * decimal integer. Allocation scans the document's existing task ids and
 * takes `max + 1` — deterministic for the same document regardless of array
 * order. */
const RENDERER_TASK_ID_PATTERN = /^t(\d+)$/

/** Canonical local-id pattern for renderer-allocated dependencies: `d` + a
 * positive decimal integer. Allocation scans the document's existing
 * dependency ids and takes `max + 1` — deterministic for the same document.
 * Imported ids that do not match the pattern (e.g. the MSPDI importer's
 * `d-{succ}-{pred}-{type}` ids) are ignored by the scan, so allocation never
 * collides with them. */
const RENDERER_DEPENDENCY_ID_PATTERN = /^d(\d+)$/

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

/**
 * Where a renderer-created task is appended. The canonical `CreateTask`
 * command (PROJECT-007) ALWAYS places a new task as the LAST child of its
 * parent — or as the last root task: the frozen command union has NO
 * insertion index/position field, so "insert immediately after a specific
 * sibling" is NOT expressible. This type deliberately models exactly the two
 * positions the engine can execute (a discriminated union, so the API cannot
 * promise a row-position insert it cannot deliver); the renderer core does
 * not invent a second insertion model (lock §9/§11).
 */
export type TaskInsertPosition =
  /** Append as the LAST root task (after every existing root). */
  | { readonly kind: 'lastRoot' }
  /** Append as the LAST child of `parentId` — after every existing sibling
   * of that parent, never between two existing siblings. The engine is the
   * validation authority: a nonexistent `parentId` surfaces as the engine's
   * `MISSING_PARENT` rejection through the session, not a builder error. */
  | { readonly kind: 'lastChildOf'; readonly parentId: TaskId }

/**
 * Builds the `CreateTask` command for a new task at the given position:
 * identity allocated deterministically from the document, canonical
 * creation defaults, and the parent wiring from the explicit position.
 * The engine recomputes all derived fields and enforces every semantic
 * rule on acceptance.
 */
export function buildCreateTaskCommand(
  document: ProjectDocument,
  position: TaskInsertPosition = { kind: 'lastRoot' },
): ProjectCommand {
  const identity = nextTaskIdentity(document)
  const task = defaultNewTask(identity)
  if (position.kind === 'lastChildOf') {
    return { type: 'CreateTask', task: { ...task, parentTaskId: position.parentId } }
  }
  return { type: 'CreateTask', task }
}

/**
 * Builds the `CreateTask` command for the outline gesture "create a new
 * task in this row's sibling group": the new task joins the ANCHOR's
 * sibling group — the same parent as the anchor (the root level when the
 * anchor is a root task). The executable position is append-as-LAST-member
 * of that group: the new task lands after the anchor's LAST sibling, NOT
 * immediately after the anchor — the frozen `CreateTask` command cannot
 * express a row-position insert between existing siblings, and the
 * renderer core never simulates one (it computes the position ONCE here so
 * two hosts cannot diverge over the anchor→parent mapping).
 *
 * Returns `undefined` when the anchor does not exist (a disabled gesture —
 * the same contract as `buildIndentCommand`/`buildOutdentCommand`; the
 * host never invents a position).
 */
export function buildCreateTaskInSiblingGroupCommand(
  document: ProjectDocument,
  anchorTaskId: TaskId,
): ProjectCommand | undefined {
  const anchor = document.tasks.find((candidate) => candidate.id === anchorTaskId)
  if (anchor === undefined) return undefined
  return buildCreateTaskCommand(
    document,
    anchor.parentTaskId !== undefined
      ? { kind: 'lastChildOf', parentId: anchor.parentTaskId }
      : { kind: 'lastRoot' },
  )
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

// ===========================================================================
// PROJECT-024 — dependency command builders
// ===========================================================================

/**
 * The deterministic identity for the next renderer-created dependency: the
 * smallest `d{n}` greater than every existing `d{n}` id (or `d1` when none).
 * Pure: the same document always yields the same identity. Ids that do not
 * match the `d{n}` pattern (imported links) are ignored — allocation never
 * collides with them and never renumbers existing links.
 */
export function nextDependencyIdentity(document: ProjectDocument): DependencyId {
  let maxId = 0
  for (const dependency of document.dependencies) {
    const match = RENDERER_DEPENDENCY_ID_PATTERN.exec(dependency.id)
    if (match !== null) {
      const numeric = Number.parseInt(match[1]!, 10)
      if (Number.isFinite(numeric) && numeric > maxId) maxId = numeric
    }
  }
  return asDependencyId(`d${maxId + 1}`)
}

/** The canonical creation defaults for a renderer-created link (presentation
 * defaults only — the Microsoft Project link defaults): FS with zero lag. */
export const DEFAULT_DEPENDENCY_TYPE: DependencyType = 'FS'
export const DEFAULT_DEPENDENCY_LAG_MINUTES = 0

/** The optional creation parameters for `buildAddDependencyCommand`. */
export interface AddDependencyOptions {
  /** The relationship type (default `FS`). */
  readonly type?: DependencyType
  /** The lag in integer working minutes (default `0`; negative = lead). */
  readonly lagMinutes?: number
}

/**
 * Builds the `AddDependency` command for the link-creation gesture from one
 * task to another: identity allocated deterministically from the document,
 * canonical defaults (FS, zero lag), and the endpoint wiring. The engine
 * enforces every semantic rule on acceptance — cycles
 * (`DEPENDENCY_CYCLE`), duplicate links (`DUPLICATE_DEPENDENCY_LINK`), the
 * summary↔own-descendant rule (`SUMMARY_DEPENDENCY`), and lag integrality
 * (`INVALID_LAG`) surface through the session as engine diagnostics.
 *
 * Returns `undefined` for the gestures the engine would ALWAYS reject on
 * endpoint structure alone (the disabled-gesture contract of
 * `buildIndentCommand`/`buildOutdentCommand`): an unknown predecessor or
 * successor, or a self-referencing link (predecessor === successor). The
 * host surfaces a disabled gesture; it never invents a command.
 */
export function buildAddDependencyCommand(
  document: ProjectDocument,
  predecessorId: TaskId,
  successorId: TaskId,
  options: AddDependencyOptions = {},
): ProjectCommand | undefined {
  const predecessor = document.tasks.some((task) => task.id === predecessorId)
  const successor = document.tasks.some((task) => task.id === successorId)
  if (!predecessor || !successor) return undefined
  if (predecessorId === successorId) return undefined
  return {
    type: 'AddDependency',
    dependency: {
      id: nextDependencyIdentity(document),
      predecessorId,
      successorId,
      type: options.type ?? DEFAULT_DEPENDENCY_TYPE,
      lagMinutes: options.lagMinutes ?? DEFAULT_DEPENDENCY_LAG_MINUTES,
    },
  }
}

/**
 * Builds the `RemoveDependency` commands for a dependency selection: one
 * command per EXISTING selected id, emitted in canonical document order
 * (deterministic; each removal looks the link up by identity, so the order
 * is a stability guarantee, not a correctness dependency). Unknown ids are
 * dropped — the gesture never invents removals.
 */
export function buildRemoveDependencySelectionCommands(
  document: ProjectDocument,
  selectedDependencyIds: readonly DependencyId[],
): ProjectCommand[] {
  const selected = new Set(selectedDependencyIds)
  if (selected.size === 0) return []
  return document.dependencies
    .filter((dependency) => selected.has(dependency.id))
    .map(
      (dependency) => ({ type: 'RemoveDependency', dependencyId: dependency.id }) as ProjectCommand,
    )
}
