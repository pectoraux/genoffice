import type { ProjectDocument, Task, TaskId } from '@genoffice/project-contracts'

export class HierarchyError extends Error {
  constructor(
    public readonly code: 'MISSING_PARENT' | 'SELF_PARENT' | 'HIERARCHY_CYCLE',
    message: string,
  ) {
    super(message)
  }
}

export interface HierarchyNode {
  task: Task
  parent?: HierarchyNode
  children: HierarchyNode[]
}

export interface TaskHierarchy {
  roots: HierarchyNode[]
  byId: Map<TaskId, HierarchyNode>
  /**
   * Canonical outline order: depth-first over the hierarchy with siblings in
   * canonical sibling order. Deterministic for a given task list.
   */
  order: TaskId[]
}

/**
 * Canonical sibling-ordering representation (PROJECT-007 decision):
 *
 * The relative order of tasks inside `ProjectDocument.tasks` IS the canonical
 * sibling ordering. Array position expresses ordering only — it is never
 * identity (identity is TaskId) and never hierarchy (hierarchy is
 * parentTaskId). This is the smallest architecture-compatible representation:
 * no extra ordering field exists on the frozen Task model, and because
 * ordering is positional an invalid sibling ordering cannot be represented.
 */
export function buildTaskHierarchy(tasks: Task[]): TaskHierarchy {
  const byId = new Map<TaskId, HierarchyNode>()
  for (const task of tasks) {
    byId.set(task.id, { task, children: [] })
  }
  const roots: HierarchyNode[] = []
  for (const task of tasks) {
    const node = byId.get(task.id)!
    if (!task.parentTaskId) {
      roots.push(node)
      continue
    }
    if (task.parentTaskId === task.id) {
      throw new HierarchyError('SELF_PARENT', `Task ${task.id} cannot parent itself`)
    }
    const parent = byId.get(task.parentTaskId)
    if (!parent) {
      throw new HierarchyError(
        'MISSING_PARENT',
        `Task ${task.id} references missing parent ${task.parentTaskId}`,
      )
    }
    parent.children.push(node)
    node.parent = parent
  }

  // Tasks unreachable from a root can only belong to a parent-chain cycle.
  const order: TaskId[] = []
  const visited = new Set<TaskId>()
  const visit = (node: HierarchyNode) => {
    visited.add(node.task.id)
    order.push(node.task.id)
    for (const child of node.children) visit(child)
  }
  for (const root of roots) visit(root)
  for (const task of tasks) {
    if (!visited.has(task.id)) {
      throw new HierarchyError(
        'HIERARCHY_CYCLE',
        `Task hierarchy contains a cycle involving ${task.id}`,
      )
    }
  }
  return { roots, byId, order }
}

export interface DerivedHierarchyFields {
  /** Hierarchy depth: root tasks are 1, children are parent + 1. */
  outlineLevel: number
  /**
   * Derived WBS code: dotted 1-based sibling numbering (for example
   * 1, 1.1, 1.2, 2, 2.1). WBS is a derived hierarchy representation and is
   * never identity — TaskId remains identity.
   */
  wbs: string
  /** Semantic summary flag: a task is a summary iff it has at least one child. */
  summary: boolean
}

/**
 * Derives canonical hierarchy fields (outlineLevel, WBS, summary) for every
 * task. Requires a valid hierarchy; invalid hierarchies throw HierarchyError
 * (document validation reports those as diagnostics instead).
 */
export function deriveHierarchyFields(tasks: Task[]): Map<TaskId, DerivedHierarchyFields> {
  const hierarchy = buildTaskHierarchy(tasks)
  const fields = new Map<TaskId, DerivedHierarchyFields>()
  const derive = (node: HierarchyNode, outlineLevel: number, wbs: string) => {
    fields.set(node.task.id, {
      outlineLevel,
      wbs,
      summary: node.children.length > 0,
    })
    node.children.forEach((child, index) => derive(child, outlineLevel + 1, `${wbs}.${index + 1}`))
  }
  hierarchy.roots.forEach((root, index) => derive(root, 1, `${index + 1}`))
  return fields
}

/**
 * Returns the tasks in canonical outline order (depth-first, siblings in
 * canonical sibling order) with recomputed derived hierarchy fields
 * (outlineLevel, WBS, summary). Sibling ordering is preserved from the
 * relative array positions of same-parent tasks. Idempotent on canonical
 * input.
 */
export function canonicalizeTasks(tasks: Task[]): Task[] {
  const hierarchy = buildTaskHierarchy(tasks)
  const result: Task[] = []
  const emit = (node: HierarchyNode, outlineLevel: number, wbs: string) => {
    result.push({
      ...node.task,
      outlineLevel,
      wbs,
      summary: node.children.length > 0,
    })
    node.children.forEach((child, index) => emit(child, outlineLevel + 1, `${wbs}.${index + 1}`))
  }
  hierarchy.roots.forEach((root, index) => emit(root, 1, `${index + 1}`))
  return result
}

/** Applies canonical task ordering and derived hierarchy fields to a document. */
export function canonicalizeDocument(document: ProjectDocument): ProjectDocument {
  return { ...document, tasks: canonicalizeTasks(document.tasks) }
}

/** Sibling list containing the node (its parent's children, or the roots). */
export function siblingsOf(hierarchy: TaskHierarchy, node: HierarchyNode): HierarchyNode[] {
  return node.parent ? node.parent.children : hierarchy.roots
}

/** The sibling immediately preceding `node` in canonical sibling order, if any. */
export function previousSiblingOf(
  hierarchy: TaskHierarchy,
  node: HierarchyNode,
): HierarchyNode | undefined {
  const siblings = siblingsOf(hierarchy, node)
  const index = siblings.indexOf(node)
  return index > 0 ? siblings[index - 1] : undefined
}

/** Task ids of `node` and all its descendants. */
export function subtreeIdsOf(node: HierarchyNode): Set<TaskId> {
  const ids = new Set<TaskId>()
  const visit = (current: HierarchyNode) => {
    ids.add(current.task.id)
    for (const child of current.children) visit(child)
  }
  visit(node)
  return ids
}

/** True when `candidate` is a strict descendant of `ancestor`. */
export function isDescendantOf(candidate: HierarchyNode, ancestor: HierarchyNode): boolean {
  let current = candidate.parent
  while (current) {
    if (current === ancestor) return true
    current = current.parent
  }
  return false
}
