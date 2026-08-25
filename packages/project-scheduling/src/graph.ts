import type { Dependency, ProjectDocument, TaskId } from '@genoffice/project-contracts'

export class DependencyGraphError extends Error {
  constructor(
    public readonly code: 'SELF_DEPENDENCY' | 'MISSING_TASK_REFERENCE' | 'CYCLE',
    message: string,
  ) {
    super(message)
  }
}

export interface DependencyGraph {
  /** Dependency lists keyed by successor task. */
  predecessors: Map<TaskId, Dependency[]>
  /** Dependency lists keyed by predecessor task. */
  successors: Map<TaskId, Dependency[]>
  /**
   * Canonical processing order respecting both dependency edges
   * (predecessor before successor) and summary hierarchy (child before
   * parent). Tasks are released in ascending task-id order, which makes the
   * order independent of input array order.
   */
  topologicalOrder: TaskId[]
}

// Plain code-unit comparison: locale-free and deterministic across hosts.
const compareIds = (a: TaskId, b: TaskId): number => (a < b ? -1 : a > b ? 1 : 0)

export function buildDependencyGraph(document: ProjectDocument): DependencyGraph {
  const taskIds = new Set(document.tasks.map((task) => task.id))
  const predecessors = new Map<TaskId, Dependency[]>()
  const successors = new Map<TaskId, Dependency[]>()
  const adjacency = new Map<TaskId, TaskId[]>()
  const indegree = new Map<TaskId, number>()
  for (const task of document.tasks) {
    predecessors.set(task.id, [])
    successors.set(task.id, [])
    adjacency.set(task.id, [])
    indegree.set(task.id, 0)
  }

  const addEdge = (from: TaskId, to: TaskId) => {
    adjacency.get(from)!.push(to)
    indegree.set(to, indegree.get(to)! + 1)
  }

  for (const dependency of document.dependencies) {
    if (dependency.predecessorId === dependency.successorId) {
      throw new DependencyGraphError(
        'SELF_DEPENDENCY',
        `Dependency ${dependency.id} is self-referential`,
      )
    }
    if (!taskIds.has(dependency.predecessorId) || !taskIds.has(dependency.successorId)) {
      throw new DependencyGraphError(
        'MISSING_TASK_REFERENCE',
        `Dependency ${dependency.id} references an unknown task`,
      )
    }
    predecessors.get(dependency.successorId)!.push(dependency)
    successors.get(dependency.predecessorId)!.push(dependency)
    addEdge(dependency.predecessorId, dependency.successorId)
  }

  for (const task of document.tasks) {
    if (task.parentTaskId === task.id) {
      throw new DependencyGraphError('CYCLE', `Task ${task.id} cannot parent itself`)
    }
    if (task.parentTaskId) {
      if (!taskIds.has(task.parentTaskId)) {
        throw new DependencyGraphError(
          'MISSING_TASK_REFERENCE',
          `Task ${task.id} references missing parent ${task.parentTaskId}`,
        )
      }
      // Children must be scheduled before their summary parent.
      addEdge(task.id, task.parentTaskId)
    }
  }

  const ready = document.tasks
    .filter((task) => indegree.get(task.id) === 0)
    .map((task) => task.id)
    .sort(compareIds)
  const topologicalOrder: TaskId[] = []
  while (ready.length) {
    const id = ready.shift()!
    topologicalOrder.push(id)
    for (const next of adjacency.get(id) ?? []) {
      const degree = indegree.get(next)! - 1
      indegree.set(next, degree)
      if (degree === 0) {
        ready.push(next)
        ready.sort(compareIds)
      }
    }
  }
  if (topologicalOrder.length !== document.tasks.length) {
    throw new DependencyGraphError(
      'CYCLE',
      'Dependency graph or summary hierarchy contains a cycle',
    )
  }
  return { predecessors, successors, topologicalOrder }
}
