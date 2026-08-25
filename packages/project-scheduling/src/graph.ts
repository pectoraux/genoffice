import type { Dependency, ProjectDocument, TaskId } from '@genoffice/project-contracts'

export class DependencyGraphError extends Error {
  constructor(public readonly code: 'SELF_DEPENDENCY' | 'MISSING_TASK_REFERENCE' | 'CYCLE', message: string) { super(message) }
}

export interface DependencyGraph {
  predecessors: Map<TaskId, Dependency[]>
  successors: Map<TaskId, Dependency[]>
  topologicalOrder: TaskId[]
}

export function buildDependencyGraph(document: ProjectDocument): DependencyGraph {
  const taskIds = new Set(document.tasks.map((task) => task.id))
  const predecessors = new Map<TaskId, Dependency[]>()
  const successors = new Map<TaskId, Dependency[]>()
  const indegree = new Map<TaskId, number>()
  for (const task of document.tasks) { predecessors.set(task.id, []); successors.set(task.id, []); indegree.set(task.id, 0) }
  for (const dependency of document.dependencies) {
    if (dependency.predecessorId === dependency.successorId) throw new DependencyGraphError('SELF_DEPENDENCY', `Dependency ${dependency.id} is self-referential`)
    if (!taskIds.has(dependency.predecessorId) || !taskIds.has(dependency.successorId)) throw new DependencyGraphError('MISSING_TASK_REFERENCE', `Dependency ${dependency.id} references an unknown task`)
    predecessors.get(dependency.successorId)!.push(dependency)
    successors.get(dependency.predecessorId)!.push(dependency)
    indegree.set(dependency.successorId, indegree.get(dependency.successorId)! + 1)
  }
  const ready = document.tasks.filter((task) => indegree.get(task.id) === 0).map((task) => task.id).sort()
  const topologicalOrder: TaskId[] = []
  while (ready.length) {
    const id = ready.shift()!
    topologicalOrder.push(id)
    for (const dependency of successors.get(id) ?? []) {
      const next = dependency.successorId
      indegree.set(next, indegree.get(next)! - 1)
      if (indegree.get(next) === 0) { ready.push(next); ready.sort() }
    }
  }
  if (topologicalOrder.length !== document.tasks.length) throw new DependencyGraphError('CYCLE', 'Dependency graph contains a cycle')
  return { predecessors, successors, topologicalOrder }
}
