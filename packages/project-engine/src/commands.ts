import type {
  ProjectCommand,
  ProjectCommandResult,
  ProjectDocument,
  Task,
  TaskId,
} from '@genoffice/project-contracts'
import { asTaskId } from '@genoffice/project-contracts'
import { validateProjectDocument, type ValidationDiagnostic } from './document.js'
import {
  HierarchyError,
  buildTaskHierarchy,
  canonicalizeTasks,
  isDescendantOf,
  previousSiblingOf,
  subtreeIdsOf,
  type HierarchyNode,
} from './hierarchy.js'

export interface CommandExecution {
  /**
   * Resulting document. On rejection this is the input document reference
   * unchanged; on acceptance it is a new document in canonical outline order
   * with recomputed derived hierarchy fields. The input document is never
   * mutated.
   */
  document: ProjectDocument
  result: ProjectCommandResult
}

const findTask = (document: ProjectDocument, taskId: TaskId): Task | undefined =>
  document.tasks.find((task) => task.id === taskId)

const rejected = (
  commandId: string,
  diagnostics: ValidationDiagnostic[],
): ProjectCommandResult => ({
  commandId,
  accepted: false,
  diagnostics,
  affectedTaskIds: [],
})

/**
 * Returns a task object with the given parent. When `parentTaskId` is
 * undefined the parentTaskId key is removed entirely so canonical documents
 * never carry a stale undefined-valued key.
 */
const withParent = (task: Task, parentTaskId: TaskId | undefined): Task => {
  if (parentTaskId !== undefined) return { ...task, parentTaskId }
  const { parentTaskId: _removed, ...rest } = task
  return rest
}

/** Serializes the (possibly mutated) hierarchy into a task array in outline order. */
const serializeHierarchy = (roots: HierarchyNode[]): Task[] => {
  const tasks: Task[] = []
  const emit = (node: HierarchyNode) => {
    tasks.push(node.task)
    for (const child of node.children) emit(child)
  }
  for (const root of roots) emit(root)
  return tasks
}

interface MutationOutcome {
  tasks: Task[]
  dependencies?: ProjectDocument['dependencies']
  assignments?: ProjectDocument['assignments']
  baselines?: ProjectDocument['baselines']
  affectedTaskIds: TaskId[]
  inverse?: ProjectCommand
}

type Mutation =
  | { kind: 'accepted'; outcome: MutationOutcome }
  | { kind: 'rejected'; diagnostics: ValidationDiagnostic[] }

function mutateForCreateTask(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'CreateTask' },
): Mutation {
  const { task } = command
  if (task.parentTaskId === task.id) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'SELF_PARENT', message: `Task ${task.id} cannot parent itself` }],
    }
  }
  if (task.parentTaskId && !findTask(document, task.parentTaskId)) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'MISSING_PARENT',
          message: `Task ${task.id} references missing parent ${task.parentTaskId}`,
        },
      ],
    }
  }
  // A new task becomes the last child of its parent (or the last root task).
  // Appending at the end of the array preserves every existing sibling order
  // and places the task last among its new siblings. Derived fields
  // (outlineLevel, WBS, summary) are recomputed by the executor; a created
  // task is always a leaf, so a summary flag on the command is normalized.
  return {
    kind: 'accepted',
    outcome: {
      tasks: [...document.tasks, task],
      affectedTaskIds: [task.id],
      inverse: { type: 'DeleteTask', taskId: task.id },
    },
  }
}

function mutateForDeleteTask(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'DeleteTask' },
): Mutation {
  const task = findTask(document, command.taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${command.taskId} does not exist` }],
    }
  }
  // Canonical deletion policy (matches Microsoft Project outline deletion):
  // deleting a task deletes its entire descendant subtree. Dependencies,
  // assignments, and baseline snapshots that reference deleted tasks are
  // removed so every accepted mutation leaves the document valid.
  const hierarchy = buildTaskHierarchy(document.tasks)
  const node = hierarchy.byId.get(command.taskId)!
  const removed = subtreeIdsOf(node)
  const dependencies = document.dependencies.filter(
    (dependency) => !removed.has(dependency.predecessorId) && !removed.has(dependency.successorId),
  )
  const assignments = document.assignments.filter((assignment) => !removed.has(assignment.taskId))
  const baselines = document.baselines.map((baseline) => {
    const keys = Object.keys(baseline.taskSnapshots).filter((key) => !removed.has(asTaskId(key)))
    if (keys.length === Object.keys(baseline.taskSnapshots).length) return baseline
    const taskSnapshots: typeof baseline.taskSnapshots = {}
    for (const key of keys) taskSnapshots[key] = baseline.taskSnapshots[key as TaskId]
    return { ...baseline, taskSnapshots }
  })
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks.filter((candidate) => !removed.has(candidate.id)),
      dependencies,
      assignments,
      baselines,
      affectedTaskIds: task.parentTaskId ? [command.taskId, task.parentTaskId] : [command.taskId],
    },
  }
}

function mutateForRenameTask(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'RenameTask' },
): Mutation {
  const task = findTask(document, command.taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${command.taskId} does not exist` }],
    }
  }
  const renamed: Task = { ...task, name: command.name }
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks.map((candidate) =>
        candidate.id === command.taskId ? renamed : candidate,
      ),
      affectedTaskIds: [command.taskId],
      inverse: { type: 'RenameTask', taskId: command.taskId, name: task.name },
    },
  }
}

function mutateForIndentTask(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'IndentTask' },
): Mutation {
  const { taskId, parentTaskId } = command
  const task = findTask(document, taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${taskId} does not exist` }],
    }
  }
  const target = findTask(document, parentTaskId)
  if (!target) {
    return {
      kind: 'rejected',
      diagnostics: [
        { code: 'MISSING_PARENT', message: `Indent target ${parentTaskId} does not exist` },
      ],
    }
  }
  const hierarchy = buildTaskHierarchy(document.tasks)
  const node = hierarchy.byId.get(taskId)!
  const targetNode = hierarchy.byId.get(parentTaskId)!
  if (parentTaskId === taskId || isDescendantOf(targetNode, node)) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'INDENT_CYCLE',
          message: `Task ${taskId} cannot be indented beneath itself or its descendant ${parentTaskId}`,
        },
      ],
    }
  }
  // Canonical indent semantics: the task becomes the last child of its
  // immediately preceding sibling. Any other target is invalid.
  const previous = previousSiblingOf(hierarchy, node)
  if (!previous) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'INVALID_INDENT_NO_SIBLING',
          message: `Task ${taskId} has no preceding sibling to indent under`,
        },
      ],
    }
  }
  if (previous.task.id !== parentTaskId) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'INVALID_INDENT_TARGET',
          message: `Task ${taskId} can only indent beneath its preceding sibling ${previous.task.id}, not ${parentTaskId}`,
        },
      ],
    }
  }
  const siblings = node.parent ? node.parent.children : hierarchy.roots
  siblings.splice(siblings.indexOf(node), 1)
  previous.children.push(node)
  node.parent = previous
  node.task = withParent(task, parentTaskId)
  return {
    kind: 'accepted',
    outcome: {
      tasks: serializeHierarchy(hierarchy.roots),
      affectedTaskIds: [taskId, parentTaskId],
      inverse: { type: 'OutdentTask', taskId },
    },
  }
}

function mutateForOutdentTask(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'OutdentTask' },
): Mutation {
  const { taskId } = command
  const task = findTask(document, taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${taskId} does not exist` }],
    }
  }
  if (!task.parentTaskId) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'INVALID_OUTDENT_ROOT',
          message: `Task ${taskId} is a root task and cannot be outdented`,
        },
      ],
    }
  }
  const hierarchy = buildTaskHierarchy(document.tasks)
  const node = hierarchy.byId.get(taskId)!
  const parent = node.parent!
  // Canonical outdent semantics: the task (with its own subtree) becomes the
  // next sibling of its former parent.
  const parentSiblings = parent.parent ? parent.parent.children : hierarchy.roots
  parent.children.splice(parent.children.indexOf(node), 1)
  parentSiblings.splice(parentSiblings.indexOf(parent) + 1, 0, node)
  node.parent = parent.parent
  node.task = withParent(task, parent.parent ? parent.parent.task.id : undefined)
  return {
    kind: 'accepted',
    outcome: {
      tasks: serializeHierarchy(hierarchy.roots),
      affectedTaskIds: [taskId, parent.task.id],
      // No inverse: a single IndentTask can only re-attach the task as the
      // former parent's LAST child, which does not restore an arbitrary
      // original sibling position. Exact undo of outdent needs a richer
      // command model and is deliberately left unclaimed here.
    },
  }
}

/**
 * Applies a semantic ProjectCommand to a ProjectDocument.
 *
 * Pure and deterministic: the same document plus the same command sequence
 * always produces the same resulting document bytes. Rejected commands return
 * the input document unchanged with diagnostics. Accepted commands leave the
 * document in canonical outline order with recomputed derived hierarchy
 * fields (outlineLevel, WBS, summary), so every accepted mutation leaves the
 * ProjectDocument valid.
 *
 * Only the PROJECT-007 hierarchy commands are executable today; other command
 * types are rejected deterministically until their work items are authorized.
 */
export function applyProjectCommand(
  document: ProjectDocument,
  command: ProjectCommand,
  commandId: string = command.type,
): CommandExecution {
  // Commands only execute against a valid document, and structural hierarchy
  // building is guaranteed by prior validation; treat structural failures as
  // deterministic rejections rather than crashes.
  try {
    const pre = validateProjectDocument(document)
    if (!pre.accepted) {
      return { document, result: rejected(commandId, pre.diagnostics) }
    }

    let mutation: Mutation
    switch (command.type) {
      case 'CreateTask':
        mutation = mutateForCreateTask(document, command)
        break
      case 'DeleteTask':
        mutation = mutateForDeleteTask(document, command)
        break
      case 'RenameTask':
        mutation = mutateForRenameTask(document, command)
        break
      case 'IndentTask':
        mutation = mutateForIndentTask(document, command)
        break
      case 'OutdentTask':
        mutation = mutateForOutdentTask(document, command)
        break
      default:
        mutation = {
          kind: 'rejected',
          diagnostics: [
            {
              code: 'UNSUPPORTED_COMMAND',
              message: `Command ${command.type} is not implemented by the hierarchy command executor`,
            },
          ],
        }
    }

    if (mutation.kind === 'rejected') {
      return { document, result: rejected(commandId, mutation.diagnostics) }
    }

    const candidate: ProjectDocument = {
      ...document,
      tasks: canonicalizeTasks(mutation.outcome.tasks),
      ...(mutation.outcome.dependencies !== undefined
        ? { dependencies: mutation.outcome.dependencies }
        : {}),
      ...(mutation.outcome.assignments !== undefined
        ? { assignments: mutation.outcome.assignments }
        : {}),
      ...(mutation.outcome.baselines !== undefined
        ? { baselines: mutation.outcome.baselines }
        : {}),
    }

    const post = validateProjectDocument(candidate)
    if (!post.accepted) {
      return { document, result: rejected(commandId, post.diagnostics) }
    }

    return {
      document: candidate,
      result: {
        commandId,
        accepted: true,
        diagnostics: [],
        affectedTaskIds: mutation.outcome.affectedTaskIds,
        inverse: mutation.outcome.inverse,
      },
    }
  } catch (error) {
    if (error instanceof HierarchyError) {
      return {
        document,
        result: rejected(commandId, [{ code: error.code, message: error.message }]),
      }
    }
    throw error
  }
}
