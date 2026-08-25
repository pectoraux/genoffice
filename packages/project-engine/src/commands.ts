import type {
  ConstraintType,
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
 * The six date-bounded constraints require a constraintDate; ASAP/ALAP never
 * use one. This constant mirrors the document validator's rule so the command
 * mutator can build a canonical task shape without silent reinterpretation.
 */
const DATE_BOUNDED_CONSTRAINTS: ReadonlySet<ConstraintType> = new Set([
  'startNoEarlierThan',
  'startNoLaterThan',
  'mustStartOn',
  'finishNoEarlierThan',
  'finishNoLaterThan',
  'mustFinishOn',
])

/**
 * Returns a task with the given optional date field. When the value is
 * undefined the key is removed entirely so canonical documents never carry a
 * stale undefined-valued key (matching the `withParent` convention).
 */
const withOptionalDate = <K extends string>(
  task: Task,
  key: K,
  value: string | undefined,
): Task => {
  if (value !== undefined) return { ...task, [key]: value }
  const { [key]: _removed, ...rest } = task
  return rest as Task
}

function mutateForSetConstraint(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'SetConstraint' },
): Mutation {
  const task = findTask(document, command.taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${command.taskId} does not exist` }],
    }
  }
  // The frozen SetConstraint command shape allows an undefined constraintType
  // (it reuses the Task field type). A constraint cannot be set without a
  // type, so reject deterministically rather than silently defaulting.
  if (command.constraintType === undefined) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'MISSING_CONSTRAINT_TYPE',
          message: `Task ${command.taskId} cannot set a constraint without a constraintType`,
        },
      ],
    }
  }
  const constraintType: ConstraintType = command.constraintType
  // The mutator stores exactly the constraint shape given and lets the
  // post-mutation validator enforce the canonical rules (missing date on a
  // date-bounded type, disallowed date on ASAP/ALAP, malformed dates). This
  // means no constraint is silently reinterpreted: MSO never collapses to
  // SNET, and ASAP with a date is rejected rather than honored.
  const previousType = task.constraintType
  const previousDate = task.constraintDate
  let updated = withOptionalDate(task, 'constraintDate', command.constraintDate)
  updated = { ...updated, constraintType }
  // Re-normalize: a date-bounded constraint must keep its date key (handled
  // above), and ASAP/ALAP must never carry a constraintDate key. If the command
  // supplied a date alongside ASAP/ALAP the validator rejects it; if it did
  // not, ensure no stale date key remains from a prior dated constraint.
  if (!DATE_BOUNDED_CONSTRAINTS.has(constraintType) && command.constraintDate === undefined) {
    updated = withOptionalDate(updated, 'constraintDate', undefined)
  }
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks.map((candidate) =>
        candidate.id === command.taskId ? updated : candidate,
      ),
      affectedTaskIds: [command.taskId],
      // Restore the previous constraint shape. A previously-unconstrained task
      // restores to the ASAP default (the no-constraint scheduling mode), so
      // the inverse is always a valid SetConstraint payload.
      inverse: {
        type: 'SetConstraint',
        taskId: command.taskId,
        constraintType: previousType ?? 'asSoonAsPossible',
        constraintDate: previousDate,
      },
    },
  }
}

function mutateForSetPercentComplete(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'SetPercentComplete' },
): Mutation {
  const task = findTask(document, command.taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${command.taskId} does not exist` }],
    }
  }
  // Summary progress is a DERIVED roll-up of the subtree; it is never set
  // directly. Rejecting here keeps the summary's stored percentComplete a
  // non-authoritative placeholder and prevents the renderer from asserting a
  // value the engine would otherwise overwrite on the next derivation.
  if (task.summary) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'SUMMARY_PROGRESS_NOT_SETTABLE',
          message: `Task ${command.taskId} is a summary; its progress is derived from children`,
        },
      ],
    }
  }
  if (
    !Number.isFinite(command.percentComplete) ||
    command.percentComplete < 0 ||
    command.percentComplete > 100
  ) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'INVALID_PERCENT_COMPLETE',
          message: `Task ${command.taskId} percentComplete ${command.percentComplete} is outside 0-100`,
        },
      ],
    }
  }
  const updated: Task = { ...task, percentComplete: command.percentComplete }
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks.map((candidate) =>
        candidate.id === command.taskId ? updated : candidate,
      ),
      affectedTaskIds: [command.taskId],
      inverse: {
        type: 'SetPercentComplete',
        taskId: command.taskId,
        percentComplete: task.percentComplete,
      },
    },
  }
}

function mutateForSetDeadline(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'SetDeadline' },
): Mutation {
  const task = findTask(document, command.taskId)
  if (!task) {
    return {
      kind: 'rejected',
      diagnostics: [{ code: 'MISSING_TASK', message: `Task ${command.taskId} does not exist` }],
    }
  }
  // A deadline is NOT a scheduling constraint: it never moves the task. The
  // mutator stores exactly the deadline given (or clears it when undefined)
  // and lets the validator reject malformed dates. No scheduling semantics
  // are invented here.
  const previousDeadline = task.deadline
  const updated = withOptionalDate(task, 'deadline', command.deadline)
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks.map((candidate) =>
        candidate.id === command.taskId ? updated : candidate,
      ),
      affectedTaskIds: [command.taskId],
      inverse: {
        type: 'SetDeadline',
        taskId: command.taskId,
        deadline: previousDeadline,
      },
    },
  }
}

/**
 * PROJECT-009 CreateBaseline.
 *
 * A baseline is an immutable snapshot of task start/finish/duration/work/cost
 * captured at a point in time. The command carries a fully-formed `Baseline`
 * value: the mutator stores it as-is and lets the post-mutation validator
 * enforce the canonical rules (duplicate baseline id, malformed `capturedAt`,
 * snapshot keys referencing missing tasks). No scheduling is performed inside
 * the mutator — building the snapshot from the current `DerivedSchedule` is the
 * scheduling package's job (`captureBaseline`), keeping the command pure and
 * deterministic.
 *
 * To preserve `Task.baseline` as a canonical reverse index, every task that has
 * a snapshot in the new baseline receives the baseline id in its `baseline`
 * array (deduplicated; existing references are preserved). This keeps the
 * invariant `task.baseline ⊆ document.baselines` bidirectionally consistent:
 * a baseline exists iff at least one task references it, and a task references
 * a baseline iff that baseline has a snapshot for it.
 *
 * Baseline immutability through hierarchy mutations is structural: baselines
 * are keyed by stable `TaskId`, and `IndentTask`/`OutdentTask`/`RenameTask`
 * never change `TaskId`, so snapshots survive those mutations unchanged.
 * `DeleteTask` already prunes dangling snapshots for deleted tasks (the
 * Microsoft Project outline-deletion behavior), so every accepted mutation
 * leaves the document's baseline state valid.
 *
 * No inverse is provided: a future `DeleteBaseline` command (PROJECT-038
 * territory: multiple-baseline management) is required to undo a capture
 * cleanly. This mirrors the existing `OutdentTask` precedent of leaving an
 * undefined inverse when a single inverse command cannot restore prior state.
 */
function mutateForCreateBaseline(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'CreateBaseline' },
): Mutation {
  const { baseline } = command
  // Early rejection for a duplicate id keeps the diagnostic surface clean and
  // avoids the post-validation pass discovering it after partial work.
  if (document.baselines.some((existing) => existing.id === baseline.id)) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'DUPLICATE_BASELINE_ID',
          message: `Baseline ${baseline.id} already exists`,
        },
      ],
    }
  }
  // Early rejection for snapshots referencing missing tasks: every snapshot
  // key must be an existing TaskId.
  const taskIds = new Set(document.tasks.map((task) => task.id))
  const missing: string[] = []
  for (const taskKey of Object.keys(baseline.taskSnapshots)) {
    if (!taskIds.has(asTaskId(taskKey))) missing.push(taskKey)
  }
  if (missing.length) {
    return {
      kind: 'rejected',
      diagnostics: missing.map((taskKey) => ({
        code: 'MISSING_TASK_REFERENCE',
        message: `Baseline ${baseline.id} references missing task ${taskKey}`,
      })),
    }
  }
  // Add the baseline id to every task that has a snapshot, preserving any
  // existing baseline references (a task may be tracked by several baselines).
  const snapshotKeys = new Set(Object.keys(baseline.taskSnapshots))
  const tasks = document.tasks.map((task) =>
    snapshotKeys.has(task.id as string) && !task.baseline.includes(baseline.id)
      ? { ...task, baseline: [...task.baseline, baseline.id] }
      : task,
  )
  return {
    kind: 'accepted',
    outcome: {
      tasks,
      baselines: [...document.baselines, baseline],
      affectedTaskIds: Object.keys(baseline.taskSnapshots) as TaskId[],
      // No inverse: a future DeleteBaseline command is required to undo a
      // capture cleanly (see the PROJECT-009 spec clarifications).
    },
  }
}

/**
 * PROJECT-010 AssignResource.
 *
 * Adds a resource/task relationship to the document as an `Assignment`. The
 * command carries a fully-formed `Assignment` value: the mutator stores it
 * as-is and lets the post-mutation validator enforce the canonical rules
 * (duplicate assignment id, missing task/resource references, invalid units,
 * duplicate task/resource pair). Early rejection for a duplicate id and for
 * missing task/resource references keeps the diagnostic surface clean and
 * avoids the post-validation pass discovering structural failures after
 * partial work.
 *
 * Assignment identity (`AssignmentId`) is stable and never array position.
 * The assignment is appended to the `assignments` array; canonical ordering
 * for scheduling output is enforced deterministically by the scheduling
 * engine (sorted by `AssignmentId`), so insertion order here does not affect
 * derived schedule bytes.
 *
 * PROJECT-010 does NOT compute assignment work or cost (those are PROJECT-011).
 * The `Assignment` carries work/cost fields in the frozen contract, but this
 * command only establishes the valid scheduling-input relationship.
 */
function mutateForAssignResource(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'AssignResource' },
): Mutation {
  const { assignment } = command
  // Early rejection for a duplicate assignment id keeps the diagnostic surface
  // clean (the post-mutation validator would also catch it, but reporting here
  // gives a single, immediate, actionable diagnostic).
  if (document.assignments.some((existing) => existing.id === assignment.id)) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'DUPLICATE_ASSIGNMENT_ID',
          message: `Assignment ${assignment.id} already exists`,
        },
      ],
    }
  }
  // Early rejection for missing task/resource references: an assignment must
  // link real entities. The post-mutation validator re-checks this, but the
  // early rejection avoids building a candidate document from broken inputs.
  if (!findTask(document, assignment.taskId)) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'MISSING_TASK_REFERENCE',
          message: `Assignment ${assignment.id} references missing task ${assignment.taskId}`,
        },
      ],
    }
  }
  if (!document.resources.some((resource) => resource.id === assignment.resourceId)) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'MISSING_RESOURCE_REFERENCE',
          message: `Assignment ${assignment.id} references missing resource ${assignment.resourceId}`,
        },
      ],
    }
  }
  // Early rejection for a duplicate task/resource pair: two assignment rows
  // cannot silently shadow the same task/resource relationship. The
  // post-mutation validator also enforces this, but the early rejection keeps
  // the rejected document byte-identical to the input.
  const pairExists = document.assignments.some(
    (existing) =>
      existing.taskId === assignment.taskId && existing.resourceId === assignment.resourceId,
  )
  if (pairExists) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'DUPLICATE_ASSIGNMENT_PAIR',
          message: `Assignment ${assignment.id} duplicates the task/resource link ${assignment.taskId}->${assignment.resourceId}`,
        },
      ],
    }
  }
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks,
      assignments: [...document.assignments, assignment],
      affectedTaskIds: [assignment.taskId],
      inverse: { type: 'UnassignResource', assignmentId: assignment.id },
    },
  }
}

/**
 * PROJECT-010 UnassignResource.
 *
 * Removes an assignment by `AssignmentId`. The removed assignment is captured
 * so the inverse (`AssignResource`) can restore the exact prior relationship.
 * Removing an assignment never moves task dates (PROJECT-010 does not
 * resource-level); it only removes the scheduling-input relationship.
 */
function mutateForUnassignResource(
  document: ProjectDocument,
  command: ProjectCommand & { type: 'UnassignResource' },
): Mutation {
  const existing = document.assignments.find((item) => item.id === command.assignmentId)
  if (!existing) {
    return {
      kind: 'rejected',
      diagnostics: [
        {
          code: 'MISSING_ASSIGNMENT',
          message: `Assignment ${command.assignmentId} does not exist`,
        },
      ],
    }
  }
  return {
    kind: 'accepted',
    outcome: {
      tasks: document.tasks,
      assignments: document.assignments.filter((item) => item.id !== command.assignmentId),
      affectedTaskIds: [existing.taskId],
      inverse: { type: 'AssignResource', assignment: existing },
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
 * The PROJECT-007 hierarchy commands plus the PROJECT-008 constraint/deadline/
 * progress commands are executable today; other command types are rejected
 * deterministically until their work items are authorized.
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
      case 'SetConstraint':
        mutation = mutateForSetConstraint(document, command)
        break
      case 'SetPercentComplete':
        mutation = mutateForSetPercentComplete(document, command)
        break
      case 'SetDeadline':
        mutation = mutateForSetDeadline(document, command)
        break
      case 'CreateBaseline':
        mutation = mutateForCreateBaseline(document, command)
        break
      case 'AssignResource':
        mutation = mutateForAssignResource(document, command)
        break
      case 'UnassignResource':
        mutation = mutateForUnassignResource(document, command)
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
