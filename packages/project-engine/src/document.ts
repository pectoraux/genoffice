import type {
  Calendar,
  CalendarPeriod,
  ProjectDocument,
  ProjectCommand,
  ProjectCommandResult,
  TaskId,
} from '@genoffice/project-contracts'

export type ValidationDiagnostic = { code: string; message: string }

function checkDuplicateIds(
  entities: Array<{ id: string }>,
  label: string,
  code: string,
  diagnostics: ValidationDiagnostic[],
): Set<string> {
  const seen = new Set<string>()
  for (const entity of entities) {
    if (seen.has(entity.id)) {
      diagnostics.push({ code, message: `Duplicate ${label} id ${entity.id}` })
    }
    seen.add(entity.id)
  }
  return seen
}

function isValidDate(value: string): boolean {
  return !Number.isNaN(new Date(value).getTime())
}

function isValidExceptionDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(value) && isValidDate(`${value}T00:00:00.000Z`)
}

function validatePeriods(
  calendarId: string,
  periods: CalendarPeriod[],
  diagnostics: ValidationDiagnostic[],
) {
  const sorted = [...periods].sort(
    (a, b) => a.startMinute - b.startMinute || a.endMinute - b.endMinute,
  )
  let previousEnd = -1
  for (const period of sorted) {
    if (!Number.isInteger(period.startMinute) || !Number.isInteger(period.endMinute)) {
      diagnostics.push({
        code: 'CALENDAR_PERIOD_MALFORMED',
        message: `Calendar ${calendarId} has a period with non-integer minute bounds`,
      })
      return
    }
    if (
      period.startMinute < 0 ||
      period.endMinute > 1440 ||
      period.startMinute >= period.endMinute
    ) {
      diagnostics.push({
        code: 'CALENDAR_PERIOD_MALFORMED',
        message: `Calendar ${calendarId} has a period outside 00:00-24:00 or with an empty interval`,
      })
      return
    }
    if (period.startMinute < previousEnd) {
      diagnostics.push({
        code: 'CALENDAR_PERIOD_MALFORMED',
        message: `Calendar ${calendarId} has overlapping periods`,
      })
      return
    }
    previousEnd = period.endMinute
  }
}

function validateCalendars(document: ProjectDocument, diagnostics: ValidationDiagnostic[]) {
  const calendarIds = checkDuplicateIds(
    document.calendars,
    'calendar',
    'DUPLICATE_CALENDAR_ID',
    diagnostics,
  )
  for (const calendar of document.calendars) {
    for (const periods of Object.values(calendar.workingWeek)) {
      validatePeriods(calendar.id as string, periods, diagnostics)
    }
    for (const exception of calendar.exceptions) {
      if (!isValidExceptionDate(exception.date)) {
        diagnostics.push({
          code: 'INVALID_DATE',
          message: `Calendar ${calendar.id} has an exception with malformed date ${exception.date}`,
        })
      }
      validatePeriods(calendar.id as string, exception.periods, diagnostics)
    }
  }
  for (const calendar of document.calendars) {
    if (calendar.baseCalendarId && !calendarIds.has(calendar.baseCalendarId)) {
      diagnostics.push({
        code: 'MISSING_BASE_CALENDAR',
        message: `Calendar ${calendar.id} inherits from missing calendar ${calendar.baseCalendarId}`,
      })
    }
  }
  // Inheritance cycle detection.
  const byId = new Map(document.calendars.map((calendar) => [calendar.id as string, calendar]))
  for (const calendar of document.calendars) {
    const seen = new Set<string>()
    let current: Calendar | undefined = calendar
    while (current?.baseCalendarId) {
      if (seen.has(current.id as string)) {
        diagnostics.push({
          code: 'CALENDAR_CYCLE',
          message: `Calendar inheritance cycle involves ${current.id}`,
        })
        break
      }
      seen.add(current.id as string)
      current = byId.get(current.baseCalendarId as string)
    }
  }
  if (!calendarIds.has(document.properties.defaultCalendarId)) {
    diagnostics.push({
      code: 'MISSING_CALENDAR',
      message: `Project default calendar ${document.properties.defaultCalendarId} is missing`,
    })
  }
}

function validateTasks(
  document: ProjectDocument,
  diagnostics: ValidationDiagnostic[],
): {
  taskIds: Set<string>
  taskById: Map<string, { id: string; parentTaskId?: string; summary: boolean }>
} {
  const taskIds = checkDuplicateIds(document.tasks, 'task', 'DUPLICATE_TASK_ID', diagnostics)
  const taskById = new Map(document.tasks.map((task) => [task.id as string, task]))
  const calendarIds = new Set(document.calendars.map((calendar) => calendar.id as string))
  const baselineIds = new Set(document.baselines.map((baseline) => baseline.id as string))
  const customFieldIds = new Set(document.customFields.map((field) => field.id as string))

  for (const task of document.tasks) {
    if (task.duration < 0) {
      diagnostics.push({
        code: 'NEGATIVE_DURATION',
        message: `Task ${task.id} has negative duration`,
      })
    }
    if (
      !Number.isFinite(task.percentComplete) ||
      task.percentComplete < 0 ||
      task.percentComplete > 100
    ) {
      diagnostics.push({
        code: 'INVALID_PERCENT_COMPLETE',
        message: `Task ${task.id} has percentComplete outside 0-100`,
      })
    }
    if (
      task.physicalPercentComplete !== undefined &&
      (!Number.isFinite(task.physicalPercentComplete) ||
        task.physicalPercentComplete < 0 ||
        task.physicalPercentComplete > 100)
    ) {
      diagnostics.push({
        code: 'INVALID_PERCENT_COMPLETE',
        message: `Task ${task.id} has physicalPercentComplete outside 0-100`,
      })
    }
    if (task.outlineLevel < 1) {
      diagnostics.push({
        code: 'INVALID_OUTLINE_LEVEL',
        message: `Task ${task.id} has outlineLevel below 1`,
      })
    }
    if (task.calendarId && !calendarIds.has(task.calendarId)) {
      diagnostics.push({
        code: 'MISSING_CALENDAR',
        message: `Task ${task.id} references missing calendar ${task.calendarId}`,
      })
    }
    if (task.parentTaskId === task.id) {
      diagnostics.push({ code: 'SELF_PARENT', message: `Task ${task.id} cannot parent itself` })
    } else if (task.parentTaskId && !taskIds.has(task.parentTaskId)) {
      diagnostics.push({
        code: 'MISSING_PARENT',
        message: `Task ${task.id} references missing parent ${task.parentTaskId}`,
      })
    }
    for (const dateField of [task.start, task.finish, task.constraintDate, task.deadline]) {
      if (dateField !== undefined && !isValidDate(dateField)) {
        diagnostics.push({
          code: 'INVALID_DATE',
          message: `Task ${task.id} has a malformed date ${dateField}`,
        })
      }
    }
    // PROJECT-008 constraint canonical validation: the six date-bounded
    // constraint types MUST carry a valid constraintDate, while ASAP/ALAP
    // never use one. This prevents silent reinterpretation (e.g. MSO stored
    // without a date being treated as SNET, or ASAP silently honoring a date).
    if (task.constraintType) {
      const dateBounded =
        task.constraintType !== 'asSoonAsPossible' && task.constraintType !== 'asLateAsPossible'
      if (dateBounded && (task.constraintDate === undefined || !isValidDate(task.constraintDate))) {
        diagnostics.push({
          code: 'MISSING_CONSTRAINT_DATE',
          message: `Task ${task.id} constraint ${task.constraintType} requires a valid constraintDate`,
        })
      }
      if (!dateBounded && task.constraintDate !== undefined) {
        diagnostics.push({
          code: 'CONSTRAINT_DATE_NOT_ALLOWED',
          message: `Task ${task.id} constraint ${task.constraintType} must not carry a constraintDate`,
        })
      }
    }
    for (const baselineId of task.baseline) {
      if (!baselineIds.has(baselineId)) {
        diagnostics.push({
          code: 'MISSING_BASELINE_REFERENCE',
          message: `Task ${task.id} references missing baseline ${baselineId}`,
        })
      }
    }
    for (const customFieldId of Object.keys(task.customFields)) {
      if (!customFieldIds.has(customFieldId)) {
        diagnostics.push({
          code: 'MISSING_CUSTOM_FIELD_REFERENCE',
          message: `Task ${task.id} references missing custom field ${customFieldId}`,
        })
      }
    }
  }

  // Task UID uniqueness: UIDs are persistent interoperability identifiers
  // (source-file round-tripping), so they must be unique per document even
  // though TaskId remains the canonical identity.
  const taskUids = new Set<number>()
  for (const task of document.tasks) {
    if (taskUids.has(task.uid)) {
      diagnostics.push({
        code: 'DUPLICATE_TASK_UID',
        message: `Task ${task.id} duplicates uid ${task.uid}`,
      })
    }
    taskUids.add(task.uid)
  }

  // Parent-chain cycle detection.
  for (const task of document.tasks) {
    const seen = new Set<string>()
    let current = taskById.get(task.id as string)
    while (current?.parentTaskId) {
      if (seen.has(current.id as string)) {
        diagnostics.push({
          code: 'PARENT_CYCLE',
          message: `Task hierarchy contains a cycle involving ${current.id}`,
        })
        break
      }
      seen.add(current.id as string)
      current = taskById.get(current.parentTaskId as string)
    }
  }

  // PROJECT-007 hierarchy invariants: outlineLevel must equal hierarchy depth
  // (root = 1, child = parent + 1) and the summary flag must equal "has at
  // least one child". Cyclic or missing-parent chains are skipped here; they
  // are already reported above.
  const parentsWithChildren = new Set<string>()
  for (const task of document.tasks) {
    if (task.parentTaskId && taskIds.has(task.parentTaskId)) {
      parentsWithChildren.add(task.parentTaskId as string)
    }
  }
  const outlineDepthOf = (task: { id: string; parentTaskId?: string }): number | undefined => {
    let depth = 1
    const seen = new Set<string>()
    let current: { id: string; parentTaskId?: string } | undefined = task
    while (current?.parentTaskId) {
      if (seen.has(current.id)) return undefined
      seen.add(current.id)
      const parent = taskById.get(current.parentTaskId as string)
      if (!parent) return undefined
      current = parent
      depth += 1
    }
    return depth
  }
  for (const task of document.tasks) {
    const depth = outlineDepthOf(task)
    if (depth !== undefined && task.outlineLevel >= 1 && task.outlineLevel !== depth) {
      diagnostics.push({
        code: 'INCONSISTENT_OUTLINE_LEVEL',
        message: `Task ${task.id} has outlineLevel ${task.outlineLevel} but hierarchy depth is ${depth}`,
      })
    }
    const hasChildren = parentsWithChildren.has(task.id as string)
    if (task.summary !== hasChildren) {
      diagnostics.push({
        code: 'INCONSISTENT_SUMMARY_FLAG',
        message: hasChildren
          ? `Task ${task.id} has children and must be a summary task`
          : `Task ${task.id} has no children and must not be flagged as a summary task`,
      })
    }
  }

  return { taskIds, taskById }
}

function validateDependencies(
  document: ProjectDocument,
  taskIds: Set<string>,
  taskById: Map<string, { id: string; parentTaskId?: string; summary: boolean }>,
  diagnostics: ValidationDiagnostic[],
) {
  checkDuplicateIds(document.dependencies, 'dependency', 'DUPLICATE_DEPENDENCY_ID', diagnostics)
  const linkKeys = new Set<string>()
  const edges: Array<{ predecessor: string; successor: string }> = []
  for (const dependency of document.dependencies) {
    if (!taskIds.has(dependency.predecessorId) || !taskIds.has(dependency.successorId)) {
      diagnostics.push({
        code: 'MISSING_TASK_REFERENCE',
        message: `Dependency ${dependency.id} references a missing task`,
      })
      continue
    }
    if (dependency.predecessorId === dependency.successorId) {
      diagnostics.push({
        code: 'SELF_DEPENDENCY',
        message: `Dependency ${dependency.id} cannot reference the same task`,
      })
      continue
    }
    if (!Number.isInteger(dependency.lagMinutes)) {
      diagnostics.push({
        code: 'INVALID_LAG',
        message: `Dependency ${dependency.id} has a non-integer lag`,
      })
      continue
    }
    const key = `${dependency.predecessorId}->${dependency.successorId}:${dependency.type}`
    if (linkKeys.has(key)) {
      diagnostics.push({
        code: 'DUPLICATE_DEPENDENCY_LINK',
        message: `Dependency ${dependency.id} duplicates a link between ${dependency.predecessorId} and ${dependency.successorId}`,
      })
      continue
    }
    linkKeys.add(key)
    edges.push({
      predecessor: dependency.predecessorId as string,
      successor: dependency.successorId as string,
    })
  }

  // Dependencies between a summary task and its own descendants are rejected:
  // the summary roll-up would create a scheduling fixpoint.
  const isAncestor = (ancestor: string, candidate: string): boolean => {
    let current = taskById.get(candidate)
    while (current?.parentTaskId) {
      if (current.parentTaskId === ancestor) return true
      current = taskById.get(current.parentTaskId)
    }
    return false
  }
  for (const edge of edges) {
    if (
      isAncestor(edge.predecessor, edge.successor) ||
      isAncestor(edge.successor, edge.predecessor)
    ) {
      diagnostics.push({
        code: 'SUMMARY_DEPENDENCY',
        message: `Dependency between ${edge.predecessor} and ${edge.successor} connects a summary task with its own descendant`,
      })
    }
  }

  // Dependency-cycle detection (Kahn's algorithm over the raw edges).
  const indegree = new Map<string, number>()
  const adjacency = new Map<string, string[]>()
  for (const task of document.tasks) {
    indegree.set(task.id as string, 0)
    adjacency.set(task.id as string, [])
  }
  for (const edge of edges) {
    adjacency.get(edge.predecessor)!.push(edge.successor)
    indegree.set(edge.successor, (indegree.get(edge.successor) ?? 0) + 1)
  }
  const queue = [...indegree.entries()].filter(([, degree]) => degree === 0).map(([id]) => id)
  let processed = 0
  while (queue.length) {
    const id = queue.shift()!
    processed += 1
    for (const next of adjacency.get(id) ?? []) {
      const degree = indegree.get(next)! - 1
      indegree.set(next, degree)
      if (degree === 0) queue.push(next)
    }
  }
  if (processed !== document.tasks.length && edges.length > 0) {
    diagnostics.push({
      code: 'DEPENDENCY_CYCLE',
      message: 'Dependency graph contains a cycle',
    })
  }
}

function validateResourcesAndAssignments(
  document: ProjectDocument,
  taskIds: Set<string>,
  diagnostics: ValidationDiagnostic[],
) {
  const resourceIds = checkDuplicateIds(
    document.resources,
    'resource',
    'DUPLICATE_RESOURCE_ID',
    diagnostics,
  )
  checkDuplicateIds(document.assignments, 'assignment', 'DUPLICATE_ASSIGNMENT_ID', diagnostics)
  const calendarIds = new Set(document.calendars.map((calendar) => calendar.id as string))
  for (const resource of document.resources) {
    if (resource.calendarId && !calendarIds.has(resource.calendarId)) {
      diagnostics.push({
        code: 'MISSING_CALENDAR',
        message: `Resource ${resource.id} references missing calendar ${resource.calendarId}`,
      })
    }
    for (const slot of resource.availability) {
      if (!isValidDate(slot.start) || (slot.finish !== undefined && !isValidDate(slot.finish))) {
        diagnostics.push({
          code: 'INVALID_DATE',
          message: `Resource ${resource.id} has a malformed availability date`,
        })
      }
    }
  }
  for (const assignment of document.assignments) {
    if (!taskIds.has(assignment.taskId)) {
      diagnostics.push({
        code: 'MISSING_TASK_REFERENCE',
        message: `Assignment ${assignment.id} references missing task ${assignment.taskId}`,
      })
    }
    if (!resourceIds.has(assignment.resourceId)) {
      diagnostics.push({
        code: 'MISSING_RESOURCE_REFERENCE',
        message: `Assignment ${assignment.id} references missing resource ${assignment.resourceId}`,
      })
    }
  }
}

export function validateProjectDocument(document: ProjectDocument): ProjectCommandResult {
  const diagnostics: ValidationDiagnostic[] = []

  for (const dateField of [
    document.properties.startDate,
    document.properties.finishDate,
    document.properties.statusDate,
  ]) {
    if (dateField !== undefined && !isValidDate(dateField)) {
      diagnostics.push({
        code: 'INVALID_DATE',
        message: `Project properties contain a malformed date ${dateField}`,
      })
    }
  }

  validateCalendars(document, diagnostics)
  const { taskIds, taskById } = validateTasks(document, diagnostics)
  validateDependencies(document, taskIds, taskById, diagnostics)
  validateResourcesAndAssignments(document, taskIds, diagnostics)

  checkDuplicateIds(document.baselines, 'baseline', 'DUPLICATE_BASELINE_ID', diagnostics)
  checkDuplicateIds(document.customFields, 'custom field', 'DUPLICATE_CUSTOM_FIELD_ID', diagnostics)
  // PROJECT-009: baseline snapshots are keyed by TaskId. Every snapshot key
  // MUST reference an existing task so a baseline never carries a dangling
  // reference (the reverse of the MISSING_BASELINE_REFERENCE check above,
  // which ensures a task's `baseline` array references real baselines).
  for (const baseline of document.baselines) {
    if (!isValidDate(baseline.capturedAt)) {
      diagnostics.push({
        code: 'INVALID_DATE',
        message: `Baseline ${baseline.id} has a malformed capturedAt date`,
      })
    }
    for (const taskKey of Object.keys(baseline.taskSnapshots)) {
      if (!taskIds.has(taskKey)) {
        diagnostics.push({
          code: 'MISSING_TASK_REFERENCE',
          message: `Baseline ${baseline.id} references missing task ${taskKey}`,
        })
      }
    }
  }

  return {
    commandId: 'validation',
    accepted: diagnostics.length === 0,
    diagnostics,
    affectedTaskIds: [],
  }
}

export function affectedTaskIds(command: ProjectCommand): TaskId[] {
  switch (command.type) {
    case 'CreateTask':
      return [command.task.id]
    case 'DeleteTask':
    case 'RenameTask':
    case 'IndentTask':
    case 'OutdentTask':
    case 'SetTaskDuration':
    case 'SetTaskStart':
    case 'SetTaskFinish':
    case 'SetConstraint':
    case 'SetDeadline':
    case 'SetPercentComplete':
      return [command.taskId]
    case 'AddDependency':
      return [command.dependency.predecessorId, command.dependency.successorId]
    case 'ChangeDependencyType':
    case 'ChangeLag':
      return []
    case 'RemoveDependency':
      return []
    case 'AssignResource':
      return [command.assignment.taskId]
    case 'UnassignResource':
      return []
    case 'CreateBaseline':
      return Object.keys(command.baseline.taskSnapshots) as TaskId[]
    case 'LevelResources':
      return command.taskIds ?? []
  }
}
