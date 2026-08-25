import type { ProjectDocument, ProjectCommand, ProjectCommandResult, TaskId } from '@genoffice/project-contracts'

export function validateProjectDocument(document: ProjectDocument): ProjectCommandResult {
  const diagnostics: Array<{ code: string; message: string }> = []
  const taskIds = new Set<string>()
  const calendarIds = new Set<string>(document.calendars.map((calendar) => calendar.id))
  const dependencyIds = new Set<string>()
  for (const task of document.tasks) {
    if (taskIds.has(task.id)) diagnostics.push({ code: 'DUPLICATE_TASK_ID', message: `Duplicate task id ${task.id}` })
    taskIds.add(task.id)
    if (task.duration < 0) diagnostics.push({ code: 'NEGATIVE_DURATION', message: `Task ${task.id} has negative duration` })
    if (task.calendarId && !calendarIds.has(task.calendarId)) diagnostics.push({ code: 'MISSING_CALENDAR', message: `Task ${task.id} references missing calendar ${task.calendarId}` })
    if (task.parentTaskId === task.id) diagnostics.push({ code: 'SELF_PARENT', message: `Task ${task.id} cannot parent itself` })
  }
  for (const dependency of document.dependencies) {
    if (dependencyIds.has(dependency.id)) diagnostics.push({ code: 'DUPLICATE_DEPENDENCY_ID', message: `Duplicate dependency id ${dependency.id}` })
    dependencyIds.add(dependency.id)
    if (!taskIds.has(dependency.predecessorId) || !taskIds.has(dependency.successorId)) diagnostics.push({ code: 'MISSING_TASK_REFERENCE', message: `Dependency ${dependency.id} references a missing task` })
    if (dependency.predecessorId === dependency.successorId) diagnostics.push({ code: 'SELF_DEPENDENCY', message: `Dependency ${dependency.id} cannot reference the same task` })
  }
  return { commandId: 'validation', accepted: diagnostics.length === 0, diagnostics, affectedTaskIds: [] }
}

export function affectedTaskIds(command: ProjectCommand): TaskId[] {
  switch (command.type) {
    case 'CreateTask': return [command.task.id]
    case 'DeleteTask': case 'RenameTask': case 'IndentTask': case 'OutdentTask': case 'SetTaskDuration': case 'SetTaskStart': case 'SetTaskFinish': case 'SetConstraint': case 'SetPercentComplete': return [command.taskId]
    case 'AddDependency': return [command.dependency.predecessorId, command.dependency.successorId]
    case 'ChangeDependencyType': case 'ChangeLag': return []
    case 'RemoveDependency': return []
    case 'AssignResource': return [command.assignment.taskId]
    case 'UnassignResource': return []
    case 'CreateBaseline': return Object.keys(command.baseline.taskSnapshots) as TaskId[]
    case 'LevelResources': return command.taskIds ?? []
  }
}
