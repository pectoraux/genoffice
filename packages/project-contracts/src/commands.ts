import type {
  AssignmentId,
  DependencyId,
  DependencyType,
  TaskId,
  WorkingMinutes,
  ProjectDocument,
} from './types.js'

export type ProjectCommand =
  | { type: 'CreateTask'; task: ProjectDocument['tasks'][number] }
  | { type: 'DeleteTask'; taskId: TaskId }
  | { type: 'RenameTask'; taskId: TaskId; name: string }
  | { type: 'IndentTask'; taskId: TaskId; parentTaskId: TaskId }
  | { type: 'OutdentTask'; taskId: TaskId }
  | { type: 'AddDependency'; dependency: ProjectDocument['dependencies'][number] }
  | { type: 'RemoveDependency'; dependencyId: DependencyId }
  | { type: 'ChangeDependencyType'; dependencyId: DependencyId; dependencyType: DependencyType }
  | { type: 'ChangeLag'; dependencyId: DependencyId; lagMinutes: number }
  | { type: 'SetTaskDuration'; taskId: TaskId; duration: WorkingMinutes }
  | { type: 'SetTaskStart'; taskId: TaskId; start: string }
  | { type: 'SetTaskFinish'; taskId: TaskId; finish: string }
  | {
      type: 'SetConstraint'
      taskId: TaskId
      constraintType: ProjectDocument['tasks'][number]['constraintType']
      constraintDate?: string
    }
  | { type: 'SetDeadline'; taskId: TaskId; deadline?: string }
  | { type: 'AssignResource'; assignment: ProjectDocument['assignments'][number] }
  | { type: 'UnassignResource'; assignmentId: AssignmentId }
  | { type: 'SetPercentComplete'; taskId: TaskId; percentComplete: number }
  | { type: 'CreateBaseline'; baseline: ProjectDocument['baselines'][number] }
  | { type: 'LevelResources'; taskIds?: TaskId[] }

export interface ProjectCommandResult {
  commandId: string
  accepted: boolean
  diagnostics: Array<{ code: string; message: string }>
  affectedTaskIds: TaskId[]
  inverse?: ProjectCommand
}

export interface JournalEntry {
  commandId: string
  command: ProjectCommand
  result: ProjectCommandResult
}
