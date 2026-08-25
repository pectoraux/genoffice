import type { JournalEntry, ProjectCommand, ProjectCommandResult } from '@genoffice/project-contracts'

export class ProjectJournal {
  private past: JournalEntry[] = []
  private future: JournalEntry[] = []

  record(command: ProjectCommand, result: ProjectCommandResult): void {
    if (!result.accepted) return
    this.past.push({ commandId: result.commandId, command, result })
    this.future = []
  }

  canUndo(): boolean { return this.past.length > 0 }
  canRedo(): boolean { return this.future.length > 0 }

  undo(): JournalEntry | undefined {
    const entry = this.past.pop()
    if (!entry) return undefined
    this.future.push(entry)
    return entry
  }

  redo(): JournalEntry | undefined {
    const entry = this.future.pop()
    if (!entry) return undefined
    this.past.push(entry)
    return entry
  }

  entries(): readonly JournalEntry[] { return this.past }
}
