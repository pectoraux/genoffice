import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import type { ProjectCommand, ProjectCommandResult } from '@genoffice/project-contracts'
import { ProjectJournal } from '../src/index.js'

const command: ProjectCommand = { type: 'RenameTask', taskId: asTaskId('t1'), name: 'Renamed' }
const result: ProjectCommandResult = {
  commandId: 'c1',
  accepted: true,
  diagnostics: [],
  affectedTaskIds: [asTaskId('t1')],
}

describe('ProjectJournal', () => {
  it('clears redo history after a new accepted command', () => {
    const journal = new ProjectJournal()
    journal.record(command, result)
    expect(journal.canUndo()).toBe(true)
    journal.undo()
    expect(journal.canRedo()).toBe(true)
    journal.record(command, result)
    expect(journal.canRedo()).toBe(false)
  })
})
