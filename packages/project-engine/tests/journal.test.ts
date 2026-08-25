import { describe, expect, it } from 'vitest'
import { ProjectJournal } from '../src/index.js'

const command = { type: 'RenameTask', taskId: 't1' as any, name: 'Renamed' } as const
const result = { commandId: 'c1', accepted: true, diagnostics: [], affectedTaskIds: ['t1'] as any[] }

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
