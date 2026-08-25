import { describe, expect, it } from 'vitest'
import type { ProjectDocument, TaskId } from '../src/index.js'

describe('project contracts', () => {
  it('uses stable branded identities and a canonical document shape', () => {
    const id = 'task-1' as TaskId
    const document: ProjectDocument = {
      schemaVersion: 1,
      properties: { id: 'p1', name: 'Demo', startDate: '2026-01-01T09:00:00.000Z' as any, defaultCalendarId: 'cal-1' as any },
      tasks: [], resources: [], assignments: [], dependencies: [], calendars: [], baselines: [], customFields: [], views: [], tables: [], filters: [], groups: [],
    }
    expect(id).toBe('task-1')
    expect(document.schemaVersion).toBe(1)
  })
})
