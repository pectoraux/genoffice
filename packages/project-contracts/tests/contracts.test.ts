import { describe, expect, it } from 'vitest'
import {
  asCalendarId,
  asISODateTime,
  asTaskId,
  plainMinutes,
  asWorkingMinutes,
} from '../src/index.js'
import type { ProjectDocument } from '../src/index.js'

describe('project contracts', () => {
  it('uses stable branded identities and a canonical document shape', () => {
    const id = asTaskId('task-1')
    const document: ProjectDocument = {
      schemaVersion: 1,
      properties: {
        id: 'p1',
        name: 'Demo',
        startDate: asISODateTime('2026-01-01T09:00:00.000Z'),
        defaultCalendarId: asCalendarId('cal-1'),
      },
      tasks: [],
      resources: [],
      assignments: [],
      dependencies: [],
      calendars: [],
      baselines: [],
      customFields: [],
      views: [],
      tables: [],
      filters: [],
      groups: [],
    }
    expect(id).toBe('task-1')
    expect(document.schemaVersion).toBe(1)
    expect(document.properties.defaultCalendarId).toBe('cal-1')
  })

  it('round-trips working-minute brand conversions', () => {
    expect(plainMinutes(asWorkingMinutes(480))).toBe(480)
    expect(plainMinutes(asWorkingMinutes(0))).toBe(0)
  })
})
