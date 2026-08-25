import { describe, expect, it } from 'vitest'
import { buildDependencyGraph, DependencyGraphError } from '../src/index.js'

const task = (id: string): any => ({ id, uid: 1, wbs: '', outlineLevel: 1, name: id, taskType: 'fixedDuration', summary: false, milestone: false, manualScheduled: false, autoScheduled: true, duration: 60, priority: 500, percentComplete: 0, work: 60, remainingWork: 60, actualWork: 0, cost: 0, actualCost: 0, remainingCost: 0, baseline: [], customFields: {}, notes: [] })
const base: any = { schemaVersion: 1, properties: { id: 'p', name: 'p', startDate: '2026-01-01T09:00:00.000Z', defaultCalendarId: 'c' }, tasks: [task('a'), task('b')], resources: [], assignments: [], dependencies: [], calendars: [{ id: 'c', name: 'c', workingWeek: {}, exceptions: [] }], baselines: [], customFields: [], views: [], tables: [], filters: [], groups: [] }

describe('dependency graph', () => {
  it('rejects self dependencies and cycles', () => {
    expect(() => buildDependencyGraph({ ...base, dependencies: [{ id: 'd', predecessorId: 'a', successorId: 'a', type: 'FS', lagMinutes: 0 }] })).toThrowError(/self-referential/)
    expect(() => buildDependencyGraph({ ...base, dependencies: [{ id: 'd1', predecessorId: 'a', successorId: 'b', type: 'FS', lagMinutes: 0 }, { id: 'd2', predecessorId: 'b', successorId: 'a', type: 'FS', lagMinutes: 0 }] })).toThrowError(/cycle/i)
  })
})
