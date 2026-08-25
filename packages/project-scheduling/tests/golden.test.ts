import { describe, expect, it } from 'vitest'
import { schedule } from '../src/index.js'
import type { ProjectDocument, TaskId } from '@genoffice/project-contracts'

const t = (id: string, name: string, duration: number, parentTaskId?: string): any => ({ id, uid: Number(id.replace(/\D/g, '')) || 1, wbs: '', outlineLevel: parentTaskId ? 2 : 1, name, taskType: 'fixedDuration', summary: false, milestone: duration === 0, manualScheduled: false, autoScheduled: true, duration, priority: 500, percentComplete: 0, work: duration, remainingWork: duration, actualWork: 0, cost: 0, actualCost: 0, remainingCost: 0, baseline: [], customFields: {}, notes: [], parentTaskId })
const calendar: any = { id: 'cal-1', name: 'Standard', workingWeek: { 0: [], 1: [{ startMinute: 540, endMinute: 1020 }], 2: [{ startMinute: 540, endMinute: 1020 }], 3: [{ startMinute: 540, endMinute: 1020 }], 4: [{ startMinute: 540, endMinute: 1020 }], 5: [{ startMinute: 540, endMinute: 1020 }], 6: [] }, exceptions: [] }
function doc(tasks: any[], dependencies: any[] = [], calendars = [calendar]): ProjectDocument { return { schemaVersion: 1, properties: { id: 'p1', name: 'Golden', startDate: '2026-08-03T09:00:00.000Z' as any, defaultCalendarId: 'cal-1' as any }, tasks, resources: [], assignments: [], dependencies, calendars, baselines: [], customFields: [], views: [], tables: [], filters: [], groups: [] } }

describe('deterministic scheduling goldens', () => {
  it('schedules a sequential FS chain across a working-day boundary', () => {
    const project = doc([t('t1', 'A', 480), t('t2', 'B', 480)], [{ id: 'd1', predecessorId: 't1', successorId: 't2', type: 'FS', lagMinutes: 0 }])
    const result = schedule(project)
    expect(result.taskSchedules['t1' as TaskId].scheduledFinish).toBe('2026-08-03T17:00:00.000Z')
    expect(result.taskSchedules['t2' as TaskId].scheduledStart).toBe('2026-08-04T09:00:00.000Z')
  })

  it('supports SS, FF, SF and lag deterministically', () => {
    const project = doc([
      t('t1', 'A', 240), t('t2', 'B', 480), t('t3', 'C', 120), t('t4', 'D', 60), t('t5', 'E', 60),
    ], [
      { id: 'd1', predecessorId: 't1', successorId: 't2', type: 'SS', lagMinutes: 60 },
      { id: 'd2', predecessorId: 't1', successorId: 't3', type: 'FF', lagMinutes: 60 },
      { id: 'd3', predecessorId: 't1', successorId: 't4', type: 'SF', lagMinutes: 0 },
      { id: 'd4', predecessorId: 't2', successorId: 't5', type: 'FS', lagMinutes: 60 },
    ])
    const a = schedule(project)
    const b = schedule(project)
    expect(a).toEqual(b)
    expect(a.taskSchedules['t2' as TaskId].scheduledStart).toBe('2026-08-03T10:00:00.000Z')
    expect(a.taskSchedules['t5' as TaskId].scheduledStart).toBe('2026-08-04T10:00:00.000Z')
  })

  it('skips a holiday exception and identifies a critical chain', () => {
    const holidayCalendar = { ...calendar, exceptions: [{ date: '2026-08-04', periods: [] }] }
    const project = doc([t('t1', 'A', 480), t('t2', 'B', 480)], [{ id: 'd1', predecessorId: 't1', successorId: 't2', type: 'FS', lagMinutes: 0 }], [holidayCalendar])
    const result = schedule(project)
    expect(result.taskSchedules['t2' as TaskId].scheduledStart).toBe('2026-08-05T09:00:00.000Z')
    expect(result.taskSchedules['t1' as TaskId].critical).toBe(true)
    expect(result.taskSchedules['t2' as TaskId].critical).toBe(true)
  })
})
