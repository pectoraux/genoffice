import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId } from '@genoffice/project-contracts'
import { createViewState, reconcileViewState, parseInstant, formatInstant } from '../src/index.js'
import {
  makeDependency,
  makeDocument,
  makeResource,
  makeTask,
  outlineDocument,
} from './fixtures.js'

describe('PROJECT-021 view state creation', () => {
  it('creates an empty selection/collapse state with a deterministic viewport from the properties window', () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      finishDate: '2026-08-14T17:00:00.000Z',
    })
    const state = createViewState(document)
    expect(state.tasks.taskIds).toEqual([])
    expect(state.dependencies).toEqual([])
    expect(state.resources).toEqual([])
    expect(state.collapsed).toEqual([])
    expect(state.viewport.start).toBe('2026-08-03T09:00:00.000Z')
    expect(state.viewport.finish).toBe('2026-08-14T17:00:00.000Z')
  })

  it('defaults to a 30-day window when no finish information exists', () => {
    const document = makeDocument({ startDate: '2026-08-03T09:00:00.000Z' })
    const state = createViewState(document)
    const start = parseInstant(state.viewport.start)
    const finish = parseInstant(state.viewport.finish)
    expect(start).toBe(Date.parse('2026-08-03T09:00:00.000Z'))
    expect(finish! - start!).toBe(30 * 24 * 60 * 60 * 1000)
  })

  it('prefers the derived schedule project window over the properties window', () => {
    const document = makeDocument({
      startDate: '2026-08-03T09:00:00.000Z',
      finishDate: '2026-12-31T17:00:00.000Z',
    })
    const schedule = {
      taskSchedules: {},
      projectStart: asISODateTime('2026-08-05T09:00:00.000Z'),
      projectFinish: asISODateTime('2026-08-21T17:00:00.000Z'),
      diagnostics: [],
    }
    const state = createViewState(document, schedule)
    expect(state.viewport.start).toBe('2026-08-05T09:00:00.000Z')
    expect(state.viewport.finish).toBe('2026-08-21T17:00:00.000Z')
  })

  it('is deterministic across repeated creation (3×)', () => {
    const document = outlineDocument()
    const first = JSON.stringify(createViewState(document))
    expect(JSON.stringify(createViewState(document))).toBe(first)
    expect(JSON.stringify(createViewState(document))).toBe(first)
  })
})

describe('PROJECT-021 view state reconciliation', () => {
  it('drops task selection, collapse, and anchor/focus references that no longer exist', () => {
    const document = outlineDocument()
    const state = {
      ...createViewState(document),
      tasks: {
        taskIds: [asTaskId('a'), asTaskId('gone')],
        anchorId: asTaskId('gone'),
        focusId: asTaskId('a'),
      },
      collapsed: [asTaskId('gone'), asTaskId('b')],
      dependencies: [],
      resources: [],
    }
    const reconciled = reconcileViewState(state, document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect('anchorId' in reconciled.tasks).toBe(false)
    expect(reconciled.tasks.focusId).toBe(asTaskId('a'))
    expect(reconciled.collapsed).toEqual([asTaskId('b')])
  })

  it('drops dependency and resource selections that no longer exist, preserving order', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' }), makeTask({ id: 't2' })],
      resources: [makeResource({ id: 'r1', name: 'R1' })],
      dependencies: [makeDependency('d1', 't1', 't2'), makeDependency('d2', 't1', 't2')],
    })
    const state = {
      ...createViewState(document),
      dependencies: ['d1', 'd-gone', 'd2'] as never,
      resources: ['r-gone'] as never,
    }
    const reconciled = reconcileViewState(state, document)
    expect(reconciled.dependencies).toEqual(['d1', 'd2'] as never)
    expect(reconciled.resources).toEqual([])
  })

  it('keeps active view references that still resolve and clears vanished ones', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      views: [{ id: 'v1' as never, name: 'Gantt', type: 'gantt' }],
      tables: [],
      filters: [],
      groups: [],
    })
    const state = {
      ...createViewState(document),
      activeViewId: 'v1' as never,
      activeTableId: 'tv-gone' as never,
    }
    const reconciled = reconcileViewState(state, document)
    expect(reconciled.activeViewId).toBe('v1' as never)
    expect('activeTableId' in reconciled).toBe(false)
  })

  it('leaves the viewport untouched (time is not an entity reference)', () => {
    const document = outlineDocument()
    const state = createViewState(document)
    const reconciled = reconcileViewState(state, makeDocument())
    expect(reconciled.viewport).toEqual(state.viewport)
  })
})

describe('PROJECT-021 instant helpers', () => {
  it('parseInstant rejects non-instants and formatInstant round-trips UTC milliseconds', () => {
    expect(parseInstant('2026-08-03T09:00:00.000Z')).toBe(Date.parse('2026-08-03T09:00:00.000Z'))
    expect(parseInstant('not-a-date')).toBeUndefined()
    expect(parseInstant('')).toBeUndefined()
    expect(formatInstant(Date.parse('2026-08-03T09:00:00.000Z'))).toBe('2026-08-03T09:00:00.000Z')
  })
})
