import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId } from '@genoffice/project-contracts'
import { createViewState, reduceViewState } from '../src/index.js'
import { makeDocument, makeTask, multiSiblingDocument, outlineDocument } from './fixtures.js'

const context = (document = outlineDocument()) => ({ document })

describe('PROJECT-021 reducer — task selection', () => {
  it('set mode selects a single task with anchor and focus', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'selectTask', taskId: asTaskId('a') }, context(document))
    expect(state.tasks.taskIds).toEqual([asTaskId('a')])
    expect(state.tasks.anchorId).toBe(asTaskId('a'))
    expect(state.tasks.focusId).toBe(asTaskId('a'))
  })

  it('toggle mode adds then removes, keeping anchor/focus consistent', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a'), mode: 'toggle' },
      context(document),
    )
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('b'), mode: 'toggle' },
      context(document),
    )
    expect(state.tasks.taskIds).toEqual([asTaskId('a'), asTaskId('b')])
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a'), mode: 'toggle' },
      context(document),
    )
    expect(state.tasks.taskIds).toEqual([asTaskId('b')])
    // Removing a task re-anchors to the last remaining selection.
    expect(state.tasks.anchorId).toBe(asTaskId('b'))
    // Removing the last selected task clears anchor and focus.
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('b'), mode: 'toggle' },
      context(document),
    )
    expect(state.tasks.taskIds).toEqual([])
    expect('anchorId' in state.tasks).toBe(false)
    expect('focusId' in state.tasks).toBe(false)
  })

  it('extend mode selects the canonical outline-order range from the anchor (both directions)', () => {
    const document = outlineDocument() // root, a, a1, b
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'selectTask', taskId: asTaskId('a') }, context(document))
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('b'), mode: 'extend' },
      context(document),
    )
    expect(state.tasks.taskIds).toEqual([asTaskId('a'), asTaskId('a1'), asTaskId('b')])
    expect(state.tasks.anchorId).toBe(asTaskId('a'))
    expect(state.tasks.focusId).toBe(asTaskId('b'))
    // Extending backwards from the anchor covers the range before it
    // (anchor stays `a`; the range root..a replaces the selection).
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('root'), mode: 'extend' },
      context(document),
    )
    expect(state.tasks.taskIds).toEqual([asTaskId('root'), asTaskId('a')])
    expect(state.tasks.anchorId).toBe(asTaskId('a'))
    expect(state.tasks.focusId).toBe(asTaskId('root'))
  })

  it('selectTasks sets the exact validated selection in first-occurrence order', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      {
        type: 'selectTasks',
        taskIds: [asTaskId('b'), asTaskId('gone'), asTaskId('a'), asTaskId('b')],
      },
      context(document),
    )
    expect(state.tasks.taskIds).toEqual([asTaskId('b'), asTaskId('a')])
    expect(state.tasks.anchorId).toBe(asTaskId('a'))
    expect(state.tasks.focusId).toBe(asTaskId('a'))
  })

  it('clearSelection clears task, dependency, and resource selections', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(state, { type: 'selectTask', taskId: asTaskId('a') }, context(document))
    state = reduceViewState(state, { type: 'clearSelection' }, context(document))
    expect(state.tasks.taskIds).toEqual([])
    expect(state.dependencies).toEqual([])
    expect(state.resources).toEqual([])
  })

  it('ignores intents referencing unknown tasks (deterministic no-op, same reference)', () => {
    const document = outlineDocument()
    const state = createViewState(document)
    const next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('nope') },
      context(document),
    )
    expect(next).toBe(state)
  })
})

describe('PROJECT-021 reducer — dependency/resource selection', () => {
  it('selects and toggles dependencies against the document', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' }), makeTask({ id: 't2' })],
      dependencies: [
        {
          id: 'd1' as never,
          predecessorId: asTaskId('t1'),
          successorId: asTaskId('t2'),
          type: 'FS',
          lagMinutes: 0,
        },
      ],
    })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: 'd1' as never },
      context(document),
    )
    expect(state.dependencies).toEqual(['d1' as never])
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: 'd1' as never, mode: 'toggle' },
      context(document),
    )
    expect(state.dependencies).toEqual([])
    const unchanged = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: 'dX' as never },
      context(document),
    )
    expect(unchanged).toBe(state)
  })
})

describe('PROJECT-021 reducer — collapse', () => {
  it('toggles collapse per task and ignores unknown tasks', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      context(document),
    )
    expect(state.collapsed).toEqual([asTaskId('a')])
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      context(document),
    )
    expect(state.collapsed).toEqual([])
    const unchanged = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('zzz') },
      context(document),
    )
    expect(unchanged).toBe(state)
  })

  it('setCollapsed adds/removes validated sets; collapseAll collapses exactly the summaries', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setCollapsed', taskIds: [asTaskId('root'), asTaskId('gone')], collapsed: true },
      context(document),
    )
    expect(state.collapsed).toEqual([asTaskId('root')])
    state = reduceViewState(
      state,
      { type: 'setCollapsed', taskIds: [asTaskId('root')], collapsed: false },
      context(document),
    )
    expect(state.collapsed).toEqual([])
    state = reduceViewState(state, { type: 'collapseAll' }, context(document))
    expect(state.collapsed).toEqual([asTaskId('root'), asTaskId('a')])
    state = reduceViewState(state, { type: 'expandAll' }, context(document))
    expect(state.collapsed).toEqual([])
  })

  it('toggleCollapse on a LEAF task is a deterministic no-op (collapsed ⊆ summaries)', () => {
    const document = outlineDocument() // leaves: a1, b; summaries: root, a
    const state = createViewState(document)
    for (const leaf of [asTaskId('a1'), asTaskId('b')]) {
      const unchanged = reduceViewState(
        state,
        { type: 'toggleCollapse', taskId: leaf },
        context(document),
      )
      expect(unchanged).toBe(state) // reference-equal: the state is returned untouched
    }
  })

  it('setCollapsed ignores LEAF ids in both directions (collapsed ⊆ summaries)', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      {
        type: 'setCollapsed',
        taskIds: [asTaskId('root'), asTaskId('a1'), asTaskId('b')],
        collapsed: true,
      },
      context(document),
    )
    expect(state.collapsed).toEqual([asTaskId('root')]) // leaves filtered out
    // Removal requests naming leaves are no-ops under the same invariant:
    state = reduceViewState(
      state,
      { type: 'setCollapsed', taskIds: [asTaskId('b')], collapsed: false },
      context(document),
    )
    expect(state.collapsed).toEqual([asTaskId('root')])
  })

  it('collapseAll selects exactly the summaries — leaves can never enter the set', () => {
    const document = multiSiblingDocument() // summaries: root1, p; leaves: a1, a2, a3, root2
    const state = reduceViewState(
      createViewState(document),
      { type: 'collapseAll' },
      context(document),
    )
    expect(state.collapsed).toEqual([asTaskId('root1'), asTaskId('p')])
    for (const leaf of [asTaskId('a1'), asTaskId('a2'), asTaskId('a3'), asTaskId('root2')]) {
      expect(state.collapsed).not.toContain(leaf)
    }
  })

  it('reconciles a collapsed summary that became a LEAF (subtree deleted, summary recomputed)', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      context(document),
    )
    expect(state.collapsed).toEqual([asTaskId('a')])
    // The document loses a's subtree: the engine recomputed `summary`, so `a`
    // is now a leaf and can no longer stay in the collapsed set.
    const afterDeletion = makeDocument({
      tasks: [
        makeTask({ id: 'root', outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({ id: 'a', parentTaskId: asTaskId('root'), outlineLevel: 2, wbs: '1.1' }),
        makeTask({ id: 'b', parentTaskId: asTaskId('root'), outlineLevel: 2, wbs: '1.2' }),
      ],
    })
    // The reducer applies reconciliation automatically on the next intent
    // dispatched against the replaced document — the leaf `a` is pruned:
    state = reduceViewState(state, { type: 'clearSelection' }, { document: afterDeletion })
    expect(state.collapsed).toEqual([])
  })
})

describe('PROJECT-021 reducer — viewport', () => {
  it('setViewport accepts well-formed windows and rejects malformed/degenerate ones', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      {
        type: 'setViewport',
        start: '2026-08-03T00:00:00.000Z',
        finish: '2026-09-03T00:00:00.000Z',
      },
      context(document),
    )
    expect(state.viewport.start).toBe('2026-08-03T00:00:00.000Z')
    const before = state
    const reversed = reduceViewState(
      state,
      {
        type: 'setViewport',
        start: '2026-09-03T00:00:00.000Z',
        finish: '2026-08-03T00:00:00.000Z',
      },
      context(document),
    )
    expect(reversed).toBe(before)
    const unparseable = reduceViewState(
      state,
      { type: 'setViewport', start: 'nope', finish: '2026-09-03T00:00:00.000Z' },
      context(document),
    )
    expect(unparseable).toBe(before)
  })

  it('scaleViewport zooms around the focus instant and keeps it stationary', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      {
        type: 'setViewport',
        start: '2026-08-01T00:00:00.000Z',
        finish: '2026-08-31T00:00:00.000Z',
      },
      context(document),
    )
    state = reduceViewState(
      state,
      { type: 'scaleViewport', factor: 0.5, focus: '2026-08-15T00:00:00.000Z' },
      context(document),
    )
    const span = Date.parse(state.viewport.finish) - Date.parse(state.viewport.start)
    expect(span).toBe(15 * 24 * 60 * 60 * 1000)
    // The focus instant stays stationary: it was 14/30 into the window, so it
    // stays 14/30 into the halved window (Aug 8 + 7d = Aug 15).
    expect(Date.parse(state.viewport.start)).toBe(Date.parse('2026-08-08T00:00:00.000Z'))
    expect(Date.parse(state.viewport.finish)).toBe(Date.parse('2026-08-23T00:00:00.000Z'))
  })

  it('scaleViewport ignores non-positive factors (deterministic no-op)', () => {
    const document = outlineDocument()
    const state = createViewState(document)
    expect(reduceViewState(state, { type: 'scaleViewport', factor: 0 }, context(document))).toBe(
      state,
    )
    expect(reduceViewState(state, { type: 'scaleViewport', factor: -2 }, context(document))).toBe(
      state,
    )
    expect(
      reduceViewState(state, { type: 'scaleViewport', factor: Number.NaN }, context(document)),
    ).toBe(state)
  })

  it('fitViewport fits the derived schedule window with deterministic padding', () => {
    const document = makeDocument({ startDate: '2026-08-03T09:00:00.000Z' })
    const schedule = {
      taskSchedules: {},
      projectStart: asISODateTime('2026-08-10T09:00:00.000Z'),
      projectFinish: asISODateTime('2026-08-20T09:00:00.000Z'),
      diagnostics: [],
    }
    const state = reduceViewState(
      createViewState(document),
      { type: 'fitViewport' },
      { document, schedule },
    )
    // 10-day span → 2% padding each side = 4.8h.
    expect(state.viewport.start).toBe('2026-08-10T04:12:00.000Z')
    expect(state.viewport.finish).toBe('2026-08-20T13:48:00.000Z')
  })
})

describe('PROJECT-021 reducer — active view references', () => {
  it('sets and clears active view definitions, ignoring unknown ids', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 't1' })],
      views: [{ id: 'v1' as never, name: 'Gantt', type: 'gantt' }],
      tables: [{ id: 'tb1' as never, name: 'Entry', columns: [] }],
    })
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'setActiveView', viewId: 'v1' as never },
      context(document),
    )
    expect(state.activeViewId).toBe('v1' as never)
    state = reduceViewState(
      state,
      { type: 'setActiveTable', tableId: 'tb1' as never },
      context(document),
    )
    expect(state.activeTableId).toBe('tb1' as never)
    const unchanged = reduceViewState(
      state,
      { type: 'setActiveView', viewId: 'vX' as never },
      context(document),
    )
    expect(unchanged).toBe(state)
    state = reduceViewState(state, { type: 'setActiveView' }, context(document))
    expect('activeViewId' in state).toBe(false)
  })
})

describe('PROJECT-021 reducer — determinism and reconciliation', () => {
  it('the same intent sequence produces byte-identical states across 3 runs', () => {
    const document = outlineDocument()
    const run = (): string => {
      let state = createViewState(document)
      state = reduceViewState(
        state,
        { type: 'selectTask', taskId: asTaskId('a') },
        context(document),
      )
      state = reduceViewState(
        state,
        { type: 'selectTask', taskId: asTaskId('b'), mode: 'extend' },
        context(document),
      )
      state = reduceViewState(
        state,
        { type: 'toggleCollapse', taskId: asTaskId('root') },
        context(document),
      )
      state = reduceViewState(state, { type: 'scaleViewport', factor: 0.5 }, context(document))
      return JSON.stringify(state)
    }
    expect(run()).toBe(run())
    expect(run()).toBe(run())
  })

  it('reducing against a smaller document reconciles dead references immediately', () => {
    const document = outlineDocument()
    let state = createViewState(document)
    state = reduceViewState(
      state,
      { type: 'selectTasks', taskIds: [asTaskId('a'), asTaskId('b')] },
      context(document),
    )
    state = reduceViewState(
      state,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      context(document),
    )
    const smaller = makeDocument({ tasks: [makeTask({ id: 'root' })] })
    state = reduceViewState(state, { type: 'expandAll' }, { document: smaller })
    expect(state.tasks.taskIds).toEqual([])
    expect(state.collapsed).toEqual([])
  })
})
