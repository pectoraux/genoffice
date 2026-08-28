import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId, type TaskId } from '@genoffice/project-contracts'
import { createViewState, reconcileViewState, parseInstant, formatInstant } from '../src/index.js'
import {
  makeDependency,
  makeDocument,
  makeResource,
  makeTask,
  outlineDocument,
} from './fixtures.js'

/** The documented `TaskSelection` contract (PROJECT-021 review round 2):
 * `anchorId`/`focusId` are ALWAYS members of `taskIds` when present —
 * `anchorId === undefined || taskIds.includes(anchorId)` (and likewise for
 * `focusId`). Reconciliation must restore this invariant from ANY input,
 * including externally restored or malformed states. */
function expectSelectionInvariant(selection: {
  readonly taskIds: readonly TaskId[]
  readonly anchorId?: TaskId
  readonly focusId?: TaskId
}): void {
  expect(
    selection.anchorId === undefined || selection.taskIds.includes(selection.anchorId),
    `anchorId ${String(selection.anchorId)} must be undefined or a member of the surviving selection`,
  ).toBe(true)
  expect(
    selection.focusId === undefined || selection.taskIds.includes(selection.focusId),
    `focusId ${String(selection.focusId)} must be undefined or a member of the surviving selection`,
  ).toBe(true)
}

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
      collapsed: [asTaskId('gone'), asTaskId('b'), asTaskId('a')],
      dependencies: [],
      resources: [],
    }
    const reconciled = reconcileViewState(state, document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect('anchorId' in reconciled.tasks).toBe(false)
    expect(reconciled.tasks.focusId).toBe(asTaskId('a'))
    // `gone` does not exist and `b` is a LEAF (collapsed ⊆ summaries);
    // only the surviving summary `a` stays:
    expect(reconciled.collapsed).toEqual([asTaskId('a')])
  })

  it('prunes collapsed LEAF entries even when the task still exists (collapsed ⊆ summaries)', () => {
    const document = outlineDocument() // summaries: root, a; leaves: a1, b
    // A restored host state that predates the invariant (persisted view
    // state is host-owned and may be stale — PROJECT-021 persistence rule):
    const restored = {
      ...createViewState(document),
      collapsed: [asTaskId('b'), asTaskId('root'), asTaskId('a1')],
    }
    const reconciled = reconcileViewState(restored, document)
    expect(reconciled.collapsed).toEqual([asTaskId('root')]) // leaves pruned
  })

  it('prunes a collapsed summary whose subtree was deleted (engine recomputed summary → leaf)', () => {
    const document = outlineDocument()
    const collapsedState = {
      ...createViewState(document),
      collapsed: [asTaskId('a')], // `a` is a summary here (a > a1)
    }
    const afterDeletion = makeDocument({
      tasks: [
        makeTask({ id: 'root', outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({ id: 'a', parentTaskId: asTaskId('root'), outlineLevel: 2, wbs: '1.1' }),
        makeTask({ id: 'b', parentTaskId: asTaskId('root'), outlineLevel: 2, wbs: '1.2' }),
      ],
    })
    const reconciled = reconcileViewState(collapsedState, afterDeletion)
    expect(reconciled.collapsed).toEqual([]) // `a` is a leaf now — pruned
  })

  it('drops a LIVE but UNSELECTED anchor (selection membership, not mere document existence)', () => {
    const document = outlineDocument() // live tasks: root, a, a1, b
    // `b` exists in the document but is NOT in the selection — a state shape
    // only external restoration (or malformed input) can produce, and one
    // the old document-existence check wrongly kept:
    const state = {
      ...createViewState(document),
      tasks: { taskIds: [asTaskId('a')], anchorId: asTaskId('b') },
    }
    const reconciled = reconcileViewState(state, document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect('anchorId' in reconciled.tasks).toBe(false)
    expectSelectionInvariant(reconciled.tasks)
  })

  it('drops a LIVE but UNSELECTED focus the same way', () => {
    const document = outlineDocument()
    const state = {
      ...createViewState(document),
      tasks: { taskIds: [asTaskId('a')], focusId: asTaskId('b') },
    }
    const reconciled = reconcileViewState(state, document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect('focusId' in reconciled.tasks).toBe(false)
    expectSelectionInvariant(reconciled.tasks)
  })

  // ---- PROJECT-031 — the focused CELL (focusId, focusField) ---------------
  it('keeps a valid focused-cell field and drops one outside the accepted four', () => {
    const document = outlineDocument()
    // A valid field on a focused+selected row survives reconciliation.
    const valid = {
      ...createViewState(document),
      tasks: {
        taskIds: [asTaskId('a')],
        anchorId: asTaskId('a'),
        focusId: asTaskId('a'),
        focusField: 'duration' as const,
      },
    }
    const kept = reconcileViewState(valid, document)
    expect(kept.tasks.focusField).toBe('duration')
    // A restored field outside the four editable fields is dropped (the
    // honest-pruning rule every restored reference follows).
    const malformed = {
      ...createViewState(document),
      tasks: {
        taskIds: [asTaskId('a')],
        anchorId: asTaskId('a'),
        focusId: asTaskId('a'),
        focusField: 'wbs' as never,
      },
    }
    const pruned = reconcileViewState(malformed, document)
    expect('focusField' in pruned.tasks).toBe(false)
  })

  it('drops the focused-cell field when its focused row is dropped (the cell is a PAIR)', () => {
    const document = outlineDocument()
    // `b` is live but NOT selected → the focus row is dropped, and with it
    // the field (a field focus without its row is not a cell).
    const state = {
      ...createViewState(document),
      tasks: { taskIds: [asTaskId('a')], focusId: asTaskId('b'), focusField: 'start' as const },
    }
    const reconciled = reconcileViewState(state, document)
    expect('focusId' in reconciled.tasks).toBe(false)
    expect('focusField' in reconciled.tasks).toBe(false)
    expectSelectionInvariant(reconciled.tasks)
  })

  it('reconciles anchor/focus when a SELECTED task is deleted (independently, keeping surviving members)', () => {
    const document = outlineDocument()
    const state = {
      ...createViewState(document),
      tasks: {
        taskIds: [asTaskId('a'), asTaskId('b')],
        anchorId: asTaskId('b'), // deleted with the document change
        focusId: asTaskId('b'),
      },
    }
    // The document loses `b` (subtree deleted + reloaded):
    const afterDeletion = makeDocument({
      tasks: [
        makeTask({ id: 'root', outlineLevel: 1, summary: true, wbs: '1' }),
        makeTask({
          id: 'a',
          parentTaskId: asTaskId('root'),
          outlineLevel: 2,
          summary: true,
          wbs: '1.1',
        }),
        makeTask({ id: 'a1', parentTaskId: asTaskId('a'), outlineLevel: 3, wbs: '1.1.1' }),
      ],
    })
    const reconciled = reconcileViewState(state, afterDeletion)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect('anchorId' in reconciled.tasks).toBe(false)
    expect('focusId' in reconciled.tasks).toBe(false)
    expectSelectionInvariant(reconciled.tasks)
    // Anchor and focus reconcile INDEPENDENTLY: a deleted anchor with a
    // surviving selected focus keeps the focus (and vice versa).
    const mixed = reconcileViewState(
      {
        ...createViewState(document),
        tasks: {
          taskIds: [asTaskId('a'), asTaskId('b')],
          anchorId: asTaskId('b'), // deleted
          focusId: asTaskId('a'), // survives, stays selected
        },
      },
      afterDeletion,
    )
    expect(mixed.tasks.taskIds).toEqual([asTaskId('a')])
    expect('anchorId' in mixed.tasks).toBe(false)
    expect(mixed.tasks.focusId).toBe(asTaskId('a'))
    expectSelectionInvariant(mixed.tasks)
  })

  it('restores the invariant from an externally malformed restored state (live anchor/focus outside the selection)', () => {
    const document = outlineDocument()
    // The review round-2 counterexample shape: taskIds [a, b] with
    // anchorId/focusId pointing at live-but-unselected tasks. Mere document
    // existence checks would keep them — violating "anchorId is always a
    // member of taskIds" on the state hosts are told to reconcile on restore.
    const restored = {
      ...createViewState(document),
      tasks: {
        taskIds: [asTaskId('a'), asTaskId('b')],
        anchorId: asTaskId('root'), // live in the document, NOT selected
        focusId: asTaskId('a1'), // live in the document, NOT selected
      },
    }
    const reconciled = reconcileViewState(restored, document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a'), asTaskId('b')])
    expect('anchorId' in reconciled.tasks).toBe(false)
    expect('focusId' in reconciled.tasks).toBe(false)
    expectSelectionInvariant(reconciled.tasks)
  })

  it('proves the final selection invariant across a battery of malformed/edge states', () => {
    const document = outlineDocument()
    const battery: Array<{
      readonly taskIds: readonly TaskId[]
      readonly anchorId?: TaskId
      readonly focusId?: TaskId
    }> = [
      // live but unselected anchor+focus:
      { taskIds: [asTaskId('a')], anchorId: asTaskId('b'), focusId: asTaskId('b') },
      // anchor/focus with an EMPTY selection (anchor live in the document):
      { taskIds: [], anchorId: asTaskId('a'), focusId: asTaskId('a') },
      // everything dead:
      { taskIds: [asTaskId('gone')], anchorId: asTaskId('gone'), focusId: asTaskId('gone') },
      // well-formed control (must stay intact):
      { taskIds: [asTaskId('a'), asTaskId('b')], anchorId: asTaskId('a'), focusId: asTaskId('b') },
      // mixed live/dead/selected/unselected:
      {
        taskIds: [asTaskId('a1'), asTaskId('b'), asTaskId('gone')],
        anchorId: asTaskId('a'), // live, unselected
        focusId: asTaskId('root'), // live, unselected
      },
    ]
    for (const tasks of battery) {
      const reconciled = reconcileViewState({ ...createViewState(document), tasks }, document)
      expectSelectionInvariant(reconciled.tasks) // the invariant itself, per case
    }
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
