/**
 * PROJECT-023 — selection/editing integration with the PROJECT-022 view
 * models and the command/session pipeline.
 *
 * The directive's integration surface: the projection reflects the
 * selection/focus/edit-target state (grid and timeline rows carry the
 * projection rows BY REFERENCE, so the flags flow into every view model);
 * pointer hit-testing resolves to the entity the host then selects through
 * the standard intents; mutations through the session reconcile the
 * selection deterministically (creation keeps it, indent/outdent keep it
 * verbatim); and malformed command values surface UNSUPPORTED_COMMAND (the
 * runtime safety net) with the document and selection untouched.
 *
 * The scheduling authority is the REAL package at the test layer only (the
 * accepted PROJECT-021 precedent).
 */
import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import type { ProjectCommand, ProjectDocument } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  applyRendererCommand,
  buildCreateTaskCommand,
  buildGanttView,
  buildIndentCommand,
  buildOutdentCommand,
  createRendererSession,
  createViewState,
  hitTestGantt,
  projectDocumentView,
  reconcileViewState,
  reduceViewState,
} from '../src/index.js'
import { ganttDocument, makeDocument, makeTask, multiSiblingDocument } from './fixtures.js'

/** The Gantt fixture (root > a > a1,a2, b, m — milestone m at row 5) with a
 * real derived schedule and a full-width row window. */
const ganttFixture = () => {
  const document = ganttDocument()
  const session = createRendererSession(document, { schedule })
  const state = createViewState(document, session.schedule)
  const view = buildGanttView(
    document,
    projectDocumentView(document, session.schedule, state),
    state,
    { firstRow: 0, visibleRows: 6, overscan: 0 },
  )
  return { document, session, state, view }
}

describe('PROJECT-023 integration — projection selection/edit reflection', () => {
  it('rows carry selected/focused as a pure echo of the view state', () => {
    const { document, session, state } = ganttFixture()
    let next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    next = reduceViewState(
      next,
      { type: 'selectTask', taskId: asTaskId('b'), mode: 'toggle' },
      { document, schedule: session.schedule },
    )

    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows.map((row) => [row.taskId, row.selected, row.focused])).toEqual([
      [asTaskId('root'), false, false],
      [asTaskId('a'), false, false],
      [asTaskId('a1'), true, false],
      [asTaskId('a2'), false, false],
      [asTaskId('b'), true, true],
      [asTaskId('m'), false, false],
    ])
  })

  it('the edit TARGET is reflected (editingField); the draft stays on the live state', () => {
    const { document, session, state } = ganttFixture()
    const next = reduceViewState(
      state,
      { type: 'beginTaskEdit', taskId: asTaskId('a2'), field: 'duration' },
      { document, schedule: session.schedule },
    )
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows[3]!.editingField).toBe('duration')
    expect(projection.rows.map((row) => row.editingField ?? null)).toEqual([
      null,
      null,
      null,
      'duration',
      null,
      null,
    ])
    // The draft text is live user input, NOT projection state.
    expect(next.editing?.draft).toBe('960')
  })

  it('reflection is deterministic across repeated projections (3× byte-identical)', () => {
    const { document, session, state } = ganttFixture()
    const next = reduceViewState(
      state,
      { type: 'moveTaskFocus', direction: 'down' },
      { document, schedule: session.schedule },
    )
    const first = JSON.stringify(projectDocumentView(document, session.schedule, next))
    expect(JSON.stringify(projectDocumentView(document, session.schedule, next))).toBe(first)
    expect(JSON.stringify(projectDocumentView(document, session.schedule, next))).toBe(first)
  })
})

describe('PROJECT-023 integration — grid row / bar / milestone → selection', () => {
  it('a grid row identifies its task: row → selectTask → the rebuilt view reflects the selection', () => {
    const { document, session, state, view } = ganttFixture()
    const gridRow = view.taskGrid.rows[2]!
    expect(gridRow.row.taskId).toBe(asTaskId('a1'))

    const next = reduceViewState(
      state,
      { type: 'selectTask', taskId: gridRow.row.taskId },
      { document, schedule: session.schedule },
    )
    const rebuilt = buildGanttView(
      document,
      projectDocumentView(document, session.schedule, next),
      next,
      { firstRow: 0, visibleRows: 6, overscan: 0 },
    )
    expect(rebuilt.taskGrid.rows[2]!.row.selected).toBe(true)
    expect(rebuilt.taskGrid.rows[2]!.row.focused).toBe(true)
    expect(rebuilt.taskGrid.rows[0]!.row.selected).toBe(false)
    // The timeline pane addresses the same row by reference — synchronized.
    expect(rebuilt.timeline.rows[2]!.row.selected).toBe(true)
  })

  it('a bar hit resolves to its task: hitTestGantt → selectTask → the selected row', () => {
    const { document, session, state, view } = ganttFixture()
    const bar = view.timeline.bars.find((candidate) => candidate.taskId === asTaskId('a1'))!
    expect(bar).toBeDefined()
    const hit = hitTestGantt(view.timeline, {
      rowIndex: bar.rowIndex,
      fraction: (bar.startFraction + bar.finishFraction) / 2,
    })
    expect(hit).toEqual({ kind: 'bar', taskId: asTaskId('a1') })

    const next = reduceViewState(
      state,
      { type: 'selectTask', taskId: hit!.taskId },
      { document, schedule: session.schedule },
    )
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows[bar.rowIndex]!.selected).toBe(true)
  })

  it('a milestone hit resolves to its task (milestones before bars)', () => {
    const { document, session, state, view } = ganttFixture()
    const milestone = view.timeline.milestones[0]!
    expect(milestone.taskId).toBe(asTaskId('m'))
    const hit = hitTestGantt(view.timeline, {
      rowIndex: milestone.rowIndex,
      fraction: milestone.fraction,
    })
    expect(hit).toEqual({ kind: 'milestone', taskId: asTaskId('m') })

    const next = reduceViewState(
      state,
      { type: 'selectTask', taskId: hit!.taskId },
      { document, schedule: session.schedule },
    )
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows[milestone.rowIndex]!.selected).toBe(true)
  })
})

describe('PROJECT-023 integration — collapsed-row selection policy', () => {
  it('hiding is not deselecting: a selected child stays selected while hidden, and re-projects on expand', () => {
    const { document, session, state } = ganttFixture()
    let next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a1') },
      { document, schedule: session.schedule },
    )
    next = reduceViewState(
      next,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )

    // The child is hidden (the visibility rule) but its selection is RETAINED.
    const collapsedProjection = projectDocumentView(document, session.schedule, next)
    expect(collapsedProjection.rows.map((row) => row.taskId)).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('b'),
      asTaskId('m'),
    ])
    expect(next.tasks.taskIds).toEqual([asTaskId('a1')])
    expect(next.tasks.anchorId).toBe(asTaskId('a1'))
    expect(next.tasks.focusId).toBe(asTaskId('a1'))

    // Expanding re-projects the child as a selected row.
    next = reduceViewState(
      next,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )
    const expanded = projectDocumentView(document, session.schedule, next)
    expect(expanded.rows[2]!.taskId).toBe(asTaskId('a1'))
    expect(expanded.rows[2]!.selected).toBe(true)
    expect(expanded.rows[2]!.focused).toBe(true)
  })

  it('range selection across a collapsed summary keeps the hidden members selected (the documented outline rule)', () => {
    const { document, session, state } = ganttFixture()
    let next = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('root') },
      { document, schedule: session.schedule },
    )
    next = reduceViewState(
      next,
      { type: 'toggleCollapse', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )
    // Shift-click root → b: the accepted 021 outline-order range [root..b]
    // includes the hidden a1/a2 — retained in the selection.
    next = reduceViewState(
      next,
      { type: 'selectTask', taskId: asTaskId('b'), mode: 'extend' },
      { document, schedule: session.schedule },
    )
    expect(next.tasks.taskIds).toEqual([
      asTaskId('root'),
      asTaskId('a'),
      asTaskId('a1'),
      asTaskId('a2'),
      asTaskId('b'),
    ])
    // The VISIBLE rows in the range are flagged; the hidden members stay
    // selected in state and re-project on expand.
    const projection = projectDocumentView(document, session.schedule, next)
    expect(projection.rows.map((row) => [row.taskId, row.selected])).toEqual([
      [asTaskId('root'), true],
      [asTaskId('a'), true],
      [asTaskId('b'), true],
      [asTaskId('m'), false],
    ])
  })
})

describe('PROJECT-023 integration — selection through the command/session pipeline', () => {
  const twoLeafDocument = (): ProjectDocument =>
    makeDocument({
      tasks: [
        makeTask({ id: 'a', name: 'Design', wbs: '1' }),
        makeTask({ id: 'b', name: 'Build', wbs: '2' }),
      ],
    })

  it('selection after task CREATION: the selection is preserved and the new row is unselected', () => {
    const document = twoLeafDocument()
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )

    const command = buildCreateTaskCommand(document, { kind: 'lastRoot' })
    const outcome = applyRendererCommand(session, command)
    expect(outcome.result.accepted).toBe(true)

    // Host reconciliation after the document replacement: a created task is
    // NOT auto-selected — the documented policy.
    const reconciled = reconcileViewState(state, outcome.session.document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a')])
    expect(reconciled.tasks.anchorId).toBe(asTaskId('a'))
    const projection = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      reconciled,
    )
    expect(projection.rows.map((row) => [row.taskId, row.selected])).toEqual([
      [asTaskId('a'), true],
      [asTaskId('b'), false],
      [asTaskId('t1'), false],
    ])
  })

  it('selection after INDENT: ids survive, the selection is preserved verbatim', () => {
    const document = multiSiblingDocument()
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a2') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a3'), mode: 'toggle' },
      { document, schedule: session.schedule },
    )

    const command = buildIndentCommand(document, asTaskId('a2'))
    expect(command).toBeDefined()
    const outcome = applyRendererCommand(session, command!)
    expect(outcome.result.accepted).toBe(true)

    // a2 became a child of a1 — the SAME TaskId, so the selection survives
    // verbatim (order, anchor, focus) through the host reconciliation.
    const reconciled = reconcileViewState(state, outcome.session.document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a2'), asTaskId('a3')])
    const projection = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      reconciled,
    )
    expect(projection.rows.map((row) => [row.taskId, row.selected])).toEqual([
      [asTaskId('root1'), false],
      [asTaskId('p'), false],
      [asTaskId('a1'), false],
      [asTaskId('a2'), true],
      [asTaskId('a3'), true],
      [asTaskId('root2'), false],
    ])
  })

  it('selection after OUTDENT: ids survive, the selection is preserved verbatim', () => {
    const document = multiSiblingDocument()
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a3') },
      { document, schedule: session.schedule },
    )

    const indent = buildIndentCommand(document, asTaskId('a3'))
    const indented = applyRendererCommand(session, indent!)
    expect(indented.result.accepted).toBe(true)

    const outdent = buildOutdentCommand(indented.session.document, asTaskId('a3'))
    expect(outdent).toBeDefined()
    const outcome = applyRendererCommand(indented.session, outdent!)
    expect(outcome.result.accepted).toBe(true)

    const reconciled = reconcileViewState(state, outcome.session.document)
    expect(reconciled.tasks.taskIds).toEqual([asTaskId('a3')])
    const projection = projectDocumentView(
      outcome.session.document,
      outcome.session.schedule,
      reconciled,
    )
    expect(projection.rows.find((row) => row.taskId === asTaskId('a3'))!.selected).toBe(true)
  })

  it('a malformed command value is rejected with the document and selection untouched (UNSUPPORTED_COMMAND runtime safety net)', () => {
    const document = twoLeafDocument()
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: asTaskId('a') },
      { document, schedule: session.schedule },
    )

    // PROJECT-024 implemented the frozen dependency commands, so the
    // UNSUPPORTED_COMMAND path now guards only malformed command VALUES
    // arriving through untyped boundaries (the pre-024 suite used
    // AddDependency as its unimplemented example).
    const command = { type: 'NonsenseCommand' } as unknown as ProjectCommand
    const outcome = applyRendererCommand(session, command)
    expect(outcome.result.accepted).toBe(false)
    expect(outcome.result.diagnostics[0]!.code).toBe('UNSUPPORTED_COMMAND')
    // The SAME session reference: nothing journaled, nothing mutated.
    expect(outcome.session).toBe(session)
    expect(outcome.session.document).toBe(document)
    expect(outcome.session.revision).toBe(0)
    // The selection state is untouched and still projects deterministically.
    expect(state.tasks.taskIds).toEqual([asTaskId('a')])
    const projection = projectDocumentView(document, session.schedule, state)
    expect(projection.rows[0]!.selected).toBe(true)
  })
})
