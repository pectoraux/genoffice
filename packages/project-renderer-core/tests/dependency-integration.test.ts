/**
 * PROJECT-024 integration — dependency editing through the command/session
 * pipeline.
 *
 * The directive's integration surface over the REAL scheduling package:
 * the full semantic path (dependency gesture → builder/intent → semantic
 * ProjectCommand → engine → ProjectDocument → schedule() → projection and
 * link-surface refresh), the link geometry refresh for every mutation
 * (create / retype / remove), the interaction-state reflection on the link
 * surface (selected / editingField — a pure echo that never changes the
 * geometry), cycle / self-reference / invalid-reference rejection with the
 * document and selection untouched, selection preservation through the
 * dependency commands, and deterministic 3×-repeat replays.
 */
import { describe, expect, it } from 'vitest'
import { asDependencyId, asTaskId } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  applyRendererCommand,
  buildAddDependencyCommand,
  buildGanttView,
  buildRemoveDependencySelectionCommands,
  commitDependencyEditThroughSession,
  createRendererSession,
  createViewState,
  projectDocumentView,
  reconcileViewState,
  reduceViewState,
} from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

const id = (value: string) => asTaskId(value)
const depId = (value: string) => asDependencyId(value)

const twoLeafDocument = (): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({ id: 'a', name: 'Design', wbs: '1' }),
      makeTask({ id: 'b', name: 'Build', wbs: '2' }),
    ],
  })

const snapshot = (value: unknown): string => JSON.stringify(value)

const fullWindow = { firstRow: 0, visibleRows: 10, overscan: 0 }

/** Builds the gantt view's link surface for a (document, schedule, state). */
const linksOf = (
  document: ProjectDocument,
  derived: Parameters<typeof projectDocumentView>[1],
  state: Parameters<typeof projectDocumentView>[2],
) =>
  buildGanttView(document, projectDocumentView(document, derived, state), state, fullWindow)
    .timeline.links

describe('PROJECT-024 integration — the full semantic path', () => {
  it('dependency gesture → command → engine → document → scheduler → projection/link refresh', () => {
    const document = twoLeafDocument()
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)

    // Before: no links; b starts at the project start.
    expect(linksOf(document, session.schedule, state)).toEqual([])
    expect(session.schedule!.taskSchedules[id('b')].scheduledStart).toBe('2026-08-03T09:00:00.000Z')

    // The link-creation gesture: the shared builder (identity + defaults).
    const command = buildAddDependencyCommand(document, id('a'), id('b'))
    expect(command).toBeDefined()
    const outcome = applyRendererCommand(session, command!)

    // The engine accepted; the document carries the link; the REAL
    // scheduler moved the successor exactly as FS derives it.
    expect(outcome.result.accepted).toBe(true)
    expect(outcome.session.document.dependencies).toEqual([makeDependency('d1', 'a', 'b', 'FS', 0)])
    expect(outcome.session.schedule!.taskSchedules[id('b')].scheduledStart).toBe(
      '2026-08-04T09:00:00.000Z',
    )

    // The projection and the link surface refreshed from the new state.
    const nextState = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      {
        document: outcome.session.document,
        schedule: outcome.session.schedule,
      },
    )
    const links = linksOf(outcome.session.document, outcome.session.schedule, nextState)
    expect(links).toHaveLength(1)
    expect(links[0]!.dependencyId).toBe(depId('d1'))
    expect(links[0]!.type).toBe('FS')
    expect(links[0]!.predecessorTaskId).toBe(id('a'))
    expect(links[0]!.successorTaskId).toBe(id('b'))
  })
})

describe('PROJECT-024 integration — link-surface refresh for every mutation', () => {
  const linkedSession = () => {
    const document = twoLeafDocument()
    const session = createRendererSession(document, { schedule })
    const withLink = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a', 'b', 'FS', 0),
    })
    expect(withLink.result.accepted).toBe(true)
    const state = createViewState(withLink.session.document, withLink.session.schedule)
    return { session: withLink.session, state }
  }

  it('a type change re-routes the link geometry (FS finish→start becomes SS start→start)', () => {
    const { session, state } = linkedSession()
    const before = linksOf(session.document, session.schedule, state)
    expect(before[0]!.from.edge).toBe('finish')
    expect(before[0]!.to.edge).toBe('start')

    const retyped = applyRendererCommand(session, {
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'SS',
    })
    expect(retyped.result.accepted).toBe(true)
    const after = linksOf(retyped.session.document, retyped.session.schedule, state)
    expect(after[0]!.type).toBe('SS')
    expect(after[0]!.from.edge).toBe('start')
    expect(after[0]!.to.edge).toBe('start')
  })

  it('a lag change keeps the link but moves the successor bar (the schedule is the authority)', () => {
    const { session, state } = linkedSession()
    const relagged = applyRendererCommand(session, {
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: 480,
    })
    expect(relagged.result.accepted).toBe(true)
    expect(relagged.session.schedule!.taskSchedules[id('b')].scheduledStart).toBe(
      '2026-08-05T09:00:00.000Z',
    )
    const links = linksOf(relagged.session.document, relagged.session.schedule, state)
    expect(links).toHaveLength(1)
    // The successor endpoint attaches at the MOVED start instant (the
    // fraction follows the schedule, never an invented position).
    expect(links[0]!.to.taskId).toBe(id('b'))
  })

  it('a removal drops the link from the surface', () => {
    const { session, state } = linkedSession()
    const removed = applyRendererCommand(session, {
      type: 'RemoveDependency',
      dependencyId: depId('d1'),
    })
    expect(removed.result.accepted).toBe(true)
    expect(linksOf(removed.session.document, removed.session.schedule, state)).toEqual([])
  })
})

describe('PROJECT-024 integration — link interaction-state reflection', () => {
  const linkedFixture = () => {
    const document = twoLeafDocument()
    const session = createRendererSession(document, { schedule })
    const withLink = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a', 'b', 'FS', 0),
    })
    const state = createViewState(withLink.session.document, withLink.session.schedule)
    return { document: withLink.session.document, session: withLink.session, state }
  }

  it('links carry selected as a pure echo of the view state; the edit TARGET is reflected, the draft is not', () => {
    const { document, session, state } = linkedFixture()
    let next = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      { document, schedule: session.schedule },
    )
    let links = linksOf(document, session.schedule, next)
    expect(links[0]!.selected).toBe(true)

    next = reduceViewState(
      next,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'lag' },
      { document, schedule: session.schedule },
    )
    links = linksOf(document, session.schedule, next)
    expect(links[0]!.editingField).toBe('lag')
    // The live DRAFT stays on the state — only the TARGET projects.
    expect(next.dependencyEditing!.draft).toBe('0')

    next = reduceViewState(
      next,
      { type: 'updateDependencyEditDraft', draft: '480' },
      { document, schedule: session.schedule },
    )
    links = linksOf(document, session.schedule, next)
    expect(links[0]!.editingField).toBe('lag')
    expect((links[0] as { draft?: string }).draft).toBeUndefined()
  })

  it('reflection changes NO geometry: with and without state the link routes are byte-identical', () => {
    const { document, session, state } = linkedFixture()
    const selected = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      { document, schedule: session.schedule },
    )
    const withoutState = linksOf(
      document,
      session.schedule,
      createViewState(document, session.schedule),
    )
    const withState = linksOf(document, session.schedule, selected)
    expect(withState[0]!.selected).toBe(true)
    expect(withoutState[0]!.selected).toBe(false)
    expect(snapshot(withState[0]!.route)).toBe(snapshot(withoutState[0]!.route))
    expect(snapshot(withState[0]!.from)).toBe(snapshot(withoutState[0]!.from))
    expect(snapshot(withState[0]!.to)).toBe(snapshot(withoutState[0]!.to))
  })

  it('reflection is deterministic across repeated projections (3× byte-identical)', () => {
    const { document, session, state } = linkedFixture()
    const selected = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'type' },
      { document, schedule: session.schedule },
    )
    const first = snapshot(linksOf(document, session.schedule, selected))
    expect(first).toBe(snapshot(linksOf(document, session.schedule, selected)))
    expect(first).toBe(snapshot(linksOf(document, session.schedule, selected)))
  })
})

describe('PROJECT-024 integration — cycle / self-reference / invalid-reference rejection', () => {
  it('a cycle-closing link is rejected through the session with the document and selection untouched', () => {
    const document = twoLeafDocument()
    const session = createRendererSession(document, { schedule })
    const forward = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a', 'b', 'FS', 0),
    })
    let state = createViewState(forward.session.document, forward.session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      { document: forward.session.document, schedule: forward.session.schedule },
    )

    // The closing gesture: b→a over the existing a→b.
    const closing = buildAddDependencyCommand(forward.session.document, id('b'), id('a'))!
    const outcome = applyRendererCommand(forward.session, closing)
    expect(outcome.result.accepted).toBe(false)
    expect(outcome.result.diagnostics.some((d) => d.code === 'DEPENDENCY_CYCLE')).toBe(true)
    // The SAME session reference: nothing journaled, nothing mutated.
    expect(outcome.session).toBe(forward.session)
    expect(outcome.session.document).toBe(forward.session.document)
    expect(outcome.session.revision).toBe(1)
    // The selection state is untouched and still projects deterministically.
    expect(state.dependencies).toEqual([depId('d1')])
    const links = linksOf(forward.session.document, forward.session.schedule, state)
    expect(links).toHaveLength(1)
    expect(links[0]!.selected).toBe(true)
  })

  it('a self-referencing gesture is disabled at the builder (never a command)', () => {
    const document = twoLeafDocument()
    expect(buildAddDependencyCommand(document, id('a'), id('a'))).toBeUndefined()
  })

  it('an invalid-reference gesture is disabled at the builder (never a command)', () => {
    const document = twoLeafDocument()
    expect(buildAddDependencyCommand(document, id('a'), id('zzz'))).toBeUndefined()
    expect(buildAddDependencyCommand(document, id('zzz'), id('b'))).toBeUndefined()
  })
})

describe('PROJECT-024 integration — selection through the dependency pipeline', () => {
  it('an accepted link creation preserves the existing task AND dependency selections', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a', wbs: '1' }),
        makeTask({ id: 'b', wbs: '2' }),
        makeTask({ id: 'c', wbs: '3' }),
      ],
      dependencies: [makeDependency('d0', 'a', 'b', 'SS', 0)],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: id('b') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d0') },
      { document, schedule: session.schedule },
    )

    const created = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a', 'c', 'FS', 0),
    })
    expect(created.result.accepted).toBe(true)
    const reconciled = reduceViewState(
      state,
      { type: 'endDependencyEdit' },
      { document: created.session.document, schedule: created.session.schedule },
    )
    // Both selections survive the accepted creation verbatim.
    expect(reconciled.tasks.taskIds).toEqual([id('b')])
    expect(reconciled.dependencies).toEqual([depId('d0')])
  })

  it('removing a link prunes its selection (the reconcile pass) while the task selection survives', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', wbs: '1' }), makeTask({ id: 'b', wbs: '2' })],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS', 0)],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: id('a') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      { document, schedule: session.schedule },
    )

    const [removal] = buildRemoveDependencySelectionCommands(document, [depId('d1')])
    const removed = applyRendererCommand(session, removal!)
    expect(removed.result.accepted).toBe(true)

    // The reconcile pass hosts run after ANY document replacement (the
    // accepted PROJECT-021 contract) prunes the removed link's selection
    // while the task selection survives.
    const reconciled = reconcileViewState(state, removed.session.document)
    expect(reconciled.dependencies).toEqual([])
    expect(reconciled.tasks.taskIds).toEqual([id('a')])
  })

  it('a dependency edit preserves the task selection alongside the dependency selection', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', wbs: '1' }), makeTask({ id: 'b', wbs: '2' })],
      dependencies: [makeDependency('d1', 'a', 'b', 'FS', 0)],
    })
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectTask', taskId: id('b') },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'lag' },
      { document, schedule: session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '480' },
      { document, schedule: session.schedule },
    )

    const outcome = commitDependencyEditThroughSession(session, state)
    expect(outcome.result!.accepted).toBe(true)
    expect(outcome.state.tasks.taskIds).toEqual([id('b')])
    expect(outcome.state.dependencies).toEqual([depId('d1')])
  })
})

describe('PROJECT-024 integration — determinism', () => {
  it('a mixed create / retype / relag / remove sequence replays 3× byte-identically', () => {
    // The fixture is built ONCE (the fixture uid allocator is module-level
    // state); every replay starts from the same immutable document value.
    const document = twoLeafDocument()
    const run = () => {
      let session = createRendererSession(document, { schedule })
      const created = applyRendererCommand(
        session,
        buildAddDependencyCommand(document, id('a'), id('b'))!,
      )
      session = created.session
      const retyped = applyRendererCommand(session, {
        type: 'ChangeDependencyType',
        dependencyId: depId('d1'),
        dependencyType: 'SS',
      })
      session = retyped.session
      const relagged = applyRendererCommand(session, {
        type: 'ChangeLag',
        dependencyId: depId('d1'),
        lagMinutes: 480,
      })
      session = relagged.session
      const state = createViewState(session.document, session.schedule)
      const links = linksOf(session.document, session.schedule, state)
      const removed = applyRendererCommand(session, {
        type: 'RemoveDependency',
        dependencyId: depId('d1'),
      })
      return {
        finalDocument: snapshot(removed.session.document),
        finalSchedule: snapshot(removed.session.schedule),
        midLinks: snapshot(links),
        revision: removed.session.revision,
      }
    }
    const first = run()
    expect(run()).toEqual(first)
    expect(run()).toEqual(first)
  })
})
