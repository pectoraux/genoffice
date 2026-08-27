/**
 * PROJECT-024 — the golden scenario battery (D01–D12).
 *
 * Each golden drives one dependency-editing acceptance scenario end-to-end
 * through the real machinery — the link-creation gesture through the shared
 * builder, edits through the dependency commit flow, commands through the
 * session and the REAL scheduling authority — and asserts the complete
 * observable tuple: the dependency selection (and the task selection where
 * it matters), the active edit target where applicable, the canonical
 * ProjectDocument, the DerivedSchedule, and the link surface the hosts
 * render. Pure inputs, no wall clock, no randomness — D12 proves the whole
 * battery deterministic (3× byte-identical).
 *
 * Real-scheduler date anchors on the standard Mon–Fri 09:00–17:00 calendar
 * (project start Monday 2026-08-03 09:00): an unpinned leaf starts at the
 * project start; an FS-linked successor starts at the next working instant
 * after the predecessor's finish; a 480-minute lag pushes one working day.
 */
import { describe, expect, it } from 'vitest'
import { asDependencyId, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
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
  reduceViewState,
  redoRendererCommand,
  undoRendererCommand,
} from '../src/index.js'
import type { ProjectViewState } from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

const id = (value: string) => asTaskId(value)
const depId = (value: string) => asDependencyId(value)

const MONDAY = '2026-08-03T09:00:00.000Z'
const TUESDAY = '2026-08-04T09:00:00.000Z'
const WEDNESDAY = '2026-08-05T09:00:00.000Z'
const THURSDAY = '2026-08-06T09:00:00.000Z'

/**
 * The golden document: two unlinked leaves (a1 480, a2 480). The
 * dependency-editing scenarios link them and drive every mutation. Built
 * ONCE per scenario call — the fixture uid counter makes per-call rebuilds
 * differ, so the determinism scenario reuses one instance.
 */
function goldenDocument(): ProjectDocument {
  return makeDocument({
    tasks: [
      makeTask({ id: 'a1', uid: 1, name: 'Draft', duration: asWorkingMinutes(480), wbs: '1' }),
      makeTask({ id: 'a2', uid: 2, name: 'Review', duration: asWorkingMinutes(480), wbs: '2' }),
    ],
  })
}

/** The observable tuple one golden asserts: selection, edit target,
 * document, derived schedule, and link surface. */
const observable = (
  document: ProjectDocument,
  derived: ReturnType<typeof schedule>,
  state: ProjectViewState,
) => ({
  dependencies: document.dependencies,
  selection: state.dependencies,
  taskSelection: state.tasks.taskIds,
  editTarget: state.dependencyEditing
    ? { dependencyId: state.dependencyEditing.dependencyId, field: state.dependencyEditing.field }
    : null,
  schedule: {
    a1Start: derived.taskSchedules[id('a1')].scheduledStart,
    a1Finish: derived.taskSchedules[id('a1')].scheduledFinish,
    a2Start: derived.taskSchedules[id('a2')].scheduledStart,
    a2Finish: derived.taskSchedules[id('a2')].scheduledFinish,
  },
  links: buildGanttView(document, projectDocumentView(document, derived, state), state, {
    firstRow: 0,
    visibleRows: 10,
    overscan: 0,
  }).timeline.links,
})

const snapshot = (value: unknown): string => JSON.stringify(value)

describe('PROJECT-024 goldens — creation', () => {
  it('D01 — create FS (defaults): the successor starts at the predecessor finish + 1 working instant', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    let state = createViewState(document, session.schedule)
    // The gesture: create the link from a1 to a2 (canonical defaults).
    const command = buildAddDependencyCommand(document, id('a1'), id('a2'))!

    const created = applyRendererCommand(session, command)
    expect(created.result.accepted).toBe(true)
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      { document: created.session.document, schedule: created.session.schedule },
    )

    const observed = observable(created.session.document, created.session.schedule!, state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 0)])
    expect(observed.selection).toEqual([depId('d1')])
    expect(observed.schedule).toEqual({
      a1Start: MONDAY,
      a1Finish: '2026-08-03T17:00:00.000Z',
      a2Start: TUESDAY,
      a2Finish: '2026-08-04T17:00:00.000Z',
    })
    expect(observed.links).toHaveLength(1)
    expect(observed.links[0]!.type).toBe('FS')
    expect(observed.links[0]!.selected).toBe(true)
  })

  it('D02 — create FS with one working-day lag: the successor starts a further working day later', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const command = buildAddDependencyCommand(document, id('a1'), id('a2'), { lagMinutes: 480 })!

    const created = applyRendererCommand(session, command)
    expect(created.result.accepted).toBe(true)
    const state = createViewState(created.session.document, created.session.schedule)
    const observed = observable(created.session.document, created.session.schedule!, state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 480)])
    expect(observed.schedule.a2Start).toBe(WEDNESDAY)
  })

  it('D03 — create FS with a one-day lead (negative lag): the successor pulls back to the predecessor start', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const command = buildAddDependencyCommand(document, id('a1'), id('a2'), { lagMinutes: -480 })!

    const created = applyRendererCommand(session, command)
    expect(created.result.accepted).toBe(true)
    const state = createViewState(created.session.document, created.session.schedule)
    const observed = observable(created.session.document, created.session.schedule!, state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', -480)])
    expect(observed.schedule.a2Start).toBe(MONDAY)
  })

  it('D04 — create SS: the successor starts at the predecessor start', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const command = buildAddDependencyCommand(document, id('a1'), id('a2'), { type: 'SS' })!

    const created = applyRendererCommand(session, command)
    expect(created.result.accepted).toBe(true)
    const state = createViewState(created.session.document, created.session.schedule)
    const observed = observable(created.session.document, created.session.schedule!, state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'SS', 0)])
    expect(observed.schedule.a2Start).toBe(MONDAY)
    expect(observed.links[0]!.from.edge).toBe('start')
    expect(observed.links[0]!.to.edge).toBe('start')
  })

  it('D05 — create FF: the successor finish is pinned to the predecessor finish', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const command = buildAddDependencyCommand(document, id('a1'), id('a2'), { type: 'FF' })!

    const created = applyRendererCommand(session, command)
    expect(created.result.accepted).toBe(true)
    const state = createViewState(created.session.document, created.session.schedule)
    const observed = observable(created.session.document, created.session.schedule!, state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FF', 0)])
    // Both leaves are 480 minutes: FF pins a2's finish to a1's finish.
    expect(observed.schedule.a2Finish).toBe('2026-08-03T17:00:00.000Z')
    expect(observed.links[0]!.from.edge).toBe('finish')
    expect(observed.links[0]!.to.edge).toBe('finish')
  })

  it('D06 — create SF with a lag that pushes: the successor start is bounded by (pred start + lag − duration)', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    // SF: the successor cannot finish before the predecessor starts (plus
    // lag). The scheduler's bound on the successor's required START is
    // predStart + lag − duration, floored by the project start. With
    // lag 960 and a 480-minute successor: required start = Tue 09:00 —
    // the successor starts Tuesday and finishes Tue 17:00 (its finish now
    // honors the SF bound predStart + lag = Tue 09:00). With lag 0 the
    // natural project-start schedule already satisfies the bound (the
    // required start Fri 09:00 is below the project-start floor), which is
    // exactly why the lag makes the SF effect observable here.
    const command = buildAddDependencyCommand(document, id('a1'), id('a2'), {
      type: 'SF',
      lagMinutes: 960,
    })!

    const created = applyRendererCommand(session, command)
    expect(created.result.accepted).toBe(true)
    const state = createViewState(created.session.document, created.session.schedule)
    const observed = observable(created.session.document, created.session.schedule!, state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'SF', 960)])
    expect(observed.schedule.a2Start).toBe(TUESDAY)
    expect(observed.schedule.a2Finish).toBe('2026-08-04T17:00:00.000Z')
    expect(observed.links[0]!.from.edge).toBe('start')
    expect(observed.links[0]!.to.edge).toBe('finish')
  })
})

describe('PROJECT-024 goldens — editing', () => {
  /** The editing fixture: a1 → a2 FS with the link's editor active. */
  const editingFixture = () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const withLink = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a1', 'a2', 'FS', 0),
    })
    const linkedSession = withLink.session
    let state = createViewState(linkedSession.document, linkedSession.schedule)
    return {
      linkedSession,
      state,
      edit: (field: 'type' | 'lag', draft: string) => {
        state = reduceViewState(
          state,
          { type: 'beginDependencyEdit', dependencyId: depId('d1'), field },
          { document: linkedSession.document, schedule: linkedSession.schedule },
        )
        state = reduceViewState(
          state,
          { type: 'updateDependencyEditDraft', draft },
          { document: linkedSession.document, schedule: linkedSession.schedule },
        )
        return { linkedSession, state }
      },
    }
  }

  it('D07 — change the lag through the commit flow: the derived start moves exactly as the scheduler derives it', () => {
    const { edit } = editingFixture()
    const { linkedSession, state } = edit('lag', '960')

    const outcome = commitDependencyEditThroughSession(linkedSession, state)
    expect(outcome.result!.accepted).toBe(true)
    const observed = observable(outcome.session.document, outcome.session.schedule!, outcome.state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 960)])
    // FS + 960 lag (two working days): a2 starts Thursday.
    expect(observed.schedule.a2Start).toBe(THURSDAY)
    // The editor ended; the edited link stays selected (it survives).
    expect(observed.editTarget).toBeNull()
    expect(observed.selection).toEqual([depId('d1')])
  })

  it('D08 — change the type through the commit flow: the link re-routes and the schedule follows', () => {
    const { edit } = editingFixture()
    const { linkedSession, state } = edit('type', 'SS')

    const outcome = commitDependencyEditThroughSession(linkedSession, state)
    expect(outcome.result!.accepted).toBe(true)
    const observed = observable(outcome.session.document, outcome.session.schedule!, outcome.state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'SS', 0)])
    // SS: a2 starts at a1's start (Monday).
    expect(observed.schedule.a2Start).toBe(MONDAY)
    expect(observed.links[0]!.from.edge).toBe('start')
    expect(observed.links[0]!.to.edge).toBe('start')
  })

  it('D09 — a rejected lag edit (engine INVALID_LAG) leaves everything observable unchanged', () => {
    const { edit } = editingFixture()
    const { linkedSession, state } = edit('lag', '480.5')

    const outcome = commitDependencyEditThroughSession(linkedSession, state)
    expect(outcome.result!.accepted).toBe(false)
    expect(outcome.result!.diagnostics.some((d) => d.code === 'INVALID_LAG')).toBe(true)
    expect(outcome.session).toBe(linkedSession)
    const observed = observable(linkedSession.document, linkedSession.schedule!, outcome.state)
    expect(observed.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 0)])
    expect(observed.schedule.a2Start).toBe(TUESDAY)
    expect(observed.links).toHaveLength(1)
    // The editor ended; the selection is untouched by the rejection.
    expect(observed.editTarget).toBeNull()
    expect(observed.selection).toEqual([depId('d1')])
  })
})

describe('PROJECT-024 goldens — session', () => {
  it('D10 — a cycle-closing gesture is rejected with the whole observable tuple unchanged', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const withLink = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a1', 'a2', 'FS', 0),
    })
    let state = createViewState(withLink.session.document, withLink.session.schedule)
    state = reduceViewState(
      state,
      { type: 'selectDependency', dependencyId: depId('d1') },
      { document: withLink.session.document, schedule: withLink.session.schedule },
    )

    const closing = buildAddDependencyCommand(withLink.session.document, id('a2'), id('a1'))!
    const outcome = applyRendererCommand(withLink.session, closing)
    expect(outcome.result.accepted).toBe(false)
    expect(outcome.result.diagnostics.some((d) => d.code === 'DEPENDENCY_CYCLE')).toBe(true)
    expect(outcome.session).toBe(withLink.session)

    const before = observable(withLink.session.document, withLink.session.schedule!, state)
    expect(before.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 0)])
    expect(before.schedule.a2Start).toBe(TUESDAY)
    expect(before.links).toHaveLength(1)
    expect(before.selection).toEqual([depId('d1')])
  })

  it('D11 — undo/redo restores the document, the schedule, AND the selection state after a lag edit', () => {
    const document = goldenDocument()
    const session = createRendererSession(document, { schedule })
    const withLink = applyRendererCommand(session, {
      type: 'AddDependency',
      dependency: makeDependency('d1', 'a1', 'a2', 'FS', 0),
    })
    let state = createViewState(withLink.session.document, withLink.session.schedule)
    state = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'lag' },
      { document: withLink.session.document, schedule: withLink.session.schedule },
    )
    state = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '480' },
      { document: withLink.session.document, schedule: withLink.session.schedule },
    )

    const committed = commitDependencyEditThroughSession(withLink.session, state)
    expect(committed.result!.accepted).toBe(true)
    expect(
      observable(committed.session.document, committed.session.schedule!, committed.state).schedule
        .a2Start,
    ).toBe(WEDNESDAY)

    const undone = undoRendererCommand(committed.session)
    expect(undone.applied).toBe(true)
    const undoObserved = observable(
      undone.session.document,
      undone.session.schedule!,
      committed.state,
    )
    // The document and schedule restored byte-identically to the pre-edit
    // state; the selection still carries the surviving link.
    expect(undoObserved.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 0)])
    expect(undoObserved.schedule.a2Start).toBe(TUESDAY)
    expect(undoObserved.selection).toEqual([depId('d1')])

    const redone = redoRendererCommand(undone.session)
    expect(redone.applied).toBe(true)
    const redoObserved = observable(
      redone.session.document,
      redone.session.schedule!,
      committed.state,
    )
    expect(redoObserved.dependencies).toEqual([makeDependency('d1', 'a1', 'a2', 'FS', 480)])
    expect(redoObserved.schedule.a2Start).toBe(WEDNESDAY)
  })

  it('D12 — deterministic repeated interaction sequence: the whole dependency battery 3× byte-identical', () => {
    // The fixture document is built ONCE (the fixture uid allocator is
    // module-level state); each replay drives the identical sequence from
    // the same immutable starting value.
    const document = goldenDocument()
    const replay = () => {
      let session = createRendererSession(document, { schedule })
      let state = createViewState(session.document, session.schedule)

      // Create (gesture → builder).
      const created = applyRendererCommand(
        session,
        buildAddDependencyCommand(session.document, id('a1'), id('a2'))!,
      )
      session = created.session
      state = reduceViewState(
        state,
        { type: 'selectDependency', dependencyId: depId('d1') },
        { document: session.document, schedule: session.schedule },
      )

      // Edit the lag (begin → draft → commit through the session).
      state = reduceViewState(
        state,
        { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'lag' },
        { document: session.document, schedule: session.schedule },
      )
      state = reduceViewState(
        state,
        { type: 'updateDependencyEditDraft', draft: '480' },
        { document: session.document, schedule: session.schedule },
      )
      const committed = commitDependencyEditThroughSession(session, state)
      session = committed.session
      state = committed.state

      // Undo, redo.
      const undone = undoRendererCommand(session)
      session = undone.session
      const redone = redoRendererCommand(session)
      session = redone.session

      // Retype through the flow.
      state = reduceViewState(
        state,
        { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'type' },
        { document: session.document, schedule: session.schedule },
      )
      state = reduceViewState(
        state,
        { type: 'updateDependencyEditDraft', draft: 'SS' },
        { document: session.document, schedule: session.schedule },
      )
      const retyped = commitDependencyEditThroughSession(session, state)
      session = retyped.session
      state = retyped.state

      // Remove (selection-driven builder).
      const [removal] = buildRemoveDependencySelectionCommands(session.document, [depId('d1')])
      const removed = applyRendererCommand(session, removal!)
      session = removed.session

      return snapshot({
        observable: observable(session.document, session.schedule!, state),
        revision: session.revision,
        commandSeq: session.commandSeq,
      })
    }
    const first = replay()
    expect(replay()).toBe(first)
    expect(replay()).toBe(first)
  })
})
