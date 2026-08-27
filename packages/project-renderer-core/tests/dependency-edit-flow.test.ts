/**
 * PROJECT-024 — the dependency-edit commit flow (command/session
 * integration).
 *
 * `commitDependencyEditThroughSession` is the one-call orchestration of a
 * committed dependency edit through the accepted pipeline: pure translation
 * → semantic command → engine → new document → re-schedule → editor end →
 * view-state reconciliation. The behavior contract mirrors the PROJECT-023
 * task flow: accepted commands journal + refresh; rejected commands return
 * the SAME session reference with the engine's verbatim diagnostics; the
 * editor always ends; the selection is preserved by construction (an
 * accepted lag/type edit keeps the edited link selected — it survives — and
 * undo/redo restores the document, the schedule, and the selection state
 * byte-identically).
 */
import { describe, expect, it } from 'vitest'
import { asDependencyId, asTaskId } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  commitDependencyEditThroughSession,
  createRendererSession,
  createViewState,
  reduceViewState,
  undoRendererCommand,
  redoRendererCommand,
} from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

const id = (value: string) => asTaskId(value)
const depId = (value: string) => asDependencyId(value)

const linkedDocument = (): ProjectDocument =>
  makeDocument({
    tasks: [makeTask({ id: 'a', wbs: '1' }), makeTask({ id: 'b', wbs: '2' })],
    dependencies: [makeDependency('d1', 'a', 'b', 'FS', 0)],
  })

/** The full editing fixture: a real session with the real scheduler, the
 * link's lag editor active, and the edited link selected. */
const editingFixture = () => {
  const document = linkedDocument()
  const session = createRendererSession(document, { schedule })
  let state = createViewState(document, session.schedule)
  state = reduceViewState(
    state,
    { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'lag' },
    { document, schedule: session.schedule },
  )
  return { document, session, state }
}

const snapshot = (value: unknown): string => JSON.stringify(value)

describe('PROJECT-024 dependency edit flow — accepted path', () => {
  it('a lag edit journals the command, ends the editor, refreshes the schedule, and keeps the link selected', () => {
    const { document, session, state } = editingFixture()
    const editing = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '480' },
      { document, schedule: session.schedule },
    )

    const outcome = commitDependencyEditThroughSession(session, editing)
    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.result!.accepted).toBe(true)
    expect(outcome.entry!.command).toEqual({
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: 480,
    })
    // The document and schedule were replaced: the successor moved one
    // working day (the real scheduler's authority).
    expect(outcome.session.document).not.toBe(document)
    expect(outcome.session.document.dependencies[0]!.lagMinutes).toBe(480)
    expect(outcome.session.schedule!.taskSchedules[id('b')].scheduledStart).toBe(
      '2026-08-05T09:00:00.000Z',
    )
    // The editor ended; the edited link survives and stays selected.
    expect(outcome.state.dependencyEditing).toBeUndefined()
    expect(outcome.state.dependencies).toEqual([depId('d1')])
  })

  it('undo/redo restores the document, the schedule, and the selection state byte-identically', () => {
    const { document, session, state } = editingFixture()
    const editing = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '480' },
      { document, schedule: session.schedule },
    )
    const outcome = commitDependencyEditThroughSession(session, editing)
    const scheduleBefore = snapshot(session.schedule)
    const selectionBefore = snapshot(outcome.state.dependencies)

    const undone = undoRendererCommand(outcome.session)
    expect(undone.applied).toBe(true)
    expect(snapshot(undone.session.document)).toBe(snapshot(document))
    expect(snapshot(undone.session.schedule)).toBe(scheduleBefore)
    // The reconciled selection against the restored document still carries
    // the edited link (it exists again) — selection state is preserved.
    expect(snapshot(outcome.state.dependencies)).toBe(selectionBefore)

    const redone = redoRendererCommand(undone.session)
    expect(redone.applied).toBe(true)
    expect(snapshot(redone.session.document)).toBe(snapshot(outcome.session.document))
    expect(snapshot(redone.session.schedule)).toBe(snapshot(outcome.session.schedule))
  })
})

describe('PROJECT-024 dependency edit flow — rejected path', () => {
  it('an engine-rejected lag (INVALID_LAG) returns the SAME session reference with verbatim diagnostics', () => {
    const { document, session, state } = editingFixture()
    // 480.5 parses as canonical decimal text (the renderer owns syntax
    // only) and dispatches; the engine is the single semantic authority.
    const editing = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '480.5' },
      { document, schedule: session.schedule },
    )

    const outcome = commitDependencyEditThroughSession(session, editing)
    expect(outcome.commit.kind).toBe('apply')
    expect(outcome.result!.accepted).toBe(false)
    expect(outcome.result!.diagnostics.some((d) => d.code === 'INVALID_LAG')).toBe(true)
    // Nothing journaled, nothing mutated; the editor still ends (the
    // rejected-commit-reverts-the-cell behavior).
    expect(outcome.session).toBe(session)
    expect(outcome.entry).toBeUndefined()
    expect(outcome.state.dependencyEditing).toBeUndefined()
    // The selection is untouched by the rejection.
    expect(outcome.state.dependencies).toEqual([depId('d1')])
  })
})

describe('PROJECT-024 dependency edit flow — no command dispatched', () => {
  it('noChange (draft equals the stored value) ends the editor and returns the same session', () => {
    const { session, state } = editingFixture()
    // The fixture's draft is the stored lag ('0') — committing it is a
    // noChange: no command, no journal, the editor ends.
    expect(state.dependencyEditing!.draft).toBe('0')
    const outcome = commitDependencyEditThroughSession(session, state)
    expect(outcome.commit).toEqual({ kind: 'noChange' })
    expect(outcome.session).toBe(session)
    expect(outcome.result).toBeUndefined()
    expect(outcome.state.dependencyEditing).toBeUndefined()
    expect(outcome.state.dependencies).toEqual([depId('d1')])
  })

  it('an unparseable draft is invalid: no command, same session, editor ends', () => {
    const { document, session, state } = editingFixture()
    const editing = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '1d' },
      { document, schedule: session.schedule },
    )
    const outcome = commitDependencyEditThroughSession(session, editing)
    expect(outcome.commit).toEqual({ kind: 'invalid', reason: 'unparseableLag' })
    expect(outcome.session).toBe(session)
    expect(outcome.state.dependencyEditing).toBeUndefined()
  })

  it('no active dependency edit is a pure no-op (same session AND state references)', () => {
    const document = linkedDocument()
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)
    const outcome = commitDependencyEditThroughSession(session, state)
    expect(outcome.commit).toEqual({ kind: 'none' })
    expect(outcome.session).toBe(session)
    expect(outcome.state).toBe(state)
  })

  it('is deterministic: the same commit attempt replays byte-identically', () => {
    const { document, session, state } = editingFixture()
    const editing = reduceViewState(
      state,
      { type: 'updateDependencyEditDraft', draft: '960' },
      { document, schedule: session.schedule },
    )
    const first = commitDependencyEditThroughSession(session, editing)
    const second = commitDependencyEditThroughSession(session, editing)
    expect(snapshot(first.session.document)).toBe(snapshot(second.session.document))
    expect(snapshot(first.session.schedule)).toBe(snapshot(second.session.schedule))
    expect(snapshot(first.state)).toBe(snapshot(second.state))
  })
})
