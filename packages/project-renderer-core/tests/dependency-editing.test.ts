/**
 * PROJECT-024 — the dependency-editing model audit.
 *
 * The host-independent editing surface for a dependency's two fields (type,
 * lag): the editable-field domain, the canonical initial drafts, the
 * deterministic draft → semantic-command translation (apply / noChange /
 * invalid / none — syntax here, semantics in the engine), and the reducer's
 * dependency-edit intents (activation selects the link and enforces the
 * single-editor rule across the task and dependency families; the reconcile
 * pass drops an edit whose target link vanished).
 */
import { describe, expect, it } from 'vitest'
import { asDependencyId, asTaskId } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { schedule } from '@genoffice/project-scheduling'
import {
  DEPENDENCY_TYPE_CODES,
  EDITABLE_DEPENDENCY_FIELDS,
  commitDependencyEdit,
  createRendererSession,
  createViewState,
  editableDependencyFields,
  initialDependencyEditDraft,
  reduceViewState,
  reconcileViewState,
} from '../src/index.js'
import type { DependencyEditing } from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

const id = (value: string) => asTaskId(value)
const depId = (value: string) => asDependencyId(value)

const linkedDocument = (): ProjectDocument =>
  makeDocument({
    tasks: [makeTask({ id: 'a', wbs: '1' }), makeTask({ id: 'b', wbs: '2' })],
    dependencies: [makeDependency('d1', 'a', 'b', 'FS', 60)],
  })

describe('PROJECT-024 dependency editing — the editable-field domain', () => {
  it('both fields are editable on every dependency (no summary/leaf distinction)', () => {
    expect(EDITABLE_DEPENDENCY_FIELDS).toEqual(['type', 'lag'])
    expect(editableDependencyFields()).toEqual(['type', 'lag'])
  })

  it('the canonical type-code domain is FS/SS/FF/SF', () => {
    expect(DEPENDENCY_TYPE_CODES).toEqual(['FS', 'SS', 'FF', 'SF'])
  })
})

describe('PROJECT-024 dependency editing — initial drafts', () => {
  it('the type draft is the stored type code; the lag draft is the stored decimal minutes', () => {
    const document = linkedDocument()
    expect(initialDependencyEditDraft(document, depId('d1'), 'type')).toBe('FS')
    expect(initialDependencyEditDraft(document, depId('d1'), 'lag')).toBe('60')
  })

  it('a missing dependency yields the empty draft (the reducer never activates it)', () => {
    expect(initialDependencyEditDraft(linkedDocument(), depId('zzz'), 'type')).toBe('')
    expect(initialDependencyEditDraft(linkedDocument(), depId('zzz'), 'lag')).toBe('')
  })
})

describe('PROJECT-024 dependency editing — the commit translation', () => {
  it('no active edit is a pure none', () => {
    expect(commitDependencyEdit(linkedDocument(), undefined)).toEqual({ kind: 'none' })
  })

  it('a missing dependency is invalid (missingDependency)', () => {
    const editing: DependencyEditing = {
      dependencyId: depId('zzz'),
      field: 'type',
      draft: 'SS',
    }
    expect(commitDependencyEdit(linkedDocument(), editing)).toEqual({
      kind: 'invalid',
      reason: 'missingDependency',
    })
  })

  it('a type draft maps 1:1 onto ChangeDependencyType for all four codes', () => {
    const document = linkedDocument()
    for (const code of ['SS', 'FF', 'SF'] as const) {
      expect(
        commitDependencyEdit(document, { dependencyId: depId('d1'), field: 'type', draft: code }),
      ).toEqual({
        kind: 'apply',
        command: {
          type: 'ChangeDependencyType',
          dependencyId: depId('d1'),
          dependencyType: code,
        },
      })
    }
  })

  it('a type draft outside the four codes is invalid (unparseableDependencyType)', () => {
    const document = linkedDocument()
    for (const draft of ['fs', 'XX', 'FS ', '', 'Finish-to-Start']) {
      expect(
        commitDependencyEdit(document, { dependencyId: depId('d1'), field: 'type', draft }),
      ).toEqual({ kind: 'invalid', reason: 'unparseableDependencyType' })
    }
  })

  it('a type draft equal to the stored type is noChange', () => {
    expect(
      commitDependencyEdit(linkedDocument(), {
        dependencyId: depId('d1'),
        field: 'type',
        draft: 'FS',
      }),
    ).toEqual({ kind: 'noChange' })
  })

  it('a lag draft parses under the canonical decimal rule onto ChangeLag', () => {
    expect(
      commitDependencyEdit(linkedDocument(), {
        dependencyId: depId('d1'),
        field: 'lag',
        draft: '480',
      }),
    ).toEqual({
      kind: 'apply',
      command: { type: 'ChangeLag', dependencyId: depId('d1'), lagMinutes: 480 },
    })
    // Negative lag (lead) parses — the engine owns the semantic domain.
    expect(
      commitDependencyEdit(linkedDocument(), {
        dependencyId: depId('d1'),
        field: 'lag',
        draft: '-480',
      }),
    ).toEqual({
      kind: 'apply',
      command: { type: 'ChangeLag', dependencyId: depId('d1'), lagMinutes: -480 },
    })
  })

  it('a fractional lag draft PARSES and becomes a command (the engine rejects it)', () => {
    // The SetTaskDuration precedent: the renderer owns text syntax only —
    // 480.5 is canonical decimal text, so it dispatches; the engine's
    // INVALID_LAG is the single semantic authority.
    expect(
      commitDependencyEdit(linkedDocument(), {
        dependencyId: depId('d1'),
        field: 'lag',
        draft: '480.5',
      }),
    ).toEqual({
      kind: 'apply',
      command: { type: 'ChangeLag', dependencyId: depId('d1'), lagMinutes: 480.5 },
    })
  })

  it('a lag draft that is not canonical decimal text is invalid (unparseableLag)', () => {
    const document = linkedDocument()
    for (const draft of ['1h', '480m', '1,5', '', ' ', '1e3', '-']) {
      expect(
        commitDependencyEdit(document, { dependencyId: depId('d1'), field: 'lag', draft }),
      ).toEqual({ kind: 'invalid', reason: 'unparseableLag' })
    }
  })

  it('a lag draft equal to the stored lag is noChange', () => {
    expect(
      commitDependencyEdit(linkedDocument(), {
        dependencyId: depId('d1'),
        field: 'lag',
        draft: '60',
      }),
    ).toEqual({ kind: 'noChange' })
  })

  it('is pure: the same inputs always yield the same outcome and never mutate the document', () => {
    const document = linkedDocument()
    const editing: DependencyEditing = { dependencyId: depId('d1'), field: 'lag', draft: '480' }
    const before = JSON.stringify(document)
    expect(commitDependencyEdit(document, editing)).toEqual(commitDependencyEdit(document, editing))
    expect(JSON.stringify(document)).toBe(before)
  })
})

describe('PROJECT-024 dependency editing — reducer intents', () => {
  const fixture = () => {
    const document = linkedDocument()
    const session = createRendererSession(document, { schedule })
    const state = createViewState(document, session.schedule)
    return { document, session, state }
  }

  it('beginDependencyEdit activates the editor with the canonical draft and selects the link', () => {
    const { document, session, state } = fixture()
    const next = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'lag' },
      { document, schedule: session.schedule },
    )
    expect(next.dependencyEditing).toEqual({
      dependencyId: depId('d1'),
      field: 'lag',
      draft: '60',
    })
    expect(next.dependencies).toEqual([depId('d1')])
  })

  it('beginDependencyEdit is a deterministic no-op for an unknown dependency', () => {
    const { document, session, state } = fixture()
    const next = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('zzz'), field: 'type' },
      { document, schedule: session.schedule },
    )
    expect(next).toBe(state)
  })

  it('updateDependencyEditDraft replaces the draft and is a no-op when nothing is edited', () => {
    const { document, session, state } = fixture()
    const editing = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'type' },
      { document, schedule: session.schedule },
    )
    const updated = reduceViewState(
      editing,
      { type: 'updateDependencyEditDraft', draft: 'SS' },
      { document, schedule: session.schedule },
    )
    expect(updated.dependencyEditing).toEqual({
      dependencyId: depId('d1'),
      field: 'type',
      draft: 'SS',
    })
    expect(
      reduceViewState(
        state,
        { type: 'updateDependencyEditDraft', draft: 'SS' },
        {
          document,
          schedule: session.schedule,
        },
      ),
    ).toBe(state)
  })

  it('endDependencyEdit drops the editor (a no-op when nothing is edited)', () => {
    const { document, session, state } = fixture()
    const editing = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'type' },
      { document, schedule: session.schedule },
    )
    const ended = reduceViewState(
      editing,
      { type: 'endDependencyEdit' },
      { document, schedule: session.schedule },
    )
    expect(ended.dependencyEditing).toBeUndefined()
    expect(ended.dependencies).toEqual([depId('d1')])
    expect(
      reduceViewState(
        state,
        { type: 'endDependencyEdit' },
        { document, schedule: session.schedule },
      ),
    ).toBe(state)
  })

  it('the single-editor rule: a dependency edit replaces an active TASK edit and vice versa', () => {
    const { document, session, state } = fixture()
    const taskEditing = reduceViewState(
      state,
      { type: 'beginTaskEdit', taskId: id('a'), field: 'taskName' },
      { document, schedule: session.schedule },
    )
    const dependencyEditing = reduceViewState(
      taskEditing,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'type' },
      { document, schedule: session.schedule },
    )
    expect(dependencyEditing.editing).toBeUndefined()
    expect(dependencyEditing.dependencyEditing).toEqual({
      dependencyId: depId('d1'),
      field: 'type',
      draft: 'FS',
    })

    const backToTask = reduceViewState(
      dependencyEditing,
      { type: 'beginTaskEdit', taskId: id('b'), field: 'taskName' },
      { document, schedule: session.schedule },
    )
    expect(backToTask.dependencyEditing).toBeUndefined()
    expect(backToTask.editing).toEqual({ taskId: id('b'), field: 'taskName', draft: 'b' })
  })

  it('reconcile drops the dependency edit when its target link no longer exists', () => {
    const { document, session, state } = fixture()
    const editing = reduceViewState(
      state,
      { type: 'beginDependencyEdit', dependencyId: depId('d1'), field: 'type' },
      { document, schedule: session.schedule },
    )
    // The document after the link's removal (an external replacement — the
    // same reconcile hosts run after any document swap).
    const withoutLink: ProjectDocument = { ...document, dependencies: [] }
    const reconciled = reconcileViewState(editing, withoutLink)
    expect(reconciled.dependencyEditing).toBeUndefined()
    expect(reconciled.dependencies).toEqual([])
  })
})
