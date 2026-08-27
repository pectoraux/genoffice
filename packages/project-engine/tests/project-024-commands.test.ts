/**
 * PROJECT-024 — dependency command audit.
 *
 * The four frozen dependency commands (AddDependency / RemoveDependency /
 * ChangeDependencyType / ChangeLag) over the canonical dependency graph:
 * acceptance shape (array edits in canonical order, affectedTaskIds,
 * inverses), the complete rejection matrix (duplicate id, missing task
 * references, self links, non-integer lag, duplicate links, the
 * summary↔own-descendant rule, cycles — direct and transitive), inverse
 * round-trips, determinism/no-mutation, and the scheduling effects through
 * the REAL scheduling package (the accepted PROJECT-013 test-layer import
 * precedent: the engine package itself never depends on the scheduler).
 */
import { describe, expect, it } from 'vitest'
import { schedule } from '@genoffice/project-scheduling'
import { applyProjectCommand, canonicalizeDocument } from '../src/index.js'
import { asDependencyId, asISODateTime, asTaskId } from '@genoffice/project-contracts'
import type { Dependency, ProjectDocument } from '@genoffice/project-contracts'
import { MONDAY, TUESDAY, WEDNESDAY, makeDocument, makeTask, wm } from './fixtures.js'

const id = (value: string) => asTaskId(value)
const depId = (value: string) => asDependencyId(value)

/** A canonical two-leaf document: a (Mon 480) and b (480, unpinned start). */
const twoLeafDocument = (): ProjectDocument =>
  canonicalizeDocument(
    makeDocument({
      tasks: [makeTask({ id: 'a', duration: wm(480) }), makeTask({ id: 'b', duration: wm(480) })],
    }),
  )

/** A canonical document with one existing FS link a→b. */
const linkedDocument = (): ProjectDocument => {
  const base = twoLeafDocument()
  const linked = applyProjectCommand(base, {
    type: 'AddDependency',
    dependency: {
      id: depId('d1'),
      predecessorId: id('a'),
      successorId: id('b'),
      type: 'FS',
      lagMinutes: 0,
    },
  })
  expect(linked.result.accepted).toBe(true)
  return linked.document
}

const makeLink = (
  linkId: string,
  predecessor: string,
  successor: string,
  type: Dependency['type'] = 'FS',
  lagMinutes = 0,
): Dependency => ({
  id: depId(linkId),
  predecessorId: id(predecessor),
  successorId: id(successor),
  type,
  lagMinutes,
})

const snapshot = (value: unknown): string => JSON.stringify(value)

// ===========================================================================
// PROJECT-024 AddDependency
// ===========================================================================

describe('PROJECT-024 AddDependency — accept', () => {
  it('appends the link in canonical order and returns the RemoveDependency inverse', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d2', 'a', 'b', 'SS'),
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.dependencies.map((dependency) => dependency.id)).toEqual([
      depId('d1'),
      depId('d2'),
    ])
    expect(exec.document.dependencies[1]).toEqual(makeLink('d2', 'a', 'b', 'SS'))
    expect(exec.result.affectedTaskIds).toEqual([id('a'), id('b')])
    expect(exec.result.inverse).toEqual({
      type: 'RemoveDependency',
      dependencyId: depId('d2'),
    })
  })

  it('never mutates the input document and is deterministic', () => {
    const document = twoLeafDocument()
    const before = snapshot(document)
    const first = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
    const second = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
    expect(snapshot(document)).toBe(before)
    expect(snapshot(first.document)).toBe(snapshot(second.document))
    expect(first.document).not.toBe(document)
  })

  it('undo via the inverse restores the document byte-identically', () => {
    const document = twoLeafDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
    const undone = applyProjectCommand(exec.document, exec.result.inverse!)
    expect(snapshot(undone.document)).toBe(snapshot(document))
  })
})

describe('PROJECT-024 AddDependency — reject', () => {
  it('rejects a duplicate dependency id', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'b', 'a'),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'DUPLICATE_DEPENDENCY_ID')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects missing task references (either endpoint)', () => {
    const document = twoLeafDocument()
    for (const link of [makeLink('d1', 'zzz', 'b'), makeLink('d1', 'a', 'zzz')]) {
      const exec = applyProjectCommand(document, { type: 'AddDependency', dependency: link })
      expect(exec.result.accepted).toBe(false)
      expect(exec.result.diagnostics.some((d) => d.code === 'MISSING_TASK_REFERENCE')).toBe(true)
      expect(exec.document).toBe(document)
    }
  })

  it('rejects a self-referencing link', () => {
    const document = twoLeafDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'a'),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'SELF_DEPENDENCY')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a non-integer lag', () => {
    const document = twoLeafDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b', 'FS', 480.5),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'INVALID_LAG')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a duplicate link (same endpoints and type)', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d2', 'a', 'b', 'FS'),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'DUPLICATE_DEPENDENCY_LINK')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a link between a summary and its own descendant (both directions)', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'c', parentTaskId: id('s'), outlineLevel: 2 }),
        ],
      }),
    )
    for (const link of [makeLink('d1', 's', 'c'), makeLink('d1', 'c', 's')]) {
      const exec = applyProjectCommand(document, { type: 'AddDependency', dependency: link })
      expect(exec.result.accepted).toBe(false)
      expect(exec.result.diagnostics.some((d) => d.code === 'SUMMARY_DEPENDENCY')).toBe(true)
      expect(exec.document).toBe(document)
    }
  })

  it('rejects a DIRECT cycle (the reverse link of an existing one)', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d2', 'b', 'a'),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'DEPENDENCY_CYCLE')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a TRANSITIVE cycle (closing a three-edge chain)', () => {
    // a→b, b→c, then closing c→a must cycle.
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      }),
    )
    const withAB = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
    expect(withAB.result.accepted).toBe(true)
    const withBC = applyProjectCommand(withAB.document, {
      type: 'AddDependency',
      dependency: makeLink('d2', 'b', 'c'),
    })
    expect(withBC.result.accepted).toBe(true)
    const closing = applyProjectCommand(withBC.document, {
      type: 'AddDependency',
      dependency: makeLink('d3', 'c', 'a'),
    })
    expect(closing.result.accepted).toBe(false)
    expect(closing.result.diagnostics.some((d) => d.code === 'DEPENDENCY_CYCLE')).toBe(true)
    expect(closing.document).toBe(withBC.document)
  })
})

// ===========================================================================
// PROJECT-024 RemoveDependency
// ===========================================================================

describe('PROJECT-024 RemoveDependency', () => {
  it('removes the link and returns the full-record AddDependency inverse', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'RemoveDependency',
      dependencyId: depId('d1'),
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.dependencies).toEqual([])
    expect(exec.result.affectedTaskIds).toEqual([id('a'), id('b')])
    expect(exec.result.inverse).toEqual({
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
  })

  it('undo via the full-record inverse restores the link byte-identically (id, endpoints, type, lag)', () => {
    const document = linkedDocument()
    const removed = applyProjectCommand(document, {
      type: 'RemoveDependency',
      dependencyId: depId('d1'),
    })
    const undone = applyProjectCommand(removed.document, removed.result.inverse!)
    expect(snapshot(undone.document)).toBe(snapshot(document))
  })

  it('rejects an unknown dependency id', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'RemoveDependency',
      dependencyId: depId('zzz'),
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'MISSING_DEPENDENCY')).toBe(true)
    expect(exec.document).toBe(document)
  })
})

// ===========================================================================
// PROJECT-024 ChangeDependencyType
// ===========================================================================

describe('PROJECT-024 ChangeDependencyType', () => {
  it('changes the type and returns the previous-type inverse', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'SS',
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.dependencies[0]!.type).toBe('SS')
    expect(exec.document.dependencies[0]!.lagMinutes).toBe(0)
    expect(exec.result.affectedTaskIds).toEqual([id('a'), id('b')])
    expect(exec.result.inverse).toEqual({
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'FS',
    })
  })

  it('accepts a same-type change as an idempotent write', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'FS',
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.dependencies[0]).toEqual(makeLink('d1', 'a', 'b'))
  })

  it('rejects an unknown dependency id', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeDependencyType',
      dependencyId: depId('zzz'),
      dependencyType: 'SS',
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'MISSING_DEPENDENCY')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a type outside the FS/SS/FF/SF domain', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'XX' as Dependency['type'],
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'INVALID_DEPENDENCY_TYPE')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a re-keying that collides with an existing link', () => {
    // d1: a→b FS, d2: a→b SS. Changing d1 to SS re-keys onto d2's link.
    const document = applyProjectCommand(linkedDocument(), {
      type: 'AddDependency',
      dependency: makeLink('d2', 'a', 'b', 'SS'),
    }).document
    const exec = applyProjectCommand(document, {
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'SS',
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'DUPLICATE_DEPENDENCY_LINK')).toBe(true)
    expect(exec.document).toBe(document)
  })
})

// ===========================================================================
// PROJECT-024 ChangeLag
// ===========================================================================

describe('PROJECT-024 ChangeLag', () => {
  it('changes the lag (negative = lead) and returns the previous-lag inverse', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: -480,
    })
    expect(exec.result.accepted).toBe(true)
    expect(exec.document.dependencies[0]!.lagMinutes).toBe(-480)
    expect(exec.document.dependencies[0]!.type).toBe('FS')
    expect(exec.result.affectedTaskIds).toEqual([id('a'), id('b')])
    expect(exec.result.inverse).toEqual({
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: 0,
    })
  })

  it('undo via the inverse restores the lag byte-identically', () => {
    const document = linkedDocument()
    const changed = applyProjectCommand(document, {
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: 960,
    })
    const undone = applyProjectCommand(changed.document, changed.result.inverse!)
    expect(snapshot(undone.document)).toBe(snapshot(document))
  })

  it('rejects an unknown dependency id', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeLag',
      dependencyId: depId('zzz'),
      lagMinutes: 480,
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'MISSING_DEPENDENCY')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('rejects a non-integer lag', () => {
    const document = linkedDocument()
    const exec = applyProjectCommand(document, {
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: 480.5,
    })
    expect(exec.result.accepted).toBe(false)
    expect(exec.result.diagnostics.some((d) => d.code === 'INVALID_LAG')).toBe(true)
    expect(exec.document).toBe(document)
  })

  it('never mutates the input document and is deterministic', () => {
    const document = linkedDocument()
    const before = snapshot(document)
    const run = () =>
      applyProjectCommand(document, {
        type: 'ChangeLag',
        dependencyId: depId('d1'),
        lagMinutes: 480,
      })
    expect(snapshot(run().document)).toBe(snapshot(run().document))
    expect(snapshot(document)).toBe(before)
  })
})

// ===========================================================================
// PROJECT-024 — scheduling effects (the REAL scheduling authority)
// ===========================================================================

describe('PROJECT-024 — scheduling effects (real scheduler)', () => {
  it('an FS link moves the unpinned successor to the predecessor finish (next working instant)', () => {
    const document = twoLeafDocument()
    const before = schedule(document)
    // Without the link, b starts at the project start (Monday 09:00).
    expect(before.taskSchedules[id('b')].scheduledStart).toBe(asISODateTime(MONDAY))

    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
    const after = schedule(exec.document)
    // a finishes Monday 17:00; FS with zero lag starts b at the next
    // working instant — Tuesday 09:00.
    expect(after.taskSchedules[id('a')].scheduledFinish).toBe(
      asISODateTime('2026-08-03T17:00:00.000Z'),
    )
    expect(after.taskSchedules[id('b')].scheduledStart).toBe(asISODateTime(TUESDAY))
  })

  it('a lag of one working day pushes the successor one further working day', () => {
    const document = applyProjectCommand(twoLeafDocument(), {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b', 'FS', 480),
    }).document
    const derived = schedule(document)
    expect(derived.taskSchedules[id('b')].scheduledStart).toBe(asISODateTime(WEDNESDAY))
  })

  it('a negative lag (lead) pulls the successor back to the predecessor start', () => {
    const document = applyProjectCommand(twoLeafDocument(), {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b', 'FS', -480),
    }).document
    const derived = schedule(document)
    expect(derived.taskSchedules[id('b')].scheduledStart).toBe(asISODateTime(MONDAY))
  })

  it('changing FS→SS moves the successor to the predecessor start', () => {
    const linked = linkedDocument()
    expect(schedule(linked).taskSchedules[id('b')].scheduledStart).toBe(asISODateTime(TUESDAY))

    const exec = applyProjectCommand(linked, {
      type: 'ChangeDependencyType',
      dependencyId: depId('d1'),
      dependencyType: 'SS',
    })
    expect(schedule(exec.document).taskSchedules[id('b')].scheduledStart).toBe(
      asISODateTime(MONDAY),
    )
  })

  it('changing the lag moves the derived start exactly as the scheduler derives it', () => {
    const linked = linkedDocument()
    const exec = applyProjectCommand(linked, {
      type: 'ChangeLag',
      dependencyId: depId('d1'),
      lagMinutes: 480,
    })
    expect(schedule(exec.document).taskSchedules[id('b')].scheduledStart).toBe(
      asISODateTime(WEDNESDAY),
    )
  })

  it('removing the link restores the unpinned successor to the project start', () => {
    const linked = linkedDocument()
    expect(schedule(linked).taskSchedules[id('b')].scheduledStart).toBe(asISODateTime(TUESDAY))

    const exec = applyProjectCommand(linked, {
      type: 'RemoveDependency',
      dependencyId: depId('d1'),
    })
    expect(schedule(exec.document).taskSchedules[id('b')].scheduledStart).toBe(
      asISODateTime(MONDAY),
    )
  })

  it('the engine itself computes no dates — only the scheduler derives them', () => {
    const document = twoLeafDocument()
    const exec = applyProjectCommand(document, {
      type: 'AddDependency',
      dependency: makeLink('d1', 'a', 'b'),
    })
    // The engine stored the semantic link; every task field is untouched.
    expect(snapshot(exec.document.tasks)).toBe(snapshot(document.tasks))
  })
})
