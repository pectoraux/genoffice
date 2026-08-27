/**
 * PROJECT-024 — dependency command-builder audit.
 *
 * The shared translation layer for the dependency gestures: deterministic
 * link-identity allocation (`d{n}`), the creation defaults (FS, zero lag —
 * the Microsoft Project link defaults), the gesture guards (unknown or
 * self-referencing endpoints are disabled gestures — `undefined`, never an
 * invented command), and the selection-driven removal builder. Pure
 * functions of the canonical document; deterministic for the same document
 * regardless of array order.
 */
import { describe, expect, it } from 'vitest'
import { asDependencyId, asTaskId } from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import {
  buildAddDependencyCommand,
  buildRemoveDependencySelectionCommands,
  nextDependencyIdentity,
} from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

const id = (value: string) => asTaskId(value)

const twoLeafDocument = (): ProjectDocument =>
  makeDocument({
    tasks: [makeTask({ id: 'a', wbs: '1' }), makeTask({ id: 'b', wbs: '2' })],
  })

describe('PROJECT-024 nextDependencyIdentity', () => {
  it('allocates d1 for a document with no dependencies', () => {
    expect(nextDependencyIdentity(twoLeafDocument())).toBe(asDependencyId('d1'))
  })

  it('allocates max+1 over the existing d{n} ids', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
      dependencies: [makeDependency('d1', 'a', 'b'), makeDependency('d3', 'a', 'b', 'SS')],
    })
    expect(nextDependencyIdentity(document)).toBe(asDependencyId('d4'))
  })

  it('ignores ids that do not match the d{n} pattern (imported links are never renumbered)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
      dependencies: [makeDependency('d-t2-t1-FS', 'a', 'b'), makeDependency('x1', 'a', 'b', 'SS')],
    })
    expect(nextDependencyIdentity(document)).toBe(asDependencyId('d1'))
  })

  it('is deterministic for the same document', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
      dependencies: [makeDependency('d2', 'a', 'b')],
    })
    expect(nextDependencyIdentity(document)).toBe(nextDependencyIdentity(document))
    expect(nextDependencyIdentity(document)).toBe(asDependencyId('d3'))
  })
})

describe('PROJECT-024 buildAddDependencyCommand', () => {
  it('builds the creation command with the canonical defaults (FS, zero lag)', () => {
    const command = buildAddDependencyCommand(twoLeafDocument(), id('a'), id('b'))!
    expect(command).toEqual({
      type: 'AddDependency',
      dependency: {
        id: asDependencyId('d1'),
        predecessorId: id('a'),
        successorId: id('b'),
        type: 'FS',
        lagMinutes: 0,
      },
    })
  })

  it('honors an explicit type and lag', () => {
    const command = buildAddDependencyCommand(twoLeafDocument(), id('a'), id('b'), {
      type: 'SS',
      lagMinutes: -480,
    })!
    expect(command).toEqual({
      type: 'AddDependency',
      dependency: {
        id: asDependencyId('d1'),
        predecessorId: id('a'),
        successorId: id('b'),
        type: 'SS',
        lagMinutes: -480,
      },
    })
  })

  it('allocates the identity over the existing links (deterministic)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
      dependencies: [makeDependency('d1', 'a', 'b')],
    })
    const command = buildAddDependencyCommand(document, id('a'), id('b'), { type: 'SS' })!
    expect((command as { dependency: { id: string } }).dependency.id).toBe(asDependencyId('d2'))
  })

  it('returns undefined for an unknown predecessor or successor (the disabled gesture)', () => {
    const document = twoLeafDocument()
    expect(buildAddDependencyCommand(document, id('zzz'), id('b'))).toBeUndefined()
    expect(buildAddDependencyCommand(document, id('a'), id('zzz'))).toBeUndefined()
  })

  it('returns undefined for a self-referencing link (the disabled gesture)', () => {
    expect(buildAddDependencyCommand(twoLeafDocument(), id('a'), id('a'))).toBeUndefined()
  })
})

describe('PROJECT-024 buildRemoveDependencySelectionCommands', () => {
  it('emits one RemoveDependency per EXISTING selected id in canonical document order', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      dependencies: [
        makeDependency('d1', 'a', 'b'),
        makeDependency('d2', 'b', 'c', 'SS'),
        makeDependency('d3', 'a', 'c', 'FF'),
      ],
    })
    const commands = buildRemoveDependencySelectionCommands(document, [
      asDependencyId('d3'),
      asDependencyId('d1'),
    ])
    expect(commands).toEqual([
      { type: 'RemoveDependency', dependencyId: asDependencyId('d1') },
      { type: 'RemoveDependency', dependencyId: asDependencyId('d3') },
    ])
  })

  it('drops unknown ids and returns [] for an empty selection', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
      dependencies: [makeDependency('d1', 'a', 'b')],
    })
    expect(buildRemoveDependencySelectionCommands(document, [asDependencyId('zzz')])).toEqual([])
    expect(buildRemoveDependencySelectionCommands(document, [])).toEqual([])
  })
})
