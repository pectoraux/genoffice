import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import type { Task } from '@genoffice/project-contracts'
import {
  HierarchyError,
  buildTaskHierarchy,
  canonicalizeDocument,
  canonicalizeTasks,
  deriveHierarchyFields,
} from '../src/index.js'
import { makeDocument, makeTask } from './fixtures.js'

const parent = (id: string) => asTaskId(id)

const fieldsOf = (tasks: Task[]) => {
  const fields = deriveHierarchyFields(tasks)
  return (id: string) => {
    const entry = fields.get(asTaskId(id))
    if (!entry) throw new Error(`missing derived fields for ${id}`)
    return entry
  }
}

describe('PROJECT-007 hierarchy derivation', () => {
  it('derives WBS for a flat task list', () => {
    const fields = fieldsOf([makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })])
    expect(fields('a').wbs).toBe('1')
    expect(fields('b').wbs).toBe('2')
    expect(fields('c').wbs).toBe('3')
  })

  it('derives WBS for a one-level hierarchy', () => {
    const fields = fieldsOf([
      makeTask({ id: 's', summary: true }),
      makeTask({ id: 'a', parentTaskId: parent('s'), outlineLevel: 2 }),
      makeTask({ id: 'b', parentTaskId: parent('s'), outlineLevel: 2 }),
    ])
    expect(fields('s').wbs).toBe('1')
    expect(fields('a').wbs).toBe('1.1')
    expect(fields('b').wbs).toBe('1.2')
  })

  it('derives WBS for two- and three-level hierarchies', () => {
    const fields = fieldsOf([
      makeTask({ id: 's1', summary: true }),
      makeTask({ id: 's2', summary: true, parentTaskId: parent('s1'), outlineLevel: 2 }),
      makeTask({ id: 's3', summary: true, parentTaskId: parent('s2'), outlineLevel: 3 }),
      makeTask({ id: 'leaf', parentTaskId: parent('s3'), outlineLevel: 4 }),
      makeTask({ id: 'b', parentTaskId: parent('s1'), outlineLevel: 2 }),
      makeTask({ id: 'r2', summary: true }),
      makeTask({ id: 'd', parentTaskId: parent('r2'), outlineLevel: 2 }),
    ])
    expect(fields('s1').wbs).toBe('1')
    expect(fields('s2').wbs).toBe('1.1')
    expect(fields('s3').wbs).toBe('1.1.1')
    expect(fields('leaf').wbs).toBe('1.1.1.1')
    expect(fields('b').wbs).toBe('1.2')
    expect(fields('r2').wbs).toBe('2')
    expect(fields('d').wbs).toBe('2.1')
  })

  it('derives WBS to arbitrary depth', () => {
    const depth = 6
    const tasks: Task[] = []
    let previous = ''
    for (let level = 1; level <= depth; level += 1) {
      const id = `t${level}`
      tasks.push(
        makeTask({
          id,
          summary: level < depth,
          ...(level > 1 ? { parentTaskId: parent(previous), outlineLevel: level } : {}),
        }),
      )
      previous = id
    }
    const fields = fieldsOf(tasks)
    expect(fields('t1').wbs).toBe('1')
    expect(fields('t2').wbs).toBe('1.1')
    expect(fields('t3').wbs).toBe('1.1.1')
    expect(fields('t4').wbs).toBe('1.1.1.1')
    expect(fields('t5').wbs).toBe('1.1.1.1.1')
    expect(fields('t6').wbs).toBe('1.1.1.1.1.1')
    expect(fields('t6').outlineLevel).toBe(depth)
  })

  it('derives outline levels from hierarchy depth (root = 1, child = parent + 1)', () => {
    const fields = fieldsOf([
      makeTask({ id: 's1', summary: true }),
      makeTask({ id: 's2', summary: true, parentTaskId: parent('s1'), outlineLevel: 2 }),
      makeTask({ id: 'x', parentTaskId: parent('s2'), outlineLevel: 3 }),
      makeTask({ id: 'root2' }),
    ])
    expect(fields('s1').outlineLevel).toBe(1)
    expect(fields('s2').outlineLevel).toBe(2)
    expect(fields('x').outlineLevel).toBe(3)
    expect(fields('root2').outlineLevel).toBe(1)
  })

  it('derives the summary flag from having children', () => {
    const fields = fieldsOf([
      makeTask({ id: 's', summary: true }),
      makeTask({ id: 'a', parentTaskId: parent('s'), outlineLevel: 2 }),
    ])
    expect(fields('s').summary).toBe(true)
    expect(fields('a').summary).toBe(false)
  })

  it('sibling ordering is positional and deterministic (never id-sorted, never identity)', () => {
    // Sibling ordering comes from relative array position. Reversing the
    // sibling order reverses the WBS numbering deterministically. There is no
    // separate order field, so an invalid sibling ordering cannot be
    // represented; identity remains the TaskId.
    const ordered = [
      makeTask({ id: 'z', summary: true }),
      makeTask({ id: 'a', parentTaskId: parent('z'), outlineLevel: 2 }),
      makeTask({ id: 'm', parentTaskId: parent('z'), outlineLevel: 2 }),
    ]
    const fields = fieldsOf(ordered)
    expect(fields('a').wbs).toBe('1.1')
    expect(fields('m').wbs).toBe('1.2')

    const reversed = [ordered[0], ordered[2], ordered[1]]
    const reversedFields = fieldsOf(reversed)
    expect(reversedFields('m').wbs).toBe('1.1')
    expect(reversedFields('a').wbs).toBe('1.2')

    // Identical input lists always derive identical output.
    expect(JSON.stringify([...deriveHierarchyFields(ordered).entries()])).toBe(
      JSON.stringify([...deriveHierarchyFields([...ordered]).entries()]),
    )
  })

  it('canonicalizeTasks emits outline order and recomputes derived fields', () => {
    // Interleaved input (children listed before their parent summary in the
    // array) still resolves to canonical outline order.
    const interleaved = [
      makeTask({ id: 'c1', parentTaskId: parent('s'), outlineLevel: 2 }),
      makeTask({ id: 'root' }),
      makeTask({ id: 'c2', parentTaskId: parent('s'), outlineLevel: 2 }),
      makeTask({ id: 's', summary: true }),
    ]
    const canonical = canonicalizeTasks(interleaved)
    expect(canonical.map((task) => task.id)).toEqual([
      asTaskId('root'),
      asTaskId('s'),
      asTaskId('c1'),
      asTaskId('c2'),
    ])
    const byId = new Map(canonical.map((task) => [task.id as string, task]))
    expect(byId.get('root')!.wbs).toBe('1')
    expect(byId.get('s')!.wbs).toBe('2')
    expect(byId.get('c1')!.wbs).toBe('2.1')
    expect(byId.get('c2')!.wbs).toBe('2.2')
    expect(byId.get('s')!.outlineLevel).toBe(1)
    expect(byId.get('c1')!.outlineLevel).toBe(2)
    expect(byId.get('s')!.summary).toBe(true)
    expect(byId.get('root')!.summary).toBe(false)
    // Canonicalization is idempotent.
    expect(canonicalizeTasks(canonical)).toEqual(canonical)
  })

  it('canonicalizeDocument canonicalizes the task list without touching other collections', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a' }),
        makeTask({ id: 's', summary: true }),
        makeTask({ id: 'b', parentTaskId: parent('s'), outlineLevel: 2 }),
      ],
    })
    const canonical = canonicalizeDocument(document)
    expect(canonical.tasks.map((task) => task.id)).toEqual([
      asTaskId('a'),
      asTaskId('s'),
      asTaskId('b'),
    ])
    expect(canonical.tasks.map((task) => task.wbs)).toEqual(['1', '2', '2.1'])
    expect(canonical.dependencies).toBe(document.dependencies)
    expect(canonical.calendars).toBe(document.calendars)
    // The input document is not mutated.
    expect(document.tasks.map((task) => task.id)).toEqual([
      asTaskId('a'),
      asTaskId('s'),
      asTaskId('b'),
    ])
    expect(document.tasks.every((task) => task.wbs === '')).toBe(true)
  })

  it('renumbers WBS after sibling removal', () => {
    const tasks = [
      makeTask({ id: 's', summary: true }),
      makeTask({ id: 'a', parentTaskId: parent('s'), outlineLevel: 2 }),
      makeTask({ id: 'b', parentTaskId: parent('s'), outlineLevel: 2 }),
      makeTask({ id: 'c', parentTaskId: parent('s'), outlineLevel: 2 }),
    ]
    const afterRemoval = canonicalizeTasks(tasks.filter((task) => task.id !== parent('b')))
    const byId = new Map(afterRemoval.map((task) => [task.id as string, task]))
    expect(byId.get('a')!.wbs).toBe('1.1')
    expect(byId.get('c')!.wbs).toBe('1.2')
  })

  it('rejects invalid hierarchies when building the tree', () => {
    expect(() =>
      buildTaskHierarchy([makeTask({ id: 'a', parentTaskId: parent('ghost') })]),
    ).toThrowError(HierarchyError)
    expect(() =>
      buildTaskHierarchy([makeTask({ id: 'a', parentTaskId: parent('a') })]),
    ).toThrowError(HierarchyError)
    expect(() =>
      buildTaskHierarchy([
        makeTask({ id: 'a', parentTaskId: parent('b') }),
        makeTask({ id: 'b', parentTaskId: parent('a') }),
      ]),
    ).toThrowError(HierarchyError)
  })
})
