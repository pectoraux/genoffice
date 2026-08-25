import { describe, expect, it } from 'vitest'
import {
  asAssignmentId,
  asBaselineId,
  asDependencyId,
  asISODateTime,
  asResourceId,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type { ProjectCommand, ProjectDocument, Task } from '@genoffice/project-contracts'
import { ProjectJournal, applyProjectCommand, canonicalizeDocument } from '../src/index.js'
import { makeDependency, makeDocument, makeTask } from './fixtures.js'

const parent = (id: string) => asTaskId(id)

const taskIn = (document: ProjectDocument, id: string): Task => {
  const task = document.tasks.find((candidate) => candidate.id === parent(id))
  if (!task) throw new Error(`missing task ${id}`)
  return task
}

const expectAccepted = (document: ProjectDocument, command: ProjectCommand) => {
  const execution = applyProjectCommand(document, command)
  if (!execution.result.accepted) {
    throw new Error(`command rejected: ${JSON.stringify(execution.result.diagnostics)}`)
  }
  return execution
}

const expectRejected = (document: ProjectDocument, command: ProjectCommand, code: string) => {
  const execution = applyProjectCommand(document, command)
  expect(execution.result.accepted).toBe(false)
  expect(execution.document).toBe(document)
  expect(execution.result.diagnostics.some((diagnostic) => diagnostic.code === code)).toBe(true)
  return execution
}

const createCommand = (overrides: Omit<Partial<Task>, 'id'> & { id: string }): ProjectCommand => ({
  type: 'CreateTask',
  task: makeTask(overrides),
})

const applyAll = (document: ProjectDocument, commands: ProjectCommand[]): ProjectDocument =>
  commands.reduce((current, command) => expectAccepted(current, command).document, document)

const parseDocument = (json: string): ProjectDocument => JSON.parse(json) as ProjectDocument

describe('PROJECT-007 CreateTask', () => {
  it('creates a root task with stable identity, derived WBS, and outline level 1', () => {
    const base = makeDocument()
    const execution = expectAccepted(base, createCommand({ id: 'first', uid: 41, name: 'First' }))
    const task = taskIn(execution.document, 'first')
    expect(task.id).toBe(parent('first'))
    expect(task.uid).toBe(41)
    expect(task.name).toBe('First')
    expect(task.outlineLevel).toBe(1)
    expect(task.wbs).toBe('1')
    expect(task.summary).toBe(false)
    expect(execution.result.affectedTaskIds).toEqual([parent('first')])
    expect(execution.result.inverse).toEqual({ type: 'DeleteTask', taskId: parent('first') })
    // Input document untouched.
    expect(base.tasks).toHaveLength(0)
  })

  it('creates a child task as the last child of its parent', () => {
    let document = applyAll(makeDocument(), [
      createCommand({ id: 's', uid: 1 }),
      createCommand({ id: 'a', uid: 2, parentTaskId: parent('s') }),
    ])
    document = expectAccepted(
      document,
      createCommand({ id: 'b', uid: 3, parentTaskId: parent('s') }),
    ).document
    expect(taskIn(document, 's').summary).toBe(true)
    expect(taskIn(document, 'a').wbs).toBe('1.1')
    expect(taskIn(document, 'b').wbs).toBe('1.2')
    expect(taskIn(document, 'a').outlineLevel).toBe(2)
    expect(taskIn(document, 'b').outlineLevel).toBe(2)
    expect(document.tasks.map((task) => task.id)).toEqual([parent('s'), parent('a'), parent('b')])
  })

  it('rejects creation under a missing parent and returns the input unchanged', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      createCommand({ id: 'x', parentTaskId: parent('ghost') }),
      'MISSING_PARENT',
    )
  })

  it('rejects creation with a duplicate task id', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(base, createCommand({ id: 'a' }), 'DUPLICATE_TASK_ID')
  })

  it('rejects creation with a duplicate uid', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a', uid: 7 })] }))
    expectRejected(base, createCommand({ id: 'b', uid: 7 }), 'DUPLICATE_TASK_UID')
  })

  it('rejects creation with a self parent', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(base, createCommand({ id: 'x', parentTaskId: parent('x') }), 'SELF_PARENT')
  })
})

describe('PROJECT-007 RenameTask', () => {
  it('renames without touching identity', () => {
    const base = canonicalizeDocument(
      makeDocument({ tasks: [makeTask({ id: 'a', uid: 5, name: 'Old' })] }),
    )
    const execution = expectAccepted(base, {
      type: 'RenameTask',
      taskId: parent('a'),
      name: 'New',
    })
    const task = taskIn(execution.document, 'a')
    expect(task.name).toBe('New')
    expect(task.id).toBe(parent('a'))
    expect(task.uid).toBe(5)
    expect(task.wbs).toBe('1')
    expect(execution.result.inverse).toEqual({
      type: 'RenameTask',
      taskId: parent('a'),
      name: 'Old',
    })
  })

  it('rejects renaming a missing task', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(base, { type: 'RenameTask', taskId: parent('ghost'), name: 'x' }, 'MISSING_TASK')
  })
})

describe('PROJECT-007 IndentTask', () => {
  const flat = () =>
    canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      }),
    )

  it('indents a task beneath its preceding sibling and derives hierarchy fields', () => {
    const execution = expectAccepted(flat(), {
      type: 'IndentTask',
      taskId: parent('c'),
      parentTaskId: parent('b'),
    })
    const document = execution.document
    expect(taskIn(document, 'c').parentTaskId).toBe(parent('b'))
    expect(taskIn(document, 'c').outlineLevel).toBe(2)
    expect(taskIn(document, 'c').wbs).toBe('2.1')
    expect(taskIn(document, 'b').summary).toBe(true)
    expect(taskIn(document, 'b').wbs).toBe('2')
    expect(taskIn(document, 'a').wbs).toBe('1')
    expect(taskIn(document, 'c').id).toBe(parent('c'))
    expect(execution.result.affectedTaskIds).toEqual([parent('c'), parent('b')])
    expect(execution.result.inverse).toEqual({ type: 'OutdentTask', taskId: parent('c') })
  })

  it('indents beneath a preceding summary as its last child', () => {
    const document = applyAll(flat(), [
      { type: 'IndentTask', taskId: parent('b'), parentTaskId: parent('a') },
      { type: 'IndentTask', taskId: parent('c'), parentTaskId: parent('a') },
    ])
    // a > [b, c]: c is the last child of summary a.
    expect(taskIn(document, 'c').parentTaskId).toBe(parent('a'))
    expect(taskIn(document, 'b').wbs).toBe('1.1')
    expect(taskIn(document, 'c').wbs).toBe('1.2')
    expect(taskIn(document, 'a').summary).toBe(true)
  })

  it('moves a whole subtree when indenting a summary task', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'a' }),
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'child', parentTaskId: parent('s'), outlineLevel: 2 }),
          makeTask({ id: 'z' }),
        ],
      }),
    )
    // Outline order is a, s, child, z — z's preceding sibling is s.
    const execution = expectAccepted(base, {
      type: 'IndentTask',
      taskId: parent('z'),
      parentTaskId: parent('s'),
    })
    const document = execution.document
    expect(taskIn(document, 'z').parentTaskId).toBe(parent('s'))
    expect(taskIn(document, 'z').wbs).toBe('2.2')
    expect(taskIn(document, 'child').wbs).toBe('2.1')
    expect(taskIn(document, 's').summary).toBe(true)
  })

  it('rejects indenting the first task with no preceding sibling', () => {
    expectRejected(
      flat(),
      { type: 'IndentTask', taskId: parent('a'), parentTaskId: parent('b') },
      'INVALID_INDENT_NO_SIBLING',
    )
  })

  it('rejects indenting beneath a target that is not the preceding sibling', () => {
    expectRejected(
      flat(),
      { type: 'IndentTask', taskId: parent('c'), parentTaskId: parent('a') },
      'INVALID_INDENT_TARGET',
    )
  })

  it('rejects indenting beneath itself', () => {
    expectRejected(
      flat(),
      { type: 'IndentTask', taskId: parent('b'), parentTaskId: parent('b') },
      'INDENT_CYCLE',
    )
  })

  it('rejects indenting beneath a descendant', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'child', parentTaskId: parent('s'), outlineLevel: 2 }),
          makeTask({ id: 'other' }),
        ],
      }),
    )
    expectRejected(
      base,
      { type: 'IndentTask', taskId: parent('s'), parentTaskId: parent('child') },
      'INDENT_CYCLE',
    )
  })

  it('rejects indenting a missing task or beneath a missing parent', () => {
    expectRejected(
      flat(),
      { type: 'IndentTask', taskId: parent('ghost'), parentTaskId: parent('a') },
      'MISSING_TASK',
    )
    expectRejected(
      flat(),
      { type: 'IndentTask', taskId: parent('b'), parentTaskId: parent('ghost') },
      'MISSING_PARENT',
    )
  })

  it('rejects commands against an invalid document', () => {
    const broken = makeDocument({
      tasks: [
        makeTask({ id: 'a', parentTaskId: parent('b') }),
        makeTask({ id: 'b', parentTaskId: parent('a') }),
      ],
    })
    expectRejected(
      broken,
      { type: 'IndentTask', taskId: parent('a'), parentTaskId: parent('b') },
      'PARENT_CYCLE',
    )
  })
})

describe('PROJECT-007 OutdentTask', () => {
  const nested = () =>
    canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'r1' }),
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'a', parentTaskId: parent('s'), outlineLevel: 2 }),
          makeTask({ id: 'b', parentTaskId: parent('s'), outlineLevel: 2 }),
          makeTask({ id: 'r2' }),
        ],
      }),
    )

  it('outdents a task to become the next sibling of its former parent', () => {
    const execution = expectAccepted(nested(), { type: 'OutdentTask', taskId: parent('b') })
    const document = execution.document
    const b = taskIn(document, 'b')
    expect(b.parentTaskId).toBeUndefined()
    expect(b.outlineLevel).toBe(1)
    expect(b.wbs).toBe('3')
    expect(taskIn(document, 's').wbs).toBe('2')
    expect(taskIn(document, 'a').wbs).toBe('2.1')
    expect(taskIn(document, 's').summary).toBe(true)
    expect(execution.result.affectedTaskIds).toEqual([parent('b'), parent('s')])
    expect(execution.result.inverse).toBeUndefined()
  })

  it('reverts a former parent to a leaf when its last child is outdented', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'only', parentTaskId: parent('s'), outlineLevel: 2 }),
        ],
      }),
    )
    const document = expectAccepted(base, { type: 'OutdentTask', taskId: parent('only') }).document
    expect(taskIn(document, 's').summary).toBe(false)
    expect(taskIn(document, 'only').outlineLevel).toBe(1)
    expect(taskIn(document, 'only').wbs).toBe('2')
    expect(taskIn(document, 's').wbs).toBe('1')
  })

  it('outdents nested tasks one level at a time, deterministically', () => {
    let document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'l1', summary: true }),
          makeTask({ id: 'l2', summary: true, parentTaskId: parent('l1'), outlineLevel: 2 }),
          makeTask({ id: 'l3', summary: true, parentTaskId: parent('l2'), outlineLevel: 3 }),
          makeTask({ id: 'leaf', parentTaskId: parent('l3'), outlineLevel: 4 }),
        ],
      }),
    )
    document = expectAccepted(document, { type: 'OutdentTask', taskId: parent('leaf') }).document
    expect(taskIn(document, 'leaf').parentTaskId).toBe(parent('l2'))
    expect(taskIn(document, 'leaf').outlineLevel).toBe(3)
    expect(taskIn(document, 'leaf').wbs).toBe('1.1.2')
    document = expectAccepted(document, { type: 'OutdentTask', taskId: parent('leaf') }).document
    expect(taskIn(document, 'leaf').parentTaskId).toBe(parent('l1'))
    expect(taskIn(document, 'leaf').outlineLevel).toBe(2)
    expect(taskIn(document, 'leaf').wbs).toBe('1.2')
    document = expectAccepted(document, { type: 'OutdentTask', taskId: parent('leaf') }).document
    expect(taskIn(document, 'leaf').parentTaskId).toBeUndefined()
    expect(taskIn(document, 'leaf').outlineLevel).toBe(1)
    expect(taskIn(document, 'leaf').wbs).toBe('2')
  })

  it('rejects outdenting a root task', () => {
    expectRejected(nested(), { type: 'OutdentTask', taskId: parent('r1') }, 'INVALID_OUTDENT_ROOT')
  })

  it('rejects outdenting a missing task', () => {
    expectRejected(nested(), { type: 'OutdentTask', taskId: parent('ghost') }, 'MISSING_TASK')
  })
})

describe('PROJECT-007 DeleteTask', () => {
  it('deletes a leaf task and renumbers WBS', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'a', parentTaskId: parent('s'), outlineLevel: 2 }),
          makeTask({ id: 'b', parentTaskId: parent('s'), outlineLevel: 2 }),
        ],
      }),
    )
    const execution = expectAccepted(base, { type: 'DeleteTask', taskId: parent('b') })
    const document = execution.document
    expect(document.tasks.map((task) => task.id)).toEqual([parent('s'), parent('a')])
    expect(taskIn(document, 'a').wbs).toBe('1.1')
    expect(taskIn(document, 's').summary).toBe(true)
    expect(execution.result.affectedTaskIds).toEqual([parent('b'), parent('s')])
  })

  it('deletes a summary task with its whole subtree (canonical policy)', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'keep' }),
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'mid', summary: true, parentTaskId: parent('s'), outlineLevel: 2 }),
          makeTask({ id: 'leaf', parentTaskId: parent('mid'), outlineLevel: 3 }),
          makeTask({ id: 'other' }),
        ],
      }),
    )
    const execution = expectAccepted(base, { type: 'DeleteTask', taskId: parent('s') })
    const document = execution.document
    expect(document.tasks.map((task) => task.id)).toEqual([parent('keep'), parent('other')])
    expect(taskIn(document, 'keep').wbs).toBe('1')
    expect(taskIn(document, 'other').wbs).toBe('2')
  })

  it('removes dependencies, assignments, and baseline snapshots referencing deleted tasks', () => {
    const resource = {
      id: asResourceId('r1'),
      uid: 1,
      name: 'Worker',
      kind: 'work' as const,
      maxUnits: 1,
      standardRate: 0,
      overtimeRate: 0,
      costPerUse: 0,
      availability: [],
    }
    const assignment = (id: string, taskId: string) => ({
      id: asAssignmentId(id),
      taskId: parent(taskId),
      resourceId: asResourceId('r1'),
      units: 1,
      work: asWorkingMinutes(0),
      actualWork: asWorkingMinutes(0),
      remainingWork: asWorkingMinutes(0),
      cost: 0,
      actualCost: 0,
      remainingCost: 0,
    })
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'x' }),
          makeTask({ id: 'y' }),
          makeTask({ id: 's', summary: true }),
          makeTask({ id: 'a', parentTaskId: parent('s'), outlineLevel: 2 }),
        ],
        dependencies: [
          makeDependency('dep-keep', 'y', 'x', 'FS'),
          makeDependency('dep-gone-1', 'x', 'a', 'FS'),
          makeDependency('dep-gone-2', 'x', 's', 'SS'),
        ],
        resources: [resource],
        assignments: [assignment('as-keep', 'x'), assignment('as-gone', 'a')],
        baselines: [
          {
            id: asBaselineId('b1'),
            name: 'Baseline',
            capturedAt: asISODateTime('2026-08-03T09:00:00.000Z'),
            taskSnapshots: {
              [parent('a')]: {
                duration: asWorkingMinutes(480),
                work: asWorkingMinutes(0),
                cost: 0,
              },
              [parent('x')]: {
                duration: asWorkingMinutes(480),
                work: asWorkingMinutes(0),
                cost: 0,
              },
            },
          },
        ],
      }),
    )
    const execution = expectAccepted(base, { type: 'DeleteTask', taskId: parent('s') })
    const document = execution.document
    expect(document.tasks.map((task) => task.id)).toEqual([parent('x'), parent('y')])
    expect(document.dependencies.map((dependency) => dependency.id)).toEqual([
      asDependencyId('dep-keep'),
    ])
    expect(document.assignments.map((item) => item.id)).toEqual([asAssignmentId('as-keep')])
    expect(Object.keys(document.baselines[0]!.taskSnapshots)).toEqual(['x'])
  })

  it('rejects deleting a missing task', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(base, { type: 'DeleteTask', taskId: parent('ghost') }, 'MISSING_TASK')
  })
})

describe('PROJECT-007 command determinism', () => {
  it('repeated identical command sequences produce byte-identical documents', () => {
    const commands: ProjectCommand[] = [
      createCommand({ id: 's', uid: 1 }),
      createCommand({ id: 'a', uid: 2, parentTaskId: parent('s') }),
      createCommand({ id: 'b', uid: 3, parentTaskId: parent('s') }),
      createCommand({ id: 'c', uid: 4 }),
      { type: 'IndentTask', taskId: parent('c'), parentTaskId: parent('s') },
      { type: 'RenameTask', taskId: parent('a'), name: 'Renamed' },
      { type: 'OutdentTask', taskId: parent('c') },
      { type: 'DeleteTask', taskId: parent('b') },
    ]
    const serialized = JSON.stringify(makeDocument())
    const first = applyAll(parseDocument(serialized), commands)
    const second = applyAll(parseDocument(serialized), commands)
    const third = applyAll(parseDocument(serialized), commands)
    expect(JSON.stringify(second)).toBe(JSON.stringify(first))
    expect(JSON.stringify(third)).toBe(JSON.stringify(first))
    expect(taskIn(first, 's').summary).toBe(true)
    expect(taskIn(first, 'a').wbs).toBe('1.1')
    expect(taskIn(first, 'c').wbs).toBe('2')
  })

  it('accepted commands never mutate the input document', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      }),
    )
    const snapshot = JSON.stringify(base)
    applyProjectCommand(base, {
      type: 'IndentTask',
      taskId: parent('c'),
      parentTaskId: parent('b'),
    })
    applyProjectCommand(base, { type: 'RenameTask', taskId: parent('a'), name: 'x' })
    expect(JSON.stringify(base)).toBe(snapshot)
  })

  it('indent then outdent restores the canonical document exactly', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' }), makeTask({ id: 'c' })],
      }),
    )
    const indent = expectAccepted(base, {
      type: 'IndentTask',
      taskId: parent('c'),
      parentTaskId: parent('b'),
    })
    expect(indent.result.inverse).toEqual({ type: 'OutdentTask', taskId: parent('c') })
    const undone = expectAccepted(indent.document, indent.result.inverse!)
    expect(JSON.stringify(undone.document)).toBe(JSON.stringify(base))
  })

  it('journal records and undoes hierarchy commands', () => {
    const journal = new ProjectJournal()
    const base = canonicalizeDocument(
      makeDocument({ tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })] }),
    )
    const command: ProjectCommand = {
      type: 'IndentTask',
      taskId: parent('b'),
      parentTaskId: parent('a'),
    }
    const execution = applyProjectCommand(base, command, 'cmd-1')
    journal.record(command, execution.result)
    expect(journal.canUndo()).toBe(true)
    const entry = journal.undo()!
    expect(entry.commandId).toBe('cmd-1')
    const restored = applyProjectCommand(execution.document, entry.result.inverse!)
    expect(JSON.stringify(restored.document)).toBe(JSON.stringify(base))
  })

  it('rejects unimplemented command types deterministically', () => {
    const base = canonicalizeDocument(
      makeDocument({ tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })] }),
    )
    expectRejected(
      base,
      { type: 'AddDependency', dependency: makeDependency('d1', 'a', 'b', 'FS') },
      'UNSUPPORTED_COMMAND',
    )
  })
})
