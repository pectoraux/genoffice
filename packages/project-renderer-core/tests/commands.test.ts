import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import type { ProjectCommand } from '@genoffice/project-contracts'
import { applyProjectCommand } from '@genoffice/project-engine'
import {
  buildCreateTaskCommand,
  buildCreateTaskInSiblingGroupCommand,
  buildDeleteSelectionCommands,
  buildIndentCommand,
  buildOutdentCommand,
  defaultNewTask,
  nextTaskIdentity,
} from '../src/index.js'
import { makeDocument, makeTask, multiSiblingDocument, outlineDocument } from './fixtures.js'

const expectAccepted = (document: ReturnType<typeof makeDocument>, command: ProjectCommand) => {
  const execution = applyProjectCommand(document, command)
  expect(execution.result.accepted).toBe(true)
  return execution
}

describe('PROJECT-021 command builders — identity allocation', () => {
  it('allocates t1/uid 1 on an empty document', () => {
    expect(nextTaskIdentity(makeDocument())).toEqual({ id: asTaskId('t1'), uid: 1 })
  })

  it('allocates max(t{n})+1 and max(uid)+1 regardless of array order', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 't9', uid: 41 }),
        makeTask({ id: 't2', uid: 7 }),
        makeTask({ id: 'other', uid: 100 }),
      ],
    })
    expect(nextTaskIdentity(document)).toEqual({ id: asTaskId('t10'), uid: 101 })
    // Deterministic across repeated calls and reversed arrays.
    expect(nextTaskIdentity(document)).toEqual({ id: asTaskId('t10'), uid: 101 })
    const reversed = makeDocument({ tasks: [...document.tasks].reverse() })
    expect(nextTaskIdentity(reversed)).toEqual({ id: asTaskId('t10'), uid: 101 })
  })

  it('defaultNewTask carries the canonical creation defaults', () => {
    const task = defaultNewTask({ id: asTaskId('t1'), uid: 1 })
    expect(task.name).toBe('New Task')
    expect(task.taskType).toBe('fixedDuration')
    expect(task.duration).toBe(480 as never)
    expect(task.priority).toBe(500)
    expect(task.percentComplete).toBe(0)
    expect(task.summary).toBe(false)
    expect(task.autoScheduled).toBe(true)
  })
})

describe('PROJECT-021 command builders — CreateTask', () => {
  it('appends as the last root task by default', () => {
    const document = outlineDocument()
    const command = buildCreateTaskCommand(document)
    const execution = expectAccepted(document, command)
    const created = execution.document.tasks[execution.document.tasks.length - 1]!
    expect(created.id).toBe(asTaskId('t1'))
    expect(created.parentTaskId).toBeUndefined()
    expect(execution.document.tasks).toHaveLength(5)
  })

  it('appends as the LAST child of the named parent — never between siblings', () => {
    const document = multiSiblingDocument() // p > a1, a2, a3
    const command = buildCreateTaskCommand(document, {
      kind: 'lastChildOf',
      parentId: asTaskId('p'),
    })
    const execution = expectAccepted(document, command)
    const created = execution.document.tasks.find((task) => task.id === asTaskId('t1'))!
    expect(created.parentTaskId).toBe(asTaskId('p'))
    // The executable CreateTask semantics: appended after the LAST sibling
    // (a3) — the frozen command union has no row-position insert, so the
    // result is NEVER between a1/a2 or a2/a3.
    const childrenOfP = execution.document.tasks
      .filter((task) => task.parentTaskId === asTaskId('p'))
      .map((task) => task.id)
    expect(childrenOfP).toEqual([asTaskId('a1'), asTaskId('a2'), asTaskId('a3'), asTaskId('t1')])
    expect(execution.document.tasks.map((task) => task.id)).toEqual([
      asTaskId('root1'),
      asTaskId('p'),
      asTaskId('a1'),
      asTaskId('a2'),
      asTaskId('a3'),
      asTaskId('t1'),
      asTaskId('root2'),
    ])
  })

  it('the sibling-group gesture appends as the LAST member of the anchor group — non-last-anchor counterexample', () => {
    // THE non-last-sibling golden (review round 1): `a2` is NOT the last
    // child of `p` (a3 follows it). The honest gesture semantics: same
    // parent as the anchor, appended after the anchor's LAST sibling —
    // NOT immediately after the anchor.
    const document = multiSiblingDocument()
    const command = buildCreateTaskInSiblingGroupCommand(document, asTaskId('a2'))
    expect(command).toBeDefined()
    expect(command!.type).toBe('CreateTask')
    const execution = expectAccepted(document, command!)
    const created = execution.document.tasks.find((task) => task.id === asTaskId('t1'))!
    expect(created.parentTaskId).toBe(asTaskId('p')) // same parent as the anchor
    const childrenOfP = execution.document.tasks
      .filter((task) => task.parentTaskId === asTaskId('p'))
      .map((task) => task.id)
    expect(childrenOfP).toEqual([asTaskId('a1'), asTaskId('a2'), asTaskId('a3'), asTaskId('t1')])
    // Explicitly NOT immediately after the anchor a2:
    const order = execution.document.tasks.map((task) => task.id)
    expect(order.indexOf(asTaskId('t1'))).not.toBe(order.indexOf(asTaskId('a2')) + 1)
    expect(order.indexOf(asTaskId('t1'))).toBe(order.indexOf(asTaskId('a3')) + 1)
  })

  it('the sibling-group gesture on a ROOT anchor appends as the last root task', () => {
    const document = multiSiblingDocument() // roots: root1, root2 — anchor root1 is NOT the last root
    const command = buildCreateTaskInSiblingGroupCommand(document, asTaskId('root1'))
    expect(command).toBeDefined()
    const execution = expectAccepted(document, command!)
    const created = execution.document.tasks.find((task) => task.id === asTaskId('t1'))!
    expect(created.parentTaskId).toBeUndefined()
    // Appended after the LAST root (root2), not immediately after root1.
    expect(execution.document.tasks.map((task) => task.id)).toEqual([
      asTaskId('root1'),
      asTaskId('p'),
      asTaskId('a1'),
      asTaskId('a2'),
      asTaskId('a3'),
      asTaskId('root2'),
      asTaskId('t1'),
    ])
  })

  it('the sibling-group gesture refuses an unknown anchor (disabled gesture)', () => {
    expect(
      buildCreateTaskInSiblingGroupCommand(outlineDocument(), asTaskId('nope')),
    ).toBeUndefined()
  })

  it('is deterministic: the same document yields the byte-identical command', () => {
    const document = multiSiblingDocument()
    expect(JSON.stringify(buildCreateTaskCommand(document))).toBe(
      JSON.stringify(buildCreateTaskCommand(document)),
    )
    expect(JSON.stringify(buildCreateTaskInSiblingGroupCommand(document, asTaskId('a2')))).toBe(
      JSON.stringify(buildCreateTaskInSiblingGroupCommand(document, asTaskId('a2'))),
    )
  })
})

describe('PROJECT-021 command builders — indent/outdent', () => {
  it('builds IndentTask onto the immediately preceding sibling and the engine accepts it', () => {
    const document = outlineDocument() // root, a, a1, b — `b`'s preceding sibling is `a`
    const command = buildIndentCommand(document, asTaskId('b'))
    expect(command).toEqual({
      type: 'IndentTask',
      taskId: asTaskId('b'),
      parentTaskId: asTaskId('a'),
    })
    expectAccepted(document, command!)
    // A second child of `a`: its preceding sibling is the first child (`a1`),
    // never an uncle or a cousin.
    const withSecondChild = makeDocument({
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
        makeTask({ id: 'a2', parentTaskId: asTaskId('a'), outlineLevel: 3, wbs: '1.1.2' }),
        makeTask({ id: 'b', parentTaskId: asTaskId('root'), outlineLevel: 2, wbs: '1.2' }),
      ],
    })
    expect(buildIndentCommand(withSecondChild, asTaskId('a2'))).toEqual({
      type: 'IndentTask',
      taskId: asTaskId('a2'),
      parentTaskId: asTaskId('a1'),
    })
  })

  it('returns undefined when there is no preceding sibling (first root task)', () => {
    const document = outlineDocument()
    expect(buildIndentCommand(document, asTaskId('root'))).toBeUndefined()
    expect(buildIndentCommand(document, asTaskId('nope'))).toBeUndefined()
  })

  it('builds OutdentTask for a child and refuses a root', () => {
    const document = outlineDocument()
    expect(buildOutdentCommand(document, asTaskId('a'))).toEqual({
      type: 'OutdentTask',
      taskId: asTaskId('a'),
    })
    expect(buildOutdentCommand(document, asTaskId('root'))).toBeUndefined()
    expect(buildOutdentCommand(document, asTaskId('nope'))).toBeUndefined()
  })
})

describe('PROJECT-021 command builders — delete selection', () => {
  it('deletes the top-most selected tasks in reverse outline order', () => {
    const document = outlineDocument() // root, a, a1, b (indices 0..3)
    // Selecting a summary AND its descendant: the ancestor's deletion covers it.
    const commands = buildDeleteSelectionCommands(document, [
      asTaskId('a1'),
      asTaskId('a'),
      asTaskId('b'),
    ])
    expect(commands).toEqual([
      { type: 'DeleteTask', taskId: asTaskId('b') }, // index 3 first (reverse order)
      { type: 'DeleteTask', taskId: asTaskId('a') }, // index 1 second; a1 covered
    ])
  })

  it('returns an empty command list for an empty selection', () => {
    expect(buildDeleteSelectionCommands(outlineDocument(), [])).toEqual([])
  })

  it('produces commands the engine accepts, leaving a valid document', () => {
    const document = outlineDocument()
    const commands = buildDeleteSelectionCommands(document, [asTaskId('a'), asTaskId('b')])
    let current = document
    for (const command of commands) current = expectAccepted(current, command).document
    expect(current.tasks.map((task) => task.id)).toEqual([asTaskId('root')])
  })

  it('ignores unknown selected ids deterministically', () => {
    const document = outlineDocument()
    const commands = buildDeleteSelectionCommands(document, [asTaskId('b'), asTaskId('ghost')])
    expect(commands).toEqual([{ type: 'DeleteTask', taskId: asTaskId('b') }])
  })
})
