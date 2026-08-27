import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import type { ProjectCommand } from '@genoffice/project-contracts'
import { applyProjectCommand } from '@genoffice/project-engine'
import {
  buildCreateTaskCommand,
  buildDeleteSelectionCommands,
  buildIndentCommand,
  buildOutdentCommand,
  defaultNewTask,
  nextTaskIdentity,
} from '../src/index.js'
import { makeDocument, makeTask, outlineDocument } from './fixtures.js'

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
  it('inserts as the last root task by default', () => {
    const document = outlineDocument()
    const command = buildCreateTaskCommand(document)
    const execution = expectAccepted(document, command)
    const created = execution.document.tasks[execution.document.tasks.length - 1]!
    expect(created.id).toBe(asTaskId('t1'))
    expect(created.parentTaskId).toBeUndefined()
    expect(execution.document.tasks).toHaveLength(5)
  })

  it('inserts as the sibling after the anchor task (same parent)', () => {
    const document = outlineDocument() // root, a, a1, b
    const command = buildCreateTaskCommand(document, { afterTaskId: asTaskId('a1') })
    const execution = expectAccepted(document, command)
    // a1's parent is `a`; the new task becomes a's last child (after a1).
    const created = execution.document.tasks.find((task) => task.id === asTaskId('t1'))!
    expect(created.parentTaskId).toBe(asTaskId('a'))
    expect(execution.document.tasks.indexOf(created)).toBe(
      execution.document.tasks.indexOf(
        execution.document.tasks.find((t) => t.id === asTaskId('a1'))!,
      ) + 1,
    )
  })

  it('inserts as the last child of the named parent', () => {
    const document = outlineDocument()
    const command = buildCreateTaskCommand(document, { parentId: asTaskId('root') })
    const execution = expectAccepted(document, command)
    const created = execution.document.tasks.find((task) => task.id === asTaskId('t1'))!
    expect(created.parentTaskId).toBe(asTaskId('root'))
  })

  it('is deterministic: the same document yields the byte-identical command', () => {
    const document = outlineDocument()
    expect(JSON.stringify(buildCreateTaskCommand(document))).toBe(
      JSON.stringify(buildCreateTaskCommand(document)),
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
