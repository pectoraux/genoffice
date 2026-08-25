import { describe, expect, it } from 'vitest'
import type { Baseline, ProjectCommand, ProjectDocument, Task } from '@genoffice/project-contracts'
import { asBaselineId, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import { applyProjectCommand, canonicalizeDocument, validateProjectDocument } from '../src/index.js'
import { makeBaseline, makeDocument, makeTask } from './fixtures.js'

const parent = (id: string) => asTaskId(id)

const taskIn = (document: ProjectDocument, id: string): Task => {
  const task = document.tasks.find((candidate) => candidate.id === parent(id))
  if (!task) throw new Error(`missing task ${id}`)
  return task
}

const baselineIn = (document: ProjectDocument, id: string): Baseline => {
  const baseline = document.baselines.find((candidate) => candidate.id === asBaselineId(id))
  if (!baseline) throw new Error(`missing baseline ${id}`)
  return baseline
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
  expect(execution.result.diagnostics.some((d) => d.code === code)).toBe(true)
  return execution
}

const snapshot = (document: ProjectDocument): string => JSON.stringify(document)
const parseDocument = (json: string): ProjectDocument => JSON.parse(json) as ProjectDocument

const CAPTURED_AT = '2026-08-03T09:00:00.000Z'

describe('PROJECT-009 CreateBaseline', () => {
  it('stores a fully-formed baseline and reports the snapshotted task ids', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const baseline = makeBaseline('b1', CAPTURED_AT, {
      a: { start: '2026-08-03T09:00:00.000Z', finish: '2026-08-03T17:00:00.000Z', duration: 480 },
    })
    const execution = expectAccepted(base, { type: 'CreateBaseline', baseline })
    expect(baselineIn(execution.document, 'b1')).toEqual(baseline)
    expect(execution.result.affectedTaskIds).toEqual([parent('a')])
    // No inverse: a future DeleteBaseline command is required to undo cleanly.
    expect(execution.result.inverse).toBeUndefined()
    // The input document is never mutated.
    expect(base.baselines).toEqual([])
    expect(taskIn(base, 'a').baseline).toEqual([])
  })

  it('adds the baseline id to every snapshotted task.baseline (reverse index)', () => {
    const base = canonicalizeDocument(
      makeDocument({ tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })] }),
    )
    const baseline = makeBaseline('b1', CAPTURED_AT, {
      a: { duration: 480 },
      b: { duration: 240 },
    })
    const execution = expectAccepted(base, { type: 'CreateBaseline', baseline })
    expect(taskIn(execution.document, 'a').baseline).toEqual([asBaselineId('b1')])
    expect(taskIn(execution.document, 'b').baseline).toEqual([asBaselineId('b1')])
  })

  it('preserves existing baseline references when a second baseline is added', () => {
    const one = makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 } })
    const withB1 = expectAccepted(
      canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] })),
      { type: 'CreateBaseline', baseline: one },
    ).document
    const two = makeBaseline('b2', CAPTURED_AT, { a: { duration: 240 } })
    const withB2 = expectAccepted(withB1, { type: 'CreateBaseline', baseline: two })
    expect(taskIn(withB2.document, 'a').baseline).toEqual([asBaselineId('b1'), asBaselineId('b2')])
    expect(withB2.document.baselines.map((b) => b.id)).toEqual([
      asBaselineId('b1'),
      asBaselineId('b2'),
    ])
  })

  it('does not duplicate the baseline id when the same baseline id is re-added to a task', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a', baseline: [asBaselineId('b1')] })],
        baselines: [makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 } })],
      }),
    )
    // Re-issuing CreateBaseline with the same id is rejected as a duplicate,
    // so the task's baseline array is never mutated.
    const execution = expectRejected(
      base,
      {
        type: 'CreateBaseline',
        baseline: makeBaseline('b1', CAPTURED_AT, { a: { duration: 240 } }),
      },
      'DUPLICATE_BASELINE_ID',
    )
    expect(taskIn(execution.document, 'a').baseline).toEqual([asBaselineId('b1')])
  })

  it('rejects a duplicate baseline id', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        baselines: [makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 } })],
      }),
    )
    expectRejected(
      base,
      {
        type: 'CreateBaseline',
        baseline: makeBaseline('b1', CAPTURED_AT, { a: { duration: 240 } }),
      },
      'DUPLICATE_BASELINE_ID',
    )
  })

  it('rejects a snapshot referencing a missing task', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const execution = expectRejected(
      base,
      {
        type: 'CreateBaseline',
        baseline: makeBaseline('b1', CAPTURED_AT, {
          a: { duration: 480 },
          ghost: { duration: 240 },
        }),
      },
      'MISSING_TASK_REFERENCE',
    )
    expect(execution.result.diagnostics.some((d) => d.message.includes('ghost'))).toBe(true)
  })

  it('rejects a malformed capturedAt via post-mutation validation', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      {
        type: 'CreateBaseline',
        baseline: makeBaseline('b1', 'not-a-date', { a: { duration: 480 } }),
      },
      'INVALID_DATE',
    )
  })

  it('accepts a baseline with an empty snapshot map (a tracking shell)', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const execution = expectAccepted(base, {
      type: 'CreateBaseline',
      baseline: makeBaseline('b1', CAPTURED_AT, {}),
    })
    expect(execution.document.baselines.length).toBe(1)
    expect(execution.result.affectedTaskIds).toEqual([])
    // No task references the baseline because no snapshots were captured.
    expect(taskIn(execution.document, 'a').baseline).toEqual([])
  })
})

describe('PROJECT-009 baseline validation surface', () => {
  it('document validation reports MISSING_TASK_REFERENCE for a dangling snapshot', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      baselines: [makeBaseline('b1', CAPTURED_AT, { ghost: { duration: 480 } })],
    })
    const codes = validateProjectDocument(document).diagnostics.map((d) => d.code)
    expect(codes).toContain('MISSING_TASK_REFERENCE')
  })

  it('document validation reports MISSING_BASELINE_REFERENCE for a task referencing a missing baseline', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', baseline: [asBaselineId('ghost')] })],
    })
    const codes = validateProjectDocument(document).diagnostics.map((d) => d.code)
    expect(codes).toContain('MISSING_BASELINE_REFERENCE')
  })

  it('a baseline + its tasks form a bidirectionally consistent valid document', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
      baselines: [makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 }, b: { duration: 240 } })],
    })
    expect(validateProjectDocument(document).accepted).toBe(true)
  })
})

describe('PROJECT-009 baseline immutability through hierarchy mutations', () => {
  const docWithBaseline = (): ProjectDocument => {
    const base = makeDocument({
      tasks: [
        makeTask({ id: 's', summary: true, duration: asWorkingMinutes(0) }),
        makeTask({ id: 'a' }),
      ],
    })
    const canonical = canonicalizeDocument(base)
    return applyProjectCommand(canonical, {
      type: 'CreateBaseline',
      baseline: makeBaseline('b1', CAPTURED_AT, {
        s: { duration: 0 },
        a: { duration: 480 },
      }),
    }).document
  }

  it('RenameTask preserves baselines and task.baseline arrays', () => {
    const doc = docWithBaseline()
    const execution = applyProjectCommand(doc, {
      type: 'RenameTask',
      taskId: parent('a'),
      name: 'A2',
    })
    expect(execution.result.accepted).toBe(true)
    expect(execution.document.baselines.length).toBe(1)
    expect(execution.document.baselines[0].id).toEqual(asBaselineId('b1'))
    expect(taskIn(execution.document, 'a').baseline).toEqual([asBaselineId('b1')])
    expect(taskIn(execution.document, 's').baseline).toEqual([asBaselineId('b1')])
  })

  it('IndentTask preserves baseline snapshots (stable TaskId identity)', () => {
    // Two sibling root tasks; indenting 'a' under its preceding sibling 'root'
    // changes hierarchy but never TaskId, so the baseline survives intact.
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'root',
            summary: true,
            duration: asWorkingMinutes(0),
            baseline: [asBaselineId('b1')],
          }),
          makeTask({ id: 'a', baseline: [asBaselineId('b1')] }),
          makeTask({ id: 'b', baseline: [asBaselineId('b1')] }),
        ],
        baselines: [makeBaseline('b1', CAPTURED_AT, { root: {}, a: {}, b: {} })],
      }),
    )
    const execution = applyProjectCommand(base, {
      type: 'IndentTask',
      taskId: parent('a'),
      parentTaskId: parent('root'),
    })
    expect(execution.result.accepted).toBe(true)
    expect(execution.document.baselines.length).toBe(1)
    expect(Object.keys(execution.document.baselines[0].taskSnapshots).sort()).toEqual([
      'a',
      'b',
      'root',
    ])
    expect(taskIn(execution.document, 'a').baseline).toEqual([asBaselineId('b1')])
    expect(taskIn(execution.document, 'a').parentTaskId).toEqual(parent('root'))
  })

  it('OutdentTask preserves baseline snapshots (stable TaskId identity)', () => {
    // 'a' is a child of 'root'; outdenting makes it a sibling of 'root' but the
    // baseline snapshot keyed by 'a' survives because TaskId never changes.
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'root',
            summary: true,
            duration: asWorkingMinutes(0),
            baseline: [asBaselineId('b1')],
          }),
          makeTask({
            id: 'a',
            parentTaskId: parent('root'),
            outlineLevel: 2,
            baseline: [asBaselineId('b1')],
          }),
        ],
        baselines: [makeBaseline('b1', CAPTURED_AT, { root: {}, a: {} })],
      }),
    )
    const execution = applyProjectCommand(base, { type: 'OutdentTask', taskId: parent('a') })
    expect(execution.result.accepted).toBe(true)
    expect(execution.document.baselines.length).toBe(1)
    expect(Object.keys(execution.document.baselines[0].taskSnapshots).sort()).toEqual(['a', 'root'])
    expect(taskIn(execution.document, 'a').baseline).toEqual([asBaselineId('b1')])
    expect(taskIn(execution.document, 'a').parentTaskId).toBeUndefined()
  })

  it('DeleteTask prunes the deleted task snapshot from every baseline', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'a', baseline: [asBaselineId('b1')] }),
          makeTask({ id: 'b', baseline: [asBaselineId('b1')] }),
        ],
        baselines: [
          makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 }, b: { duration: 240 } }),
        ],
      }),
    )
    const execution = applyProjectCommand(base, { type: 'DeleteTask', taskId: parent('a') })
    expect(execution.result.accepted).toBe(true)
    // 'a' is gone; its snapshot is pruned from b1. 'b' remains tracked.
    expect(execution.document.tasks.map((t) => t.id)).toEqual([parent('b')])
    expect(Object.keys(execution.document.baselines[0].taskSnapshots)).toEqual(['b'])
    expect(taskIn(execution.document, 'b').baseline).toEqual([asBaselineId('b1')])
  })
})

describe('PROJECT-009 command determinism', () => {
  it('repeated CreateBaseline sequences produce byte-identical documents', () => {
    const build = (start: ProjectDocument): ProjectDocument => {
      const withTask = applyProjectCommand(start, {
        type: 'CreateTask',
        task: makeTask({ id: 'a', uid: 1 }),
      }).document
      return applyProjectCommand(withTask, {
        type: 'CreateBaseline',
        baseline: makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 } }),
      }).document
    }
    const empty = parseDocument(snapshot(canonicalizeDocument(makeDocument())))
    const first = build(empty)
    const second = build(empty)
    expect(snapshot(second)).toBe(snapshot(first))
    expect(taskIn(first, 'a').baseline).toEqual([asBaselineId('b1')])
    expect(first.baselines[0].id).toEqual(asBaselineId('b1'))
  })

  it('accepted CreateBaseline never mutates the input document', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const before = snapshot(base)
    applyProjectCommand(base, {
      type: 'CreateBaseline',
      baseline: makeBaseline('b1', CAPTURED_AT, { a: { duration: 480 } }),
    })
    expect(snapshot(base)).toBe(before)
  })
})
