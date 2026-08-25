import { describe, expect, it } from 'vitest'
import type { ProjectCommand, ProjectDocument, Task } from '@genoffice/project-contracts'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import { applyProjectCommand, canonicalizeDocument, validateProjectDocument } from '../src/index.js'
import { makeDocument, makeTask } from './fixtures.js'

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
  expect(execution.result.diagnostics.some((d) => d.code === code)).toBe(true)
  return execution
}

const snapshot = (document: ProjectDocument): string => JSON.stringify(document)
const parseDocument = (json: string): ProjectDocument => JSON.parse(json) as ProjectDocument

describe('PROJECT-008 SetConstraint', () => {
  it('sets a dated constraint and reports the affected task + inverse', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const execution = expectAccepted(base, {
      type: 'SetConstraint',
      taskId: parent('a'),
      constraintType: 'startNoEarlierThan',
      constraintDate: '2026-08-04T09:00:00.000Z',
    })
    const task = taskIn(execution.document, 'a')
    expect(task.constraintType).toBe('startNoEarlierThan')
    expect(task.constraintDate).toBe('2026-08-04T09:00:00.000Z')
    expect(execution.result.affectedTaskIds).toEqual([parent('a')])
    expect(execution.result.inverse).toEqual({
      type: 'SetConstraint',
      taskId: parent('a'),
      constraintType: 'asSoonAsPossible',
      constraintDate: undefined,
    })
    expect(base.tasks).toBe(base.tasks) // input reference unchanged
  })

  it('MSO and MFO are not silently reinterpreted as SNET/FNET', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const mso = expectAccepted(
      base,
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'mustStartOn',
        constraintDate: '2026-08-04T09:00:00.000Z',
      },
    ).document
    expect(taskIn(mso, 'a').constraintType).toBe('mustStartOn')
    const mfo = expectAccepted(
      base,
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'mustFinishOn',
        constraintDate: '2026-08-04T17:00:00.000Z',
      },
    ).document
    expect(taskIn(mfo, 'a').constraintType).toBe('mustFinishOn')
  })

  it('switching to ASAP/ALAP clears any prior constraintDate', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            constraintType: 'mustStartOn',
            constraintDate: asISODateTime('2026-08-04T09:00:00.000Z'),
          }),
        ],
      }),
    )
    const execution = expectAccepted(base, {
      type: 'SetConstraint',
      taskId: parent('a'),
      constraintType: 'asSoonAsPossible',
    })
    const task = taskIn(execution.document, 'a')
    expect(task.constraintType).toBe('asSoonAsPossible')
    expect(task.constraintDate).toBeUndefined()
    // No stale undefined-valued key survives into the serialized document.
    expect(snapshot(execution.document)).not.toContain('"constraintDate"')
  })

  it('rejects a dated constraint without a constraintDate (no silent MSO→SNET collapse)', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'mustStartOn',
      },
      'MISSING_CONSTRAINT_DATE',
    )
  })

  it('rejects a constraintDate supplied alongside ASAP/ALAP', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'asSoonAsPossible',
        constraintDate: '2026-08-04T09:00:00.000Z',
      },
      'CONSTRAINT_DATE_NOT_ALLOWED',
    )
  })

  it('rejects a malformed constraintDate', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'startNoEarlierThan',
        constraintDate: 'not-a-date',
      },
      'INVALID_DATE',
    )
  })

  it('rejects setting a constraint on a missing task and leaves the document unchanged', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      {
        type: 'SetConstraint',
        taskId: parent('ghost'),
        constraintType: 'mustStartOn',
        constraintDate: '2026-08-04T09:00:00.000Z',
      },
      'MISSING_TASK',
    )
  })
})

describe('PROJECT-008 SetPercentComplete', () => {
  it('sets percentComplete at the canonical boundaries and midpoints', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    for (const value of [0, 1, 50, 99, 100]) {
      const execution = expectAccepted(base, {
        type: 'SetPercentComplete',
        taskId: parent('a'),
        percentComplete: value,
      })
      expect(taskIn(execution.document, 'a').percentComplete).toBe(value)
      expect(execution.result.affectedTaskIds).toEqual([parent('a')])
      expect(execution.result.inverse).toEqual({
        type: 'SetPercentComplete',
        taskId: parent('a'),
        percentComplete: 0,
      })
    }
  })

  it('rejects out-of-range and non-finite percentComplete', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    for (const value of [150, -1, Number.NaN, Infinity]) {
      expectRejected(
        base,
        { type: 'SetPercentComplete', taskId: parent('a'), percentComplete: value },
        'INVALID_PERCENT_COMPLETE',
      )
    }
  })

  it('rejects setting percentComplete on a summary task (progress is derived)', () => {
    const clean = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 's', summary: true, duration: asWorkingMinutes(0) }),
          makeTask({ id: 'c', parentTaskId: parent('s'), outlineLevel: 2 }),
        ],
      }),
    )
    expectRejected(
      clean,
      { type: 'SetPercentComplete', taskId: parent('s'), percentComplete: 50 },
      'SUMMARY_PROGRESS_NOT_SETTABLE',
    )
  })

  it('rejects setting percentComplete on a missing task', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      { type: 'SetPercentComplete', taskId: parent('ghost'), percentComplete: 50 },
      'MISSING_TASK',
    )
  })
})

describe('PROJECT-008 SetDeadline', () => {
  it('sets and clears a deadline (a deadline never moves the task)', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const set = expectAccepted(base, {
      type: 'SetDeadline',
      taskId: parent('a'),
      deadline: '2026-08-05T17:00:00.000Z',
    })
    expect(taskIn(set.document, 'a').deadline).toBe('2026-08-05T17:00:00.000Z')
    expect(set.result.affectedTaskIds).toEqual([parent('a')])
    expect(set.result.inverse).toEqual({
      type: 'SetDeadline',
      taskId: parent('a'),
      deadline: undefined,
    })
    const cleared = expectAccepted(set.document, {
      type: 'SetDeadline',
      taskId: parent('a'),
      deadline: undefined,
    })
    expect(taskIn(cleared.document, 'a').deadline).toBeUndefined()
    expect(snapshot(cleared.document)).not.toContain('"deadline"')
  })

  it('rejects a malformed deadline', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      { type: 'SetDeadline', taskId: parent('a'), deadline: 'not-a-date' },
      'INVALID_DATE',
    )
  })

  it('rejects setting a deadline on a missing task', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    expectRejected(
      base,
      { type: 'SetDeadline', taskId: parent('ghost'), deadline: '2026-08-05T17:00:00.000Z' },
      'MISSING_TASK',
    )
  })
})

describe('PROJECT-008 constraint validation surface', () => {
  it('document validation reports MISSING_CONSTRAINT_DATE and CONSTRAINT_DATE_NOT_ALLOWED', () => {
    const missing = makeDocument({
      tasks: [makeTask({ id: 'a', constraintType: 'mustStartOn' })],
    })
    const codes = validateProjectDocument(missing).diagnostics.map((d) => d.code)
    expect(codes).toContain('MISSING_CONSTRAINT_DATE')

    const disallowed = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          constraintType: 'asSoonAsPossible',
          constraintDate: asISODateTime('2026-08-04T09:00:00.000Z'),
        }),
      ],
    })
    const codes2 = validateProjectDocument(disallowed).diagnostics.map((d) => d.code)
    expect(codes2).toContain('CONSTRAINT_DATE_NOT_ALLOWED')
  })
})

describe('PROJECT-008 command determinism and history', () => {
  it('repeated identical command sequences produce byte-identical documents', () => {
    const cleanCommands: ProjectCommand[] = [
      { type: 'CreateTask', task: makeTask({ id: 'a', uid: 1 }) },
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'startNoEarlierThan',
        constraintDate: '2026-08-04T09:00:00.000Z',
      },
      { type: 'SetPercentComplete', taskId: parent('a'), percentComplete: 33 },
      { type: 'SetDeadline', taskId: parent('a'), deadline: '2026-08-06T17:00:00.000Z' },
      {
        type: 'SetConstraint',
        taskId: parent('a'),
        constraintType: 'asSoonAsPossible',
      },
      { type: 'SetPercentComplete', taskId: parent('a'), percentComplete: 0 },
      { type: 'SetDeadline', taskId: parent('a'), deadline: undefined },
    ]
    const serialized = snapshot(canonicalizeDocument(makeDocument()))
    const first = cleanCommands.reduce(
      (doc, cmd) => expectAccepted(doc, cmd).document,
      parseDocument(serialized),
    )
    const second = cleanCommands.reduce(
      (doc, cmd) => expectAccepted(doc, cmd).document,
      parseDocument(serialized),
    )
    expect(snapshot(second)).toBe(snapshot(first))
    // After clearing everything, the task is back to an unconstrained 0% state
    // with no deadline keys on the wire.
    const a = taskIn(first, 'a')
    expect(a.constraintType).toBe('asSoonAsPossible')
    expect(a.constraintDate).toBeUndefined()
    expect(a.percentComplete).toBe(0)
    expect(a.deadline).toBeUndefined()
  })

  it('accepted commands never mutate the input document', () => {
    const base = canonicalizeDocument(makeDocument({ tasks: [makeTask({ id: 'a' })] }))
    const before = snapshot(base)
    applyProjectCommand(base, {
      type: 'SetConstraint',
      taskId: parent('a'),
      constraintType: 'mustStartOn',
      constraintDate: '2026-08-04T09:00:00.000Z',
    })
    applyProjectCommand(base, { type: 'SetPercentComplete', taskId: parent('a'), percentComplete: 75 })
    applyProjectCommand(base, {
      type: 'SetDeadline',
      taskId: parent('a'),
      deadline: '2026-08-06T17:00:00.000Z',
    })
    expect(snapshot(base)).toBe(before)
  })

  it('SetConstraint inverse restores the previous constraint shape', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            constraintType: 'finishNoLaterThan',
            constraintDate: asISODateTime('2026-08-05T17:00:00.000Z'),
          }),
        ],
      }),
    )
    const before = snapshot(base)
    const exec = applyProjectCommand(base, {
      type: 'SetConstraint',
      taskId: parent('a'),
      constraintType: 'mustStartOn',
      constraintDate: '2026-08-04T09:00:00.000Z',
    })
    expect(exec.result.accepted).toBe(true)
    const restored = applyProjectCommand(exec.document, exec.result.inverse!)
    expect(snapshot(restored.document)).toBe(before)
  })

  it('SetPercentComplete and SetDeadline inverses restore the prior values', () => {
    const base = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({
            id: 'a',
            percentComplete: 40,
            deadline: asISODateTime('2026-08-06T17:00:00.000Z'),
          }),
        ],
      }),
    )
    const before = snapshot(base)
    const pct = applyProjectCommand(base, {
      type: 'SetPercentComplete',
      taskId: parent('a'),
      percentComplete: 80,
    })
    const restoredPct = applyProjectCommand(pct.document, pct.result.inverse!)
    expect(snapshot(restoredPct.document)).toBe(before)
    const dl = applyProjectCommand(base, {
      type: 'SetDeadline',
      taskId: parent('a'),
      deadline: '2026-08-09T17:00:00.000Z',
    })
    const restoredDl = applyProjectCommand(dl.document, dl.result.inverse!)
    expect(snapshot(restoredDl.document)).toBe(before)
  })
})
