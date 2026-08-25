import { describe, expect, it } from 'vitest'
import type { ProjectCommand, ProjectDocument } from '@genoffice/project-contracts'
import { asAssignmentId, asWorkingMinutes } from '@genoffice/project-contracts'
import { applyProjectCommand, validateProjectDocument } from '../src/index.js'
import { canonicalizeDocument } from '../src/index.js'
import { makeAssignment, makeDocument, makeResource, makeTask } from './fixtures.js'

const expectValid = (document: ProjectDocument) => {
  const result = validateProjectDocument(document)
  if (!result.accepted) {
    throw new Error(`document invalid: ${JSON.stringify(result.diagnostics)}`)
  }
  return result
}

const expectRejectedWith = (document: ProjectDocument, code: string) => {
  const result = validateProjectDocument(document)
  expect(result.accepted).toBe(false)
  expect(result.diagnostics.some((d) => d.code === code)).toBe(true)
  return result
}

const expectAcceptedCmd = (document: ProjectDocument, command: ProjectCommand) => {
  const execution = applyProjectCommand(document, command)
  if (!execution.result.accepted) {
    throw new Error(`command rejected: ${JSON.stringify(execution.result.diagnostics)}`)
  }
  return execution
}

const expectRejectedCmd = (document: ProjectDocument, command: ProjectCommand, code: string) => {
  const execution = applyProjectCommand(document, command)
  expect(execution.result.accepted).toBe(false)
  expect(execution.document).toBe(document)
  expect(execution.result.diagnostics.some((d) => d.code === code)).toBe(true)
  return execution
}

// ---------------------------------------------------------------------------
// PROJECT-011 SetAssignmentUnits command
// ---------------------------------------------------------------------------

describe('PROJECT-011 SetAssignmentUnits command', () => {
  it('1. accepts a valid units change and preserves the assignment identity', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1', kind: 'work' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const execution = expectAcceptedCmd(document, {
      type: 'SetAssignmentUnits',
      assignmentId: asAssignmentId('as1'),
      units: 0.5,
    })
    const updated = execution.document.assignments.find((a) => a.id === asAssignmentId('as1'))!
    expect(updated.units).toBe(0.5)
    // Identity preserved.
    expect(updated.id).toEqual(asAssignmentId('as1'))
    expect(updated.taskId).toEqual(document.tasks[0].id)
    expect(updated.resourceId).toEqual(document.resources[0].id)
  })

  it('2. rejects a missing assignment with MISSING_ASSIGNMENT (document unchanged)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    expectRejectedCmd(
      document,
      {
        type: 'SetAssignmentUnits',
        assignmentId: asAssignmentId('ghost'),
        units: 0.5,
      },
      'MISSING_ASSIGNMENT',
    )
  })

  it('3. rejects a non-finite units value with INVALID_ASSIGNMENT_UNITS', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    expectRejectedCmd(
      document,
      {
        type: 'SetAssignmentUnits',
        assignmentId: asAssignmentId('as1'),
        units: Number.NaN,
      },
      'INVALID_ASSIGNMENT_UNITS',
    )
  })

  it('4. rejects a negative units value with INVALID_ASSIGNMENT_UNITS', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    expectRejectedCmd(
      document,
      {
        type: 'SetAssignmentUnits',
        assignmentId: asAssignmentId('as1'),
        units: -0.5,
      },
      'INVALID_ASSIGNMENT_UNITS',
    )
  })

  it('5. accepts a zero-unit assignment (the model permits units = 0)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const execution = expectAcceptedCmd(document, {
      type: 'SetAssignmentUnits',
      assignmentId: asAssignmentId('as1'),
      units: 0,
    })
    const updated = execution.document.assignments.find((a) => a.id === asAssignmentId('as1'))!
    expect(updated.units).toBe(0)
  })

  it('6. produces a correct inverse that restores the previous units value', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const execution = applyProjectCommand(document, {
      type: 'SetAssignmentUnits',
      assignmentId: asAssignmentId('as1'),
      units: 0.5,
    })
    expect(execution.result.accepted).toBe(true)
    expect(execution.result.inverse).toEqual({
      type: 'SetAssignmentUnits',
      assignmentId: asAssignmentId('as1'),
      units: 1,
    })
    // Applying the inverse restores the original units.
    const restored = applyProjectCommand(execution.document, execution.result.inverse!)
    expect(restored.result.accepted).toBe(true)
    const restoredAssignment = restored.document.assignments.find(
      (a) => a.id === asAssignmentId('as1'),
    )!
    expect(restoredAssignment.units).toBe(1)
  })

  it('7. leaves the document valid after a units change (post-mutation validation)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    const execution = expectAcceptedCmd(document, {
      type: 'SetAssignmentUnits',
      assignmentId: asAssignmentId('as1'),
      units: 2,
    })
    expectValid(execution.document)
  })

  it('8. accepts a 200% assignment (units = 2.0) without leveling', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1', maxUnits: 1 })],
      assignments: [makeAssignment('as1', 'a', 'r1', { units: 1 })],
    })
    // 200% is over-allocated (units > maxUnits) but the engine accepts it;
    // leveling is PROJECT-013 territory.
    const execution = expectAcceptedCmd(document, {
      type: 'SetAssignmentUnits',
      assignmentId: asAssignmentId('as1'),
      units: 2,
    })
    expect(execution.document.assignments[0].units).toBe(2)
  })
})

// ---------------------------------------------------------------------------
// PROJECT-011 work/cost validation
// ---------------------------------------------------------------------------

describe('PROJECT-011 work/cost validation', () => {
  it('9. rejects a task with negative work (INVALID_WORK)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', work: asWorkingMinutes(-100) })],
    })
    expectRejectedWith(document, 'INVALID_WORK')
  })

  it('10. rejects a task with negative actualWork (INVALID_ACTUAL_WORK)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', actualWork: asWorkingMinutes(-50) })],
    })
    expectRejectedWith(document, 'INVALID_ACTUAL_WORK')
  })

  it('11. rejects a task with negative cost (INVALID_COST)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', cost: -100 })],
    })
    expectRejectedWith(document, 'INVALID_COST')
  })

  it('12. rejects a task with negative actualCost (INVALID_ACTUAL_COST)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a', actualCost: -50 })],
    })
    expectRejectedWith(document, 'INVALID_ACTUAL_COST')
  })

  it('13. rejects a task where actualWork exceeds work (INCONSISTENT_WORK)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          work: asWorkingMinutes(100),
          actualWork: asWorkingMinutes(200),
          remainingWork: asWorkingMinutes(0),
        }),
      ],
    })
    expectRejectedWith(document, 'INCONSISTENT_WORK')
  })

  it('14. rejects a task where actualCost exceeds cost (INCONSISTENT_COST)', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          cost: 100,
          actualCost: 200,
          remainingCost: 0,
        }),
      ],
    })
    expectRejectedWith(document, 'INCONSISTENT_COST')
  })

  it('15. rejects an assignment with negative work (INVALID_WORK)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { work: asWorkingMinutes(-50) })],
    })
    expectRejectedWith(document, 'INVALID_WORK')
  })

  it('16. rejects an assignment with negative cost (INVALID_COST)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1', { cost: -100 })],
    })
    expectRejectedWith(document, 'INVALID_COST')
  })

  it('17. rejects an assignment where actualWork exceeds work (INCONSISTENT_WORK)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [
        makeAssignment('as1', 'a', 'r1', {
          work: asWorkingMinutes(100),
          actualWork: asWorkingMinutes(150),
        }),
      ],
    })
    expectRejectedWith(document, 'INCONSISTENT_WORK')
  })

  it('18. accepts a valid document with zero work/cost (the fixture default)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' })],
      assignments: [makeAssignment('as1', 'a', 'r1')],
    })
    expectValid(document)
  })

  it('19. accepts a cost-resource assignment where cost is authoritative and actuals are zero', () => {
    // A cost resource sets cost=500 but leaves actualCost/remainingCost at 0
    // (the scheduler will derive them). The "exceeds" check passes (0 <= 500).
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'c1', kind: 'cost' })],
      assignments: [makeAssignment('as1', 'a', 'c1', { cost: 500 })],
    })
    expectValid(document)
  })

  it('20. canonicalizeDocument preserves work/cost fields through canonicalization', () => {
    const document = makeDocument({
      tasks: [
        makeTask({
          id: 'a',
          work: asWorkingMinutes(480),
          cost: 400,
          actualWork: asWorkingMinutes(240),
          remainingWork: asWorkingMinutes(240),
          actualCost: 200,
          remainingCost: 200,
        }),
      ],
    })
    const canonical = canonicalizeDocument(document)
    const task = canonical.tasks[0]
    expect(task.work).toEqual(asWorkingMinutes(480))
    expect(task.cost).toBe(400)
    expect(task.actualWork).toEqual(asWorkingMinutes(240))
    expect(task.remainingWork).toEqual(asWorkingMinutes(240))
    expect(task.actualCost).toBe(200)
    expect(task.remainingCost).toBe(200)
  })
})
