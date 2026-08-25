import { describe, expect, it } from 'vitest'
import type { ProjectCommand, ProjectDocument, Resource } from '@genoffice/project-contracts'
import {
  asAssignmentId,
  asCalendarId,
  asISODateTime,
  asResourceId,
  asTaskId,
} from '@genoffice/project-contracts'
import { applyProjectCommand, canonicalizeDocument, validateProjectDocument } from '../src/index.js'
import { makeAssignment, makeCalendar, makeDocument, makeResource, makeTask } from './fixtures.js'

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

const resourceIn = (document: ProjectDocument, id: string): Resource => {
  const resource = document.resources.find((r) => r.id === asResourceId(id))
  if (!resource) throw new Error(`missing resource ${id}`)
  return resource
}

// ---------------------------------------------------------------------------
// Required tests 1-3: Work / Material / Cost resource creation + validation
// ---------------------------------------------------------------------------

describe('PROJECT-010 resource type creation and validation', () => {
  it('1. accepts a valid work resource with max units, rates, and cost-per-use', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        makeResource({
          id: 'r1',
          name: 'Engineer',
          kind: 'work',
          maxUnits: 1,
          standardRate: 50,
          overtimeRate: 75,
          costPerUse: 0,
        }),
      ],
    })
    expectValid(document)
    expect(resourceIn(document, 'r1').kind).toBe('work')
  })

  it('2. accepts a valid material resource with a standard rate and cost-per-use', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        makeResource({
          id: 'm1',
          name: 'Concrete',
          kind: 'material',
          maxUnits: 0,
          standardRate: 100,
          overtimeRate: 0,
          costPerUse: 0,
        }),
      ],
    })
    expectValid(document)
    expect(resourceIn(document, 'm1').kind).toBe('material')
  })

  it('3. accepts a valid cost resource (pure cost category, no work capacity)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        makeResource({
          id: 'c1',
          name: 'Travel',
          kind: 'cost',
          maxUnits: 0,
          standardRate: 0,
          overtimeRate: 0,
          costPerUse: 500,
        }),
      ],
    })
    expectValid(document)
    expect(resourceIn(document, 'c1').kind).toBe('cost')
  })
})

// ---------------------------------------------------------------------------
// Required tests 4-5: stable ResourceId, duplicate ResourceId rejection
// ---------------------------------------------------------------------------

describe('PROJECT-010 resource identity', () => {
  it('4. preserves stable ResourceId across the brand boundary', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' }), makeResource({ id: 'r2' })],
    })
    expectValid(document)
    const ids = document.resources.map((r) => r.id)
    expect(ids).toEqual([asResourceId('r1'), asResourceId('r2')])
    // Identity is the branded id, never array position.
    expect(resourceIn(document, 'r2').id).toEqual(asResourceId('r2'))
  })

  it('5. rejects a duplicate ResourceId', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1' }), makeResource({ id: 'r1' })],
    })
    expectRejectedWith(document, 'DUPLICATE_RESOURCE_ID')
  })
})

// ---------------------------------------------------------------------------
// Required tests 6-11: max units, standard rate, overtime rate, cost-per-use
// ---------------------------------------------------------------------------

describe('PROJECT-010 resource numeric scheduling inputs', () => {
  // `baseResource` takes the resource id as the first argument (plain string,
  // branded at the makeResource boundary) and optional numeric overrides, so
  // the brand boundary stays the single canonical promotion point.
  const baseResource = (id: string, overrides: Partial<Omit<Resource, 'id'>> = {}): Resource =>
    makeResource({ ...overrides, id })

  it('6. accepts valid max units (finite, non-negative)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        baseResource('r1', { maxUnits: 0 }),
        baseResource('r2', { maxUnits: 1 }),
        baseResource('r3', { maxUnits: 2.5 }),
      ],
    })
    expectValid(document)
  })

  it('7. rejects invalid max units (negative or non-finite)', () => {
    const docNegative = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [baseResource('r1', { maxUnits: -1 })],
    })
    expectRejectedWith(docNegative, 'INVALID_MAX_UNITS')
    const docNaN = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [baseResource('r1', { maxUnits: Number.NaN })],
    })
    expectRejectedWith(docNaN, 'INVALID_MAX_UNITS')
  })

  it('8. accepts valid standard rate (finite, non-negative)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        baseResource('r1', { standardRate: 0 }),
        baseResource('r2', { standardRate: 42.5 }),
      ],
    })
    expectValid(document)
  })

  it('9. rejects an invalid standard rate (negative)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [baseResource('r1', { standardRate: -10 })],
    })
    expectRejectedWith(document, 'INVALID_RATE')
  })

  it('10. accepts a valid overtime rate (finite, non-negative)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [baseResource('r1', { overtimeRate: 75 })],
    })
    expectValid(document)
  })

  it('11. accepts a valid cost-per-use (finite, non-negative)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [baseResource('r1', { costPerUse: 500 })],
    })
    expectValid(document)
    // Negative cost-per-use is rejected (no silent coercion).
    const bad = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [baseResource('r1', { costPerUse: -1 })],
    })
    expectRejectedWith(bad, 'INVALID_COST_PER_USE')
  })
})

// ---------------------------------------------------------------------------
// Required tests 12-13: resource calendar reference, missing calendar rejection
// ---------------------------------------------------------------------------

describe('PROJECT-010 resource calendar references', () => {
  it('12. accepts a resource that references a defined calendar', () => {
    const resourceCalendar = makeCalendar('rescal')
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      calendars: [makeCalendar('standard'), resourceCalendar],
      resources: [makeResource({ id: 'r1', calendarId: resourceCalendar.id })],
    })
    expectValid(document)
  })

  it('13. rejects a resource that references a missing calendar', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1', calendarId: asCalendarId('ghost') })],
    })
    expectRejectedWith(document, 'MISSING_CALENDAR')
  })
})

// ---------------------------------------------------------------------------
// Required test 15: resource calendar cycle rejection (engine validation)
// ---------------------------------------------------------------------------

describe('PROJECT-010 resource calendar inheritance cycles', () => {
  it('15. rejects a resource calendar whose inheritance forms a cycle', () => {
    // Two calendars that inherit from each other form a cycle. A resource
    // referencing one of them makes the cycle scheduling-relevant; validation
    // rejects with CALENDAR_CYCLE regardless of the resource reference.
    const cycleA = makeCalendar('cycleA', { baseCalendarId: asCalendarId('cycleB') })
    const cycleB = makeCalendar('cycleB', { baseCalendarId: asCalendarId('cycleA') })
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      calendars: [makeCalendar('standard'), cycleA, cycleB],
      resources: [makeResource({ id: 'r1', calendarId: cycleA.id })],
    })
    expectRejectedWith(document, 'CALENDAR_CYCLE')
  })
})

// ---------------------------------------------------------------------------
// Required test 16: availability window validation
// ---------------------------------------------------------------------------

describe('PROJECT-010 availability windows', () => {
  it('16. validates availability windows: rejects zero-length/inverted ranges and bad units', () => {
    const start = asISODateTime('2026-08-03T09:00:00.000Z')
    const finish = asISODateTime('2026-09-03T09:00:00.000Z')
    // Valid window.
    const valid = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1', availability: [{ start, finish, units: 1 }] })],
    })
    expectValid(valid)
    // Zero-length (finish == start) is rejected.
    const zeroLength = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1', availability: [{ start, finish: start, units: 1 }] })],
    })
    expectRejectedWith(zeroLength, 'INVALID_AVAILABILITY_RANGE')
    // Inverted (finish before start) is rejected.
    const inverted = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        makeResource({ id: 'r1', availability: [{ start: finish, finish: start, units: 1 }] }),
      ],
    })
    expectRejectedWith(inverted, 'INVALID_AVAILABILITY_RANGE')
    // Negative units rejected.
    const badUnits = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [makeResource({ id: 'r1', availability: [{ start, finish, units: -1 }] })],
    })
    expectRejectedWith(badUnits, 'INVALID_AVAILABILITY_UNITS')
    // Malformed date rejected.
    const badDate = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [MakeResourceBadDate()],
    })
    expectRejectedWith(badDate, 'INVALID_DATE')
  })

  it('accepts an open-ended availability window (start with no finish)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'a' })],
      resources: [
        makeResource({
          id: 'r1',
          availability: [{ start: asISODateTime('2026-08-03T09:00:00.000Z'), units: 1 }],
        }),
      ],
    })
    expectValid(document)
  })
})

// Helper to build a resource with a malformed availability start date without
// tripping the typed fixture helper. Kept local so the brand boundary stays
// the single canonical promotion point.
function MakeResourceBadDate(): Resource {
  return {
    ...makeResource({ id: 'r1' }),
    availability: [{ start: asISODateTime('not-a-date'), units: 1 }],
  }
}

// ---------------------------------------------------------------------------
// Required tests 18-23: AssignResource / UnassignResource commands
// ---------------------------------------------------------------------------

describe('PROJECT-010 AssignResource command', () => {
  const baseDocument = (): ProjectDocument =>
    canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        resources: [makeResource({ id: 'r1' })],
      }),
    )

  it('18. creates an assignment linking a task and a resource', () => {
    const document = baseDocument()
    const assignment = makeAssignment('as1', 'a', 'r1', { units: 1 })
    const execution = expectAcceptedCmd(document, { type: 'AssignResource', assignment })
    expect(execution.document.assignments).toHaveLength(1)
    expect(execution.document.assignments[0]).toEqual(assignment)
    expect(execution.result.affectedTaskIds).toEqual([asTaskId('a')])
    // Inverse is UnassignResource for the new assignment id.
    expect(execution.result.inverse).toEqual({
      type: 'UnassignResource',
      assignmentId: assignment.id,
    })
    // Input document is never mutated.
    expect(document.assignments).toEqual([])
  })

  it('19. rejects a duplicate AssignmentId', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        resources: [makeResource({ id: 'r1' })],
        assignments: [makeAssignment('as1', 'a', 'r1')],
      }),
    )
    expectRejectedCmd(
      document,
      { type: 'AssignResource', assignment: makeAssignment('as1', 'a', 'r1') },
      'DUPLICATE_ASSIGNMENT_ID',
    )
  })

  it('20. rejects an assignment referencing a missing task', () => {
    const document = baseDocument()
    expectRejectedCmd(
      document,
      { type: 'AssignResource', assignment: makeAssignment('as1', 'ghost', 'r1') },
      'MISSING_TASK_REFERENCE',
    )
  })

  it('21. rejects an assignment referencing a missing resource', () => {
    const document = baseDocument()
    expectRejectedCmd(
      document,
      { type: 'AssignResource', assignment: makeAssignment('as1', 'a', 'ghost') },
      'MISSING_RESOURCE_REFERENCE',
    )
  })

  it('rejects a duplicate task/resource pair (two rows shadowing the same link)', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        resources: [makeResource({ id: 'r1' })],
        assignments: [makeAssignment('as1', 'a', 'r1')],
      }),
    )
    expectRejectedCmd(
      document,
      { type: 'AssignResource', assignment: makeAssignment('as2', 'a', 'r1') },
      'DUPLICATE_ASSIGNMENT_PAIR',
    )
  })

  it('rejects an assignment with invalid units (negative)', () => {
    const document = baseDocument()
    const execution = applyProjectCommand(document, {
      type: 'AssignResource',
      assignment: makeAssignment('as1', 'a', 'r1', { units: -0.5 }),
    })
    // The mutator accepts (units are validated post-mutation); the
    // post-mutation validator rejects with INVALID_ASSIGNMENT_UNITS.
    expect(execution.result.accepted).toBe(false)
    expect(execution.result.diagnostics.some((d) => d.code === 'INVALID_ASSIGNMENT_UNITS')).toBe(
      true,
    )
  })
})

describe('PROJECT-010 UnassignResource command', () => {
  it('removes an assignment by id and provides an inverse', () => {
    const assignment = makeAssignment('as1', 'a', 'r1')
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        resources: [makeResource({ id: 'r1' })],
        assignments: [assignment],
      }),
    )
    const execution = expectAcceptedCmd(document, {
      type: 'UnassignResource',
      assignmentId: assignment.id,
    })
    expect(execution.document.assignments).toEqual([])
    expect(execution.result.affectedTaskIds).toEqual([asTaskId('a')])
    expect(execution.result.inverse).toEqual({
      type: 'AssignResource',
      assignment,
    })
    expect(document.assignments).toHaveLength(1)
  })

  it('rejects unassigning a missing assignment id', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' })],
        resources: [makeResource({ id: 'r1' })],
      }),
    )
    expectRejectedCmd(
      document,
      { type: 'UnassignResource', assignmentId: asResourceId('ghost') as never },
      'MISSING_ASSIGNMENT',
    )
  })
})

// ---------------------------------------------------------------------------
// Required tests 22-23: assignment validity after hierarchy mutation/deletion
// ---------------------------------------------------------------------------

describe('PROJECT-010 assignment validity after task mutations', () => {
  it('22. assignments survive hierarchy mutations that preserve TaskId', () => {
    // IndentTask/OutdentTask/RenameTask never change TaskId, so assignments
    // survive those mutations unchanged.
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [
          makeTask({ id: 'parent', summary: true, outlineLevel: 1 }),
          makeTask({ id: 'child', parentTaskId: asTaskId('parent'), outlineLevel: 2 }),
        ],
        resources: [makeResource({ id: 'r1' })],
        assignments: [makeAssignment('as1', 'child', 'r1')],
      }),
    )
    // Rename the parent task — assignments referencing the child survive.
    const renamed = applyProjectCommand(document, {
      type: 'RenameTask',
      taskId: asTaskId('parent'),
      name: 'Phase 1',
    })
    expect(renamed.result.accepted).toBe(true)
    expect(renamed.document.assignments).toHaveLength(1)
    expect(renamed.document.assignments[0].taskId).toEqual(asTaskId('child'))
    // Document stays valid.
    expectValid(renamed.document)
  })

  it('23. deleting a task removes its assignments (no dangling references)', () => {
    const document = canonicalizeDocument(
      makeDocument({
        tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
        resources: [makeResource({ id: 'r1' })],
        assignments: [makeAssignment('as1', 'a', 'r1'), makeAssignment('as2', 'b', 'r1')],
      }),
    )
    const execution = applyProjectCommand(document, { type: 'DeleteTask', taskId: asTaskId('a') })
    expect(execution.result.accepted).toBe(true)
    // Assignment as1 (task a) is removed; as2 (task b) survives.
    expect(execution.document.assignments).toHaveLength(1)
    expect(execution.document.assignments[0].id).toEqual(asAssignmentId('as2'))
    // The surviving assignment still references a real task.
    expectValid(execution.document)
  })
})
