import { describe, expect, it } from 'vitest'
import { validateProjectDocument } from '../src/index.js'
import {
  asAssignmentId,
  asBaselineId,
  asCalendarId,
  asCustomFieldId,
  asISODateTime,
  asResourceId,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type {
  Assignment,
  Baseline,
  CustomField,
  Dependency,
  Resource,
  Task,
} from '@genoffice/project-contracts'
import { makeCalendar, makeDependency, makeDocument, makeTask, standardWeek } from './fixtures.js'

const codes = (result: ReturnType<typeof validateProjectDocument>): string[] =>
  result.diagnostics.map((diagnostic) => diagnostic.code)

const validate = (...parts: Parameters<typeof makeDocument>) =>
  validateProjectDocument(makeDocument(...parts))

describe('project document validation', () => {
  it('accepts a well-formed document', () => {
    const result = validate({
      tasks: [
        makeTask({ id: 's', summary: true, duration: asWorkingMinutes(0) }),
        makeTask({ id: 'c', parentTaskId: asTaskId('s'), outlineLevel: 2 }),
        makeTask({ id: 'x' }),
      ],
      dependencies: [makeDependency('d1', 'x', 'c', 'FS', 60)],
      resources: [
        {
          id: asResourceId('r1'),
          uid: 1,
          name: 'Worker',
          kind: 'work',
          maxUnits: 1,
          standardRate: 0,
          overtimeRate: 0,
          costPerUse: 0,
          availability: [{ start: asISODateTime('2026-08-03T09:00:00.000Z'), units: 1 }],
        },
      ],
      assignments: [
        {
          id: asAssignmentId('as1'),
          taskId: asTaskId('c'),
          resourceId: asResourceId('r1'),
          units: 1,
          work: asWorkingMinutes(480),
          actualWork: asWorkingMinutes(0),
          remainingWork: asWorkingMinutes(480),
          cost: 0,
          actualCost: 0,
          remainingCost: 0,
        },
      ],
    })
    expect(result.accepted).toBe(true)
    expect(result.diagnostics).toEqual([])
  })

  it('rejects duplicate ids across every entity collection', () => {
    expect(codes(validate({ tasks: [makeTask({ id: 'a' }), makeTask({ id: 'a' })] }))).toContain(
      'DUPLICATE_TASK_ID',
    )
    const resource = (id: string): Resource => ({
      id: asResourceId(id),
      uid: 1,
      name: id,
      kind: 'work',
      maxUnits: 1,
      standardRate: 0,
      overtimeRate: 0,
      costPerUse: 0,
      availability: [],
    })
    expect(
      codes(
        validate({ resources: [resource('r1'), resource('r1')], tasks: [makeTask({ id: 'a' })] }),
      ),
    ).toContain('DUPLICATE_RESOURCE_ID')
    const assignment = (id: string): Assignment => ({
      id: asAssignmentId(id),
      taskId: asTaskId('a'),
      resourceId: asResourceId('r1'),
      units: 1,
      work: asWorkingMinutes(0),
      actualWork: asWorkingMinutes(0),
      remainingWork: asWorkingMinutes(0),
      cost: 0,
      actualCost: 0,
      remainingCost: 0,
    })
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' })],
          resources: [resource('r1')],
          assignments: [assignment('x1'), assignment('x1')],
        }),
      ),
    ).toContain('DUPLICATE_ASSIGNMENT_ID')
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
          dependencies: [
            makeDependency('d1', 'a', 'b', 'FS'),
            makeDependency('d1', 'a', 'b', 'SS'),
          ],
        }),
      ),
    ).toContain('DUPLICATE_DEPENDENCY_ID')
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' })],
          calendars: [makeCalendar('standard'), makeCalendar('standard')],
        }),
      ),
    ).toContain('DUPLICATE_CALENDAR_ID')
    const baseline = (id: string): Baseline => ({
      id: asBaselineId(id),
      name: id,
      capturedAt: asISODateTime('2026-08-03T09:00:00.000Z'),
      taskSnapshots: {},
    })
    expect(
      codes(
        validate({ tasks: [makeTask({ id: 'a' })], baselines: [baseline('b1'), baseline('b1')] }),
      ),
    ).toContain('DUPLICATE_BASELINE_ID')
    const field = (id: string): CustomField => ({ id: asCustomFieldId(id), name: id, type: 'text' })
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a' })], customFields: [field('f1'), field('f1')] })),
    ).toContain('DUPLICATE_CUSTOM_FIELD_ID')
  })

  it('rejects duplicate dependency links between the same pair and type', () => {
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
          dependencies: [
            makeDependency('d1', 'a', 'b', 'FS'),
            makeDependency('d2', 'a', 'b', 'FS', 60),
          ],
        }),
      ),
    ).toContain('DUPLICATE_DEPENDENCY_LINK')
  })

  it('rejects hierarchy problems: cycles, self parents, missing parents', () => {
    expect(
      codes(
        validate({
          tasks: [
            makeTask({ id: 'a', parentTaskId: asTaskId('b') }),
            makeTask({ id: 'b', parentTaskId: asTaskId('a') }),
          ],
        }),
      ),
    ).toContain('PARENT_CYCLE')
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', parentTaskId: asTaskId('a') })] })),
    ).toContain('SELF_PARENT')
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', parentTaskId: asTaskId('ghost') })] })),
    ).toContain('MISSING_PARENT')
  })

  it('rejects duplicate task uids while keeping uid distinct from TaskId', () => {
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', uid: 9 }), makeTask({ id: 'b', uid: 9 })] })),
    ).toContain('DUPLICATE_TASK_UID')
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', uid: 9 }), makeTask({ id: 'b', uid: 10 })] })),
    ).toEqual([])
  })

  it('rejects outline levels inconsistent with hierarchy depth', () => {
    expect(
      codes(
        validate({
          tasks: [
            makeTask({ id: 'root', summary: true }),
            makeTask({ id: 'c', parentTaskId: asTaskId('root'), outlineLevel: 3 }),
          ],
        }),
      ),
    ).toContain('INCONSISTENT_OUTLINE_LEVEL')
    expect(
      codes(
        validate({
          tasks: [
            makeTask({ id: 'root', summary: true }),
            makeTask({ id: 'mid', summary: true, parentTaskId: asTaskId('root'), outlineLevel: 2 }),
            makeTask({ id: 'leaf', parentTaskId: asTaskId('mid'), outlineLevel: 3 }),
          ],
        }),
      ),
    ).toEqual([])
  })

  it('rejects summary flags inconsistent with child relationships', () => {
    // Summary flagged without children.
    expect(codes(validate({ tasks: [makeTask({ id: 'a', summary: true })] }))).toContain(
      'INCONSISTENT_SUMMARY_FLAG',
    )
    // Children present but not flagged as summary.
    expect(
      codes(
        validate({
          tasks: [
            makeTask({ id: 's', summary: false }),
            makeTask({ id: 'c', parentTaskId: asTaskId('s'), outlineLevel: 2 }),
          ],
        }),
      ),
    ).toContain('INCONSISTENT_SUMMARY_FLAG')
  })

  it('rejects dependency cycles, self links, missing references, and bad lag', () => {
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
          dependencies: [
            makeDependency('d1', 'a', 'b', 'FS'),
            makeDependency('d2', 'b', 'a', 'FS'),
          ],
        }),
      ),
    ).toContain('DEPENDENCY_CYCLE')
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' })],
          dependencies: [makeDependency('d1', 'a', 'a', 'FS')],
        }),
      ),
    ).toContain('SELF_DEPENDENCY')
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' })],
          dependencies: [makeDependency('d1', 'a', 'x', 'FS')],
        }),
      ),
    ).toContain('MISSING_TASK_REFERENCE')
    const fractional: Dependency = { ...makeDependency('d1', 'a', 'b', 'FS'), lagMinutes: 0.5 }
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
          dependencies: [fractional],
        }),
      ),
    ).toContain('INVALID_LAG')
  })

  it('rejects dependencies between a summary and its own descendant', () => {
    expect(
      codes(
        validate({
          tasks: [
            makeTask({ id: 's', summary: true }),
            makeTask({ id: 'c', parentTaskId: asTaskId('s'), outlineLevel: 2 }),
          ],
          dependencies: [makeDependency('d1', 'c', 's', 'FS')],
        }),
      ),
    ).toContain('SUMMARY_DEPENDENCY')
  })

  it('rejects negative durations and out-of-range percentages', () => {
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', duration: asWorkingMinutes(-60) })] })),
    ).toContain('NEGATIVE_DURATION')
    expect(codes(validate({ tasks: [makeTask({ id: 'a', percentComplete: 150 })] }))).toContain(
      'INVALID_PERCENT_COMPLETE',
    )
    expect(codes(validate({ tasks: [makeTask({ id: 'a', percentComplete: -5 })] }))).toContain(
      'INVALID_PERCENT_COMPLETE',
    )
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', physicalPercentComplete: 101 })] })),
    ).toContain('INVALID_PERCENT_COMPLETE')
    expect(codes(validate({ tasks: [makeTask({ id: 'a', outlineLevel: 0 })] }))).toContain(
      'INVALID_OUTLINE_LEVEL',
    )
  })

  it('rejects calendar problems: missing refs, cycles, malformed periods, bad exceptions', () => {
    const missingDefault = makeDocument({ tasks: [makeTask({ id: 'a' })] })
    missingDefault.properties.defaultCalendarId = asCalendarId('ghost')
    expect(codes(validateProjectDocument(missingDefault))).toContain('MISSING_CALENDAR')

    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', calendarId: asCalendarId('ghost') })] })),
    ).toContain('MISSING_CALENDAR')

    const first = makeCalendar('first', { baseCalendarId: asCalendarId('second') })
    const second = makeCalendar('second', { baseCalendarId: asCalendarId('first') })
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a' })], calendars: [first, second] })),
    ).toContain('CALENDAR_CYCLE')

    const orphan = makeCalendar('orphan', { baseCalendarId: asCalendarId('ghost') })
    expect(codes(validate({ tasks: [makeTask({ id: 'a' })], calendars: [orphan] }))).toContain(
      'MISSING_BASE_CALENDAR',
    )

    const reversed = makeCalendar('reversed', {
      workingWeek: { ...standardWeek(), 1: [{ startMinute: 1020, endMinute: 540 }] },
    })
    expect(codes(validate({ tasks: [makeTask({ id: 'a' })], calendars: [reversed] }))).toContain(
      'CALENDAR_PERIOD_MALFORMED',
    )

    const overlapping = makeCalendar('overlapping', {
      workingWeek: {
        ...standardWeek(),
        2: [
          { startMinute: 540, endMinute: 720 },
          { startMinute: 700, endMinute: 900 },
        ],
      },
    })
    expect(codes(validate({ tasks: [makeTask({ id: 'a' })], calendars: [overlapping] }))).toContain(
      'CALENDAR_PERIOD_MALFORMED',
    )

    const badException = makeCalendar('exception', {
      exceptions: [{ date: '08/04/2026', periods: [] }],
    })
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a' })], calendars: [badException] })),
    ).toContain('INVALID_DATE')
  })

  it('rejects missing assignment, baseline, and custom-field references', () => {
    const resource: Resource = {
      id: asResourceId('r1'),
      uid: 1,
      name: 'r',
      kind: 'work',
      maxUnits: 1,
      standardRate: 0,
      overtimeRate: 0,
      costPerUse: 0,
      availability: [],
    }
    const assignment: Assignment = {
      id: asAssignmentId('as1'),
      taskId: asTaskId('ghost'),
      resourceId: asResourceId('r1'),
      units: 1,
      work: asWorkingMinutes(0),
      actualWork: asWorkingMinutes(0),
      remainingWork: asWorkingMinutes(0),
      cost: 0,
      actualCost: 0,
      remainingCost: 0,
    }
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' })],
          resources: [resource],
          assignments: [assignment],
        }),
      ),
    ).toContain('MISSING_TASK_REFERENCE')

    const assignment2: Assignment = {
      ...assignment,
      id: asAssignmentId('as2'),
      taskId: asTaskId('a'),
      resourceId: asResourceId('ghost'),
    }
    expect(
      codes(
        validate({
          tasks: [makeTask({ id: 'a' })],
          resources: [resource],
          assignments: [assignment2],
        }),
      ),
    ).toContain('MISSING_RESOURCE_REFERENCE')

    const taskWithBaseline: Task = {
      ...makeTask({ id: 'a' }),
      baseline: [asBaselineId('missing-baseline')],
    }
    expect(codes(validate({ tasks: [taskWithBaseline] }))).toContain('MISSING_BASELINE_REFERENCE')

    const taskWithField: Task = {
      ...makeTask({ id: 'a' }),
      customFields: { [asCustomFieldId('missing-field')]: 'value' } as Task['customFields'],
    }
    expect(codes(validate({ tasks: [taskWithField] }))).toContain('MISSING_CUSTOM_FIELD_REFERENCE')
  })

  it('rejects malformed date strings anywhere they appear', () => {
    expect(
      codes(validate({ tasks: [makeTask({ id: 'a', start: asISODateTime('not-a-date') })] })),
    ).toContain('INVALID_DATE')
    const badStart = makeDocument({ tasks: [makeTask({ id: 'a' })] })
    badStart.properties.startDate = asISODateTime('2026-13-45T99:00:00.000Z')
    expect(codes(validateProjectDocument(badStart))).toContain('INVALID_DATE')
  })
})
