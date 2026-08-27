import { describe, expect, it } from 'vitest'
import { asISODateTime, asTaskId, asWorkingMinutes } from '@genoffice/project-contracts'
import type { DerivedSchedule, ProjectDocument } from '@genoffice/project-contracts'
import {
  EDITABLE_TASK_FIELDS,
  commitTaskEdit,
  editableTaskFields,
  initialTaskEditDraft,
  isTaskFieldEditable,
} from '../src/index.js'
import { makeDocument, makeTask, outlineDocument } from './fixtures.js'

const leafDocument = (): ProjectDocument =>
  makeDocument({
    tasks: [
      makeTask({
        id: 'a',
        name: 'Design',
        duration: asWorkingMinutes(480),
        start: asISODateTime('2026-08-03T09:00:00.000Z'),
        finish: asISODateTime('2026-08-03T17:00:00.000Z'),
      }),
    ],
  })

const summaryFixture = (): ProjectDocument => outlineDocument() // root(summary) > a(summary) > a1, b

const scheduleOf = (
  entries: Record<string, { duration?: number; scheduledStart?: string; scheduledFinish?: string }>,
): DerivedSchedule =>
  ({
    projectStart: '2026-08-03T09:00:00.000Z',
    projectFinish: '2026-08-07T17:00:00.000Z',
    taskSchedules: Object.fromEntries(
      Object.entries(entries).map(([id, entry]) => [
        id,
        {
          taskId: asTaskId(id),
          totalSlack: 0,
          freeSlack: 0,
          critical: false,
          duration: asWorkingMinutes(entry.duration ?? 480),
          ...(entry.scheduledStart !== undefined ? { scheduledStart: entry.scheduledStart } : {}),
          ...(entry.scheduledFinish !== undefined
            ? { scheduledFinish: entry.scheduledFinish }
            : {}),
        },
      ]),
    ),
    diagnostics: [],
  }) as unknown as DerivedSchedule

// ===========================================================================
// Editability
// ===========================================================================

describe('PROJECT-023 editing — editable fields', () => {
  it('a leaf task edits name, duration, start, and finish (canonical grid order)', () => {
    const document = leafDocument()
    expect(editableTaskFields(document.tasks[0]!)).toEqual([
      'taskName',
      'duration',
      'start',
      'finish',
    ])
  })

  it('a summary task edits ONLY its name — scheduling values are derived roll-ups', () => {
    const document = summaryFixture()
    expect(editableTaskFields(document.tasks[0]!)).toEqual(['taskName'])
    expect(isTaskFieldEditable(document.tasks[0]!, 'duration')).toBe(false)
    expect(isTaskFieldEditable(document.tasks[0]!, 'start')).toBe(false)
    expect(isTaskFieldEditable(document.tasks[0]!, 'finish')).toBe(false)
    expect(isTaskFieldEditable(document.tasks[0]!, 'taskName')).toBe(true)
  })

  it('a milestone leaf keeps the full leaf set (zero-duration edits are the milestone domain)', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'm', milestone: true, duration: asWorkingMinutes(0) })],
    })
    expect(editableTaskFields(document.tasks[0]!)).toEqual(EDITABLE_TASK_FIELDS)
  })
})

// ===========================================================================
// Initial draft
// ===========================================================================

describe('PROJECT-023 editing — initial draft (edit what the cell shows)', () => {
  it('the name draft is the task name verbatim', () => {
    const document = leafDocument()
    expect(initialTaskEditDraft(document, undefined, asTaskId('a'), 'taskName')).toBe('Design')
  })

  it('the duration draft follows the schedule-first precedence, falling back to the task field', () => {
    const document = leafDocument()
    expect(initialTaskEditDraft(document, undefined, asTaskId('a'), 'duration')).toBe('480')
    const scheduled = scheduleOf({ a: { duration: 960 } })
    expect(initialTaskEditDraft(document, scheduled, asTaskId('a'), 'duration')).toBe('960')
  })

  it('start/finish drafts are the SCHEDULED instants and EMPTY without a schedule (never invented)', () => {
    const document = leafDocument()
    // No schedule: the grid start/finish cells are empty — the draft starts
    // empty even though the task carries stored start/finish echoes.
    expect(initialTaskEditDraft(document, undefined, asTaskId('a'), 'start')).toBe('')
    expect(initialTaskEditDraft(document, undefined, asTaskId('a'), 'finish')).toBe('')
    const scheduled = scheduleOf({
      a: {
        scheduledStart: '2026-08-03T09:00:00.000Z',
        scheduledFinish: '2026-08-03T17:00:00.000Z',
      },
    })
    expect(initialTaskEditDraft(document, scheduled, asTaskId('a'), 'start')).toBe(
      '2026-08-03T09:00:00.000Z',
    )
    expect(initialTaskEditDraft(document, scheduled, asTaskId('a'), 'finish')).toBe(
      '2026-08-03T17:00:00.000Z',
    )
  })

  it('an unknown task deterministically yields an empty draft', () => {
    const document = leafDocument()
    expect(initialTaskEditDraft(document, undefined, asTaskId('ghost'), 'taskName')).toBe('')
  })
})

// ===========================================================================
// Commit translation
// ===========================================================================

describe('PROJECT-023 editing — commit translation (draft → semantic command)', () => {
  it('a name edit becomes a verbatim RenameTask command', () => {
    const document = leafDocument()
    const commit = commitTaskEdit(document, {
      taskId: asTaskId('a'),
      field: 'taskName',
      draft: 'Design v2',
    })
    expect(commit).toEqual({
      kind: 'apply',
      command: { type: 'RenameTask', taskId: asTaskId('a'), name: 'Design v2' },
    })
  })

  it('an unchanged name is a noChange (no no-op command is ever journaled)', () => {
    const document = leafDocument()
    expect(
      commitTaskEdit(document, { taskId: asTaskId('a'), field: 'taskName', draft: 'Design' }),
    ).toEqual({ kind: 'noChange' })
  })

  it('a duration edit parses canonical decimal text into SetTaskDuration', () => {
    const document = leafDocument()
    expect(
      commitTaskEdit(document, { taskId: asTaskId('a'), field: 'duration', draft: '960' }),
    ).toEqual({
      kind: 'apply',
      command: {
        type: 'SetTaskDuration',
        taskId: asTaskId('a'),
        duration: asWorkingMinutes(960),
      },
    })
  })

  it('negative and fractional duration drafts PARSE — the engine owns the semantic rejection', () => {
    const document = leafDocument()
    // Syntax parses; the ENGINE rejects with INVALID_DURATION (the renderer
    // owns text syntax only — the single-validation-authority split).
    expect(
      commitTaskEdit(document, { taskId: asTaskId('a'), field: 'duration', draft: '-480' }),
    ).toEqual({
      kind: 'apply',
      command: {
        type: 'SetTaskDuration',
        taskId: asTaskId('a'),
        duration: asWorkingMinutes(-480),
      },
    })
    expect(
      commitTaskEdit(document, { taskId: asTaskId('a'), field: 'duration', draft: '480.5' }),
    ).toEqual({
      kind: 'apply',
      command: {
        type: 'SetTaskDuration',
        taskId: asTaskId('a'),
        duration: asWorkingMinutes(480.5),
      },
    })
  })

  it('non-canonical duration text is a deterministic invalid (no command dispatched)', () => {
    const document = leafDocument()
    for (const draft of ['', '  480', '480 ', '1e3', '0x1E0', '480,5', '8h', 'PT8H', '4_80']) {
      expect(
        commitTaskEdit(document, { taskId: asTaskId('a'), field: 'duration', draft }),
        `draft ${JSON.stringify(draft)}`,
      ).toEqual({ kind: 'invalid', reason: 'unparseableDuration' })
    }
  })

  it('an unchanged duration is a noChange', () => {
    const document = leafDocument()
    expect(
      commitTaskEdit(document, { taskId: asTaskId('a'), field: 'duration', draft: '480' }),
    ).toEqual({ kind: 'noChange' })
  })

  it('start/finish edits pass the draft through VERBATIM (the engine validates dates)', () => {
    const document = leafDocument()
    expect(
      commitTaskEdit(document, {
        taskId: asTaskId('a'),
        field: 'start',
        draft: '2026-08-04T09:00:00.000Z',
      }),
    ).toEqual({
      kind: 'apply',
      command: {
        type: 'SetTaskStart',
        taskId: asTaskId('a'),
        start: '2026-08-04T09:00:00.000Z',
      },
    })
    // Even a malformed date passes through — the engine's INVALID_DATE is
    // the single date validator.
    expect(
      commitTaskEdit(document, { taskId: asTaskId('a'), field: 'finish', draft: 'not-a-date' }),
    ).toEqual({
      kind: 'apply',
      command: { type: 'SetTaskFinish', taskId: asTaskId('a'), finish: 'not-a-date' },
    })
  })

  it('a start/finish draft equal to the STORED field is a noChange (displayed-value pins still apply)', () => {
    const document = leafDocument() // stored start = 08-03T09:00, stored finish = 08-03T17:00
    expect(
      commitTaskEdit(document, {
        taskId: asTaskId('a'),
        field: 'start',
        draft: '2026-08-03T09:00:00.000Z',
      }),
    ).toEqual({ kind: 'noChange' })
    expect(
      commitTaskEdit(document, {
        taskId: asTaskId('a'),
        field: 'finish',
        draft: '2026-08-03T17:00:00.000Z',
      }),
    ).toEqual({ kind: 'noChange' })
    // A task WITHOUT a stored start: an empty draft is the noChange; a
    // typed instant becomes a real pin command.
    const bare = makeDocument({ tasks: [makeTask({ id: 'bare' })] })
    expect(commitTaskEdit(bare, { taskId: asTaskId('bare'), field: 'start', draft: '' })).toEqual({
      kind: 'noChange',
    })
    expect(
      commitTaskEdit(bare, {
        taskId: asTaskId('bare'),
        field: 'start',
        draft: '2026-08-04T09:00:00.000Z',
      }),
    ).toEqual({
      kind: 'apply',
      command: {
        type: 'SetTaskStart',
        taskId: asTaskId('bare'),
        start: '2026-08-04T09:00:00.000Z',
      },
    })
  })

  it('a missing task is a deterministic missingTask invalid', () => {
    const document = leafDocument()
    expect(
      commitTaskEdit(document, {
        taskId: asTaskId('ghost'),
        field: 'taskName',
        draft: 'X',
      }),
    ).toEqual({ kind: 'invalid', reason: 'missingTask' })
  })

  it('no active edit is the none outcome', () => {
    const document = leafDocument()
    expect(commitTaskEdit(document, undefined)).toEqual({ kind: 'none' })
  })

  it('is pure: identical inputs produce identical outcomes and the document is never mutated', () => {
    const document = leafDocument()
    const editing = { taskId: asTaskId('a'), field: 'duration', draft: '960' } as const
    const before = JSON.stringify(document)
    const first = commitTaskEdit(document, editing)
    const second = commitTaskEdit(document, editing)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(JSON.stringify(document)).toBe(before)
  })
})
