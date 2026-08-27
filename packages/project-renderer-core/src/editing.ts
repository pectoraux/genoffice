/**
 * PROJECT-023 — the task-editing model.
 *
 * The host-independent editing surface over the PROJECT-022 grid fields:
 * which (row, field) pairs can enter edit mode, the canonical draft text an
 * editor starts from, and the deterministic translation of a committed draft
 * into the semantic `ProjectCommand` the session applies through the engine
 * (architecture-lock §9 — a renderer may not invent Project semantics from
 * raw state mutation; EVERY edit leaves here as a command value).
 *
 * Layer split (the engine is the single semantic validation authority):
 *
 * - **Syntax is parsed here, and only here**: a duration draft must be a
 *   canonical decimal number (strict, locale-free text — no whitespace, no
 *   exponent, no thousands separators, `.` as the only decimal separator);
 *   unparseable text is a deterministic `invalid` commit that dispatches NO
 *   command. Task names and ISO instants pass through VERBATIM — the engine's
 *   `INVALID_DATE` (and any future name rule) fires from the engine, never
 *   duplicated in the renderer.
 * - **Semantics stay in the engine**: negativity, fractionality, summary
 *   roll-ups (`SUMMARY_DURATION_NOT_SETTABLE`, `SUMMARY_FINISH_NOT_SETTABLE`)
 *   and date validity are the engine's rejections, surfaced to the host
 *   through the session's `ProjectCommandResult.diagnostics`.
 *
 * Editability (a projection-layer fact derived from the canonical task, never
 * invented semantics): `taskName` is editable on EVERY row; `duration`,
 * `start`, and `finish` are editable on LEAF rows only — a summary's
 * scheduling values are derived roll-ups of its subtree (the scheduler
 * overwrites the stored fields on every derivation, so an accepted summary
 * edit would be an invisible edit; the `SetPercentComplete` precedent).
 *
 * The draft protocol is CANONICAL TEXT, initialized from the DISPLAYED cell
 * value (the PROJECT-022 schedule-first precedence — the user edits what the
 * grid shows): the name verbatim, the duration as decimal working minutes,
 * start/finish as ISO instants (empty when the cell is empty). Label, number,
 * and date FORMATTING for display remains a host/locale concern; the edit
 * draft is the one place canonical text is required, because it is protocol,
 * not presentation.
 */
import { asWorkingMinutes } from '@genoffice/project-contracts'
import type {
  DerivedSchedule,
  ProjectCommand,
  ProjectDocument,
  Task,
  TaskId,
} from '@genoffice/project-contracts'

/** The editable grid fields (the PROJECT-022 `TaskGridField` ids for the
 * four task-identity/scheduling fields this increment covers — percent
 * complete, constraints, and deadlines remain direct command construction,
 * the PROJECT-021 rule; dependency editing is PROJECT-024). */
export type EditableTaskField = 'taskName' | 'duration' | 'start' | 'finish'

/** Every editable field in canonical grid order. */
export const EDITABLE_TASK_FIELDS: readonly EditableTaskField[] = [
  'taskName',
  'duration',
  'start',
  'finish',
]

/**
 * The active cell edit: the target (task + field) and the canonical draft
 * text. One edit at a time (beginning a new edit replaces an active one —
 * the reducer enforces it). The draft is user input en route to a command —
 * it is NOT a scheduling-derived value cached in renderer state (lock §11):
 * nothing reads it except the commit translation.
 */
export interface TaskEditing {
  readonly taskId: TaskId
  readonly field: EditableTaskField
  readonly draft: string
}

/**
 * The deterministic reason a committed edit cannot become a command. These
 * are SYNTAX/structural failures the renderer owns; semantic rejections
 * (`INVALID_DURATION`, `INVALID_DATE`, `MISSING_TASK`,
 * `SUMMARY_*_NOT_SETTABLE`) come from the engine through the session result.
 */
export type TaskEditInvalidReason =
  /** The editing target no longer exists in the live document. */
  | 'missingTask'
  /** The duration draft is not canonical decimal text. */
  | 'unparseableDuration'

/**
 * The outcome of committing an active edit:
 *
 * - `apply` — the draft became a semantic command; the host applies it
 *   through the session (`./edit-flow.js` does exactly this).
 * - `noChange` — the built payload equals the task's current field value;
 *   dispatching would journal a no-op, so nothing is applied.
 * - `invalid` — a syntax/structural failure (no command is dispatched).
 * - `none` — no edit is active.
 */
export type TaskEditCommit =
  | { readonly kind: 'apply'; readonly command: ProjectCommand }
  | { readonly kind: 'noChange' }
  | { readonly kind: 'invalid'; readonly reason: TaskEditInvalidReason }
  | { readonly kind: 'none' }

/** Canonical decimal text for a duration draft: an optional minus, integer
 * digits, an optional `.`-separated fraction — nothing else (no whitespace,
 * exponent notation, or locale separators; the core is locale-free). */
const CANONICAL_DECIMAL = /^-?\d+(\.\d+)?$/

const findTask = (document: ProjectDocument, taskId: TaskId): Task | undefined =>
  document.tasks.find((task) => task.id === taskId)

/**
 * The editable fields for a task in canonical grid order: `taskName` on every
 * row; the leaf scheduling inputs (`duration`, `start`, `finish`) on leaf
 * rows only. Summary scheduling values are derived roll-ups.
 */
export function editableTaskFields(task: Task): readonly EditableTaskField[] {
  return task.summary ? ['taskName'] : ['taskName', 'duration', 'start', 'finish']
}

/** Whether (task, field) can enter edit mode (the `editableTaskFields`
 * membership test as a predicate — the reducer gates `beginTaskEdit` with
 * it and hosts use it to decide whether to offer the gesture at all). */
export function isTaskFieldEditable(task: Task, field: EditableTaskField): boolean {
  return editableTaskFields(task).includes(field)
}

/**
 * The canonical draft text an edit of (task, field) starts from — the
 * DISPLAYED cell value (the PROJECT-022 schedule-first precedence): the name
 * verbatim; the duration as decimal minutes of `schedule.duration ?? task.duration`;
 * start/finish as the schedule's scheduled instants, EMPTY when the cell is
 * empty (no schedule — dates are never invented). Deterministic and pure.
 */
export function initialTaskEditDraft(
  document: ProjectDocument,
  schedule: DerivedSchedule | undefined,
  taskId: TaskId,
  field: EditableTaskField,
): string {
  const task = findTask(document, taskId)
  if (task === undefined) return ''
  if (field === 'taskName') return task.name
  const taskSchedule = schedule?.taskSchedules[taskId]
  if (field === 'duration') {
    return String(taskSchedule?.duration ?? task.duration)
  }
  if (field === 'start') {
    return taskSchedule?.scheduledStart ?? ''
  }
  return taskSchedule?.scheduledFinish ?? ''
}

/**
 * Translates an active edit's draft into the semantic command (or the
 * deterministic reason it cannot become one). Pure: the same
 * `(document, editing)` always yields the same outcome; the document is
 * never mutated (a command VALUE is produced — applying it is the session's
 * job, `./edit-flow.js`).
 *
 * Field mappings (1:1 onto the frozen command union):
 *
 * - `taskName` → `RenameTask` — the draft VERBATIM (any string; name
 *   validity is the engine's authority).
 * - `duration` → `SetTaskDuration` — the draft parsed under the canonical
 *   decimal rule; unparseable text is `invalid` (`unparseableDuration`).
 *   Negative and fractional values PARSE and become commands — the engine
 *   rejects them with `INVALID_DURATION` (the engine is the single
 *   validation authority; the renderer owns text syntax only).
 * - `start` → `SetTaskStart` and `finish` → `SetTaskFinish` — the draft
 *   VERBATIM; the engine's `INVALID_DATE` rejects malformed instants.
 *
 * `noChange`: the built payload equals the task's CURRENT stored field
 * (name === task.name, duration === task.duration, start === task.start,
 * finish === task.finish) — nothing is dispatched, so no no-op command is
 * ever journaled. Note this compares against the STORED field, not the
 * displayed cell: a start edit initialized from the derived `scheduledStart`
 * of a dependency-delayed task still produces a real command (it PINS the
 * candidate start to the displayed instant).
 */
export function commitTaskEdit(
  document: ProjectDocument,
  editing: TaskEditing | undefined,
): TaskEditCommit {
  if (editing === undefined) return { kind: 'none' }
  const task = findTask(document, editing.taskId)
  if (task === undefined) return { kind: 'invalid', reason: 'missingTask' }

  switch (editing.field) {
    case 'taskName': {
      if (editing.draft === task.name) return { kind: 'noChange' }
      return {
        kind: 'apply',
        command: { type: 'RenameTask', taskId: editing.taskId, name: editing.draft },
      }
    }
    case 'duration': {
      if (!CANONICAL_DECIMAL.test(editing.draft)) {
        return { kind: 'invalid', reason: 'unparseableDuration' }
      }
      const minutes = Number(editing.draft)
      if (minutes === (task.duration as number)) return { kind: 'noChange' }
      return {
        kind: 'apply',
        command: {
          type: 'SetTaskDuration',
          taskId: editing.taskId,
          duration: asWorkingMinutes(minutes),
        },
      }
    }
    case 'start': {
      if (editing.draft === (task.start ?? '')) return { kind: 'noChange' }
      return {
        kind: 'apply',
        command: { type: 'SetTaskStart', taskId: editing.taskId, start: editing.draft },
      }
    }
    case 'finish': {
      if (editing.draft === (task.finish ?? '')) return { kind: 'noChange' }
      return {
        kind: 'apply',
        command: { type: 'SetTaskFinish', taskId: editing.taskId, finish: editing.draft },
      }
    }
    default: {
      // Exhaustiveness guard: every editable field maps to a commit outcome.
      const exhaustive: never = editing.field
      return exhaustive
    }
  }
}
