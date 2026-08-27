/**
 * PROJECT-021 — the renderer session controller.
 *
 * The host-independent control layer that binds the semantic engine to a
 * host's view: it owns the current canonical `ProjectDocument`, the
 * authoritative `DerivedSchedule` of that exact document (when a scheduler
 * is wired), and the command history with snapshot undo/redo. Both hosts
 * (Electron desktop, web) drive the SAME controller, so command application,
 * rescheduling, and undo/redo semantics are identical everywhere
 * (R-009 + verification-matrix "command-to-renderer integration").
 *
 * Scheduling authority (architecture-lock §3/§6): the renderer core NEVER
 * schedules. The scheduler is INJECTED as a structurally-typed runner —
 * hosts pass `schedule` from `@genoffice/project-scheduling` (its
 * `DerivedSchedule` is structurally assignable), exactly like the accepted
 * PROJECT-020 compatibility-pipeline precedent. A session without a runner
 * simply carries no schedule; the projection then renders the document
 * without schedule values (never invented ones).
 *
 * Undo/redo is snapshot-based: every journaled entry carries the immutable
 * document/schedule pair from before and after the command. Because
 * `applyProjectCommand` never mutates its input (accepted commands produce
 * a new document), the snapshots are stable references and undo restores
 * the EXACT prior state — byte-identical, total over every accepted command
 * (including the deliberately non-invertible `OutdentTask`,
 * `CreateBaseline`, and `LevelResources`, for which the engine documents
 * that undo "requires a host-level document snapshot" — this controller IS
 * that shared host layer). The engine's per-command `inverse` values remain
 * available on the journal entries for hosts that want command-level undo
 * semantics; the controller's canonical undo is the snapshot restore.
 *
 * Identity: command ids are `c{n}` with `n` allocated monotonically and only
 * on ACCEPTED commands — deterministic for a given command sequence.
 */
import type {
  DerivedSchedule,
  JournalEntry,
  ProjectCommand,
  ProjectCommandResult,
  ProjectDocument,
} from '@genoffice/project-contracts'
import { applyProjectCommand } from '@genoffice/project-engine'

/**
 * The injected scheduling runner. `schedule` from the scheduling package
 * satisfies this type structurally; the renderer-core package itself stays
 * scheduling-free (no static scheduling dependency).
 */
export type ScheduleRunner = (document: ProjectDocument) => DerivedSchedule

export interface ProjectRendererSessionOptions {
  /** The injected scheduler. Absent → the session carries no schedule. */
  readonly schedule?: ScheduleRunner
}

/**
 * One journal entry: the command, its result, and the immutable document /
 * schedule snapshots from before and after acceptance. The engine's
 * `ProjectCommandResult.inverse` (when present) is carried on `result` for
 * hosts that want command-level undo semantics.
 */
export interface RendererSessionEntry {
  readonly commandId: string
  readonly command: ProjectCommand
  readonly result: ProjectCommandResult
  readonly before: Readonly<{ document: ProjectDocument; schedule?: DerivedSchedule }>
  readonly after: Readonly<{ document: ProjectDocument; schedule?: DerivedSchedule }>
}

/**
 * The shared renderer session. An immutable value: every operation returns a
 * new session (or the same reference for rejected commands / no-op
 * undo/redo); hosts never mutate it. The `scheduleRunner` field is a
 * function and intentionally not part of the serializable state — hosts that
 * persist sessions serialize the document/journal and re-inject the runner
 * on restore.
 */
export interface ProjectRendererSession {
  readonly document: ProjectDocument
  /** The derived schedule of `document` (present iff a runner is wired). */
  readonly schedule?: DerivedSchedule
  readonly past: readonly RendererSessionEntry[]
  readonly future: readonly RendererSessionEntry[]
  /** Increments on every document change (accepted command, undo, redo). */
  readonly revision: number
  /** The number of ACCEPTED commands so far (drives `c{n}` id allocation). */
  readonly commandSeq: number
  /** The injected scheduler (a function; not serializable state). */
  readonly scheduleRunner?: ScheduleRunner
}

export interface ApplyRendererCommandOutcome {
  /** The next session (the SAME reference when the command was rejected). */
  readonly session: ProjectRendererSession
  readonly result: ProjectCommandResult
  /** The journaled entry (undefined iff the command was rejected). */
  readonly entry?: RendererSessionEntry
}

export interface UndoRedoOutcome {
  /** The next session (the SAME reference when nothing was undone/redone). */
  readonly session: ProjectRendererSession
  /** The entry that was undone/redone (undefined on no-op). */
  readonly entry?: RendererSessionEntry
  readonly applied: boolean
}

/**
 * Creates a session for a document. When a scheduler is injected, the
 * initial `DerivedSchedule` is computed immediately (the schedule of the
 * given document). Pure: the same inputs produce the same session state.
 */
export function createRendererSession(
  document: ProjectDocument,
  options: ProjectRendererSessionOptions = {},
): ProjectRendererSession {
  const schedule = options.schedule !== undefined ? options.schedule(document) : undefined
  return {
    document,
    ...(schedule !== undefined ? { schedule } : {}),
    past: [],
    future: [],
    revision: 0,
    commandSeq: 0,
    ...(options.schedule !== undefined ? { scheduleRunner: options.schedule } : {}),
  }
}

/**
 * Applies one semantic command through the canonical engine. Rejected
 * commands leave the session untouched (same reference) with the engine's
 * diagnostics surfaced in `result`. Accepted commands produce the next
 * document, re-schedule it through the injected runner (when wired), journal
 * the entry with before/after snapshots, clear the redo future, and bump the
 * revision.
 */
export function applyRendererCommand(
  session: ProjectRendererSession,
  command: ProjectCommand,
): ApplyRendererCommandOutcome {
  const commandId = `c${session.commandSeq + 1}`
  const execution = applyProjectCommand(session.document, command, commandId)
  if (!execution.result.accepted) {
    return { session, result: execution.result }
  }
  const schedule =
    session.scheduleRunner !== undefined ? session.scheduleRunner(execution.document) : undefined
  const before: Readonly<{ document: ProjectDocument; schedule?: DerivedSchedule }> = {
    document: session.document,
    ...(session.schedule !== undefined ? { schedule: session.schedule } : {}),
  }
  const after: Readonly<{ document: ProjectDocument; schedule?: DerivedSchedule }> = {
    document: execution.document,
    ...(schedule !== undefined ? { schedule } : {}),
  }
  const entry: RendererSessionEntry = {
    commandId,
    command,
    result: execution.result,
    before,
    after,
  }
  const next: ProjectRendererSession = {
    document: execution.document,
    ...(schedule !== undefined ? { schedule } : {}),
    past: [...session.past, entry],
    future: [],
    revision: session.revision + 1,
    commandSeq: session.commandSeq + 1,
    ...(session.scheduleRunner !== undefined ? { scheduleRunner: session.scheduleRunner } : {}),
  }
  return { session: next, result: execution.result, entry }
}

/**
 * Undoes the most recent accepted command by restoring its `before`
 * snapshot — the exact prior document AND its exact derived schedule
 * (byte-identical; no re-execution, no inverse-command drift). A no-op
 * (same session reference, `applied: false`) when the history is empty.
 */
export function undoRendererCommand(session: ProjectRendererSession): UndoRedoOutcome {
  const entry = session.past[session.past.length - 1]
  if (entry === undefined) return { session, applied: false }
  const next: ProjectRendererSession = {
    document: entry.before.document,
    ...(entry.before.schedule !== undefined ? { schedule: entry.before.schedule } : {}),
    past: session.past.slice(0, -1),
    future: [...session.future, entry],
    revision: session.revision + 1,
    commandSeq: session.commandSeq,
    ...(session.scheduleRunner !== undefined ? { scheduleRunner: session.scheduleRunner } : {}),
  }
  return { session: next, entry, applied: true }
}

/**
 * Redoes the most recently undone command by restoring its `after`
 * snapshot. A no-op when there is nothing to redo.
 */
export function redoRendererCommand(session: ProjectRendererSession): UndoRedoOutcome {
  const entry = session.future[session.future.length - 1]
  if (entry === undefined) return { session, applied: false }
  const next: ProjectRendererSession = {
    document: entry.after.document,
    ...(entry.after.schedule !== undefined ? { schedule: entry.after.schedule } : {}),
    past: [...session.past, entry],
    future: session.future.slice(0, -1),
    revision: session.revision + 1,
    commandSeq: session.commandSeq,
    ...(session.scheduleRunner !== undefined ? { scheduleRunner: session.scheduleRunner } : {}),
  }
  return { session: next, entry, applied: true }
}

/** Whether an undo is available (history non-empty). */
export function canUndoRendererCommand(session: ProjectRendererSession): boolean {
  return session.past.length > 0
}

/** Whether a redo is available (future non-empty). */
export function canRedoRendererCommand(session: ProjectRendererSession): boolean {
  return session.future.length > 0
}

/**
 * The session's history as plain engine `JournalEntry` values (commandId,
 * command, result) — the shared shape hosts already know from
 * `@genoffice/project-engine`'s journal model, without the snapshots.
 */
export function rendererSessionJournal(session: ProjectRendererSession): readonly JournalEntry[] {
  return session.past.map((entry) => ({
    commandId: entry.commandId,
    command: entry.command,
    result: entry.result,
  }))
}
