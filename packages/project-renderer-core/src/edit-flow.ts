/**
 * PROJECT-023 — the cell-edit commit flow (command/session integration).
 *
 * The one-call orchestration of a committed cell edit through the accepted
 * command/session pipeline, so two hosts cannot diverge over the protocol:
 *
 * ```text
 * user gesture (Enter in the active editor)
 * → commitTaskEdit(session.document, state.editing)   — pure translation
 * → applyRendererCommand(session, command)            — semantic command
 * → applyProjectCommand (project-engine)              — the engine accepts
 * → ProjectDocument (new immutable document)          — canonical state
 * → injected scheduler                                — DerivedSchedule
 * → reduceViewState(endTaskEdit + reconcile)          — view-state refresh
 * → projectDocumentView (the hosts' projection refresh)
 * ```
 *
 * Every step is the accepted PROJECT-021 machinery; this module only composes
 * it. The renderer NEVER mutates a document — the only path from an edit to
 * the canonical state is the semantic `ProjectCommand` through the engine
 * (architecture-lock §9).
 *
 * Deterministic behavior contract:
 *
 * - `apply` + ACCEPTED command → the command is journaled, the document and
 *   schedule are replaced, the editor ends, and the view state is
 *   RECONCILED against the new document — selection preservation after an
 *   accepted mutation is by construction (surviving selection kept, dead
 *   references pruned, anchor/focus kept when they survive).
 * - `apply` + REJECTED command (e.g. `INVALID_DURATION`, `INVALID_DATE`) →
 *   the session is returned as the SAME reference (nothing journaled, the
 *   engine's diagnostics surfaced verbatim on `result`), the editor still
 *   ends (the Microsoft Project behavior: a rejected commit reverts the
 *   cell; the host renders a toast from the diagnostics), and the view
 *   state is reconciled against the unchanged document — selection
 *   preservation after a rejected mutation is likewise by construction.
 * - `noChange` / `invalid` → no command is dispatched at all; the editor
 *   ends; the session is the same reference.
 * - `none` (no edit active) → a pure no-op: the same session AND state
 *   references.
 */
import type { ProjectCommandResult } from '@genoffice/project-contracts'
import type { ProjectViewState } from './state.js'
import { reduceViewState } from './reduce.js'
import {
  type ProjectRendererSession,
  applyRendererCommand,
  type RendererSessionEntry,
} from './session.js'
import { type TaskEditCommit, commitTaskEdit } from './editing.js'

/** The outcome of committing the active cell edit through the session. */
export interface TaskEditFlowOutcome {
  /** The next session — the SAME reference when no accepted command was
   * applied (rejected command, invalid edit, no change, or no edit). */
  readonly session: ProjectRendererSession
  /** The next view state: the editor ended and the state reconciled against
   * `session.document` (the SAME reference when no edit was active). */
  readonly state: ProjectViewState
  /** The pure translation outcome (`commitTaskEdit`). */
  readonly commit: TaskEditCommit
  /** The engine result of the dispatched command — present only when a
   * command was dispatched (`commit.kind === 'apply'`), with the engine's
   * verbatim diagnostics whether accepted or rejected. */
  readonly result?: ProjectCommandResult
  /** The journaled entry — present only when the command was ACCEPTED. */
  readonly entry?: RendererSessionEntry
}

/**
 * Commits the active cell edit (`state.editing`) through the session: pure
 * translation, semantic command application, editor end, and view-state
 * reconciliation against the resulting document — the full pipeline in one
 * deterministic call. See the module docs for the behavior contract.
 */
export function commitTaskEditThroughSession(
  session: ProjectRendererSession,
  state: ProjectViewState,
): TaskEditFlowOutcome {
  const commit = commitTaskEdit(session.document, state.editing)

  if (commit.kind === 'none') {
    return { session, state, commit }
  }

  if (commit.kind === 'apply') {
    const outcome = applyRendererCommand(session, commit.command)
    const nextSession = outcome.session
    // The editor ALWAYS ends on a commit attempt (accepted or rejected) and
    // the reducer reconciles the state against the session's CURRENT
    // document — the unchanged document for a rejection, the new document
    // for an acceptance (surviving selection kept, dead references pruned).
    const nextState = reduceViewState(
      state,
      { type: 'endTaskEdit' },
      { document: nextSession.document, schedule: nextSession.schedule },
    )
    return outcome.entry !== undefined
      ? {
          session: nextSession,
          state: nextState,
          commit,
          result: outcome.result,
          entry: outcome.entry,
        }
      : { session: nextSession, state: nextState, commit, result: outcome.result }
  }

  // noChange / invalid: no command dispatched; the editor ends and the
  // reducer reconciles against the unchanged document.
  const nextState = reduceViewState(
    state,
    { type: 'endTaskEdit' },
    { document: session.document, schedule: session.schedule },
  )
  return { session, state: nextState, commit }
}
