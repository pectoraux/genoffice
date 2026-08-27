/**
 * PROJECT-024 — the dependency-editing model.
 *
 * The host-independent editing surface for a dependency's two editable
 * fields — the relationship `type` (FS/SS/FF/SF) and the `lag` (integer
 * working minutes; negative lag IS a lead, the canonical scheduler domain).
 * It mirrors the PROJECT-023 task-editing model exactly: which fields can
 * enter edit mode, the canonical draft text an editor starts from, and the
 * deterministic translation of a committed draft into the semantic
 * `ProjectCommand` the session applies through the engine (lock §9 — every
 * edit leaves here as a command value; the renderer NEVER mutates a
 * document).
 *
 * Layer split (the engine is the single semantic validation authority):
 *
 * - **Syntax is parsed here, and only here**: a `type` draft must be exactly
 *   one of the four canonical type codes (uppercase, no whitespace — the
 *   canonical protocol text, not a localized label; label FORMATTING for
 *   display stays a host concern); a `lag` draft must be a canonical decimal
 *   number (strict, locale-free — no whitespace, exponent, or separators).
 *   Unparseable text is a deterministic `invalid` commit that dispatches NO
 *   command.
 * - **Semantics stay in the engine**: integer lag enforcement
 *   (`INVALID_LAG`), duplicate-link re-keying on a type change
 *   (`DUPLICATE_DEPENDENCY_LINK`), and reference validity
 *   (`MISSING_DEPENDENCY`) are the engine's rejections, surfaced through the
 *   session's `ProjectCommandResult.diagnostics`.
 *
 * The draft protocol is CANONICAL TEXT, initialized from the dependency's
 * stored values: the type code verbatim, the lag as decimal minutes. Both
 * fields are editable on EVERY dependency (there is no summary/leaf
 * distinction in the dependency domain).
 */
import type {
  DependencyId,
  DependencyType,
  ProjectCommand,
  ProjectDocument,
} from '@genoffice/project-contracts'

/** The editable dependency fields. */
export type EditableDependencyField = 'type' | 'lag'

/** Every editable dependency field in canonical order. */
export const EDITABLE_DEPENDENCY_FIELDS: readonly EditableDependencyField[] = ['type', 'lag']

/** The canonical type-code domain (the draft protocol text for `type`). */
export const DEPENDENCY_TYPE_CODES: readonly DependencyType[] = ['FS', 'SS', 'FF', 'SF']

/**
 * The active dependency edit: the target (dependency + field) and the
 * canonical draft text. One edit at a time (beginning a new edit replaces an
 * active one — the reducer enforces it). The draft is user input en route to
 * a command, NOT canonical state (lock §11): nothing reads it except the
 * commit translation.
 */
export interface DependencyEditing {
  readonly dependencyId: DependencyId
  readonly field: EditableDependencyField
  readonly draft: string
}

/**
 * The deterministic reason a committed dependency edit cannot become a
 * command. These are SYNTAX/structural failures the renderer owns; semantic
 * rejections (`INVALID_LAG`, `DUPLICATE_DEPENDENCY_LINK`,
 * `MISSING_DEPENDENCY`) come from the engine through the session result.
 */
export type DependencyEditInvalidReason =
  /** The editing target no longer exists in the live document. */
  | 'missingDependency'
  /** The type draft is not one of the four canonical type codes. */
  | 'unparseableDependencyType'
  /** The lag draft is not canonical decimal text. */
  | 'unparseableLag'

/**
 * The outcome of committing an active dependency edit:
 *
 * - `apply` — the draft became a semantic command; the host applies it
 *   through the session (`./edit-flow.js` does exactly this).
 * - `noChange` — the built payload equals the dependency's current value;
 *   dispatching would journal a no-op, so nothing is applied.
 * - `invalid` — a syntax/structural failure (no command is dispatched).
 * - `none` — no dependency edit is active.
 */
export type DependencyEditCommit =
  | { readonly kind: 'apply'; readonly command: ProjectCommand }
  | { readonly kind: 'noChange' }
  | { readonly kind: 'invalid'; readonly reason: DependencyEditInvalidReason }
  | { readonly kind: 'none' }

/** Canonical decimal text for a lag draft: an optional minus, integer
 * digits, an optional `.`-separated fraction — nothing else (no whitespace,
 * exponent notation, or locale separators; the core is locale-free). */
const CANONICAL_DECIMAL = /^-?\d+(\.\d+)?$/

const findDependency = (
  document: ProjectDocument,
  dependencyId: DependencyId,
): ProjectDocument['dependencies'][number] | undefined =>
  document.dependencies.find((dependency) => dependency.id === dependencyId)

/**
 * The editable fields for a dependency: both, always (there is no
 * summary/leaf distinction in the dependency domain — every dependency's
 * type and lag are engine-settable).
 */
export function editableDependencyFields(): readonly EditableDependencyField[] {
  return EDITABLE_DEPENDENCY_FIELDS
}

/**
 * The canonical draft text an edit of (dependency, field) starts from: the
 * type code verbatim, the lag as decimal minutes of the stored
 * `lagMinutes`. Deterministic and pure; empty when the dependency does not
 * exist (the reducer never activates an edit for a missing dependency — the
 * same self-consistency contract as `initialTaskEditDraft`).
 */
export function initialDependencyEditDraft(
  document: ProjectDocument,
  dependencyId: DependencyId,
  field: EditableDependencyField,
): string {
  const dependency = findDependency(document, dependencyId)
  if (dependency === undefined) return ''
  return field === 'type' ? dependency.type : String(dependency.lagMinutes)
}

/**
 * Translates an active dependency edit's draft into the semantic command
 * (or the deterministic reason it cannot become one). Pure: the same
 * `(document, editing)` always yields the same outcome; the document is
 * never mutated (a command VALUE is produced — applying it is the session's
 * job, `./edit-flow.js`).
 *
 * Field mappings (1:1 onto the frozen command union):
 *
 * - `type` → `ChangeDependencyType` — the draft must be exactly one of
 *   `FS`/`SS`/`FF`/`SF`; anything else is `invalid`
 *   (`unparseableDependencyType`).
 * - `lag` → `ChangeLag` — the draft parsed under the canonical decimal
 *   rule; unparseable text is `invalid` (`unparseableLag`). Fractional
 *   values PARSE and become commands — the engine rejects them with
 *   `INVALID_LAG` (the engine is the single validation authority; the
 *   renderer owns text syntax only — the `SetTaskDuration` precedent).
 *
 * `noChange`: the built payload equals the dependency's CURRENT stored
 * field — nothing is dispatched, so no no-op command is ever journaled.
 */
export function commitDependencyEdit(
  document: ProjectDocument,
  editing: DependencyEditing | undefined,
): DependencyEditCommit {
  if (editing === undefined) return { kind: 'none' }
  const dependency = findDependency(document, editing.dependencyId)
  if (dependency === undefined) return { kind: 'invalid', reason: 'missingDependency' }

  switch (editing.field) {
    case 'type': {
      if (!DEPENDENCY_TYPE_CODES.includes(editing.draft as DependencyType)) {
        return { kind: 'invalid', reason: 'unparseableDependencyType' }
      }
      if (editing.draft === dependency.type) return { kind: 'noChange' }
      return {
        kind: 'apply',
        command: {
          type: 'ChangeDependencyType',
          dependencyId: editing.dependencyId,
          dependencyType: editing.draft as DependencyType,
        },
      }
    }
    case 'lag': {
      if (!CANONICAL_DECIMAL.test(editing.draft)) {
        return { kind: 'invalid', reason: 'unparseableLag' }
      }
      const minutes = Number(editing.draft)
      if (minutes === dependency.lagMinutes) return { kind: 'noChange' }
      return {
        kind: 'apply',
        command: {
          type: 'ChangeLag',
          dependencyId: editing.dependencyId,
          lagMinutes: minutes,
        },
      }
    }
    default: {
      // Exhaustiveness guard: every editable field maps to a commit outcome.
      const exhaustive: never = editing.field
      return exhaustive
    }
  }
}
