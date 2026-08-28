/**
 * PROJECT-027 — the host keyboard/menu translation layer.
 *
 * THE one mapping from native input (DOM KeyboardEvent essentials, native
 * menu command ids) to the canonical renderer-core vocabulary: view
 * `ProjectViewIntent` values, host document actions (executed through the
 * renderer-core command builders), edit-flow actions, session history
 * actions, and file actions. Pure functions — the app controller
 * (`app.ts`) executes the returned actions; NOTHING here touches the DOM,
 * the session, or a document.
 *
 * This is the entire "host keyboard/menu translation" the PROJECT-027 work
 * item names: two tables (`translateKeyDown`, `translateMenuCommand`) and
 * the zoom constants. The host invents no Project semantics — every
 * semantic action is expressed in the accepted renderer-core/intent
 * vocabulary.
 */
import type { ProjectViewIntent } from '@genoffice/project-renderer-core'
import type { MenuCommandId } from '../shared/ipc.js'

/**
 * A host action: what the controller should do with one native input. The
 * `document`/`edit`/`history`/`file` kinds are executed by the controller
 * through the renderer-core builders/flows (they need the live session);
 * `intent` values are dispatched straight through the reducer.
 */
export type HostAction =
  | { readonly kind: 'none' }
  | { readonly kind: 'intent'; readonly intent: ProjectViewIntent }
  | {
      readonly kind: 'document'
      readonly action: 'createTask' | 'deleteSelection' | 'indentSelection' | 'outdentSelection'
    }
  | { readonly kind: 'edit'; readonly action: 'beginEditName' | 'commit' | 'cancel' }
  | { readonly kind: 'history'; readonly action: 'undo' | 'redo' }
  | { readonly kind: 'file'; readonly action: 'new' | 'open' | 'save' | 'saveAs' }
  | { readonly kind: 'view'; readonly action: 'collapseSelection' | 'expandSelection' }

/** The presentation-only viewport zoom factors (fraction of the window). */
export const ZOOM_IN_FACTOR = 0.8
export const ZOOM_OUT_FACTOR = 1.25

/** The essential keyboard facts the translation consumes (a DOM
 * KeyboardEvent reduced to plain data — testable without a DOM). */
export interface KeyInput {
  readonly key: string
  readonly ctrlOrMeta: boolean
  readonly shift: boolean
  readonly alt: boolean
}

/** The live editor mode (which keys the editor owns). */
export interface TranslationMode {
  /** A cell edit is active: Enter/Escape belong to the editor and every
   * other key passes through to the input (native caret/undo/text keys). */
  readonly editing: boolean
}

const NONE: HostAction = { kind: 'none' }

const focusMove = (direction: 'up' | 'down' | 'first' | 'last', extend: boolean): HostAction => ({
  kind: 'intent',
  intent: { type: 'moveTaskFocus', direction, ...(extend ? { extend: true } : {}) },
})

/**
 * Translates one keyboard input. Order of precedence:
 * editing mode → editor keys only; then modifier shortcuts; then plain
 * navigation/activation keys; anything else is `none` (the host never
 * invents bindings for keys the accepted vocabulary does not name).
 */
export function translateKeyDown(input: KeyInput, mode: TranslationMode): HostAction {
  if (mode.editing) {
    if (input.key === 'Enter') return { kind: 'edit', action: 'commit' }
    if (input.key === 'Escape') return { kind: 'edit', action: 'cancel' }
    return NONE
  }

  // ---- modifier shortcuts (the menu accelerators' single execution path) --
  if (input.ctrlOrMeta && !input.alt) {
    switch (input.key.toLowerCase()) {
      case 'z':
        return { kind: 'history', action: input.shift ? 'redo' : 'undo' }
      case 'y':
        return { kind: 'history', action: 'redo' }
      case 'n':
        return { kind: 'file', action: 'new' }
      case 'o':
        return { kind: 'file', action: 'open' }
      case 's':
        return { kind: 'file', action: input.shift ? 'saveAs' : 'save' }
      case 't':
        return { kind: 'document', action: 'createTask' }
      case '=':
      case '+':
        return { kind: 'intent', intent: { type: 'scaleViewport', factor: ZOOM_IN_FACTOR } }
      case '-':
      case '_':
        return { kind: 'intent', intent: { type: 'scaleViewport', factor: ZOOM_OUT_FACTOR } }
      case 'f':
        if (input.shift) return { kind: 'intent', intent: { type: 'fitViewport' } }
        return NONE
      default:
        return NONE
    }
  }

  // ---- outline gestures (MS Project conventions) --------------------------
  if (input.alt && input.shift) {
    if (input.key === 'ArrowRight') return { kind: 'document', action: 'indentSelection' }
    if (input.key === 'ArrowLeft') return { kind: 'document', action: 'outdentSelection' }
    if (input.key === '-') return { kind: 'view', action: 'collapseSelection' }
    if (input.key === '+' || input.key === '=') return { kind: 'view', action: 'expandSelection' }
  }
  if (input.alt) {
    if (input.key === 'ArrowRight') return { kind: 'document', action: 'indentSelection' }
    if (input.key === 'ArrowLeft') return { kind: 'document', action: 'outdentSelection' }
  }

  // ---- plain navigation / activation ---------------------------------------
  switch (input.key) {
    case 'ArrowUp':
      return focusMove('up', input.shift)
    case 'ArrowDown':
      return focusMove('down', input.shift)
    case 'Home':
      return focusMove('first', input.shift)
    case 'End':
      return focusMove('last', input.shift)
    case 'Enter':
    case 'F2':
      return { kind: 'edit', action: 'beginEditName' }
    case 'Delete':
      return { kind: 'document', action: 'deleteSelection' }
    case 'Insert':
      return { kind: 'document', action: 'createTask' }
    default:
      return NONE
  }
}

/**
 * Translates one native menu command id. The menu is transport: each id maps
 * to exactly the same action vocabulary the keyboard table produces, so the
 * two input paths are one translation (the discipline suite pins the
 * equality on shared ids).
 */
export function translateMenuCommand(command: MenuCommandId): HostAction {
  switch (command) {
    case 'file.new':
      return { kind: 'file', action: 'new' }
    case 'file.open':
      return { kind: 'file', action: 'open' }
    case 'file.save':
      return { kind: 'file', action: 'save' }
    case 'file.saveAs':
      return { kind: 'file', action: 'saveAs' }
    case 'edit.undo':
      return { kind: 'history', action: 'undo' }
    case 'edit.redo':
      return { kind: 'history', action: 'redo' }
    case 'edit.deleteTask':
      return { kind: 'document', action: 'deleteSelection' }
    case 'task.create':
      return { kind: 'document', action: 'createTask' }
    case 'task.indent':
      return { kind: 'document', action: 'indentSelection' }
    case 'task.outdent':
      return { kind: 'document', action: 'outdentSelection' }
    case 'view.zoomIn':
      return { kind: 'intent', intent: { type: 'scaleViewport', factor: ZOOM_IN_FACTOR } }
    case 'view.zoomOut':
      return { kind: 'intent', intent: { type: 'scaleViewport', factor: ZOOM_OUT_FACTOR } }
    case 'view.fit':
      return { kind: 'intent', intent: { type: 'fitViewport' } }
    case 'view.collapse':
      return { kind: 'view', action: 'collapseSelection' }
    case 'view.expand':
      return { kind: 'view', action: 'expandSelection' }
  }
}
