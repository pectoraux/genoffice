/**
 * PROJECT-021 — renderer intents.
 *
 * Hosts translate platform input events (mouse, keyboard, touch, menus) into
 * these host-independent intent values and dispatch them through
 * `reduceViewState`. The renderer core never sees a DOM event, a React
 * synthetic event, or an Electron menu id — host-specific APIs remain
 * outside the renderer (R-009).
 *
 * Intents are split into two families:
 *
 * - **View intents** (this module): change the interaction/view state. They
 *   are pure view-state transitions, validated against the live document.
 * - **Document mutation**: hosts build semantic `ProjectCommand` values
 *   (contracts) — directly for field edits, or through the structural
 *   builders in `./commands.js` for outline gestures whose canonical mapping
 *   (preceding-sibling indent, subtree deletion, identity allocation) must
 *   not be re-implemented per host — and apply them through the session
 *   controller (`./session.js`). Architecture-lock §9: a renderer may not
 *   invent Project semantics from raw state mutation; every mutation is a
 *   semantic command.
 */
import type {
  DependencyId,
  ProjectFilterId,
  ProjectGroupId,
  ProjectTableId,
  ProjectViewId,
  ResourceId,
  TaskId,
} from '@genoffice/project-contracts'

/** How a select intent combines with the existing selection. */
export type SelectMode =
  /** Replace the selection (plain click). */
  | 'set'
  /** Add/remove the entity, keeping the rest (ctrl/cmd click). */
  | 'toggle'
  /** Select the outline-order range from the anchor to the target
   * (shift click); replaces the selection with the range. */
  | 'extend'

export type ProjectViewIntent =
  // ---- task selection ----
  | { type: 'selectTask'; taskId: TaskId; mode?: SelectMode }
  | { type: 'selectTasks'; taskIds: readonly TaskId[] }
  | { type: 'clearSelection' }
  // ---- dependency / resource selection ----
  | { type: 'selectDependency'; dependencyId: DependencyId; mode?: SelectMode }
  | { type: 'selectResource'; resourceId: ResourceId; mode?: SelectMode }
  // ---- outline collapse ----
  | { type: 'toggleCollapse'; taskId: TaskId }
  | { type: 'setCollapsed'; taskIds: readonly TaskId[]; collapsed: boolean }
  | { type: 'collapseAll' }
  | { type: 'expandAll' }
  // ---- timeline viewport ----
  | { type: 'setViewport'; start: string; finish: string }
  /** Scale the visible window by `factor` around `focus` (default: the window
   * midpoint). `factor < 1` zooms in (smaller window), `> 1` zooms out. */
  | { type: 'scaleViewport'; factor: number; focus?: string }
  /** Fit the window to the project span (derived schedule window when
   * available, else the canonical properties window). */
  | { type: 'fitViewport' }
  // ---- active canonical view definitions ----
  | { type: 'setActiveView'; viewId?: ProjectViewId }
  | { type: 'setActiveTable'; tableId?: ProjectTableId }
  | { type: 'setActiveFilter'; filterId?: ProjectFilterId }
  | { type: 'setActiveGroup'; groupId?: ProjectGroupId }
