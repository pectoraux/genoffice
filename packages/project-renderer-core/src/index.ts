/**
 * @genoffice/project-renderer-core — the shared Project renderer boundary
 * (PROJECT-021/022/023).
 *
 * Host-independent projection and control layer over the canonical Project
 * domain: view state + intents + reducer, the document/schedule view
 * projection, deterministic timeline math, structural command builders,
 * the renderer session controller (command application, snapshot
 * undo/redo, injected scheduling), the PROJECT-022 Gantt view models —
 * the virtualized, synchronized task grid and timeline surfaces (bars,
 * milestones, dependency links) in fraction space — and the PROJECT-023
 * selection/editing surface (keyboard focus navigation, cell-edit
 * activation, and the deterministic draft → semantic-command commit flow
 * through the session). Hosts (Electron desktop, web) render these
 * projections and dispatch these intents; the renderer core owns NO
 * Project semantics — scheduling authority stays with the scheduling
 * engine, semantic mutations stay `ProjectCommand` values through the
 * engine, and this package imports only `@genoffice/project-contracts`
 * and `@genoffice/project-engine` (architecture-lock §3/§9/§11/§13).
 */
export {
  type ProjectViewState,
  type TaskSelection,
  type TimelineViewport,
  createViewState,
  parseInstant,
  formatInstant,
  reconcileViewState,
} from './state.js'
export { type ProjectViewIntent, type SelectMode, type MoveFocusDirection } from './intents.js'
export { type ViewReducerContext, reduceViewState } from './reduce.js'
export {
  type TimeAxisBand,
  type TimeAxisLevel,
  MIN_VIEWPORT_SPAN_MS,
  MAX_VIEWPORT_SPAN_MS,
  FIT_PADDING_RATIO,
  FIT_MIN_TOTAL_SPAN_MS,
  axisLevelForSpan,
  buildTimeAxis,
  fitViewport,
  projectWindow,
  scaleViewport,
  viewportFraction,
  viewportInstant,
} from './timeline.js'
export {
  type ProjectTaskRow,
  type ProjectViewProjection,
  projectDocumentView,
} from './projection.js'
export {
  type TaskInsertPosition,
  buildCreateTaskCommand,
  buildCreateTaskInSiblingGroupCommand,
  buildDeleteSelectionCommands,
  buildIndentCommand,
  buildOutdentCommand,
  defaultNewTask,
  nextTaskIdentity,
} from './commands.js'
export {
  type ApplyRendererCommandOutcome,
  type ProjectRendererSession,
  type ProjectRendererSessionOptions,
  type RendererSessionEntry,
  type ScheduleRunner,
  type UndoRedoOutcome,
  applyRendererCommand,
  canRedoRendererCommand,
  canUndoRendererCommand,
  createRendererSession,
  redoRendererCommand,
  rendererSessionJournal,
  undoRendererCommand,
} from './session.js'
// ---- PROJECT-022 — Gantt / task grid / timeline view models ----
export {
  type ProjectRowWindow,
  type RowWindowInput,
  buildRowWindow,
  rowWindowIsEmpty,
} from './views/virtualization.js'
export {
  DEFAULT_TASK_GRID_COLUMNS,
  type ProjectGridColumn,
  type ProjectGridCell,
  type ProjectGridPredecessorLink,
  type ProjectGridRow,
  type ProjectTaskGrid,
  type TaskGridField,
  buildTaskGrid,
} from './views/grid.js'
export { type ProjectGanttBar, type ProjectGanttBars, buildGanttBars } from './views/gantt-bars.js'
export {
  type ProjectMilestone,
  type ProjectMilestones,
  buildMilestones,
} from './views/milestones.js'
export {
  type ProjectDependencies,
  type ProjectDependencyLink,
  type ProjectLinkEndpoint,
  type ProjectLinkPoint,
  buildDependencies,
} from './views/dependencies.js'
export { type ProjectTimeline, type ProjectTimelineRow, buildTimeline } from './views/timeline.js'
export {
  type GanttHitPoint,
  type GanttHitTarget,
  type GanttViewLayout,
  type ProjectGanttView,
  buildGanttView,
  hitTestGantt,
} from './views/gantt-view.js'
// ---- PROJECT-023 — selection / editing ----
export {
  type EditableTaskField,
  type TaskEditCommit,
  type TaskEditInvalidReason,
  type TaskEditing,
  EDITABLE_TASK_FIELDS,
  commitTaskEdit,
  editableTaskFields,
  initialTaskEditDraft,
  isTaskFieldEditable,
} from './editing.js'
export { type TaskEditFlowOutcome, commitTaskEditThroughSession } from './edit-flow.js'
