/**
 * @genoffice/project-renderer-core — the shared Project renderer boundary
 * (PROJECT-021).
 *
 * Host-independent projection and control layer over the canonical Project
 * domain: view state + intents + reducer, the document/schedule view
 * projection, deterministic timeline math, structural command builders, and
 * the renderer session controller (command application, snapshot
 * undo/redo, injected scheduling). Hosts (Electron desktop, web) render
 * these projections and dispatch these intents; the renderer core owns NO
 * Project semantics — scheduling authority stays with the scheduling engine,
 * semantic mutations stay `ProjectCommand` values through the engine, and
 * this package imports only `@genoffice/project-contracts` and
 * `@genoffice/project-engine` (architecture-lock §3/§9/§11/§13).
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
export { type ProjectViewIntent, type SelectMode } from './intents.js'
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
