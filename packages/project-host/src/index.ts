/**
 * @genoffice/project-host — the shared Project host binding (PROJECT-028).
 *
 * The public surface both shells consume:
 *
 * - `createProjectApp` — the host controller: one `ProjectRendererSession`
 *   (the canonical scheduling authority injected through `bindings.ts`),
 *   one `ProjectViewState`, the projection + `buildGanttView` pipeline with
 *   the canonical calendar/allocation inputs threaded per render, the
 *   document open/save flows over the canonical file adapters, dirty
 *   tracking by document identity, and the close-guard handshake.
 * - `createUI` — the DOM rendering layer over the pure view models (which
 *   mounts the shared ribbon, PROJECT-029).
 * - `createRibbon` / `RIBBON_TABS` / `RIBBON_COMMAND_IDS` — the shared
 *   tabbed command surface both hosts render through `createUI` (the
 *   ribbon vocabulary is the shared `MENU_COMMAND_IDS` exactly).
 * - `confirmUnsavedChanges` / `createTaskInformationDialog` — the shared
 *   dialog layer (PROJECT-030): the unsaved-changes dialog and the Task
 *   Information dialog, one implementation rendered by both hosts.
 * - `translateKeyDown` / `translateMenuCommand` — the input translation
 *   tables (keyboard + menu converge on one action vocabulary).
 * - `newProjectDocument` / import/export composition — the document flows
 *   over the canonical `@genoffice/project-file` adapters.
 * - `scheduleRunner` / `workingTimeQuery` / `allocationQuery` — the three
 *   canonical injection seams (PROJECT-021/025/026).
 * - the `ProjectHostBridge` transport contract the shells implement.
 *
 * The `./styles.css` subpath carries the deterministic host stylesheet both
 * entries mount.
 */
export {
  createProjectApp,
  type AppDependencies,
  type HostAppState,
  type ProjectHostApp,
  type StatusMessage,
} from './app.js'
export { allocationQuery, scheduleRunner, workingTimeQuery } from './bindings.js'
export {
  type HostAppInfo,
  MENU_COMMAND_IDS,
  type MenuCommandId,
  type NativeReadResult,
  type OpenFileSelection,
  PROJECT_FILE_FILTERS,
  type ProjectHostBridge,
} from './bridge.js'
export {
  confirmUnsavedChanges,
  createTaskInformationDialog,
  type DiscardChoice,
  type TaskInformationCallbacks,
  type TaskInformationDialog,
  type TaskInformationInput,
  type TaskInformationRequest,
  type TaskInformationResult,
} from './dialogs.js'
export {
  createRibbon,
  RIBBON_COMMAND_IDS,
  type Ribbon,
  type RibbonCallbacks,
  type RibbonState,
  RIBBON_TABS,
} from './ribbon.js'
export {
  adapterForFormat,
  defaultFileNameFor,
  exportDocumentBytes,
  extensionForFormat,
  formatForPath,
  importDocumentBytes,
  type HostFileFormat,
  type ImportOutcome,
  type ImportedProject,
  newProjectDocument,
  STANDARD_CALENDAR_ID,
} from './document.js'
export {
  type HostAction,
  type KeyInput,
  translateKeyDown,
  translateMenuCommand,
  type TranslationMode,
  ZOOM_IN_FACTOR,
  ZOOM_OUT_FACTOR,
} from './translate.js'
export {
  createUI,
  RESOURCE_LABEL_WIDTH,
  RESOURCE_ROW_HEIGHT,
  ROW_HEIGHT,
  type UI,
  type UICallbacks,
  type UIUpdateInputs,
} from './ui.js'
