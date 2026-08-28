/**
 * The Project host controller — the shared binding both shells run
 * (established as the desktop host controller at PROJECT-027; moved to the
 * shared `@genoffice/project-host` layer at PROJECT-028 so the web shell
 * runs the exact same controller).
 *
 * Binds the shared renderer core to a host window: one
 * `ProjectRendererSession` (the canonical scheduling authority injected),
 * one `ProjectViewState`, the projection + `buildGanttView` pipeline with
 * the canonical calendar/allocation inputs threaded, the document
 * open/save flows over the canonical file adapters, dirty tracking by
 * document IDENTITY (sessions are immutable and undo restores the exact
 * prior reference — the honest dirty signal), and the close-guard
 * handshake.
 *
 * The controller owns NO Project semantics: every mutation is a semantic
 * command built by the renderer-core builders and applied through the
 * session; every view change is an intent through the reducer; every
 * scheduling value is the injected authority's. The UI module (`ui.ts`)
 * maps the resulting view models to DOM; this module never touches a
 * pixel. The host bridge is injected (the Electron preload surface in the
 * desktop shell, the web bridge in the browser shell, an in-memory fake in
 * tests).
 */
import type {
  ImportDiagnostic,
  ProjectCommand,
  ProjectDocument,
  TaskId,
} from '@genoffice/project-contracts'
import {
  canRedoRendererCommand,
  canUndoRendererCommand,
  commitDependencyEditThroughSession,
  commitTaskEditThroughSession,
  createRendererSession,
  createViewState,
  initialTaskEditDraft,
  isTaskFieldEditable,
  nextTaskIdentity,
  projectDocumentView,
  reconcileViewState,
  redoRendererCommand,
  reduceViewState,
  undoRendererCommand,
  applyRendererCommand,
  buildGanttView,
  buildCreateTaskCommand,
  buildDeleteSelectionCommands,
  buildIndentCommand,
  buildOutdentCommand,
} from '@genoffice/project-renderer-core'
import type {
  CalendarViewInput,
  ProjectRendererSession,
  ProjectViewState,
  ProjectViewIntent,
  ResourceViewInput,
} from '@genoffice/project-renderer-core'
import type { MenuCommandId, ProjectHostBridge } from './bridge.js'
import { allocationQuery, scheduleRunner, workingTimeQuery } from './bindings.js'
import {
  defaultFileNameFor,
  exportDocumentBytes,
  formatForPath,
  importDocumentBytes,
  newProjectDocument,
} from './document.js'
import type { HostFileFormat } from './document.js'
import { translateKeyDown, translateMenuCommand } from './translate.js'
import type { HostAction, KeyInput } from './translate.js'
import { confirmUnsavedChanges, createTaskInformationDialog } from './dialogs.js'
import type { TaskInformationResult } from './dialogs.js'
import { createUI } from './ui.js'
import type { UI } from './ui.js'

export interface StatusMessage {
  readonly kind: 'info' | 'error'
  readonly text: string
}

/** Status-line labels for the editable fields (presentation only — the
 * focused-cell edit refusal surfaces WHICH field, the "Cannot indent"
 * precedent). */
const EDIT_FIELD_LABELS: Record<string, string> = {
  taskName: 'Task Name',
  duration: 'Duration',
  start: 'Start',
  finish: 'Finish',
}

export interface HostAppState {
  readonly session: ProjectRendererSession
  readonly viewState: ProjectViewState
  readonly filePath: string | null
  readonly format: HostFileFormat
  /** The document reference at the last save — dirty tracking by identity. */
  readonly savedDocument: ProjectDocument
  readonly diagnostics: readonly ImportDiagnostic[]
  readonly status: StatusMessage
}

export interface AppDependencies {
  readonly bridge: ProjectHostBridge
  readonly root: HTMLElement
  readonly initialDocument?: ProjectDocument
}

/** The app surface consumed by the host entries and the unit tests. */
export interface ProjectHostApp {
  readonly state: HostAppState
  readonly dirty: boolean
  start(): void
  keydown(input: KeyInput): void
  menuCommand(command: MenuCommandId): void
  /** Awaitable dispatch (file actions are async; everything else settles
   * in the same microtask). */
  execute(action: HostAction): Promise<void>
  openPath(path: string, bytes: Uint8Array): void
  /** The close-guard decision flow (used by onCloseRequested + tests). */
  handleCloseRequested(): Promise<void>
}

export function createProjectApp(deps: AppDependencies): ProjectHostApp {
  const { bridge, root } = deps
  const initial = deps.initialDocument ?? newProjectDocument()

  let state: HostAppState = {
    session: createRendererSession(initial, { schedule: scheduleRunner }),
    viewState: createViewState(initial),
    filePath: null,
    format: 'gproj',
    savedDocument: initial,
    diagnostics: [],
    status: { kind: 'info', text: 'Ready' },
  }

  const ui: UI = createUI(root, {
    onRowClick: (taskId, mode) => dispatchIntent({ type: 'selectTask', taskId, mode }),
    onLinkClick: (dependencyId, mode) =>
      dispatchIntent({ type: 'selectDependency', dependencyId, mode }),
    onCellActivate: (taskId, field) => dispatchIntent({ type: 'beginTaskEdit', taskId, field }),
    onLinkActivate: (dependencyId, field) =>
      dispatchIntent({ type: 'beginDependencyEdit', dependencyId, field }),
    onDraftChange: (draft) => {
      if (state.viewState.editing !== undefined) {
        dispatchIntent({ type: 'updateTaskEditDraft', draft })
      } else if (state.viewState.dependencyEditing !== undefined) {
        dispatchIntent({ type: 'updateDependencyEditDraft', draft })
      }
    },
    onScroll: (firstRow) => {
      uiState.firstRow = firstRow
      render()
    },
    onWidthChange: (width) => {
      uiState.timelineWidth = width
      render()
    },
    // The shared ribbon (PROJECT-029): control activation flows through the
    // SAME menu-command path as menu activation — one translation table
    // (translateMenuCommand), one execution pipeline; the ribbon invents
    // no dispatch of its own.
    onRibbonCommand: (command) => menuCommand(command),
  })

  // The shared Task Information dialog (PROJECT-030) — one instance, the
  // controller's only dialog surface. The unsaved-changes dialog is the
  // layer's stateless prompt (created per consultation).
  const taskDialog = createTaskInformationDialog()

  /** Host layout state (pixels): scroll window + pane width. Never view state. */
  const uiState = { firstRow: 0, timelineWidth: 1024 }

  const dirty = (): boolean => state.session.document !== state.savedDocument

  // ---- the render pipeline ----------------------------------------------

  function render(): void {
    const document = state.session.document
    const schedule = state.session.schedule
    const projection = projectDocumentView(document, schedule, state.viewState)
    // The canonical per-render inputs: the injected working-time and
    // allocation authorities (PROJECT-025/026) — the host passes the
    // injections, never an implementation. The resource surface threads
    // the session's CURRENT schedule (never a stale or re-derived one) and
    // is absent when no scheduler produced one — never invented.
    const calendar: CalendarViewInput = { workingTime: workingTimeQuery }
    const resources: ResourceViewInput | undefined =
      schedule === undefined ? undefined : { allocation: allocationQuery, schedule }
    const view = buildGanttView(
      document,
      projection,
      state.viewState,
      {
        firstRow: uiState.firstRow,
        visibleRows: ui.visibleRowCapacity,
        overscan: 2,
      },
      calendar,
      resources,
    )
    const anyEditing =
      state.viewState.editing !== undefined || state.viewState.dependencyEditing !== undefined
    ui.update({
      document,
      projection,
      view,
      viewState: state.viewState,
      filePath: state.filePath,
      format: state.format,
      dirty: dirty(),
      status: state.status,
      diagnostics: state.diagnostics,
      canUndo: canUndoRendererCommand(state.session),
      canRedo: canRedoRendererCommand(state.session),
      editing: anyEditing,
    })
  }

  // ---- intent + command plumbing ----------------------------------------

  function dispatchIntent(intent: ProjectViewIntent): void {
    const next = reduceViewState(state.viewState, intent, {
      document: state.session.document,
      schedule: state.session.schedule,
    })
    if (next === state.viewState) return
    state = { ...state, viewState: next }
    render()
  }

  /** Applies one command through the session; surfaces engine rejections. */
  function applyCommand(command: ProjectCommand): void {
    const outcome = applyRendererCommand(state.session, command)
    if (!outcome.result.accepted) {
      const first = outcome.result.diagnostics[0]
      state = {
        ...state,
        status: {
          kind: 'error',
          text: first === undefined ? 'Command rejected' : `${first.code}: ${first.message}`,
        },
      }
      render()
      return
    }
    state = {
      ...state,
      session: outcome.session,
      viewState: reconcileViewState(state.viewState, outcome.session.document),
      status: { kind: 'info', text: 'Ready' },
    }
    render()
  }

  // ---- document actions (renderer-core builders — never raw commands) ----

  function executeDocumentAction(
    action: 'createTask' | 'deleteSelection' | 'indentSelection' | 'outdentSelection',
  ): void {
    const document = state.session.document
    if (action === 'createTask') {
      // The new task's identity is deterministic — select + focus it after.
      const identity = nextTaskIdentity(document)
      applyCommand(buildCreateTaskCommand(document))
      dispatchIntent({ type: 'selectTask', taskId: identity.id })
      return
    }
    const selected = state.viewState.tasks.taskIds
    if (action === 'deleteSelection') {
      for (const command of buildDeleteSelectionCommands(document, selected)) {
        applyCommand(command)
      }
      if (selected.length === 0) {
        state = { ...state, status: { kind: 'info', text: 'No task selected' } }
        render()
      }
      return
    }
    const target = state.viewState.tasks.focusId ?? selected[selected.length - 1]
    if (target === undefined) {
      state = { ...state, status: { kind: 'info', text: 'No task selected' } }
      render()
      return
    }
    const command =
      action === 'indentSelection'
        ? buildIndentCommand(document, target)
        : buildOutdentCommand(document, target)
    if (command === undefined) {
      state = {
        ...state,
        status: {
          kind: 'info',
          text: action === 'indentSelection' ? 'Cannot indent' : 'Cannot outdent',
        },
      }
      render()
      return
    }
    applyCommand(command)
  }

  // ---- edit flows (the PROJECT-023/024 commit paths) ---------------------

  /** Commits the ACTIVE task edit through the canonical one-call flow and
   * folds the outcome into the controller state. Returns the failure text
   * when the commit did not apply (the cell editor's status line and the
   * Task Information dialog surface the same reason). */
  function commitActiveTaskEdit(): { ok: boolean; reason?: string } {
    const outcome = commitTaskEditThroughSession(state.session, state.viewState)
    state = {
      ...state,
      session: outcome.session,
      viewState: reconcileViewState(outcome.state, outcome.session.document),
    }
    if (outcome.commit.kind === 'invalid') {
      const text = `Invalid edit: ${outcome.commit.reason}`
      state = { ...state, status: { kind: 'error', text } }
      render()
      return { ok: false, reason: text }
    }
    if (outcome.result !== undefined && !outcome.result.accepted) {
      const first = outcome.result.diagnostics[0]
      const text = first === undefined ? 'Command rejected' : `${first.code}: ${first.message}`
      state = { ...state, status: { kind: 'error', text } }
      render()
      return { ok: false, reason: text }
    }
    state = { ...state, status: { kind: 'info', text: 'Ready' } }
    render()
    return { ok: true }
  }

  function executeEditAction(action: 'beginEditFocusedCell' | 'commit' | 'cancel'): void {
    if (action === 'beginEditFocusedCell') {
      const selected = state.viewState.tasks.taskIds
      const target = state.viewState.tasks.focusId ?? selected[selected.length - 1]
      if (target === undefined) {
        state = { ...state, status: { kind: 'info', text: 'No task selected' } }
        render()
        return
      }
      // The focused CELL (PROJECT-031): (focusId, focusField); an absent
      // field is the implicit taskName — the pre-031 Enter/F2 behavior.
      const field = state.viewState.tasks.focusField ?? 'taskName'
      const before = state.viewState
      dispatchIntent({ type: 'beginTaskEdit', taskId: target, field })
      // The 023 editability rule lives in the reducer (a summary row's
      // scheduling fields are derived roll-ups and never begin an edit):
      // its reference-equal no-op is surfaced honestly — the "Cannot
      // indent" precedent, never a silent nothing.
      if (state.viewState === before) {
        state = {
          ...state,
          status: { kind: 'info', text: `${EDIT_FIELD_LABELS[field]} is not editable on this row` },
        }
        render()
      }
      return
    }
    if (action === 'cancel') {
      if (state.viewState.editing !== undefined) dispatchIntent({ type: 'endTaskEdit' })
      else if (state.viewState.dependencyEditing !== undefined) {
        dispatchIntent({ type: 'endDependencyEdit' })
      }
      return
    }
    // commit: the active edit (task or dependency) through its flow.
    if (state.viewState.editing !== undefined) {
      commitActiveTaskEdit()
      return
    }
    if (state.viewState.dependencyEditing !== undefined) {
      const outcome = commitDependencyEditThroughSession(state.session, state.viewState)
      state = {
        ...state,
        session: outcome.session,
        viewState: reconcileViewState(outcome.state, outcome.session.document),
        status:
          outcome.commit.kind === 'invalid'
            ? { kind: 'error', text: `Invalid edit: ${outcome.commit.reason}` }
            : { kind: 'info', text: 'Ready' },
      }
      render()
    }
  }

  // ---- history ------------------------------------------------------------

  function executeHistoryAction(action: 'undo' | 'redo'): void {
    const outcome =
      action === 'undo' ? undoRendererCommand(state.session) : redoRendererCommand(state.session)
    if (!outcome.applied) {
      state = {
        ...state,
        status: { kind: 'info', text: action === 'undo' ? 'Nothing to undo' : 'Nothing to redo' },
      }
      render()
      return
    }
    state = {
      ...state,
      session: outcome.session,
      viewState: reconcileViewState(state.viewState, outcome.session.document),
      status: { kind: 'info', text: 'Ready' },
    }
    render()
  }

  // ---- file flows (the canonical adapters over the bridge transport) -----

  function loadDocument(
    document: ProjectDocument,
    filePath: string | null,
    format: HostFileFormat,
    diagnostics: readonly ImportDiagnostic[],
  ): void {
    // A loaded document ends any open Task Information dialog — its
    // target context is gone (the modal rule's honest completion).
    taskDialog.close()
    const session = createRendererSession(document, { schedule: scheduleRunner })
    state = {
      session,
      viewState: createViewState(document, session.schedule),
      filePath,
      format,
      savedDocument: document,
      diagnostics,
      status: { kind: 'info', text: filePath === null ? 'New project' : `Opened ${filePath}` },
    }
    uiState.firstRow = 0
    render()
  }

  function openPath(path: string, bytes: Uint8Array): void {
    const outcome = importDocumentBytes(path, bytes)
    if (outcome.kind === 'error') {
      state = {
        ...state,
        diagnostics: outcome.diagnostics,
        status: { kind: 'error', text: `Open failed: ${outcome.message}` },
      }
      render()
      return
    }
    loadDocument(
      outcome.imported.document,
      path,
      outcome.imported.format,
      outcome.imported.diagnostics,
    )
  }

  /** A failed host read (oversized/missing/unreadable): the current
   * document survives untouched — the failed read's bytes never existed
   * on this side of the bridge. */
  function failOpen(message: string): void {
    state = { ...state, status: { kind: 'error', text: `Open failed: ${message}` } }
    render()
  }

  /** Saves; returns whether the document is now persisted. */
  async function saveFlow(saveAs: boolean): Promise<boolean> {
    let path = state.filePath
    let format = state.format
    if (saveAs || path === null) {
      const picked = await bridge.pickSaveFile(defaultFileNameFor(state.session.document, format))
      if (picked === null) return false
      path = picked
      format = formatForPath(picked) ?? format
    }
    const exported = exportDocumentBytes(state.session.document, format)
    const write = await bridge.writeFile(path, exported.bytes)
    if (!write.ok) {
      state = {
        ...state,
        status: { kind: 'error', text: `Save failed: ${write.error ?? 'write error'}` },
      }
      render()
      return false
    }
    state = {
      ...state,
      filePath: path,
      format,
      savedDocument: state.session.document,
      diagnostics: exported.diagnostics,
      status: { kind: 'info', text: `Saved ${path}` },
    }
    render()
    return true
  }

  /** The unsaved-changes gate; false = the caller must stop (cancelled). */
  async function confirmUnsaved(): Promise<boolean> {
    if (!dirty()) return true
    // The SHARED dialog (PROJECT-030) — one implementation in both hosts;
    // the pre-030 bridge transport surface is gone.
    const choice = await confirmUnsavedChanges(state.session.document.properties.name)
    if (choice === 'cancel') return false
    if (choice === 'save') return saveFlow(false)
    return true
  }

  async function executeFileAction(action: 'new' | 'open' | 'save' | 'saveAs'): Promise<void> {
    if (action === 'new') {
      if (!(await confirmUnsaved())) return
      loadDocument(newProjectDocument(), null, 'gproj', [])
      return
    }
    if (action === 'open') {
      if (!(await confirmUnsaved())) return
      const selection = await bridge.pickOpenFile()
      if (selection === null) return
      if (!selection.read.ok) {
        failOpen(selection.read.error)
        return
      }
      openPath(selection.path, selection.read.bytes)
      return
    }
    await saveFlow(action === 'saveAs')
  }

  // ---- the shared Task Information dialog (PROJECT-030) -------------------

  /** Runs the dialog's collected drafts through the CANONICAL per-field
   * commit flow — the same begin → draft → `commitTaskEditThroughSession`
   * pipeline the cell editor runs, one field at a time (the single-editor
   * rule preserved by construction: the dialog never holds two view-state
   * edits at once). Each accepted field is ONE semantic command
   * (RenameTask / SetTaskDuration) through the session with journal +
   * reschedule + reconcile — the dialog operates on commands, never
   * beside them. A refused field stops the sequence: fields already
   * applied stay applied (they were real commands — undoable); the
   * dialog surfaces the reason and stays open for the fix or cancel. */
  function commitDialogFields(
    taskId: TaskId,
    request: { name: string; duration: string },
  ): TaskInformationResult {
    const fields: ReadonlyArray<'taskName' | 'duration'> = ['taskName', 'duration']
    for (const field of fields) {
      const draft = field === 'taskName' ? request.name : request.duration
      dispatchIntent({ type: 'beginTaskEdit', taskId, field })
      // A non-editable field (a summary's duration) never begins an edit:
      // the disabled input's unchanged draft is an honest no-op.
      if (state.viewState.editing === undefined) continue
      dispatchIntent({ type: 'updateTaskEditDraft', draft })
      const outcome = commitActiveTaskEdit()
      if (!outcome.ok) return { ok: false, reason: outcome.reason ?? 'Command rejected' }
    }
    state = { ...state, status: { kind: 'info', text: 'Ready' } }
    render()
    return { ok: true }
  }

  /** Opens (or refreshes) the shared Task Information dialog for the
   * focused/selected task — the same target rule indent/outdent use. */
  function openTaskInformation(): void {
    const selected = state.viewState.tasks.taskIds
    const target = state.viewState.tasks.focusId ?? selected[selected.length - 1]
    const task =
      target === undefined
        ? undefined
        : state.session.document.tasks.find((candidate) => candidate.id === target)
    if (target === undefined || task === undefined) {
      state = { ...state, status: { kind: 'info', text: 'No task selected' } }
      render()
      return
    }
    // The dialog supersedes the cell editor: any active edit ends first
    // (its uncommitted draft is discarded — the dialog opens clean).
    if (state.viewState.editing !== undefined) dispatchIntent({ type: 'endTaskEdit' })
    else if (state.viewState.dependencyEditing !== undefined) {
      dispatchIntent({ type: 'endDependencyEdit' })
    }
    const document = state.session.document
    const schedule = state.session.schedule
    taskDialog.open(
      {
        taskId: task.id,
        name: initialTaskEditDraft(document, schedule, task.id, 'taskName'),
        duration: initialTaskEditDraft(document, schedule, task.id, 'duration'),
        durationEditable: isTaskFieldEditable(task, 'duration'),
        start: initialTaskEditDraft(document, schedule, task.id, 'start'),
        finish: initialTaskEditDraft(document, schedule, task.id, 'finish'),
      },
      {
        onCommit: (request) => commitDialogFields(task.id, request),
        onCancelled: () => {},
      },
    )
  }

  // ---- action dispatch -----------------------------------------------------

  async function execute(action: HostAction): Promise<void> {
    switch (action.kind) {
      case 'none':
        return
      case 'intent':
        dispatchIntent(action.intent)
        return
      case 'document':
        executeDocumentAction(action.action)
        return
      case 'edit':
        executeEditAction(action.action)
        return
      case 'history':
        executeHistoryAction(action.action)
        return
      case 'file':
        await executeFileAction(action.action)
        return
      case 'dialog':
        openTaskInformation()
        return
      case 'view': {
        const taskIds = state.viewState.tasks.taskIds
        dispatchIntent({
          type: 'setCollapsed',
          taskIds,
          collapsed: action.action === 'collapseSelection',
        })
        return
      }
    }
  }

  function keydown(input: KeyInput): void {
    // The open Task Information dialog owns the keyboard (the menu bar's
    // open-dropdown precedent): the app's translation path is suspended
    // until it closes. The dialog's own listener handles Enter/Escape.
    if (taskDialog.isOpen()) return
    const editing =
      state.viewState.editing !== undefined || state.viewState.dependencyEditing !== undefined
    void execute(translateKeyDown(input, { editing }))
  }

  function menuCommand(command: MenuCommandId): void {
    // The open Task Information dialog is modal to the app's COMMAND
    // surface: host chrome activation is suspended — except the dialog's
    // own open command, which refreshes it from the CURRENT document (the
    // honest re-open). The close handshake is NOT a command: a host close
    // request still runs its full save/discard/cancel flow while the
    // dialog is open (a window close is never blockable by an in-page
    // dialog), and an external open request (argv/drag-drop) loads and
    // closes the dialog (loadDocument).
    if (taskDialog.isOpen() && command !== 'task.information') return
    void execute(translateMenuCommand(command))
  }

  async function handleCloseRequested(): Promise<void> {
    if (!dirty()) {
      bridge.approveClose()
      return
    }
    const choice = await confirmUnsavedChanges(state.session.document.properties.name)
    if (choice === 'cancel') return
    if (choice === 'save') {
      const saved = await saveFlow(false)
      if (!saved) return
    }
    bridge.approveClose()
  }

  function start(): void {
    bridge.onMenuCommand(menuCommand)
    bridge.onCloseRequested(() => void handleCloseRequested())
    bridge.onOpenRequested(async (path) => {
      if (!(await confirmUnsaved())) return
      // Read errors are values (the bridge contract): an oversized,
      // missing, or unreadable path surfaces the transport error and the
      // current document survives — uncapped bytes never arrive here.
      const read = await bridge.readFile(path)
      if (!read.ok) {
        failOpen(read.error)
        return
      }
      openPath(path, read.bytes)
    })
    render()
  }

  return {
    get state() {
      return state
    },
    get dirty() {
      return dirty()
    },
    start,
    keydown,
    menuCommand,
    execute,
    openPath,
    handleCloseRequested,
  }
}
