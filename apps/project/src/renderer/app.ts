/**
 * PROJECT-027 — the Project desktop host controller.
 *
 * Binds the shared renderer core to a desktop window: one
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
 * pixel. The desktop bridge is injected (the preload surface in
 * production, an in-memory fake in tests).
 */
import type {
  ImportDiagnostic,
  ProjectCommand,
  ProjectDocument,
} from '@genoffice/project-contracts'
import {
  canRedoRendererCommand,
  canUndoRendererCommand,
  commitDependencyEditThroughSession,
  commitTaskEditThroughSession,
  createRendererSession,
  createViewState,
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
import type { MenuCommandId, ProjectDesktopBridge } from '../shared/ipc.js'
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
import { createUI } from './ui.js'
import type { UI } from './ui.js'

export interface StatusMessage {
  readonly kind: 'info' | 'error'
  readonly text: string
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
  readonly bridge: ProjectDesktopBridge
  readonly root: HTMLElement
  readonly initialDocument?: ProjectDocument
}

/** The app surface consumed by the renderer entry and the unit tests. */
export interface ProjectDesktopApp {
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

export function createProjectDesktopApp(deps: AppDependencies): ProjectDesktopApp {
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
  })

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

  function executeEditAction(action: 'beginEditName' | 'commit' | 'cancel'): void {
    if (action === 'beginEditName') {
      const target = state.viewState.tasks.focusId
      if (target === undefined) {
        state = { ...state, status: { kind: 'info', text: 'No task selected' } }
        render()
        return
      }
      dispatchIntent({ type: 'beginTaskEdit', taskId: target, field: 'taskName' })
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
      const outcome = commitTaskEditThroughSession(state.session, state.viewState)
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
    const choice = await bridge.confirmDiscard(state.session.document.properties.name)
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
      openPath(selection.path, selection.bytes)
      return
    }
    await saveFlow(action === 'saveAs')
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
    const editing =
      state.viewState.editing !== undefined || state.viewState.dependencyEditing !== undefined
    void execute(translateKeyDown(input, { editing }))
  }

  function menuCommand(command: MenuCommandId): void {
    void execute(translateMenuCommand(command))
  }

  async function handleCloseRequested(): Promise<void> {
    if (!dirty()) {
      bridge.approveClose()
      return
    }
    const choice = await bridge.confirmDiscard(state.session.document.properties.name)
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
      try {
        const bytes = await bridge.readFile(path)
        openPath(path, bytes)
      } catch (error) {
        state = {
          ...state,
          status: { kind: 'error', text: `Open failed: ${String(error)}` },
        }
        render()
      }
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
