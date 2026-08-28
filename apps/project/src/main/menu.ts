/**
 * PROJECT-027 — the Project desktop host's native menu.
 *
 * Transport only: every item carries a `MenuCommandId` and forwards the
 * activation to the focused window's renderer. No item interprets Project
 * state (labels are static; enable/disable is not document-derived here).
 *
 * PROJECT-031 — presentation parity: this table MIRRORS the ONE shared
 * menu presentation table (`MENU_PRESENTATION` in
 * `@genoffice/project-host`, which the main process may not import — the
 * frozen native-transport boundary): the labels are exactly the shared
 * table's, and each native accelerator equals the shared table's display
 * string under the one documented display→native rule (`Ctrl+` ⇔
 * `CmdOrCtrl+`; every other form is identical in both conventions). The
 * desktop discipline suite pins the mirror equality; the shared table is
 * the single source (the web menu bar renders it directly and the shared
 * ribbon's tooltips display its accelerator strings). `edit.redo` displays
 * `CmdOrCtrl+Y` — the canonical shared display form `Ctrl+Y` (the shared
 * keyboard table binds BOTH Ctrl+Y and Ctrl+Shift+Z; only the display
 * converges across hosts).
 *
 * Every accelerator is DISPLAYED but deliberately NOT registered
 * (`registerAccelerator: false`): the renderer's keyboard translation
 * (the shared `@genoffice/project-host` translation table) is the single
 * execution path for accelerator
 * keys, so an active cell editor keeps its own keys (native text undo, caret
 * movement, text deletion) instead of menu commands firing mid-edit. Menu
 * clicks still arrive through the menu-command channel.
 */
import type { WebContents } from 'electron'
import { PROJECT_IPC } from '../shared/ipc.js'
import type { MenuCommandId } from '../shared/ipc.js'

export interface MenuCommandSpec {
  readonly id: MenuCommandId
  readonly label: string
  readonly accelerator?: string
}

const fileCommands: readonly MenuCommandSpec[] = [
  { id: 'file.new', label: 'New Project', accelerator: 'CmdOrCtrl+N' },
  { id: 'file.open', label: 'Open Project…', accelerator: 'CmdOrCtrl+O' },
  { id: 'file.save', label: 'Save', accelerator: 'CmdOrCtrl+S' },
  { id: 'file.saveAs', label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S' },
]

const editCommands: readonly MenuCommandSpec[] = [
  { id: 'edit.undo', label: 'Undo', accelerator: 'CmdOrCtrl+Z' },
  { id: 'edit.redo', label: 'Redo', accelerator: 'CmdOrCtrl+Y' },
  { id: 'edit.deleteTask', label: 'Delete Task', accelerator: 'Delete' },
]

const taskCommands: readonly MenuCommandSpec[] = [
  { id: 'task.create', label: 'New Task', accelerator: 'Insert' },
  // Opens the shared Task Information dialog (PROJECT-030). No displayed
  // accelerator: the shared keyboard table owns execution and binds no key
  // to it — menu/ribbon activation are the firing surfaces.
  { id: 'task.information', label: 'Task Information…' },
  { id: 'task.indent', label: 'Indent Task', accelerator: 'Alt+Shift+Right' },
  { id: 'task.outdent', label: 'Outdent Task', accelerator: 'Alt+Shift+Left' },
]

const viewCommands: readonly MenuCommandSpec[] = [
  { id: 'view.zoomIn', label: 'Zoom In', accelerator: 'CmdOrCtrl+=' },
  { id: 'view.zoomOut', label: 'Zoom Out', accelerator: 'CmdOrCtrl+-' },
  { id: 'view.fit', label: 'Fit to Project', accelerator: 'CmdOrCtrl+Shift+F' },
  { id: 'view.collapse', label: 'Collapse Selected', accelerator: 'Alt+Shift+Minus' },
  { id: 'view.expand', label: 'Expand Selected', accelerator: 'Alt+Shift+Plus' },
]

/**
 * The native menu's section structure — the desktop MIRROR of the shared
 * menu presentation table (PROJECT-031). Exported for the discipline
 * suite's parity pin (labels exactly the shared table's; native
 * accelerators = the shared display strings under the `Ctrl+` ⇔
 * `CmdOrCtrl+` rule).
 */
export const PROJECT_MENU_SECTIONS: ReadonlyArray<{
  readonly label: string
  readonly items: readonly MenuCommandSpec[]
}> = [
  { label: 'File', items: fileCommands },
  { label: 'Edit', items: editCommands },
  { label: 'Task', items: taskCommands },
  { label: 'View', items: viewCommands },
]

const itemFor = (spec: MenuCommandSpec, send: (command: MenuCommandId) => void) => ({
  id: spec.id,
  label: spec.label,
  ...(spec.accelerator !== undefined ? { accelerator: spec.accelerator } : {}),
  // Displayed, never registered: the renderer keyboard path owns execution.
  registerAccelerator: false,
  click: () => send(spec.id),
})

/** Forwards a menu command to a window (the menu click path). */
export function sendMenuCommand(webContents: WebContents, command: MenuCommandId): void {
  webContents.send(PROJECT_IPC.menuCommand, command)
}

/**
 * Builds the application menu template. `send` is the activation sink — in
 * production it forwards to the focused window; the template stays pure so
 * the menu construction is trivially auditable. The sections come from the
 * exported `PROJECT_MENU_SECTIONS` mirror (the shared presentation table's
 * structure — one source of truth for the section order).
 */
export function projectMenuTemplate(send: (command: MenuCommandId) => void) {
  return PROJECT_MENU_SECTIONS.map((section) => ({
    label: section.label,
    submenu: section.items.map((spec) => itemFor(spec, send)),
  }))
}

/** The menu command ids in template order (used by the discipline suite). */
export const MENU_COMMAND_IDS: readonly MenuCommandId[] = PROJECT_MENU_SECTIONS.flatMap((section) =>
  section.items.map((spec) => spec.id),
)
