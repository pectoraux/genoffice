/**
 * PROJECT-031 — the shared menu PRESENTATION table.
 *
 * The ONE source for how the shared command vocabulary is PRESENTED in the
 * host menus: section structure (File → Edit → Task → View), item labels,
 * and the accelerator DISPLAY strings. Presentation only — every id is a
 * transport `MenuCommandId` from the shared bridge contract and executes
 * nothing here (the shared translation table owns execution, the menus'
 * single-translation-path discipline).
 *
 * Consumers (the presentation-parity lockstep — the PROJECT-028 scope note
 * deferred to this increment):
 *
 * - the WEB menu bar (`apps/project-web/src/menu.ts`) renders this table
 *   directly — the web shell keeps NO label/accelerator literal of its own;
 * - the DESKTOP native menu (`apps/project/src/main/menu.ts`) mirrors this
 *   table in the main process (main/preload may import NO `@genoffice/*`
 *   package — the frozen transport boundary), and the desktop discipline
 *   suite pins the mirror EQUAL to this table: labels exactly, and each
 *   native accelerator equal to this table's display string under the one
 *   documented display→native rule (`Ctrl+` ⇔ `CmdOrCtrl+`; every other
 *   form — `Alt+…`, `Delete`, `Insert` — is identical in both conventions).
 *
 * The canonical display form of `edit.redo` is `Ctrl+Y` (the form the web
 * menu and the shared ribbon already displayed, and one of the two bindings
 * the shared keyboard table owns — Ctrl+Shift+Z remains bound, only the
 * DISPLAY converges; pre-031 the desktop menu displayed CmdOrCtrl+Shift+Z
 * while the web menu displayed Ctrl+Y, a cross-host presentation defect
 * this table eliminates).
 *
 * The shared ribbon displays these accelerator strings in its tooltips (the
 * PROJECT-029 rule: "the menu bar's displayed forms") through
 * `menuAcceleratorFor` — the ribbon keeps its own control LABELS (a ribbon
 * is not a menu; the two surfaces' label conventions differ legitimately),
 * but its displayed accelerators are the menu's, one source.
 *
 * Determinism: the table is build-time constant data — no wall clock, no
 * randomness, no date computation (the generic discipline scans cover it).
 */
import type { MenuCommandId } from './bridge.js'

/** One menu item's presentation: the shared command id, its label, and the
 * canonical accelerator DISPLAY string (the shared `Ctrl…` form; absent
 * when the shared keyboard table binds no key to the command — the
 * `task.information` precedent). */
export interface MenuPresentationItem {
  readonly id: MenuCommandId
  readonly label: string
  readonly accelerator?: string
}

/** One menu section (top-level menu): its label and items. */
export interface MenuPresentationSection {
  readonly label: string
  readonly items: readonly MenuPresentationItem[]
}

/**
 * The shared menu presentation — the complete command vocabulary (all 16
 * ids) in canonical menu order. The vocabulary lockstep (this table's ids =
 * `MENU_COMMAND_IDS` exactly) is pinned by the discipline suite.
 */
export const MENU_PRESENTATION: readonly MenuPresentationSection[] = [
  {
    label: 'File',
    items: [
      { id: 'file.new', label: 'New Project', accelerator: 'Ctrl+N' },
      { id: 'file.open', label: 'Open Project…', accelerator: 'Ctrl+O' },
      { id: 'file.save', label: 'Save', accelerator: 'Ctrl+S' },
      { id: 'file.saveAs', label: 'Save As…', accelerator: 'Ctrl+Shift+S' },
    ],
  },
  {
    label: 'Edit',
    items: [
      { id: 'edit.undo', label: 'Undo', accelerator: 'Ctrl+Z' },
      { id: 'edit.redo', label: 'Redo', accelerator: 'Ctrl+Y' },
      { id: 'edit.deleteTask', label: 'Delete Task', accelerator: 'Delete' },
    ],
  },
  {
    label: 'Task',
    items: [
      { id: 'task.create', label: 'New Task', accelerator: 'Insert' },
      // Opens the shared Task Information dialog (PROJECT-030). No
      // displayed accelerator — the shared keyboard table binds no key to
      // it; menu/ribbon activation are the firing surfaces.
      { id: 'task.information', label: 'Task Information…' },
      { id: 'task.indent', label: 'Indent Task', accelerator: 'Alt+Shift+Right' },
      { id: 'task.outdent', label: 'Outdent Task', accelerator: 'Alt+Shift+Left' },
    ],
  },
  {
    label: 'View',
    items: [
      { id: 'view.zoomIn', label: 'Zoom In', accelerator: 'Ctrl+=' },
      { id: 'view.zoomOut', label: 'Zoom Out', accelerator: 'Ctrl+-' },
      { id: 'view.fit', label: 'Fit to Project', accelerator: 'Ctrl+Shift+F' },
      { id: 'view.collapse', label: 'Collapse Selected', accelerator: 'Alt+Shift+Minus' },
      { id: 'view.expand', label: 'Expand Selected', accelerator: 'Alt+Shift+Plus' },
    ],
  },
]

/** The presentation table's command ids in canonical menu order (the
 * lockstep pin against `MENU_COMMAND_IDS`). */
export const MENU_PRESENTATION_COMMAND_IDS: readonly MenuCommandId[] = MENU_PRESENTATION.flatMap(
  (section) => section.items.map((item) => item.id),
)

const LABEL_BY_COMMAND: ReadonlyMap<MenuCommandId, string> = new Map(
  MENU_PRESENTATION.flatMap((section) => section.items.map((item) => [item.id, item.label])),
)

const ACCELERATOR_BY_COMMAND: ReadonlyMap<MenuCommandId, string> = new Map(
  MENU_PRESENTATION.flatMap((section) => itemAcceleratorEntries(section)),
)

function itemAcceleratorEntries(section: MenuPresentationSection): Array<[MenuCommandId, string]> {
  const entries: Array<[MenuCommandId, string]> = []
  for (const item of section.items) {
    if (item.accelerator !== undefined) entries.push([item.id, item.accelerator])
  }
  return entries
}

/** The canonical menu label of one shared command (every id has one — the
 * vocabulary lockstep guarantees it). */
export function menuLabelFor(command: MenuCommandId): string {
  return LABEL_BY_COMMAND.get(command) ?? command
}

/** The canonical accelerator DISPLAY string of one shared command, or
 * `undefined` when the shared keyboard table binds no key to it (the
 * displayed-accelerator rule: the display never invents a binding the
 * shared translation table does not own). */
export function menuAcceleratorFor(command: MenuCommandId): string | undefined {
  return ACCELERATOR_BY_COMMAND.get(command)
}
