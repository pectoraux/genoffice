/**
 * PROJECT-029 — the shared Project ribbon.
 *
 * The tabbed command surface BOTH hosts render — implemented ONCE in the
 * shared host binding (`@genoffice/project-host`), mounted by the shared
 * DOM layer (`ui.ts`) above the workspace, identical on desktop (Electron
 * renderer) and web (browser) by construction: the shells add NO ribbon
 * code of their own (the discipline suites pin it).
 *
 * Transport discipline — the exact rule the menus follow: every control
 * carries a `MenuCommandId` from the shared contract vocabulary and
 * activation forwards through the injected callback, which the controller
 * wires to its `menuCommand` path — the SAME `translateMenuCommand` table
 * the native menu, the DOM menu bar, and the keyboard path flow through
 * (one translation; the ribbon invents no Project semantics, no second
 * dispatch, no command construction).
 *
 * The ribbon reflects FOUR presentation echoes and nothing else —
 * `canUndo` / `canRedo` (the journal's honest flags), `dirty`
 * (the document-identity probe), `hasSelection` (the view state's task
 * selection) — each consumed as a plain boolean the controller's update
 * pipeline already computes. The ribbon reads no canonical state at all:
 * a disabled control is presentation, never a semantic
 * verdict (the controller remains the authority — an always-enabled
 * control with nothing selected surfaces the controller's status answer,
 * exactly like the same menu command).
 *
 * Accelerators are DISPLAYED (tooltip parity with the menu bar) and
 * execute nothing here — the shared keyboard translation table is the
 * single execution path for accelerator keys, so an active cell editor
 * keeps its own keys.
 */
import type { MenuCommandId } from './bridge.js'

/** One ribbon control: a shared command id + its static presentation. */
export interface RibbonControlSpec {
  readonly id: MenuCommandId
  readonly label: string
  /** The accelerator display (shared `Ctrl…` form — presentation only). */
  readonly accelerator?: string
}

/** One ribbon group: a labeled cluster of controls. */
export interface RibbonGroupSpec {
  readonly label: string
  readonly controls: readonly RibbonControlSpec[]
}

/** One ribbon tab: a panel of groups. */
export interface RibbonTabSpec {
  readonly id: 'task' | 'view' | 'file'
  readonly label: string
  readonly groups: readonly RibbonGroupSpec[]
}

/**
 * The shared ribbon structure — the canonical menu vocabulary (all 15
 * commands) arranged in the ribbon's tab→group→control shape. Tab order
 * follows the Project convention (the semantic Task tab first); the
 * vocabulary lockstep (ribbon ids = `MENU_COMMAND_IDS` exactly) is pinned
 * by the discipline suite.
 */
export const RIBBON_TABS: readonly RibbonTabSpec[] = [
  {
    id: 'task',
    label: 'Task',
    groups: [
      {
        label: 'Tasks',
        controls: [
          { id: 'task.create', label: 'New Task', accelerator: 'Insert' },
          { id: 'edit.deleteTask', label: 'Delete Task', accelerator: 'Delete' },
        ],
      },
      {
        label: 'Schedule',
        controls: [
          { id: 'task.indent', label: 'Indent Task', accelerator: 'Alt+Shift+Right' },
          { id: 'task.outdent', label: 'Outdent Task', accelerator: 'Alt+Shift+Left' },
        ],
      },
      {
        label: 'History',
        controls: [
          { id: 'edit.undo', label: 'Undo', accelerator: 'Ctrl+Z' },
          { id: 'edit.redo', label: 'Redo', accelerator: 'Ctrl+Y' },
        ],
      },
    ],
  },
  {
    id: 'view',
    label: 'View',
    groups: [
      {
        label: 'Zoom',
        controls: [
          { id: 'view.zoomIn', label: 'Zoom In', accelerator: 'Ctrl+=' },
          { id: 'view.zoomOut', label: 'Zoom Out', accelerator: 'Ctrl+-' },
          { id: 'view.fit', label: 'Fit to Project', accelerator: 'Ctrl+Shift+F' },
        ],
      },
      {
        label: 'Outline',
        controls: [
          { id: 'view.collapse', label: 'Collapse Selected', accelerator: 'Alt+Shift+Minus' },
          { id: 'view.expand', label: 'Expand Selected', accelerator: 'Alt+Shift+Plus' },
        ],
      },
    ],
  },
  {
    id: 'file',
    label: 'File',
    groups: [
      {
        label: 'Document',
        controls: [
          { id: 'file.new', label: 'New Project', accelerator: 'Ctrl+N' },
          { id: 'file.open', label: 'Open Project…', accelerator: 'Ctrl+O' },
          { id: 'file.save', label: 'Save', accelerator: 'Ctrl+S' },
          { id: 'file.saveAs', label: 'Save As…', accelerator: 'Ctrl+Shift+S' },
        ],
      },
    ],
  },
]

/** The ribbon's command vocabulary in ribbon order (used by the
 * discipline suite's lockstep against `MENU_COMMAND_IDS`). */
export const RIBBON_COMMAND_IDS: readonly MenuCommandId[] = RIBBON_TABS.flatMap((tab) =>
  tab.groups.flatMap((group) => group.controls.map((control) => control.id)),
)

/** The four presentation echoes the ribbon reflects (never semantics). */
export interface RibbonState {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly dirty: boolean
  readonly hasSelection: boolean
}

export interface RibbonCallbacks {
  /** Control activation — the controller routes it through the SAME
   * `translateMenuCommand` path as menu activation (one translation). */
  onCommand(command: MenuCommandId): void
}

export interface Ribbon {
  /** Reflects the four echoes onto the controls (enablement + dirty). */
  update(state: RibbonState): void
}

const el = (tag: string, className?: string): HTMLElement => {
  const node = document.createElement(tag)
  if (className !== undefined) node.className = className
  return node
}

/** The presentation-only enablement rule: which echoes gate which command.
 * Every command not listed is always enabled. */
const ENABLED_BY: Partial<Record<MenuCommandId, keyof RibbonState>> = {
  'edit.undo': 'canUndo',
  'edit.redo': 'canRedo',
  'edit.deleteTask': 'hasSelection',
  'task.indent': 'hasSelection',
  'task.outdent': 'hasSelection',
}

/**
 * Builds the ribbon into `container`. The tab strip and all three panels
 * are built once; tab selection is ribbon-owned presentation state (the
 * menu-bar dropdown precedent); `update` reflects the echoes only.
 */
export function createRibbon(container: HTMLElement, callbacks: RibbonCallbacks): Ribbon {
  container.innerHTML = ''
  const root = el('div', 'gp-ribbon')
  root.dataset.testid = 'ribbon'
  root.setAttribute('aria-label', 'Project ribbon')

  const tabStrip = el('div', 'gp-ribbon-tabs')
  tabStrip.dataset.testid = 'ribbon-tabs'
  tabStrip.setAttribute('role', 'tablist')
  tabStrip.setAttribute('aria-label', 'Project ribbon tabs')

  const panels = el('div', 'gp-ribbon-panels')
  const buttons = new Map<MenuCommandId, HTMLButtonElement>()
  let dirtyMarker: HTMLElement | null = null

  for (const tab of RIBBON_TABS) {
    const tabButton = document.createElement('button')
    tabButton.type = 'button'
    tabButton.className = 'gp-ribbon-tab'
    tabButton.dataset.testid = 'ribbon-tab'
    tabButton.dataset.tab = tab.id
    tabButton.setAttribute('role', 'tab')
    tabButton.id = `gp-ribbon-tab-${tab.id}`
    tabButton.setAttribute('aria-controls', `gp-ribbon-panel-${tab.id}`)
    tabButton.setAttribute('aria-selected', String(tab.id === 'task'))
    tabButton.textContent = tab.label
    tabStrip.appendChild(tabButton)

    const panel = el('div', 'gp-ribbon-panel')
    panel.dataset.testid = 'ribbon-panel'
    panel.dataset.tab = tab.id
    panel.id = `gp-ribbon-panel-${tab.id}`
    panel.setAttribute('role', 'tabpanel')
    panel.setAttribute('aria-labelledby', `gp-ribbon-tab-${tab.id}`)
    if (tab.id !== 'task') panel.hidden = true

    for (const group of tab.groups) {
      const groupEl = el('div', 'gp-ribbon-group')
      groupEl.dataset.testid = 'ribbon-group'
      groupEl.dataset.group = group.label
      const controls = el('div', 'gp-ribbon-controls')
      for (const control of group.controls) {
        const button = document.createElement('button')
        button.type = 'button'
        button.className = 'gp-ribbon-button'
        button.dataset.testid = 'ribbon-button'
        button.dataset.command = control.id
        button.title =
          control.accelerator === undefined
            ? control.label
            : `${control.label} (${control.accelerator})`
        const label = el('span', 'gp-ribbon-button-label')
        label.textContent = control.label
        button.appendChild(label)
        if (control.id === 'file.save') {
          const marker = el('span', 'gp-ribbon-dirty')
          marker.dataset.testid = 'ribbon-dirty-indicator'
          marker.dataset.dirty = 'false'
          marker.textContent = ''
          button.appendChild(marker)
          dirtyMarker = marker
        }
        // The disabled property is the gate (native activation is blocked);
        // the handler re-checks it so dispatched synthetic events (tests)
        // can never bypass a disabled control either.
        button.addEventListener('click', () => {
          if (button.disabled) return
          callbacks.onCommand(control.id)
        })
        buttons.set(control.id, button)
        controls.appendChild(button)
      }
      const groupLabel = el('span', 'gp-ribbon-group-label')
      groupLabel.textContent = group.label
      groupEl.append(controls, groupLabel)
      panel.appendChild(groupEl)
    }
    panels.appendChild(panel)

    tabButton.addEventListener('click', () => {
      for (const other of RIBBON_TABS) {
        const otherPanel = panels.querySelector(`[data-tab="${other.id}"]`)
        if (otherPanel instanceof HTMLElement) otherPanel.hidden = other.id !== tab.id
        const otherTab = tabStrip.querySelector(`[data-tab="${other.id}"]`)
        if (otherTab instanceof HTMLElement) {
          otherTab.setAttribute('aria-selected', String(other.id === tab.id))
        }
      }
    })
  }

  root.append(tabStrip, panels)
  container.appendChild(root)

  function update(state: RibbonState): void {
    for (const [command, button] of buttons) {
      const gate = ENABLED_BY[command]
      const enabled = gate === undefined ? true : state[gate]
      button.disabled = !enabled
      button.setAttribute('aria-disabled', String(!enabled))
    }
    if (dirtyMarker !== null) {
      dirtyMarker.dataset.dirty = String(state.dirty)
      dirtyMarker.textContent = state.dirty ? '●' : ''
    }
  }

  return { update }
}
