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
 * keeps its own keys. The displayed strings come from the ONE shared menu
 * presentation table (PROJECT-031: `menuAcceleratorFor` — the menu bar's
 * displayed forms, one source); the ribbon keeps its own control LABELS
 * (a ribbon is not a menu), but its displayed accelerators are the
 * menu's.
 */
import type { MenuCommandId } from './bridge.js'
import { menuAcceleratorFor } from './menu-presentation.js'

/** One ribbon control: a shared command id + its label. The tooltip's
 * displayed accelerator is the shared menu presentation's (never a
 * ribbon-private string). */
export interface RibbonControlSpec {
  readonly id: MenuCommandId
  readonly label: string
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
 * The shared ribbon structure — the canonical menu vocabulary (all 16
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
          { id: 'task.create', label: 'New Task' },
          { id: 'task.information', label: 'Task Information' },
          { id: 'edit.deleteTask', label: 'Delete Task' },
        ],
      },
      {
        label: 'Schedule',
        controls: [
          { id: 'task.indent', label: 'Indent Task' },
          { id: 'task.outdent', label: 'Outdent Task' },
        ],
      },
      {
        label: 'History',
        controls: [
          { id: 'edit.undo', label: 'Undo' },
          { id: 'edit.redo', label: 'Redo' },
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
          { id: 'view.zoomIn', label: 'Zoom In' },
          { id: 'view.zoomOut', label: 'Zoom Out' },
          { id: 'view.fit', label: 'Fit to Project' },
        ],
      },
      {
        label: 'Outline',
        controls: [
          { id: 'view.collapse', label: 'Collapse Selected' },
          { id: 'view.expand', label: 'Expand Selected' },
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
          { id: 'file.new', label: 'New Project' },
          { id: 'file.open', label: 'Open Project…' },
          { id: 'file.save', label: 'Save' },
          { id: 'file.saveAs', label: 'Save As…' },
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
 * menu-bar dropdown precedent) with ONE echo-driven exception — the
 * context-sensitive promotion (PROJECT-031): the appearance of a selection
 * (the hasSelection echo's false → true transition) promotes the Task tab,
 * the tab whose controls address the selection; no demotion fires when the
 * selection disappears, and a manual tab choice holds while the context
 * stays. `update` reflects the echoes + the promotion transition only.
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
        // The displayed accelerator is the shared MENU presentation's form
        // (PROJECT-031 — one source; the ribbon owns no accelerator
        // string of its own).
        const accelerator = menuAcceleratorFor(control.id)
        button.title =
          accelerator === undefined ? control.label : `${control.label} (${accelerator})`
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

    tabButton.addEventListener('click', () => selectTab(tab.id))
  }

  /** Selects one tab (the single tab-switching site: manual clicks and the
   * context-sensitive promotion below share it). */
  function selectTab(tabId: RibbonTabSpec['id']): void {
    for (const other of RIBBON_TABS) {
      const otherPanel = panels.querySelector(`[data-tab="${other.id}"]`)
      if (otherPanel instanceof HTMLElement) otherPanel.hidden = other.id !== tabId
      const otherTab = tabStrip.querySelector(`[data-tab="${other.id}"]`)
      if (otherTab instanceof HTMLElement) {
        otherTab.setAttribute('aria-selected', String(other.id === tabId))
      }
    }
  }

  // The previous hasSelection echo — the context-sensitive promotion fires
  // only on the context's APPEARANCE (the false → true transition).
  let lastHasSelection = false

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
    // Context-sensitive tab promotion (PROJECT-031 — the scope the 029
    // discipline deferred): when the selection context APPEARS, the Task
    // tab — the tab whose controls address the selection — is promoted
    // (the Office contextual convention). No demotion when the context
    // disappears, and a manual tab choice holds while the context stays:
    // promotion fires ONLY on the transition, so it never fights the user.
    if (state.hasSelection && !lastHasSelection) selectTab('task')
    lastHasSelection = state.hasSelection
  }

  return { update }
}
