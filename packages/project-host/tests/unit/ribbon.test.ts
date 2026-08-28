/**
 * The shared ribbon battery (PROJECT-029).
 *
 * Two layers:
 *
 * 1. the ribbon module directly (`createRibbon` on a bare container) —
 *    structure (the tab→group→control shape over the complete shared
 *    command vocabulary), tab switching, activation forwarding, the four
 *    presentation echoes (enablement + dirty), and build determinism;
 * 2. the REAL controller integration (`createProjectApp` on jsdom + the
 *    in-memory bridge fake) — the ribbon mounted by the shared DOM layer,
 *    controls driving the REAL engine paths (the renderer-core builders →
 *    engine → scheduler) and the REAL save paths (the canonical adapter
 *    export over the bridge), the one-translation parity with menu
 *    activation, and the echo updates through real app state changes.
 */
import { beforeEach, describe, expect, it } from 'vitest'
import { fitViewport } from '@genoffice/project-renderer-core'
import { MENU_COMMAND_IDS } from '../../src/bridge.js'
import type { MenuCommandId } from '../../src/bridge.js'
import { createRibbon } from '../../src/ribbon.js'
import { RIBBON_COMMAND_IDS, RIBBON_TABS } from '../../src/ribbon.js'
import type { RibbonState } from '../../src/ribbon.js'
import { createProjectApp } from '../../src/app.js'
import type { ProjectHostBridge } from '../../src/bridge.js'

// ---- shared helpers ---------------------------------------------------------

const ALL_OFF: RibbonState = { canUndo: false, canRedo: false, dirty: false, hasSelection: false }
const ALL_ON: RibbonState = { canUndo: true, canRedo: true, dirty: true, hasSelection: true }

const mount = (): HTMLElement => {
  const host = document.createElement('div')
  document.body.appendChild(host)
  return host
}

const ribbonButtons = (): HTMLElement[] =>
  [...document.querySelectorAll('[data-testid="ribbon-button"]')] as HTMLElement[]

const ribbonButton = (command: MenuCommandId): HTMLButtonElement =>
  document.querySelector(
    `[data-testid="ribbon-button"][data-command="${command}"]`,
  ) as HTMLButtonElement

const ribbonTab = (tab: string): HTMLButtonElement =>
  document.querySelector(`[data-testid="ribbon-tab"][data-tab="${tab}"]`) as HTMLButtonElement

const ribbonPanel = (tab: string): HTMLElement =>
  document.querySelector(`[data-testid="ribbon-panel"][data-tab="${tab}"]`) as HTMLElement

const click = (command: MenuCommandId): void => {
  ribbonButton(command).click()
}

const switchTab = (tab: string): void => {
  ribbonTab(tab).click()
}

/** The in-memory bridge fake (the app.test.ts harness shape). */
function fakeBridge() {
  const files = new Map<string, Uint8Array>()
  const calls = {
    pickSaveFile: 0,
    writeFile: [] as string[],
  }
  const bridge: ProjectHostBridge = {
    pickOpenFile: async () => null,
    pickSaveFile: async () => {
      calls.pickSaveFile += 1
      return 'saved.gproj'
    },
    readFile: async () => ({ ok: false, error: 'ENOENT' }),
    writeFile: async (path, bytes) => {
      calls.writeFile.push(path)
      files.set(path, bytes)
      return { ok: true }
    },
    appInfo: async () => ({ platform: 'test', version: '0.1.0' }),
    onMenuCommand: () => {},
    onCloseRequested: () => {},
    approveClose: () => {},
    onOpenRequested: () => {},
  }
  return { bridge, files, calls }
}

beforeEach(() => {
  document.body.innerHTML = ''
})

// ---- layer 1: the ribbon module ---------------------------------------------

describe('the ribbon vocabulary (transport discipline)', () => {
  it('carries every shared menu command exactly once (the complete 16)', () => {
    expect(new Set(RIBBON_COMMAND_IDS).size).toBe(RIBBON_COMMAND_IDS.length)
    expect([...RIBBON_COMMAND_IDS].sort()).toEqual([...MENU_COMMAND_IDS].sort())
    expect(RIBBON_COMMAND_IDS).toHaveLength(16)
  })

  it('structures the vocabulary as tabs → groups → controls', () => {
    expect(RIBBON_TABS.map((tab) => tab.id)).toEqual(['task', 'view', 'file'])
    for (const tab of RIBBON_TABS) {
      expect(tab.groups.length).toBeGreaterThan(0)
      for (const group of tab.groups) {
        expect(group.controls.length).toBeGreaterThan(0)
        for (const control of group.controls) {
          // Every control is in the shared vocabulary (type-level + runtime).
          expect(MENU_COMMAND_IDS).toContain(control.id)
          expect(control.label.length).toBeGreaterThan(0)
        }
      }
    }
  })
})

describe('the ribbon structure', () => {
  it('renders the three tabs with the Task tab active by default', () => {
    createRibbon(mount(), { onCommand: () => {} })
    expect(ribbonTab('task')).not.toBeNull()
    expect(ribbonTab('view')).not.toBeNull()
    expect(ribbonTab('file')).not.toBeNull()
    expect(ribbonTab('task').getAttribute('aria-selected')).toBe('true')
    expect(ribbonTab('view').getAttribute('aria-selected')).toBe('false')
    expect(ribbonPanel('task').hidden).toBe(false)
    expect(ribbonPanel('view').hidden).toBe(true)
    expect(ribbonPanel('file').hidden).toBe(true)
  })

  it('renders all 16 controls with command ids, labels, and accelerator tooltips', () => {
    createRibbon(mount(), { onCommand: () => {} })
    const buttons = ribbonButtons()
    expect(buttons).toHaveLength(16)
    const ids = buttons.map((button) => button.dataset.command)
    expect(new Set(ids).size).toBe(16)
    for (const button of buttons) {
      const label = button.querySelector('.gp-ribbon-button-label')?.textContent ?? ''
      expect(label.length).toBeGreaterThan(0)
      // Every tooltip carries the label (accelerator display parity).
      expect(button.title).toContain(label)
    }
    // Accelerator display samples (the menu bar's displayed forms).
    expect(ribbonButton('task.create').title).toBe('New Task (Insert)')
    expect(ribbonButton('file.save').title).toBe('Save (Ctrl+S)')
    expect(ribbonButton('view.fit').title).toBe('Fit to Project (Ctrl+Shift+F)')
  })

  it('renders the groups with their labels', () => {
    createRibbon(mount(), { onCommand: () => {} })
    const groups = [...document.querySelectorAll('[data-testid="ribbon-group"]')] as HTMLElement[]
    expect(groups.map((group) => group.dataset.group)).toEqual([
      'Tasks',
      'Schedule',
      'History',
      'Zoom',
      'Outline',
      'Document',
    ])
  })

  it('switches panels on tab clicks (one visible panel, aria-selected moves)', () => {
    createRibbon(mount(), { onCommand: () => {} })
    switchTab('view')
    expect(ribbonTab('view').getAttribute('aria-selected')).toBe('true')
    expect(ribbonTab('task').getAttribute('aria-selected')).toBe('false')
    expect(ribbonPanel('view').hidden).toBe(false)
    expect(ribbonPanel('task').hidden).toBe(true)
    switchTab('file')
    expect(ribbonPanel('file').hidden).toBe(false)
    expect(ribbonPanel('view').hidden).toBe(true)
    switchTab('task')
    expect(ribbonPanel('task').hidden).toBe(false)
    expect(ribbonTab('task').getAttribute('aria-selected')).toBe('true')
  })
})

describe('ribbon activation', () => {
  it('forwards every enabled control click as its command id', () => {
    const seen: MenuCommandId[] = []
    const ribbon = createRibbon(mount(), { onCommand: (command) => seen.push(command) })
    ribbon.update(ALL_ON)
    for (const id of MENU_COMMAND_IDS) {
      ribbonButton(id).click()
    }
    expect(seen).toEqual([...MENU_COMMAND_IDS])
  })

  it('never dispatches from a disabled control (synthetic events included)', () => {
    const seen: MenuCommandId[] = []
    const ribbon = createRibbon(mount(), { onCommand: (command) => seen.push(command) })
    ribbon.update(ALL_OFF)
    for (const id of [
      'edit.undo',
      'edit.redo',
      'edit.deleteTask',
      'task.indent',
      'task.outdent',
    ] as const) {
      // A dispatched synthetic click cannot bypass the disabled gate.
      ribbonButton(id).dispatchEvent(new MouseEvent('click', { bubbles: true }))
    }
    expect(seen).toEqual([])
  })
})

describe('the four presentation echoes', () => {
  it('gates undo/redo on the session flags and the task commands on selection', () => {
    const ribbon = createRibbon(mount(), { onCommand: () => {} })
    ribbon.update(ALL_OFF)
    expect(ribbonButton('edit.undo').disabled).toBe(true)
    expect(ribbonButton('edit.redo').disabled).toBe(true)
    expect(ribbonButton('edit.deleteTask').disabled).toBe(true)
    expect(ribbonButton('task.indent').disabled).toBe(true)
    expect(ribbonButton('task.outdent').disabled).toBe(true)
    expect(ribbonButton('task.create').disabled).toBe(false)
    expect(ribbonButton('file.save').disabled).toBe(false)
    expect(ribbonButton('view.fit').disabled).toBe(false)

    ribbon.update(ALL_ON)
    expect(ribbonButton('edit.undo').disabled).toBe(false)
    expect(ribbonButton('edit.redo').disabled).toBe(false)
    expect(ribbonButton('edit.deleteTask').disabled).toBe(false)
    expect(ribbonButton('task.indent').disabled).toBe(false)
    expect(ribbonButton('task.outdent').disabled).toBe(false)
  })

  it('mirrors aria-disabled with the disabled property', () => {
    const ribbon = createRibbon(mount(), { onCommand: () => {} })
    ribbon.update(ALL_OFF)
    expect(ribbonButton('edit.undo').getAttribute('aria-disabled')).toBe('true')
    ribbon.update(ALL_ON)
    expect(ribbonButton('edit.undo').getAttribute('aria-disabled')).toBe('false')
  })

  it('reflects the dirty echo on the Save control only', () => {
    const ribbon = createRibbon(mount(), { onCommand: () => {} })
    ribbon.update(ALL_OFF)
    const marker = document.querySelector('[data-testid="ribbon-dirty-indicator"]') as HTMLElement
    expect(marker).not.toBeNull()
    expect(marker.dataset.dirty).toBe('false')
    expect(marker.textContent).toBe('')
    ribbon.update({ ...ALL_OFF, dirty: true })
    expect(marker.dataset.dirty).toBe('true')
    expect(marker.textContent).toBe('●')
    ribbon.update({ ...ALL_ON, dirty: false })
    expect(marker.dataset.dirty).toBe('false')
    expect(marker.textContent).toBe('')
  })

  it('echo updates never touch the tab selection (presentation state is separate)', () => {
    const ribbon = createRibbon(mount(), { onCommand: () => {} })
    switchTab('view')
    ribbon.update(ALL_ON)
    ribbon.update(ALL_OFF)
    expect(ribbonPanel('view').hidden).toBe(false)
    expect(ribbonTab('view').getAttribute('aria-selected')).toBe('true')
  })
})

describe('ribbon build determinism', () => {
  it('two fresh instances produce byte-identical chrome', () => {
    const first = document.createElement('div')
    const second = document.createElement('div')
    const ribbonA = createRibbon(first, { onCommand: () => {} })
    const ribbonB = createRibbon(second, { onCommand: () => {} })
    ribbonA.update(ALL_ON)
    ribbonB.update(ALL_ON)
    expect(second.innerHTML).toBe(first.innerHTML)
    ribbonA.update(ALL_OFF)
    ribbonB.update(ALL_OFF)
    expect(second.innerHTML).toBe(first.innerHTML)
  })
})

// ---- layer 2: the REAL controller integration --------------------------------

describe('the ribbon through the real controller (engine/save paths)', () => {
  it('is mounted by the shared DOM layer inside the app skeleton', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    const appRoot = document.querySelector('[data-testid="project-app"]') as HTMLElement
    const ribbon = document.querySelector('[data-testid="ribbon"]') as HTMLElement
    expect(ribbon).not.toBeNull()
    expect(appRoot.contains(ribbon)).toBe(true)
    // The ribbon sits above the workspace and carries the full vocabulary.
    expect(appRoot.querySelector('[data-testid="ribbon-button"]')).not.toBeNull()
    expect(ribbonButtons()).toHaveLength(16)
  })

  it('boot echoes: nothing to undo/redo, no selection, clean document', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    expect(ribbonButton('edit.undo').disabled).toBe(true)
    expect(ribbonButton('edit.redo').disabled).toBe(true)
    expect(ribbonButton('edit.deleteTask').disabled).toBe(true)
    expect(ribbonButton('task.indent').disabled).toBe(true)
    expect(
      (document.querySelector('[data-testid="ribbon-dirty-indicator"]') as HTMLElement).dataset
        .dirty,
    ).toBe('false')
  })

  it('New Task creates through the REAL engine path and enables the selection echoes', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    const created = app.state.session.document.tasks
    expect(created).toHaveLength(1)
    // The scheduler derived the schedule (the real session pipeline).
    expect(Object.keys(app.state.session.schedule?.taskSchedules ?? {})).toHaveLength(1)
    // The new task is selected → the selection echoes flip on.
    expect(ribbonButton('edit.deleteTask').disabled).toBe(false)
    expect(ribbonButton('task.indent').disabled).toBe(false)
    expect(ribbonButton('task.outdent').disabled).toBe(false)
    // A row rendered (the projection pipeline ran).
    expect(document.querySelector('[data-testid="task-row"]')).not.toBeNull()
  })

  it('Delete Task removes through the REAL engine path and disables the echoes', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    click('edit.deleteTask')
    expect(app.state.session.document.tasks).toHaveLength(0)
    expect(ribbonButton('edit.deleteTask').disabled).toBe(true)
    expect(document.querySelector('[data-testid="task-row"]')).toBeNull()
  })

  it('Undo/Redo controls drive the REAL session journal', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    expect(ribbonButton('edit.undo').disabled).toBe(false)
    click('task.create')
    expect(app.state.session.document.tasks).toHaveLength(2)
    click('edit.undo')
    expect(app.state.session.document.tasks).toHaveLength(1)
    expect(ribbonButton('edit.redo').disabled).toBe(false)
    click('edit.redo')
    expect(app.state.session.document.tasks).toHaveLength(2)
    click('edit.undo')
    click('edit.undo')
    expect(app.state.session.document.tasks).toHaveLength(0)
    // Empty journal: both history controls disabled again.
    expect(ribbonButton('edit.undo').disabled).toBe(true)
    expect(ribbonButton('edit.redo').disabled).toBe(false)
  })

  it('Save runs the REAL save path over the bridge (adapter bytes, dirty cleared)', async () => {
    const { bridge, calls, files } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    const dirtyMarker = document.querySelector(
      '[data-testid="ribbon-dirty-indicator"]',
    ) as HTMLElement
    expect(dirtyMarker.dataset.dirty).toBe('true')
    click('file.save')
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(calls.writeFile).toEqual(['saved.gproj'])
    const written = files.get('saved.gproj')
    expect(written).toBeDefined()
    // The bytes are the canonical .gproj export (the JSON envelope, magic
    // "gproj" — the same envelope the adapter round-trip asserts).
    const envelope = JSON.parse(new TextDecoder().decode(written!)) as { format: string }
    expect(envelope.format).toBe('gproj')
    // Saved: the dirty echo clears.
    expect(dirtyMarker.dataset.dirty).toBe('false')
  })

  it('Save As consults the save picker first (the Save As path)', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('file.saveAs')
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(calls.pickSaveFile).toBe(1)
    expect(calls.writeFile).toEqual(['saved.gproj'])
    expect(app.state.filePath).toBe('saved.gproj')
  })

  it('menu activation and ribbon activation are ONE translation (parity)', () => {
    const { bridge } = fakeBridge()
    const appA = createProjectApp({ bridge, root: mount() })
    appA.start()
    click('task.create')
    click('task.create')
    const viaRibbon = appA.state.session.document

    document.body.innerHTML = ''
    const appB = createProjectApp({ bridge, root: mount() })
    appB.start()
    appB.menuCommand('task.create')
    appB.menuCommand('task.create')
    const viaMenu = appB.state.session.document

    // Identical shape through the shared translation table (names + ids).
    expect(viaRibbon.tasks.map((task) => task.id)).toEqual(viaMenu.tasks.map((task) => task.id))
    expect(viaRibbon.tasks.map((task) => task.name)).toEqual(viaMenu.tasks.map((task) => task.name))
  })

  it('View-tab controls dispatch view intents through the reducer', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    const before = app.state.viewState.viewport
    switchTab('view')
    click('view.zoomOut')
    const after = app.state.viewState.viewport
    // Zoom out widens the span (the shared ZOOM_OUT_FACTOR).
    expect(Date.parse(after.finish) - Date.parse(after.start)).toBeGreaterThan(
      Date.parse(before.finish) - Date.parse(before.start),
    )
    // Fit restores the canonical fitViewport() result (the authority).
    click('view.fit')
    expect(app.state.viewState.viewport).toEqual(
      fitViewport(app.state.session.document, app.state.session.schedule),
    )
  })

  it('Collapse Selected drives the view-state collapse through the ribbon', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    click('task.create')
    // Indent the second task under the first (a summary appears).
    const second = app.state.session.document.tasks[1]!.id
    app.execute({ kind: 'intent', intent: { type: 'selectTask', taskId: second, mode: 'set' } })
    click('task.indent')
    // Root tasks live at outlineLevel 1 (the MS Project convention) — the
    // indent moves the second task under the first (level 2).
    expect(app.state.session.document.tasks[1]!.outlineLevel).toBe(2)
    // Select the parent, collapse through the ribbon → the child hides.
    const parent = app.state.session.document.tasks[0]!.id
    app.execute({ kind: 'intent', intent: { type: 'selectTask', taskId: parent, mode: 'set' } })
    const rowsBefore = document.querySelectorAll('[data-testid="task-row"]').length
    switchTab('view')
    click('view.collapse')
    const rowsAfter = document.querySelectorAll('[data-testid="task-row"]').length
    expect(rowsAfter).toBe(rowsBefore - 1)
    // Expand restores the row.
    click('view.expand')
    expect(document.querySelectorAll('[data-testid="task-row"]').length).toBe(rowsBefore)
  })

  it('tab selection survives renders driven by app state changes', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    switchTab('file')
    click('task.create')
    expect(app.state.session.document.tasks).toHaveLength(1)
    // The ribbon re-rendered (echoes updated) but the File panel stays open.
    expect(ribbonPanel('file').hidden).toBe(false)
    expect(ribbonTab('file').getAttribute('aria-selected')).toBe('true')
  })

  it('the dirty echo tracks real document identity through undo-to-saved', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    click('task.create')
    click('file.save')
    return new Promise<void>((resolve) => {
      setTimeout(() => {
        const marker = document.querySelector(
          '[data-testid="ribbon-dirty-indicator"]',
        ) as HTMLElement
        expect(marker.dataset.dirty).toBe('false')
        click('task.create')
        expect(marker.dataset.dirty).toBe('true')
        // Undo back to the saved reference: the identity probe clears.
        click('edit.undo')
        expect(marker.dataset.dirty).toBe('false')
        resolve()
      }, 0)
    })
  })
})
