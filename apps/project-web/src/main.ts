/**
 * PROJECT-028 — the Project web host entry: mounts the SAME shared host
 * binding the desktop shell runs (`@genoffice/project-host`) into the
 * browser page. The web bridge is the browser transport; the DOM menu bar
 * is the web chrome (activation forwards through the bridge's
 * menu-command path); keyboard events translate through the shared host
 * table and dispatch through the controller — the single execution path,
 * exactly like the desktop entry.
 *
 * This module imports NO Node or Electron API (the discipline suite scans
 * the whole web host source; the built bundle is scanned in CI).
 */
import '@genoffice/project-host/styles.css'
import './web.css'
import { createProjectApp, translateKeyDown } from '@genoffice/project-host'
import type { KeyInput } from '@genoffice/project-host'
import { createWebBridge } from './web-bridge.js'
import type { WebBridge } from './web-bridge.js'
import { createMenuBar } from './menu.js'

const root = document.getElementById('root')
const menuContainer = document.getElementById('menu-bar')
if (root === null || menuContainer === null) {
  throw new Error('Missing #root/#menu-bar mount elements')
}

const bridge = createWebBridge()
const app = createProjectApp({ bridge, root })

// The host bridge surface, exposed for host diagnostics — the exact parity
// of the desktop preload's `window.projectDesktop` (the E2E drives the
// transport read through it, the desktop battery's precedent).
declare global {
  interface Window {
    projectWeb: WebBridge
  }
}
window.projectWeb = bridge

// The beforeunload guard's synchronous dirty answer. The guard is purely
// synchronous — it never initiates the controller's async close handshake
// during unload (see web-bridge.ts); the in-app close request is the one
// firing path for that handshake.
bridge.setDirtyProbe(() => app.dirty)

// The DOM menu bar: activation forwards through the bridge's menu-command
// path (the native menu's IPC channel equivalent).
const menuBar = createMenuBar(menuContainer, (command) => bridge.dispatchMenuCommand(command))

// Keyboard: the single translation path (the menu bar displays accelerators
// but executes nothing — the shared table owns execution).
window.addEventListener(
  'keydown',
  (event) => {
    // While a menu dropdown is open, the menu owns the keyboard (the
    // native-menu behavior); the app never sees the key.
    if (menuBar.isOpen()) {
      event.preventDefault()
      return
    }
    const input: KeyInput = {
      key: event.key,
      ctrlOrMeta: event.ctrlKey || event.metaKey,
      shift: event.shiftKey,
      alt: event.altKey,
    }
    const editing =
      app.state.viewState.editing !== undefined ||
      app.state.viewState.dependencyEditing !== undefined
    // The translation decides; when it yields an action the browser default
    // (page scroll, caret navigation) is suppressed.
    const action = translateKeyDown(input, { editing })
    if (action.kind !== 'none') event.preventDefault()
    app.keydown(input)
  },
  { capture: true },
)

// Drag-and-drop open: the web analog of the desktop argv/second-instance
// open — the dropped File is staged in the bridge and the request flows
// through the controller's onOpenRequested path (unsaved-changes gate +
// bounded read + adapter import, all in the SHARED controller code).
window.addEventListener('dragover', (event) => {
  if (event.dataTransfer?.types.includes('Files') === true) {
    event.preventDefault()
    document.body.classList.add('gp-web-dragging')
  }
})
window.addEventListener('dragleave', (event) => {
  if (event.relatedTarget === null) document.body.classList.remove('gp-web-dragging')
})
window.addEventListener('drop', (event) => {
  if (event.dataTransfer === null) return
  event.preventDefault()
  document.body.classList.remove('gp-web-dragging')
  const file = event.dataTransfer.files[0]
  if (file === undefined) return
  bridge.stageExternalFile(file)
})

app.start()
