/**
 * PROJECT-027 — the renderer entry: mounts the Project desktop host into
 * the window. The desktop bridge is the preload surface; keyboard events
 * translate through the host table and dispatch through the controller.
 */
import './styles.css'
import { createProjectDesktopApp } from './app.js'
import { translateKeyDown } from './translate.js'
import type { ProjectDesktopBridge } from '../shared/ipc.js'

declare global {
  interface Window {
    projectDesktop: ProjectDesktopBridge
  }
}

const root = document.getElementById('root')
if (root === null) {
  throw new Error('Missing #root mount element')
}

const bridge: ProjectDesktopBridge = window.projectDesktop
const app = createProjectDesktopApp({ bridge, root })

// Keyboard: the single translation path (the native menu accelerators are
// displayed but not registered — see src/main/menu.ts).
window.addEventListener(
  'keydown',
  (event) => {
    const input = {
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

// window.onbeforeunload is NOT used: the close guard lives in the main
// process ('close' event) + the controller's close-requested handler.

app.start()
