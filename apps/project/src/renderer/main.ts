/**
 * The Project desktop renderer entry (PROJECT-027; the shared binding
 * consumed from `@genoffice/project-host` since PROJECT-028): mounts the
 * host controller into the window. The desktop bridge is the preload
 * surface; keyboard events translate through the shared host table and
 * dispatch through the controller.
 */
import '@genoffice/project-host/styles.css'
import { createProjectApp, translateKeyDown } from '@genoffice/project-host'
import type { KeyInput } from '@genoffice/project-host'
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
const app = createProjectApp({ bridge, root })

// Keyboard: the single translation path (the native menu accelerators are
// displayed but not registered — see src/main/menu.ts).
window.addEventListener(
  'keydown',
  (event) => {
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

// window.onbeforeunload is NOT used: the close guard lives in the main
// process ('close' event) + the controller's close-requested handler.

app.start()
