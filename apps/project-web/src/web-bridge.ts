/**
 * PROJECT-028 — the Project web transport bridge.
 *
 * The browser implementation of the shared `ProjectHostBridge` contract
 * (`@genoffice/project-host`): the same typed surface the Electron preload
 * exposes on desktop, over browser primitives only —
 *
 * - `pickOpenFile`  → a real `<input type="file">` picker (the accept
 *                     vocabulary derived from `PROJECT_FILE_FILTERS`) whose
 *                     chosen File crosses the ONE bounded web read;
 * - `pickSaveFile`  → the browser's download flow IS the save dialog (the
 *                     default name is returned; `writeFile` downloads);
 * - `writeFile`     → a Blob + anchor download under the file name;
 * - `readFile`      → the bounded read of a STAGED external file
 *                     (drag-and-drop — the web analog of argv/second
 *                     instance opens on desktop);
 * - `confirmDiscard`→ a three-button DOM dialog (Save / Don't Save /
 *                     Cancel — `window.confirm` is two-button and cannot
 *                     express the contract);
 * - `onCloseRequested`/`approveClose` → the `beforeunload` guard: the
 *                     browser cannot await an async handshake during
 *                     unload, so the bridge consults a registered DIRTY
 *                     PROBE synchronously and confirms leaving natively
 *                     when the document is dirty; the controller's
 *                     save/discard/cancel flow still governs every in-app
 *                     destructive action (New/Open over unsaved changes);
 * - `onMenuCommand` → the DOM menu bar's activation path (the web analog
 *                     of the native menu's forwarding).
 *
 * THE one bounded web read: every File the bridge delivers — the picker
 * path AND the external/staged path — crosses ONE helper (`readCapped`):
 * the File's size is checked BEFORE a byte is read (the browser analog of
 * the desktop's stat-first rejection), then read whole (a File is already
 * a host-managed, size-known object — there is no window to race). Results
 * are `NativeReadResult` VALUES on every surface: errors are never thrown
 * and the error variant carries NO bytes — the controller can never
 * receive uncapped file contents, structurally rather than
 * conventionally (the PROJECT-027 desktop transport invariant, mirrored).
 *
 * This module imports NO Node or Electron API (the discipline suite scans
 * the whole web host source; the built bundle is scanned in CI).
 */
import type {
  DiscardChoice,
  HostAppInfo,
  MenuCommandId,
  NativeReadResult,
  OpenFileSelection,
  ProjectHostBridge,
} from '@genoffice/project-host'
import { PROJECT_FILE_FILTERS } from '@genoffice/project-host'

/**
 * The web transport cap — the same 100 MiB value the desktop transport
 * helper and the canonical adapters carry (defense in depth: three
 * independent layers, each refusing oversized input on its own terms).
 * Defined exactly once in the web host, here.
 */
export const MAX_WEB_FILE_BYTES = 104_857_600

/** The web host version echo (presentation only). */
const WEB_APP_VERSION = '0.1.0'

/** The accept attribute for the open picker, derived from the shared
 * filter vocabulary (never a second filter list). */
const ACCEPT_FROM_FILTERS = PROJECT_FILE_FILTERS.map((filter) =>
  filter.extensions.map((extension) => `.${extension}`),
)
  .flat()
  .join(',')

/** The single bounded web read (both read surfaces route through it).
 * Exported for direct unit testing — the desktop transport's
 * `boundedReadFile` precedent. */
export async function readCapped(file: File): Promise<NativeReadResult> {
  if (file.size > MAX_WEB_FILE_BYTES) {
    return { ok: false, error: `File exceeds the ${MAX_WEB_FILE_BYTES} byte limit` }
  }
  try {
    const buffer = await file.arrayBuffer()
    return { ok: true, bytes: new Uint8Array(buffer) }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

/** The file name part of a path-like string (presentation only). */
const baseName = (path: string): string => path.split(/[\\/]/).pop() ?? path

/**
 * The web bridge surface: the shared `ProjectHostBridge` plus the two
 * web-only wiring seams (the dirty probe for the synchronous beforeunload
 * guard, and the menu-bar/external-file dispatch paths).
 */
export interface WebBridge extends ProjectHostBridge {
  /** Registers the synchronous dirty answer the beforeunload guard
   * consults (the web entry registers `() => app.dirty`). */
  setDirtyProbe(probe: () => boolean): void
  /** The DOM menu bar's activation path — forwards to the controller's
   * registered `onMenuCommand` handler (the native menu's channel). */
  dispatchMenuCommand(command: MenuCommandId): void
  /** Stages an externally provided File (drag-and-drop) and requests its
   * open through the controller's `onOpenRequested` handler — the web
   * analog of the desktop argv/second-instance open: the unsaved-changes
   * gate, the bounded read, and the adapter import all flow through the
   * SHARED controller path. */
  stageExternalFile(file: File): void
}

export function createWebBridge(): WebBridge {
  let menuHandler: ((command: MenuCommandId) => void) | undefined
  let closeHandler: (() => void) | undefined
  let openHandler: ((path: string) => void) | undefined
  let dirtyProbe: (() => boolean) | undefined
  let closeApproved = false
  /** Staged external files by name (the readFile surface's store). */
  const stagedFiles = new Map<string, File>()

  // ---- the beforeunload close guard ------------------------------------
  // The browser cannot await the controller's async save/discard/cancel
  // handshake during unload: the guard consults the dirty probe
  // synchronously and asks the user natively. The controller's
  // handleCloseRequested still runs (its dialog is moot mid-unload), and
  // every in-app destructive path (New/Open over unsaved changes) consults
  // the same confirmDiscard dialog the desktop uses.
  window.addEventListener('beforeunload', (event) => {
    closeHandler?.()
    if (closeApproved) return
    if (dirtyProbe?.() === true) {
      event.preventDefault()
      // The legacy property some engines require for the prompt to appear.
      event.returnValue = ''
    }
  })

  const bridge: WebBridge = {
    async pickOpenFile(): Promise<OpenFileSelection | null> {
      const input = document.createElement('input')
      input.type = 'file'
      input.accept = ACCEPT_FROM_FILTERS
      input.dataset.testid = 'open-picker'
      input.style.display = 'none'
      document.body.appendChild(input)
      try {
        const file = await chooseFile(input)
        if (file === null) return null
        return { path: file.name, read: await readCapped(file) }
      } finally {
        input.remove()
      }
    },

    async pickSaveFile(defaultName: string): Promise<string | null> {
      // The browser's download flow is the save dialog; the name is the
      // shared document-flow default (the controller keeps the opened
      // file's name for plain saves).
      return defaultName
    },

    async readFile(path: string): Promise<NativeReadResult> {
      const file = stagedFiles.get(path)
      if (file === undefined) {
        return { ok: false, error: `ENOENT: no such file or directory, '${path}'` }
      }
      return readCapped(file)
    },

    async writeFile(path: string, bytes: Uint8Array): Promise<{ ok: boolean; error?: string }> {
      try {
        const blob = new Blob([bytes as BlobPart], { type: 'application/octet-stream' })
        const url = URL.createObjectURL(blob)
        const anchor = document.createElement('a')
        anchor.href = url
        anchor.download = baseName(path)
        anchor.dataset.testid = 'save-download'
        document.body.appendChild(anchor)
        anchor.click()
        anchor.remove()
        // The download starts synchronously on click; revoke late so the
        // transfer is never cut short.
        setTimeout(() => URL.revokeObjectURL(url), 60_000)
        return { ok: true }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },

    confirmDiscard(projectName: string): Promise<DiscardChoice> {
      return confirmDiscardDialog(projectName)
    },

    async appInfo(): Promise<HostAppInfo> {
      return { platform: navigator.platform || 'web', version: WEB_APP_VERSION }
    },

    onMenuCommand(handler: (command: MenuCommandId) => void): void {
      menuHandler = handler
    },

    onCloseRequested(handler: () => void): void {
      closeHandler = handler
    },

    approveClose(): void {
      closeApproved = true
    },

    onOpenRequested(handler: (path: string) => void): void {
      openHandler = handler
    },

    setDirtyProbe(probe: () => boolean): void {
      dirtyProbe = probe
    },

    dispatchMenuCommand(command: MenuCommandId): void {
      menuHandler?.(command)
    },

    stageExternalFile(file: File): void {
      stagedFiles.set(file.name, file)
      openHandler?.(file.name)
    },
  }
  return bridge
}

/** Opens the (already-mounted) file input and resolves the chosen File,
 * or null when the picker is cancelled. */
function chooseFile(input: HTMLInputElement): Promise<File | null> {
  return new Promise((resolve) => {
    let settled = false
    const done = (file: File | null): void => {
      if (settled) return
      settled = true
      resolve(file)
    }
    input.addEventListener('change', () => done(input.files?.[0] ?? null))
    input.addEventListener('cancel', () => done(null))
    input.click()
  })
}

/**
 * The three-button unsaved-changes dialog (Save / Don't Save / Cancel) —
 * the DOM analog of the native desktop dialog. Escape cancels (the native
 * dialog's keyboard behavior); focus starts on Save.
 */
function confirmDiscardDialog(projectName: string): Promise<DiscardChoice> {
  return new Promise((resolve) => {
    const overlay = document.createElement('div')
    overlay.dataset.testid = 'discard-dialog'
    overlay.className = 'gp-web-dialog-overlay'

    const dialog = document.createElement('div')
    dialog.className = 'gp-web-dialog'
    dialog.setAttribute('role', 'alertdialog')
    dialog.setAttribute('aria-modal', 'true')
    dialog.setAttribute('aria-label', 'Unsaved changes')

    const title = document.createElement('div')
    title.className = 'gp-web-dialog-title'
    title.textContent = `Save changes to '${projectName}'?`

    const buttons = document.createElement('div')
    buttons.className = 'gp-web-dialog-buttons'

    let settled = false
    const answer = (choice: DiscardChoice): void => {
      if (settled) return
      settled = true
      overlay.remove()
      document.removeEventListener('keydown', onKey)
      resolve(choice)
    }

    const addButton = (label: string, choice: DiscardChoice, testid: string): void => {
      const button = document.createElement('button')
      button.type = 'button'
      button.dataset.testid = testid
      button.textContent = label
      button.addEventListener('click', () => answer(choice))
      buttons.appendChild(button)
    }
    addButton('Save', 'save', 'discard-save')
    addButton("Don't Save", 'discard', 'discard-dont-save')
    addButton('Cancel', 'cancel', 'discard-cancel')

    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        event.stopPropagation()
        answer('cancel')
      }
    }
    document.addEventListener('keydown', onKey, { capture: true })

    dialog.appendChild(title)
    dialog.appendChild(buttons)
    overlay.appendChild(dialog)
    document.body.appendChild(overlay)
    overlay.querySelector<HTMLButtonElement>('[data-testid="discard-save"]')?.focus()
  })
}
