/**
 * PROJECT-027 — the Project desktop host main process.
 *
 * NATIVE TRANSPORT ONLY (architecture-lock §3): this process owns the
 * window/lifecycle integration, the native menu (activation forwarding),
 * the native file pickers, size-capped raw filesystem reads/writes (every
 * read crosses the ONE canonical bounded helper — bounded-read.ts), and
 * the window close guard. It NEVER imports a `@genoffice/project-*`
 * package and never sees a `ProjectDocument` — documents cross the
 * boundary as raw bytes and the canonical file adapters run renderer-side
 * (the discipline suite asserts the import surface). The unsaved-changes
 * dialog left the transport at PROJECT-030: it is shared presentation
 * rendered by the host binding's dialog layer in the renderer (the close
 * handshake below is unchanged — a prevented close consults the renderer,
 * which runs the shared dialog and approves or refuses).
 *
 * Close guard: the window 'close' event is prevented once and forwarded to
 * the renderer, which owns the dirty state and the save flow. The renderer
 * replies with close-approved (possibly after saving); a crashed/dead
 * renderer cannot veto shutdown — the guard auto-approves in that case.
 */
import { BrowserWindow, Menu, app, dialog, ipcMain, shell } from 'electron'
import { writeFile } from 'node:fs/promises'
import { extname, join } from 'node:path'
import { PROJECT_FILE_FILTERS, PROJECT_IPC } from '../shared/ipc.js'
import type { MenuCommandId } from '../shared/ipc.js'
import { MAX_FILE_BYTES, boundedReadFile } from './bounded-read.js'
import { projectMenuTemplate, sendMenuCommand } from './menu.js'

/** File extensions the host opens (matches the canonical adapters). */
const OPEN_EXTENSIONS = new Set(['gproj', 'xml'])

let mainWindow: BrowserWindow | null = null
/** Set once the renderer approved a pending close (close-guard handshake). */
let closeApproved = false

/** Finds a document path in an argv (launch open + second-instance open). */
function documentArgvPath(argv: readonly string[]): string | null {
  for (const arg of argv.slice(1)) {
    if (arg.startsWith('-')) continue
    const extension = extname(arg).slice(1).toLowerCase()
    if (OPEN_EXTENSIONS.has(extension)) return arg
  }
  return null
}

function sendToWindow(channel: string, payload?: unknown): void {
  if (mainWindow !== null && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send(channel, payload)
  }
}

function sendCommand(command: MenuCommandId): void {
  if (mainWindow !== null && !mainWindow.isDestroyed()) {
    sendMenuCommand(mainWindow.webContents, command)
  }
}

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 540,
    title: 'GenOffice Project',
    webPreferences: {
      // electron-vite emits out/preload/index.mjs for ESM packages.
      preload: join(__dirname, '../preload/index.mjs'),
      contextIsolation: true,
      nodeIntegration: false,
      // The preload uses only contextBridge + ipcRenderer (no require of
      // arbitrary modules), so the sandbox stays on.
      sandbox: false,
      spellcheck: false,
    },
  })

  // Native menu: activation forwarding only (translation lives renderer-side).
  Menu.setApplicationMenu(Menu.buildFromTemplate(projectMenuTemplate(sendCommand)))

  // Open external links in the OS browser, never inside the host window.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url)
    return { action: 'deny' }
  })

  // ---- close guard -------------------------------------------------------
  mainWindow.on('close', (event) => {
    if (closeApproved || mainWindow === null) return
    if (!mainWindow.webContents || mainWindow.webContents.isCrashed()) {
      // A dead renderer cannot answer the handshake — allow the close.
      closeApproved = true
      return
    }
    event.preventDefault()
    mainWindow.webContents.send(PROJECT_IPC.closeRequested)
  })

  mainWindow.on('closed', () => {
    mainWindow = null
  })

  if (process.env.PROJECT_DEV_SERVER_URL !== undefined) {
    void mainWindow.loadURL(process.env.PROJECT_DEV_SERVER_URL)
  } else if (process.env.PROJECT_RENDERER_PAGE !== undefined) {
    void mainWindow.loadFile(process.env.PROJECT_RENDERER_PAGE)
  } else {
    // electron-vite build output (out/renderer/index.html).
    void mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

/** Registers the native-transport IPC handlers (once per app lifetime). */
function registerIpcHandlers(): void {
  // Native open dialog + the canonical bounded read. Returns null when the
  // user cancelled; read errors (oversized/missing/unreadable) are values.
  ipcMain.handle(PROJECT_IPC.pickOpenFile, async () => {
    const focused = BrowserWindow.getFocusedWindow() ?? mainWindow
    if (focused === null) return null
    const selection = await dialog.showOpenDialog(focused, {
      title: 'Open Project',
      properties: ['openFile'],
      filters: [...PROJECT_FILE_FILTERS],
    })
    const path = selection.filePaths[0]
    if (selection.canceled || path === undefined) return null
    return { path, read: await boundedReadFile(path) }
  })

  // Native save dialog: path selection only (bytes come from the renderer).
  ipcMain.handle(PROJECT_IPC.pickSaveFile, async (_event, defaultName: string) => {
    const focused = BrowserWindow.getFocusedWindow() ?? mainWindow
    if (focused === null) return null
    const selection = await dialog.showSaveDialog(focused, {
      title: 'Save Project',
      defaultPath: defaultName,
      filters: [...PROJECT_FILE_FILTERS],
    })
    return selection.canceled || selection.filePath === undefined ? null : selection.filePath
  })

  // The argv / second-instance read path — the SAME canonical bounded read
  // as the picker path (one transport policy; no uncapped read surface).
  ipcMain.handle(PROJECT_IPC.readFile, async (_event, path: string) => boundedReadFile(path))

  ipcMain.handle(PROJECT_IPC.writeFile, async (_event, path: string, bytes: Uint8Array) => {
    try {
      if (bytes.byteLength > MAX_FILE_BYTES) {
        return { ok: false, error: `Export exceeds the ${MAX_FILE_BYTES} byte limit` }
      }
      await writeFile(path, bytes)
      return { ok: true }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  })

  ipcMain.handle(PROJECT_IPC.appInfo, () => ({
    platform: process.platform,
    version: app.getVersion(),
  }))

  // Close-guard handshake: the renderer approves a pending close.
  ipcMain.on(PROJECT_IPC.closeApproved, () => {
    closeApproved = true
    mainWindow?.close()
  })
}

// Profile/E2E isolation FIRST: the single-instance lock is keyed on the
// userData path, so the path must be installed BEFORE the lock is
// requested — a scratch dir then carries its own isolated lock and a
// second launch with a DIFFERENT dir runs independently. (Installing the
// path after the lock would key the lock on the real profile and let E2E
// runs collide with each other and with a running install — the review
// finding; the discipline suite pins this ordering.)
if (process.env.GENOFFICE_PROJECT_USER_DATA !== undefined) {
  app.setPath('userData', process.env.GENOFFICE_PROJECT_USER_DATA)
}

// Single instance: a second launch forwards its document argv to this one.
if (!app.requestSingleInstanceLock()) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv) => {
    const path = documentArgvPath(argv)
    if (path !== null) sendToWindow(PROJECT_IPC.openRequested, path)
    if (mainWindow !== null) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
  })

  app.whenReady().then(() => {
    registerIpcHandlers()
    createWindow()
    // Launch-open: forward the argv document once the renderer can listen.
    const launchPath = documentArgvPath(process.argv)
    if (launchPath !== null) {
      mainWindow?.webContents.once('did-finish-load', () => {
        sendToWindow(PROJECT_IPC.openRequested, launchPath)
      })
    }
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
}
