/**
 * PROJECT-027 — the Project desktop host main process.
 *
 * NATIVE TRANSPORT ONLY (architecture-lock §3): this process owns the
 * window/lifecycle integration, the native menu (activation forwarding),
 * the native file dialogs, raw filesystem reads/writes, and the window
 * close guard. It NEVER imports a `@genoffice/project-*` package and never
 * sees a `ProjectDocument` — documents cross the boundary as raw bytes and
 * the canonical file adapters run renderer-side (the discipline suite
 * asserts the import surface).
 *
 * Close guard: the window 'close' event is prevented once and forwarded to
 * the renderer, which owns the dirty state and the save flow. The renderer
 * replies with close-approved (possibly after saving); a crashed/dead
 * renderer cannot veto shutdown — the guard auto-approves in that case.
 */
import { BrowserWindow, Menu, app, dialog, ipcMain, shell } from 'electron'
import { readFile, writeFile } from 'node:fs/promises'
import { basename, extname, join } from 'node:path'
import { PROJECT_FILE_FILTERS, PROJECT_IPC } from '../shared/ipc.js'
import type { MenuCommandId } from '../shared/ipc.js'
import { projectMenuTemplate, sendMenuCommand } from './menu.js'

/** File extensions the host opens (matches the canonical adapters). */
const OPEN_EXTENSIONS = new Set(['gproj', 'xml'])

/** The maximum size the host will read or write (defense in depth; the
 * canonical adapters carry their own caps — this is transport hygiene). */
const MAX_FILE_BYTES = 100 * 1024 * 1024

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
  // Native open dialog + raw read. Returns null when the user cancelled.
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
    const bytes = await readFile(path)
    if (bytes.byteLength > MAX_FILE_BYTES) {
      return { path, error: `File exceeds the ${MAX_FILE_BYTES} byte limit` }
    }
    return { path, bytes }
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

  ipcMain.handle(PROJECT_IPC.readFile, async (_event, path: string) => readFile(path))

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

  // Unsaved-changes dialog: Save / Don't Save / Cancel.
  ipcMain.handle(PROJECT_IPC.confirmDiscard, async (_event, projectName: string) => {
    const focused = BrowserWindow.getFocusedWindow() ?? mainWindow
    if (focused === null) return 'cancel'
    const title = basename(projectName) || 'Untitled'
    const result = await dialog.showMessageBox(focused, {
      type: 'warning',
      buttons: ['Save', "Don't Save", 'Cancel'],
      defaultId: 0,
      cancelId: 2,
      message: `Save changes to “${title}” before closing?`,
      detail: 'Unsaved changes will be lost if you don’t save them.',
    })
    if (result.response === 0) return 'save'
    if (result.response === 1) return 'discard'
    return 'cancel'
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

  // E2E scratch isolation (the shell's GENOFFICE_USER_DATA pattern).
  if (process.env.GENOFFICE_PROJECT_USER_DATA !== undefined) {
    app.setPath('userData', process.env.GENOFFICE_PROJECT_USER_DATA)
  }

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
