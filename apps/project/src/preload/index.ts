/**
 * PROJECT-027 — the Project desktop host preload bridge.
 *
 * Exposes the typed `window.projectDesktop` surface (contextBridge) and
 * nothing else. The renderer never touches `electron`, `ipcRenderer`, or
 * Node APIs: everything crosses through this typed transport boundary, so
 * the renderer binding is unit-testable under jsdom with an in-memory fake
 * of the same interface.
 */
import { contextBridge, ipcRenderer } from 'electron'
import { PROJECT_IPC } from '../shared/ipc.js'
import type { MenuCommandId, ProjectDesktopBridge } from '../shared/ipc.js'

const bridge: ProjectDesktopBridge = {
  pickOpenFile: () => ipcRenderer.invoke(PROJECT_IPC.pickOpenFile),

  pickSaveFile: (defaultName: string) => ipcRenderer.invoke(PROJECT_IPC.pickSaveFile, defaultName),

  readFile: (path: string) => ipcRenderer.invoke(PROJECT_IPC.readFile, path),

  writeFile: (path: string, bytes: Uint8Array) =>
    ipcRenderer.invoke(PROJECT_IPC.writeFile, path, bytes),

  confirmDiscard: (projectName: string) =>
    ipcRenderer.invoke(PROJECT_IPC.confirmDiscard, projectName),

  appInfo: () => ipcRenderer.invoke(PROJECT_IPC.appInfo),

  onMenuCommand: (handler) => {
    ipcRenderer.on(PROJECT_IPC.menuCommand, (_event, command: MenuCommandId) => {
      handler(command)
    })
  },

  onCloseRequested: (handler) => {
    ipcRenderer.on(PROJECT_IPC.closeRequested, () => {
      handler()
    })
  },

  approveClose: () => {
    ipcRenderer.send(PROJECT_IPC.closeApproved)
  },

  onOpenRequested: (handler) => {
    ipcRenderer.on(PROJECT_IPC.openRequested, (_event, path: string) => {
      handler(path)
    })
  },
}

contextBridge.exposeInMainWorld('projectDesktop', bridge)
