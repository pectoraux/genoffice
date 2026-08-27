/**
 * PROJECT-027 — the Project desktop host's IPC contract.
 *
 * The typed boundary between the Electron main process (native transport),
 * the preload bridge (contextBridge surface), and the renderer (the shared
 * renderer-core binding). This module owns CHANNEL IDENTIFIERS, MENU COMMAND
 * IDS, and the BRIDGE INTERFACE — transport vocabulary only. It contains no
 * Project semantics: every semantic value that crosses the bridge is either
 * raw file bytes (the canonical file adapters run renderer-side) or an
 * opaque path string.
 *
 * The main process never imports a `@genoffice/project-*` package (the
 * architecture discipline suite asserts this): dialogs, file reads/writes,
 * and menu forwarding are native transport, not Project logic.
 */

/** Structured-clone channel identifiers for `ipcMain`/`ipcRenderer`. */
export const PROJECT_IPC = {
  /** Renderer → main (invoke): native open dialog + file read. */
  pickOpenFile: 'project:pick-open-file',
  /** Renderer → main (invoke): native save dialog (path selection only). */
  pickSaveFile: 'project:pick-save-file',
  /** Renderer → main (invoke): read raw bytes (argv/second-instance opens). */
  readFile: 'project:read-file',
  /** Renderer → main (invoke): write raw bytes to a path. */
  writeFile: 'project:write-file',
  /** Renderer → main (invoke): unsaved-changes dialog (Save/Discard/Cancel). */
  confirmDiscard: 'project:confirm-discard',
  /** Renderer → main (invoke): host platform/version echo. */
  appInfo: 'project:app-info',
  /** Main → renderer (send): a native menu command was activated. */
  menuCommand: 'project:menu-command',
  /** Main → renderer (send): the window close was requested (close guard). */
  closeRequested: 'project:close-requested',
  /** Renderer → main (send): the renderer approves the close. */
  closeApproved: 'project:close-approved',
  /** Main → renderer (send): open the document at this path (argv /
   * second instance). */
  openRequested: 'project:open-requested',
} as const

/**
 * The native menu command vocabulary. Menu ids are TRANSPORT identifiers:
 * the renderer's translation layer (`src/renderer/translate.ts`) maps each
 * id to the canonical renderer-core intent/command action — the host menu
 * never names a Project semantic operation directly.
 */
export type MenuCommandId =
  | 'file.new'
  | 'file.open'
  | 'file.save'
  | 'file.saveAs'
  | 'edit.undo'
  | 'edit.redo'
  | 'edit.deleteTask'
  | 'task.create'
  | 'task.indent'
  | 'task.outdent'
  | 'view.zoomIn'
  | 'view.zoomOut'
  | 'view.fit'
  | 'view.collapse'
  | 'view.expand'

/**
 * The result of a bounded native read: capped raw bytes, or the transport
 * error (oversized / missing / unreadable). Errors are VALUES, never
 * throws — and the error variant carries NO bytes, so the renderer can
 * NEVER receive uncapped file contents: one `NativeReadResult` is the only
 * shape file content crosses the bridge in, on every read surface.
 */
export type NativeReadResult =
  { readonly ok: true; readonly bytes: Uint8Array } | { readonly ok: false; readonly error: string }

/** The result of the native open flow: the chosen path plus its bounded
 * read (null = the user cancelled the dialog). */
export interface OpenFileSelection {
  readonly path: string
  readonly read: NativeReadResult
}

/** The unsaved-changes dialog answer. */
export type DiscardChoice = 'save' | 'discard' | 'cancel'

/** Host platform/version echo (presentation only, never semantics). */
export interface DesktopAppInfo {
  readonly platform: NodeJS.Platform
  readonly version: string
}

/**
 * The typed surface the preload exposes as `window.projectDesktop`. The
 * renderer depends on THIS interface (injected into the host controller),
 * never on `electron` — which keeps the renderer binding unit-testable
 * under jsdom with an in-memory fake.
 */
export interface ProjectDesktopBridge {
  /** Native open dialog (`.gproj` / MSPDI XML filters) + the bounded read. */
  pickOpenFile(): Promise<OpenFileSelection | null>
  /** Native save dialog; returns the chosen path or null (cancelled). */
  pickSaveFile(defaultName: string): Promise<string | null>
  /** The bounded read at an absolute path (argv / second-instance opens).
   * Read errors are returned, never thrown — and never uncapped bytes. */
  readFile(path: string): Promise<NativeReadResult>
  /** Writes raw bytes atomically-enough for the desktop flow; errors are
   * returned, never thrown, so the renderer's status surface can show them. */
  writeFile(path: string, bytes: Uint8Array): Promise<{ ok: boolean; error?: string }>
  /** The unsaved-changes dialog: Save / Don't Save / Cancel. */
  confirmDiscard(projectName: string): Promise<DiscardChoice>
  /** Host platform/version echo. */
  appInfo(): Promise<DesktopAppInfo>
  /** Native menu command activation (main → renderer). */
  onMenuCommand(handler: (command: MenuCommandId) => void): void
  /** The window close was requested; the renderer decides (close guard). */
  onCloseRequested(handler: () => void): void
  /** Approves a pending close (renderer → main). */
  approveClose(): void
  /** An external open request: argv at launch, or a second instance. */
  onOpenRequested(handler: (path: string) => void): void
}

/** The dialog filter vocabulary shared by the open/save native dialogs. */
export const PROJECT_FILE_FILTERS: Array<{ name: string; extensions: string[] }> = [
  { name: 'GenOffice Project Document', extensions: ['gproj'] },
  { name: 'MSPDI XML (MS Project interchange)', extensions: ['xml'] },
]
