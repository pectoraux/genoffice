/**
 * The Project host transport contract — host-neutral (PROJECT-028).
 *
 * The typed boundary between the host controller (this package) and each
 * shell's transport. It contains TRANSPORT vocabulary only: menu command
 * ids, file-picker/read/write surfaces, and the close/open event
 * subscriptions. It contains no Project semantics and (since PROJECT-030,
 * the shared dialogs increment) no dialog surfaces either: the modal
 * dialogs are shared PRESENTATION rendered by this package's dialog layer
 * (`src/dialogs.ts`) in BOTH hosts — every semantic value that crosses a
 * bridge is either raw file bytes (the canonical file adapters run
 * host-side, in the controller) or an opaque path/name string.
 *
 * Two implementations exist, one per shell (architecture-lock §3):
 *
 * - the Electron preload bridge (`apps/project/src/preload`) crossing
 *   `contextBridge` IPC to the native main-process transport;
 * - the web bridge (`apps/project-web/src/web-bridge.ts`) over browser
 *   primitives (the File API, Blob downloads, `beforeunload`).
 *
 * The controller depends on THIS interface (injected at construction),
 * never on Electron or Node — which keeps the entire host binding
 * unit-testable under jsdom with an in-memory fake of this interface.
 */

/**
 * The host menu command vocabulary. Menu ids are TRANSPORT identifiers: the
 * shared translation layer (`src/translate.ts`) maps each id to the
 * canonical renderer-core intent/command action — a host menu never names a
 * Project semantic operation directly.
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
  | 'task.information'
  | 'task.indent'
  | 'task.outdent'
  | 'view.zoomIn'
  | 'view.zoomOut'
  | 'view.fit'
  | 'view.collapse'
  | 'view.expand'

/**
 * The complete host menu command vocabulary, in canonical menu order
 * (File → Edit → Task → View). Every host menu — the native desktop menu,
 * the web menu bar — carries exactly these ids; the shared translation
 * table covers every one of them (the discipline suites pin the lockstep).
 * `task.information` (PROJECT-030) opens the shared Task Information
 * dialog — the semantic dialog that operates on commands.
 */
export const MENU_COMMAND_IDS: readonly MenuCommandId[] = [
  'file.new',
  'file.open',
  'file.save',
  'file.saveAs',
  'edit.undo',
  'edit.redo',
  'edit.deleteTask',
  'task.create',
  'task.information',
  'task.indent',
  'task.outdent',
  'view.zoomIn',
  'view.zoomOut',
  'view.fit',
  'view.collapse',
  'view.expand',
]

/**
 * The result of a bounded host read: capped raw bytes, or the transport
 * error (oversized / missing / unreadable). Errors are VALUES, never
 * throws — and the error variant carries NO bytes, so the controller can
 * NEVER receive uncapped file contents: one `NativeReadResult` is the only
 * shape file content crosses a bridge in, on every read surface, in every
 * host.
 */
export type NativeReadResult =
  { readonly ok: true; readonly bytes: Uint8Array } | { readonly ok: false; readonly error: string }

/** The result of the host open flow: the chosen path plus its bounded
 * read (null = the user cancelled the picker). */
export interface OpenFileSelection {
  readonly path: string
  readonly read: NativeReadResult
}

/** Host platform/version echo (presentation only, never semantics). */
export interface HostAppInfo {
  readonly platform: string
  readonly version: string
}

/**
 * The typed bridge surface every host exposes to the controller. The
 * controller is constructed with an implementation of THIS interface and
 * never touches `electron`, `node:*`, or any host API directly.
 */
export interface ProjectHostBridge {
  /** The host open picker (`.gproj` / MSPDI XML filters) + the bounded read. */
  pickOpenFile(): Promise<OpenFileSelection | null>
  /** The host save picker; returns the chosen path/name or null (cancelled). */
  pickSaveFile(defaultName: string): Promise<string | null>
  /** The bounded read at a path/name (external open requests: argv /
   * second-instance on desktop, drag-and-drop on web). Read errors are
   * returned, never thrown — and never uncapped bytes. */
  readFile(path: string): Promise<NativeReadResult>
  /** Writes raw bytes; errors are returned, never thrown, so the
   * controller's status surface can show them. */
  writeFile(path: string, bytes: Uint8Array): Promise<{ ok: boolean; error?: string }>
  /** Host platform/version echo. */
  appInfo(): Promise<HostAppInfo>
  /** Host menu command activation (host chrome → controller). */
  onMenuCommand(handler: (command: MenuCommandId) => void): void
  /** A host close was requested; the controller decides (close guard). */
  onCloseRequested(handler: () => void): void
  /** Approves a pending close (controller → host). */
  approveClose(): void
  /** An external open request: argv at launch or a second instance
   * (desktop), drag-and-drop (web). */
  onOpenRequested(handler: (path: string) => void): void
}

/** The dialog filter vocabulary shared by the host open/save pickers. */
export const PROJECT_FILE_FILTERS: Array<{ name: string; extensions: string[] }> = [
  { name: 'GenOffice Project Document', extensions: ['gproj'] },
  { name: 'MSPDI XML (MS Project interchange)', extensions: ['xml'] },
]
