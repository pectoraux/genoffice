import { defineConfig } from 'electron-vite'

/**
 * GenOffice Project desktop host build (PROJECT-027).
 *
 * The main process and preload carry NO `@genoffice/*` runtime dependencies:
 * the main process is native transport only (electron + node builtins), and
 * the preload imports the shared IPC module for channel identifiers — a
 * TypeScript source module that must be bundled (the workspace packages ship
 * TS source with no compiled entry point, the apps/docs + apps/sheets
 * pattern). The renderer binds the shared Project packages and bundles them
 * the same way.
 *
 * Every native-menu accelerator is DISPLAYED but not registered
 * (`registerAccelerator: false`, see src/main/menu.ts): the renderer's
 * keyboard translation is the single execution path for accelerator keys, so
 * an active cell editor can consume its own keys (native text undo, caret
 * movement) without menu commands firing mid-edit.
 */
export default defineConfig({
  main: {},
  preload: {},
  renderer: {},
})
