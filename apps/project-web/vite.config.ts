import { defineConfig } from 'vite'

/**
 * GenOffice Project web host build (PROJECT-028).
 *
 * A plain browser build of the web shell: the entry mounts the shared
 * @genoffice/project-host binding (workspace TS source, consumed the same
 * way the office web app consumes @genoffice/ui) — the renderer core,
 * engine, scheduler, and file adapters are bundled from source with NO
 * Node/Electron polyfills, so an accidental `node:`/`electron` import in
 * the browser-side implementation fails the build instead of silently
 * shipping. The dist bundle is additionally scanned in CI (the static
 * "no Node/Electron in the browser-side implementation" proof).
 */
export default defineConfig({
  build: {
    outDir: 'dist',
  },
  server: {
    port: 5190,
    strictPort: true,
  },
  preview: {
    port: 5190,
    strictPort: true,
  },
})
