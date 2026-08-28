/**
 * The Project desktop host architecture discipline suite (PROJECT-027; the
 * renderer-binding scans moved with the modules to
 * `@genoffice/project-host`'s own suite at PROJECT-028).
 *
 * Static source guards pinning the DESKTOP transport's structural contract:
 *
 * 1. The MAIN process and PRELOAD are native transport only — NO
 *    `@genoffice/*` imports of any kind (direct or through the shared IPC
 *    module, which stays self-contained so the main-process bundle never
 *    pulls a `@genoffice/*` package).
 * 2. The native menu never registers an accelerator (the single keyboard
 *    translation path) and its ids are the complete `MenuCommandId`
 *    vocabulary — in lockstep with the shared contract's `MENU_COMMAND_IDS`.
 * 3. The preload bridge and the shared IPC contract stay in lockstep (every
 *    channel crossed), and the desktop bridge contract is structurally
 *    identical to the shared `ProjectHostBridge` the controller consumes.
 * 4. The corrected transport invariants (the ONE bounded read helper; the
 *    userData path installed BEFORE the single-instance lock).
 * 5. The renderer entry consumes the SHARED host binding — the desktop
 *    shell never re-implements controller/DOM/translation surface.
 *
 * Raw-source scans (vitest `?raw` imports) — the accepted discipline-suite
 * pattern of the Project packages.
 */
import { describe, expect, it } from 'vitest'
import mainIndex from '../../src/main/index.ts?raw'
import mainMenu from '../../src/main/menu.ts?raw'
import boundedRead from '../../src/main/bounded-read.ts?raw'
import preloadIndex from '../../src/preload/index.ts?raw'
import sharedIpc from '../../src/shared/ipc.ts?raw'
import mainEntry from '../../src/renderer/main.ts?raw'
import { MENU_COMMAND_IDS } from '../../src/main/menu.js'
import { PROJECT_IPC } from '../../src/shared/ipc.js'
import { MENU_COMMAND_IDS as SHARED_MENU_COMMAND_IDS } from '@genoffice/project-host'
import type { ProjectHostBridge } from '@genoffice/project-host'
import type { ProjectDesktopBridge } from '../../src/shared/ipc.js'

describe('the main process and preload are native transport only', () => {
  it('no main/preload module imports any @genoffice package', () => {
    const importPattern = /from\s+['"]@genoffice\//
    for (const [name, source] of Object.entries({
      'main/index.ts': mainIndex,
      'main/menu.ts': mainMenu,
      'preload/index.ts': preloadIndex,
      'shared/ipc.ts': sharedIpc,
    })) {
      expect(
        importPattern.test(source),
        `${name} must not import @genoffice packages (native transport only — the main-process bundle stays package-free)`,
      ).toBe(false)
    }
  })

  it('the main process owns the Electron surface (app/window/dialog/ipc) and node fs', () => {
    expect(mainIndex.includes("from 'electron'")).toBe(true)
    expect(mainIndex.includes("from 'node:fs/promises'")).toBe(true)
  })
})

describe('the native menu discipline', () => {
  it('every menu item displays its accelerator but registers NONE (single keyboard path)', () => {
    // All items are built through `itemFor`, which pins
    // `registerAccelerator: false`; the flag must never be flipped to true
    // (that would route accelerator keys around the editor-aware
    // translation table).
    expect(mainMenu.includes('registerAccelerator: false')).toBe(true)
    expect(mainMenu.includes('registerAccelerator: true')).toBe(false)
    expect(mainMenu.includes('registerAccelerator')).toBe(true)
  })

  it('the menu vocabulary is exactly MenuCommandId', () => {
    expect(new Set(MENU_COMMAND_IDS).size).toBe(MENU_COMMAND_IDS.length)
    const ids = [...mainMenu.matchAll(/id: '((?:file|edit|task|view)\.[a-zA-Z]+)'/g)].map(
      (match) => match[1]!,
    )
    expect(ids.sort()).toEqual([...MENU_COMMAND_IDS].sort())
  })

  it('the desktop menu vocabulary is in lockstep with the shared host contract', () => {
    expect([...MENU_COMMAND_IDS].sort()).toEqual([...SHARED_MENU_COMMAND_IDS].sort())
  })
})

describe('the bridge and the IPC contract stay in lockstep', () => {
  it('every PROJECT_IPC channel is crossed by the preload bridge', () => {
    for (const channelProperty of Object.keys(PROJECT_IPC)) {
      expect(
        preloadIndex.includes(`PROJECT_IPC.${channelProperty}`),
        `the preload must wire the "${channelProperty}" channel`,
      ).toBe(true)
    }
  })

  it('every PROJECT_IPC channel is served by the main process', () => {
    for (const channelProperty of Object.keys(PROJECT_IPC)) {
      const served =
        mainIndex.includes(`PROJECT_IPC.${channelProperty}`) ||
        mainMenu.includes(`PROJECT_IPC.${channelProperty}`)
      expect(served, `the main process must serve the "${channelProperty}" channel`).toBe(true)
    }
  })

  it('the preload exposes exactly the contextBridge surface', () => {
    expect(preloadIndex.includes('contextBridge.exposeInMainWorld')).toBe(true)
    expect((preloadIndex.match(/ipcRenderer\.invoke/g) ?? []).length).toBe(6)
  })

  it('the desktop bridge contract is structurally identical to the shared ProjectHostBridge', () => {
    // Compile-time structural equivalence, production direction (the
    // PROJECT-020 injected-runner precedent): the self-contained desktop
    // contract is assignable to the host-neutral contract the shared
    // controller consumes — the desktop bridge can never drift from the
    // contract `createProjectApp` owns. (The desktop `appInfo` narrows
    // `platform` to NodeJS.Platform, a subtype of the contract's string —
    // method covariance.)
    const desktopBridgeSatisfiesHost: (bridge: ProjectDesktopBridge) => ProjectHostBridge = (
      bridge,
    ) => bridge
    expect(desktopBridgeSatisfiesHost).toBeDefined()
  })
})

describe('the canonical bounded native read (PROJECT-027 correction)', () => {
  it('every native read routes through the ONE bounded helper (both IPC surfaces)', () => {
    expect(mainIndex.includes('boundedReadFile(')).toBe(true)
    // Exactly two call sites: the pickOpenFile handler and the readFile
    // (argv/second-instance) handler — the two read surfaces share one policy.
    expect((mainIndex.match(/boundedReadFile\(/g) ?? []).length).toBe(2)
  })

  it('the raw Node read primitive is not imported anywhere in the host main', () => {
    // With the raw readFile import gone, no handler can bypass the cap.
    expect(/import\s*\{[^}]*\breadFile\b[^}]*\}\s*from\s*'node:fs\/promises'/.test(mainIndex)).toBe(
      false,
    )
    expect(/\breadFile\(/.test(mainIndex)).toBe(false)
  })

  it('the transport cap is defined exactly once (bounded-read.ts)', () => {
    expect(boundedRead.includes('export const MAX_FILE_BYTES')).toBe(true)
    for (const [name, source] of Object.entries({
      'main/index.ts': mainIndex,
      'main/menu.ts': mainMenu,
      'preload/index.ts': preloadIndex,
      'shared/ipc.ts': sharedIpc,
      'renderer/main.ts': mainEntry,
    })) {
      expect(
        source.includes('MAX_FILE_BYTES ='),
        `${name} must not re-define the transport cap (the helper owns it)`,
      ).toBe(false)
    }
  })

  it('the bounded helper is pure Node transport (no Electron, no @genoffice)', () => {
    expect(/from 'electron'/.test(boundedRead)).toBe(false)
    expect(/from '@genoffice\//.test(boundedRead)).toBe(false)
  })

  it('userData is installed BEFORE the single-instance lock is requested', () => {
    // The lock is keyed on the userData path: installing the path after the
    // lock request would key every scratch/isolated profile's lock on the
    // REAL profile (the review finding — E2E isolation would be fictional).
    const setPathAt = mainIndex.indexOf("app.setPath('userData'")
    const lockAt = mainIndex.indexOf('app.requestSingleInstanceLock()')
    expect(setPathAt).toBeGreaterThanOrEqual(0)
    expect(lockAt).toBeGreaterThan(setPathAt)
  })
})

describe('the renderer entry consumes the shared host binding (PROJECT-028)', () => {
  it('the entry mounts the controller and stylesheet from @genoffice/project-host', () => {
    expect(mainEntry.includes("from '@genoffice/project-host'")).toBe(true)
    expect(mainEntry.includes("import '@genoffice/project-host/styles.css'")).toBe(true)
    expect(mainEntry.includes('createProjectApp')).toBe(true)
  })

  it('keyboard events flow through the translation table (capture, no raw handlers)', () => {
    expect(mainEntry.includes('translateKeyDown')).toBe(true)
    expect(mainEntry.includes("'keydown'")).toBe(true)
  })

  it('the entry never touches Electron or Node APIs directly', () => {
    expect(/from ['"](electron|node:)/.test(mainEntry)).toBe(false)
    expect(mainEntry.includes('require(')).toBe(false)
  })

  it('the desktop shell does not re-implement the shared host binding', () => {
    // The renderer surface is the ENTRY ONLY — the controller, DOM layer,
    // translation tables, document flows, and scheduling bindings live in
    // the shared package both shells consume.
    expect(mainEntry.includes('function createUI')).toBe(false)
    expect(mainEntry.includes('function createProjectApp')).toBe(false)
  })
})

describe('the shared ribbon (PROJECT-029)', () => {
  it('the desktop shell adds NO ribbon of its own — the shared binding owns it', () => {
    // The renderer surface is the ENTRY ONLY: the shared host binding's
    // DOM layer mounts the ribbon in BOTH shells; the desktop never
    // re-renders, re-interprets, or augments it.
    for (const [name, source] of [
      ['renderer/main.ts', mainEntry],
      ['main/menu.ts', mainMenu],
      ['main/index.ts', mainIndex],
      ['preload/index.ts', preloadIndex],
      ['shared/ipc.ts', sharedIpc],
    ] as const) {
      expect(source.includes('gp-ribbon'), `${name} must not build ribbon DOM`).toBe(false)
      expect(
        source.includes('createRibbon'),
        `${name} must not construct a ribbon (the shared binding renders it)`,
      ).toBe(false)
      expect(
        source.includes('RIBBON_'),
        `${name} must not define a private ribbon vocabulary (the shared contract owns it)`,
      ).toBe(false)
    }
  })
})
