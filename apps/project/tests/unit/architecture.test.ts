/**
 * PROJECT-027 — the desktop host architecture discipline suite.
 *
 * Static source guards pinning the host's structural contract:
 *
 * 1. The MAIN process and PRELOAD are native transport only — NO
 *    `@genoffice/*` imports of any kind.
 * 2. `bindings.ts` is the ONLY renderer module importing the scheduling
 *    package; the host never re-implements scheduling/calendar/allocation
 *    semantics (no working-time primitives, no leveler, no date arithmetic
 *    anywhere in host src).
 * 3. The host never constructs a `ProjectCommand` literal — commands are
 *    built by the renderer-core builders/edit flows only.
 * 4. The native menu never registers an accelerator (the single keyboard
 *    translation path) and its ids are the complete `MenuCommandId`
 *    vocabulary.
 * 5. The preload bridge and the shared IPC contract stay in lockstep (every
 *    channel crossed).
 *
 * Raw-source scans (vitest `?raw` imports) — the accepted discipline-suite
 * pattern of the Project packages.
 */
import { describe, expect, it } from 'vitest'
import mainIndex from '../../src/main/index.ts?raw'
import mainMenu from '../../src/main/menu.ts?raw'
import preloadIndex from '../../src/preload/index.ts?raw'
import sharedIpc from '../../src/shared/ipc.ts?raw'
import bindings from '../../src/renderer/bindings.ts?raw'
import appSource from '../../src/renderer/app.ts?raw'
import documentSource from '../../src/renderer/document.ts?raw'
import translateSource from '../../src/renderer/translate.ts?raw'
import uiSource from '../../src/renderer/ui.ts?raw'
import mainEntry from '../../src/renderer/main.ts?raw'
import { MENU_COMMAND_IDS } from '../../src/main/menu.js'
import { PROJECT_IPC } from '../../src/shared/ipc.js'
const rendererSources: Record<string, string> = {
  'app.ts': appSource,
  'document.ts': documentSource,
  'translate.ts': translateSource,
  'ui.ts': uiSource,
  'main.ts': mainEntry,
}

const allHostSources: Record<string, string> = {
  'main/index.ts': mainIndex,
  'main/menu.ts': mainMenu,
  'preload/index.ts': preloadIndex,
  'shared/ipc.ts': sharedIpc,
  'renderer/bindings.ts': bindings,
  ...rendererSources,
}

describe('the main process and preload are native transport only', () => {
  it('no main/preload module imports any @genoffice package', () => {
    const importPattern = /from\s+['"]@genoffice\//
    for (const [name, source] of Object.entries({
      'main/index.ts': mainIndex,
      'main/menu.ts': mainMenu,
      'preload/index.ts': preloadIndex,
    })) {
      expect(
        importPattern.test(source),
        `${name} must not import @genoffice packages (native transport only)`,
      ).toBe(false)
    }
  })

  it('the main process owns the Electron surface (app/window/dialog/ipc) and node fs', () => {
    expect(mainIndex.includes("from 'electron'")).toBe(true)
    expect(mainIndex.includes("from 'node:fs/promises'")).toBe(true)
  })
})

describe('the renderer binding surface', () => {
  it('bindings.ts is the ONLY module importing the scheduling package', () => {
    for (const [name, source] of Object.entries(rendererSources)) {
      expect(
        source.includes('@genoffice/project-scheduling'),
        `${name} must not import the scheduling package (bindings.ts is the single site)`,
      ).toBe(false)
    }
    expect(bindings.includes('@genoffice/project-scheduling')).toBe(true)
  })

  it('bindings.ts wires exactly the three canonical injections and nothing else', () => {
    const schedulingImports = [
      ...bindings.matchAll(/import\s*\{([^}]*)\}\s*from\s*'@genoffice\/project-scheduling'/g),
    ]
      .flatMap((match) => match[1]!.split(','))
      .map((name) => name.trim())
      .filter((name) => name.length > 0 && !name.startsWith('type'))
    expect(schedulingImports.sort()).toEqual([
      'resolveCalendar',
      'resourceAllocations',
      'schedule',
      'workingIntervals',
    ])
    // The three documented injection seams (PROJECT-021/025/026).
    expect(bindings).toContain('ScheduleRunner')
    expect(bindings).toContain('CalendarWorkingTimeQuery')
    expect(bindings).toContain('ResourceAllocationQuery')
  })

  it('no host module re-implements scheduling/calendar/allocation semantics', () => {
    const semanticMarkers = [
      'addWorkingTime',
      'subtractWorkingTime',
      'workingDuration',
      'isWorking(',
      'levelResources',
      'registerLeveler',
    ]
    for (const [name, source] of Object.entries(allHostSources)) {
      for (const marker of semanticMarkers) {
        expect(
          source.includes(marker),
          `${name} must not touch the canonical semantic primitive "${marker}"`,
        ).toBe(false)
      }
    }
  })

  it('no host module computes dates (the canonical time model is never re-derived)', () => {
    const dateMarkers = [
      'Date.now(',
      'new Date(',
      'Date.UTC(',
      'getUTCFullYear',
      'getUTCMonth',
      'getUTCDate',
      'toISOString(',
      'getTime()',
    ]
    for (const [name, source] of Object.entries(allHostSources)) {
      for (const marker of dateMarkers) {
        expect(
          source.includes(marker),
          `${name} must not compute dates ("${marker}") — display formats slice ISO strings`,
        ).toBe(false)
      }
    }
  })

  it('the host never constructs a ProjectCommand literal', () => {
    const commandTypes = [
      'CreateTask',
      'DeleteTask',
      'RenameTask',
      'IndentTask',
      'OutdentTask',
      'AddDependency',
      'RemoveDependency',
      'ChangeDependencyType',
      'ChangeLag',
      'SetTaskDuration',
      'SetTaskStart',
      'SetTaskFinish',
      'SetConstraint',
      'SetDeadline',
      'AssignResource',
      'UnassignResource',
      'SetAssignmentUnits',
      'SetPercentComplete',
      'CreateBaseline',
      'LevelResources',
    ]
    const commandLiteral = new RegExp(`\\{\\s*type:\\s*'(?:${commandTypes.join('|')})'`)
    for (const [name, source] of Object.entries({
      ...rendererSources,
      'renderer/bindings.ts': bindings,
    })) {
      expect(
        commandLiteral.test(source),
        `${name} must not hand-build ProjectCommand values (renderer-core builders only)`,
      ).toBe(false)
    }
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
})

describe('the renderer entry wiring', () => {
  it('keyboard events flow through the translation table (capture, no raw handlers)', () => {
    expect(mainEntry.includes('translateKeyDown')).toBe(true)
    expect(mainEntry.includes("'keydown'")).toBe(true)
  })

  it('the renderer never touches Electron or Node APIs directly', () => {
    for (const [name, source] of Object.entries(rendererSources)) {
      expect(
        /from ['"](electron|node:)/.test(source),
        `${name} must not import Electron/Node APIs (the bridge is the only transport)`,
      ).toBe(false)
      expect(source.includes('require('), `${name} must not use CommonJS require`).toBe(false)
    }
  })
})
