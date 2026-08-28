/**
 * The shared host binding architecture discipline suite (PROJECT-028).
 *
 * Static source guards pinning the `@genoffice/project-host` structural
 * contract — the scans the PROJECT-027 desktop suite ran over the renderer
 * binding, moved with the modules and extended to the whole package:
 *
 * 1. `bindings.ts` is the ONLY module importing the scheduling package; the
 *    host binding never re-implements scheduling/calendar/allocation
 *    semantics (no working-time primitives, no leveler, no date arithmetic).
 * 2. The host binding never constructs a `ProjectCommand` literal —
 *    commands are built by the renderer-core builders/edit flows only.
 * 3. NO module in the package imports Electron or Node APIs (the package is
 *    the browser-safe binding both shells load; the shells' transports are
 *    injected through the bridge contract).
 * 4. The package's static dependencies are exactly the canonical four
 *    (contracts, engine-consumed renderer-core, file, scheduling) — no
 *    React, no Electron, no host packages.
 *
 * Raw-source scans (vitest `?raw` imports) — the accepted discipline-suite
 * pattern of the Project packages.
 */
import { describe, expect, it } from 'vitest'
import { MENU_COMMAND_IDS } from '../../src/bridge.js'
import packageManifestRaw from '../../package.json?raw'
import appSource from '../../src/app.ts?raw'
import bindings from '../../src/bindings.ts?raw'
import bridge from '../../src/bridge.ts?raw'
import documentSource from '../../src/document.ts?raw'
import indexSource from '../../src/index.ts?raw'
import ribbonSource from '../../src/ribbon.ts?raw'
import translateSource from '../../src/translate.ts?raw'
import uiSource from '../../src/ui.ts?raw'

const packageSources: Record<string, string> = {
  'app.ts': appSource,
  'bindings.ts': bindings,
  'bridge.ts': bridge,
  'document.ts': documentSource,
  'index.ts': indexSource,
  'ribbon.ts': ribbonSource,
  'translate.ts': translateSource,
  'ui.ts': uiSource,
}

describe('the renderer binding surface', () => {
  it('bindings.ts is the ONLY module importing the scheduling package', () => {
    for (const [name, source] of Object.entries(packageSources)) {
      if (name === 'bindings.ts') continue
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
    for (const [name, source] of Object.entries(packageSources)) {
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
    for (const [name, source] of Object.entries(packageSources)) {
      for (const marker of dateMarkers) {
        expect(
          source.includes(marker),
          `${name} must not compute dates ("${marker}") — display formats slice ISO strings`,
        ).toBe(false)
      }
    }
  })

  it('the host binding never constructs a ProjectCommand literal', () => {
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
    for (const [name, source] of Object.entries(packageSources)) {
      expect(
        commandLiteral.test(source),
        `${name} must not hand-build ProjectCommand values (renderer-core builders only)`,
      ).toBe(false)
    }
  })
})

describe('the package is browser-safe (the binding both shells load)', () => {
  it('no module imports Electron or Node APIs directly', () => {
    for (const [name, source] of Object.entries(packageSources)) {
      expect(
        /from ['"](electron|node:)/.test(source),
        `${name} must not import Electron/Node APIs (the bridge is the only transport)`,
      ).toBe(false)
      expect(source.includes('require('), `${name} must not use CommonJS require`).toBe(false)
    }
  })

  it('the package manifest carries exactly the canonical four dependencies', () => {
    const manifest = JSON.parse(packageManifestRaw) as { dependencies?: Record<string, string> }
    expect(Object.keys(manifest.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-file',
      '@genoffice/project-renderer-core',
      '@genoffice/project-scheduling',
    ])
  })

  it('the public surface exports the documented host binding vocabulary', () => {
    for (const name of [
      'createProjectApp',
      'ProjectHostApp',
      'HostAppState',
      'scheduleRunner',
      'workingTimeQuery',
      'allocationQuery',
      'ProjectHostBridge',
      'MENU_COMMAND_IDS',
      'PROJECT_FILE_FILTERS',
      'newProjectDocument',
      'importDocumentBytes',
      'exportDocumentBytes',
      'translateKeyDown',
      'translateMenuCommand',
      'createUI',
      'createRibbon',
      'RIBBON_COMMAND_IDS',
      'RIBBON_TABS',
      'RibbonState',
    ]) {
      expect(indexSource.includes(name), `the public surface must export "${name}"`).toBe(true)
    }
  })

  it('the menu command vocabulary is complete and unique (15 ids)', () => {
    expect(MENU_COMMAND_IDS.length).toBe(15)
    expect(new Set(MENU_COMMAND_IDS).size).toBe(15)
    for (const id of MENU_COMMAND_IDS) {
      expect(bridge).toContain(`'${id}'`)
    }
  })
})

describe('the shared ribbon (PROJECT-029)', () => {
  it('ribbon.ts imports ONLY the shared command vocabulary (no document, schedule, or projection type)', () => {
    const imports = [...ribbonSource.matchAll(/from '([^']+)'/g)].map((match) => match[1]!)
    expect(imports).toEqual(['./bridge.js'])
  })

  it('the controller routes ribbon activation through the SAME menu-command path (one translation)', () => {
    expect(appSource).toContain('onRibbonCommand: (command) => menuCommand(command)')
  })

  it('the shared DOM layer mounts the ribbon above the workspace (both hosts inherit it)', () => {
    expect(uiSource).toContain("const ribbonHost = el('div', 'gp-ribbon-host')")
    expect(uiSource).toContain('app.append(ribbonHost, workspace, statusbar)')
    expect(uiSource).toContain('ribbon.update({')
  })

  it('the ribbon reflects exactly the four presentation echoes — no other state input', () => {
    // The RibbonState shape: the four booleans, nothing else.
    expect(ribbonSource).toContain('readonly canUndo: boolean')
    expect(ribbonSource).toContain('readonly canRedo: boolean')
    expect(ribbonSource).toContain('readonly dirty: boolean')
    expect(ribbonSource).toContain('readonly hasSelection: boolean')
    // The module reads no engine/session/projection surface.
    for (const marker of ['session', 'document.tasks', 'schedule', 'projection', 'buildGantt']) {
      expect(
        ribbonSource.includes(marker),
        `ribbon.ts must not read "${marker}" (echoes only)`,
      ).toBe(false)
    }
  })
})
