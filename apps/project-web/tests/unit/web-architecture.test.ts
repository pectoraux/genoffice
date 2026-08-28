/**
 * PROJECT-028 — the web host architecture discipline suite.
 *
 * Static source guards pinning the web shell's structural contract:
 *
 * 1. NO module in the web host imports Electron or Node APIs — the
 *    browser-side implementation is structurally incapable of reaching a
 *    Node/Electron surface (architecture-lock §3; the CI web-e2e job
 *    additionally scans the BUILT bundle, and the E2E boots that bundle in
 *    a real chromium).
 * 2. The web host has NO scheduling authority and constructs no
 *    `ProjectCommand` literal: semantics live in the shared host binding
 *    (`@genoffice/project-host`), which the entry consumes verbatim.
 * 3. The web bridge is the ONE transport: the shared `ProjectHostBridge`
 *    contract implemented over browser primitives, with the ONE bounded
 *    web read routed by both read surfaces (the size check precedes any
 *    read; the error variant carries no bytes).
 * 4. The web menu bar's vocabulary is EXACTLY the shared
 *    `MENU_COMMAND_IDS` (transport lockstep with the desktop native menu).
 *
 * Raw-source scans (vitest `?raw` imports) — the accepted discipline-suite
 * pattern of the Project packages.
 */
import { describe, expect, it } from 'vitest'
import mainEntry from '../../src/main.ts?raw'
import menuSource from '../../src/menu.ts?raw'
import webBridgeSource from '../../src/web-bridge.ts?raw'
import webCssSource from '../../src/web.css?raw'
import { createWebBridge } from '../../src/web-bridge.js'
import type { WebBridge } from '../../src/web-bridge.js'
import type { ProjectHostBridge } from '@genoffice/project-host'
import { MENU_COMMAND_IDS } from '@genoffice/project-host'
import { WEB_MENU_COMMAND_IDS } from '../../src/menu.js'

const webSources: Record<string, string> = {
  'main.ts': mainEntry,
  'menu.ts': menuSource,
  'web-bridge.ts': webBridgeSource,
}

describe('no Node/Electron in the browser-side implementation', () => {
  it('no web host module imports Electron or Node APIs', () => {
    for (const [name, source] of Object.entries(webSources)) {
      expect(
        /from ['"](electron|node:)/.test(source),
        `${name} must not import Electron/Node APIs (the browser shell is transport only)`,
      ).toBe(false)
      expect(source.includes('require('), `${name} must not use CommonJS require`).toBe(false)
      expect(
        /\bprocess\.(env|cwd|platform)\b/.test(source),
        `${name} must not read the Node process surface`,
      ).toBe(false)
    }
  })

  it('the web host declares only the shared host binding as a runtime dependency', () => {
    expect(webSources['main.ts']).toContain("from '@genoffice/project-host'")
    for (const [name, source] of Object.entries(webSources)) {
      expect(
        source.includes('@genoffice/project-scheduling'),
        `${name} must not import the scheduling package (the web host has no scheduling authority — the shared binding injects it)`,
      ).toBe(false)
      expect(
        source.includes('@genoffice/project-engine'),
        `${name} must not import the engine package (semantics live in the shared binding)`,
      ).toBe(false)
    }
  })
})

describe('the shared binding is consumed verbatim', () => {
  it('the entry mounts the shared controller + stylesheet and the keyboard path is the shared table', () => {
    expect(mainEntry.includes("from '@genoffice/project-host'")).toBe(true)
    expect(mainEntry.includes("import '@genoffice/project-host/styles.css'")).toBe(true)
    expect(mainEntry.includes('createProjectApp')).toBe(true)
    expect(mainEntry.includes('translateKeyDown')).toBe(true)
    expect(mainEntry.includes("'keydown'")).toBe(true)
    // The entry never re-implements host binding surface.
    expect(mainEntry.includes('function createUI')).toBe(false)
    expect(mainEntry.includes('function createProjectApp')).toBe(false)
  })

  it('the entry wires the dirty probe, the menu dispatch, and the drag-and-drop open', () => {
    expect(mainEntry.includes('setDirtyProbe')).toBe(true)
    expect(mainEntry.includes('dispatchMenuCommand')).toBe(true)
    expect(mainEntry.includes('stageExternalFile')).toBe(true)
    expect(mainEntry.includes("'drop'")).toBe(true)
  })

  it('the web bridge satisfies the shared ProjectHostBridge contract', () => {
    // Compile-time: the created bridge IS a ProjectHostBridge (plus the
    // web-only seams) — the same typed surface the desktop preload exposes.
    const bridge: WebBridge = createWebBridge()
    const asHostBridge: ProjectHostBridge = bridge
    expect(asHostBridge).toBeDefined()
  })
})

describe('the ONE bounded web read (the PROJECT-027 desktop invariant, mirrored)', () => {
  it('the cap is defined exactly once (web-bridge.ts)', () => {
    expect(webBridgeSource.includes('export const MAX_WEB_FILE_BYTES')).toBe(true)
    for (const [name, source] of Object.entries(webSources)) {
      if (name === 'web-bridge.ts') continue
      expect(
        source.includes('MAX_WEB_FILE_BYTES ='),
        `${name} must not re-define the web transport cap (web-bridge.ts owns it)`,
      ).toBe(false)
    }
  })

  it('both read surfaces route through the ONE readCapped helper', () => {
    expect(webBridgeSource.includes('await readCapped(file)')).toBe(true)
    // Exactly two readCapped call sites: the picker path and the
    // staged-external (readFile) path.
    expect((webBridgeSource.match(/readCapped\(/g) ?? []).length).toBe(3)
  })

  it('the size check precedes any read (the stat-first analog)', () => {
    const sizeCheckAt = webBridgeSource.indexOf('file.size > MAX_WEB_FILE_BYTES')
    const readAt = webBridgeSource.indexOf('await file.arrayBuffer()')
    expect(sizeCheckAt).toBeGreaterThanOrEqual(0)
    expect(readAt).toBeGreaterThan(sizeCheckAt)
  })

  it('no web host module computes dates or builds a ProjectCommand literal', () => {
    const dateMarkers = ['Date.now(', 'new Date(', 'Date.UTC(', 'toISOString(', 'getTime()']
    for (const [name, source] of Object.entries(webSources)) {
      for (const marker of dateMarkers) {
        expect(
          source.includes(marker),
          `${name} must not compute dates ("${marker}") — display formats slice ISO strings`,
        ).toBe(false)
      }
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
      expect(commandLiteral.test(source), `${name} must not hand-build ProjectCommand values`).toBe(
        false,
      )
    }
  })
})

describe('the menu transport lockstep', () => {
  it('the web menu vocabulary is exactly the shared MENU_COMMAND_IDS', () => {
    expect([...WEB_MENU_COMMAND_IDS].sort()).toEqual([...MENU_COMMAND_IDS].sort())
    // The dropdown items carry the shared ids as data attributes (the E2E's
    // activation path).
    expect(menuSource.includes('data-menu-id') || menuSource.includes('menuId')).toBe(true)
  })
})

describe('the shared ribbon (PROJECT-029)', () => {
  it('the web shell adds NO ribbon of its own — the shared binding owns it', () => {
    // The shared host binding's DOM layer mounts the ribbon in BOTH
    // shells; the web chrome stays the menu bar + the transport bridge.
    for (const [name, source] of Object.entries(webSources)) {
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
    // The entry still mounts ONLY the shared binding (which carries the
    // ribbon) — no additional host package is pulled in.
    expect(webSources['main.ts']).toContain("from '@genoffice/project-host'")
  })
})

describe('the shared dialogs (PROJECT-030)', () => {
  it('the web shell adds NO dialog of its own — the shared binding owns them', () => {
    // The shared host binding's dialog layer renders the unsaved-changes
    // dialog and the Task Information dialog in BOTH shells; the web
    // sources carry no dialog DOM, no dialog constructor, and no dialog
    // testids of their own.
    for (const [name, source] of Object.entries(webSources)) {
      expect(source.includes('gp-dialog'), `${name} must not build dialog DOM`).toBe(false)
      expect(
        source.includes('gp-web-dialog'),
        `${name} must not carry the pre-030 private web dialog`,
      ).toBe(false)
      expect(
        source.includes('confirmDiscardDialog'),
        `${name} must not implement a dialog (the shared binding renders them)`,
      ).toBe(false)
      expect(
        source.includes('confirmUnsavedChanges'),
        `${name} must not construct the shared dialog directly (the controller owns it)`,
      ).toBe(false)
      expect(
        source.includes('createTaskInformationDialog'),
        `${name} must not construct the shared task dialog (the controller owns it)`,
      ).toBe(false)
      for (const testid of ['discard-save', 'task-info-ok']) {
        expect(
          source.includes(testid),
          `${name} must not own dialog testids ("${testid}" — the shared layer's)`,
        ).toBe(false)
      }
    }
  })

  it('the web stylesheet carries no dialog chrome (the shared stylesheet owns it)', () => {
    expect(webCssSource.includes('gp-web-dialog')).toBe(false)
    expect(webCssSource.includes('gp-dialog')).toBe(false)
  })
})
