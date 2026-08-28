/**
 * PROJECT-027 — the host controller battery (jsdom + the in-memory bridge
 * fake + the REAL engine/scheduler/adapters through the app's own
 * bindings).
 *
 * Proves the desktop binding end-to-end at the unit layer: boot render,
 * task creation through the renderer-core builder, cell editing through the
 * canonical edit flow with the SCHEDULE moving as the authority derives it,
 * undo/redo, dirty tracking by document identity, open/save through the
 * canonical adapters over the bridge, the close-guard handshake, menu
 * command dispatch, and keyboard navigation.
 */
import { beforeEach, describe, expect, it } from 'vitest'
import { asTaskId, plainMinutes } from '@genoffice/project-contracts'
import { fitViewport } from '@genoffice/project-renderer-core'
import type {
  DiscardChoice,
  MenuCommandId,
  NativeReadResult,
  OpenFileSelection,
  ProjectDesktopBridge,
} from '../../src/shared/ipc.js'
import { createProjectDesktopApp } from '../../src/renderer/app.js'
import type { ProjectDesktopApp } from '../../src/renderer/app.js'
import { exportDocumentBytes, newProjectDocument } from '../../src/renderer/document.js'

const flush = async (): Promise<void> => {
  await new Promise((resolve) => setTimeout(resolve, 0))
}

/** The in-memory bridge fake: config-driven behavior, observable calls.
 * The read surfaces mirror the REAL contract: `NativeReadResult` values
 * (errors are values, never throws — the real main process returns the
 * bounded-read outcome for every read, so the fake's `readFile` never
 * rejects and never fabricates uncapped bytes). */
interface BridgeConfig {
  readonly pickOpenFile?: () => Promise<OpenFileSelection | null>
  readonly pickSavePath?: string | null
  readonly discardChoice?: DiscardChoice
  /** Overrides the in-memory readFile behavior (default: files map / ENOENT). */
  readonly readFile?: (path: string) => NativeReadResult
}

function fakeBridge(config: BridgeConfig = {}) {
  const files = new Map<string, Uint8Array>()
  const calls = {
    pickOpenFile: 0,
    pickSaveFile: 0,
    readFile: [] as string[],
    writeFile: [] as string[],
    confirmDiscard: [] as string[],
    approveClose: 0,
  }
  const bridge: ProjectDesktopBridge = {
    pickOpenFile: async () => {
      calls.pickOpenFile += 1
      return config.pickOpenFile ? config.pickOpenFile() : null
    },
    pickSaveFile: async () => {
      calls.pickSaveFile += 1
      return config.pickSavePath ?? null
    },
    readFile: async (path) => {
      calls.readFile.push(path)
      if (config.readFile) return config.readFile(path)
      const bytes = files.get(path)
      if (bytes === undefined) {
        return { ok: false, error: `ENOENT: no such file or directory, open '${path}'` }
      }
      return { ok: true, bytes }
    },
    writeFile: async (path, bytes) => {
      calls.writeFile.push(path)
      files.set(path, bytes)
      return { ok: true }
    },
    confirmDiscard: async (projectName) => {
      calls.confirmDiscard.push(projectName)
      return config.discardChoice ?? 'cancel'
    },
    appInfo: async () => ({ platform: 'linux', version: '0.1.0' }),
    onMenuCommand: () => {},
    onCloseRequested: () => {},
    approveClose: () => {
      calls.approveClose += 1
    },
    onOpenRequested: () => {},
  }
  return { bridge, files, calls }
}

const mount = (): HTMLElement => {
  const host = document.createElement('div')
  document.body.appendChild(host)
  return host
}

const rows = (): HTMLElement[] =>
  [...document.querySelectorAll('[data-testid="task-row"]')] as HTMLElement[]

const rowText = (row: HTMLElement): string => row.textContent ?? ''

const statusText = (): string =>
  (document.querySelector('[data-testid="status-text"]') as HTMLElement | null)?.textContent ?? ''

const dirtyIndicator = (): string | null =>
  (document.querySelector('[data-testid="dirty-indicator"]') as HTMLElement | null)?.dataset
    .dirty ?? null

const key = (
  app: ProjectDesktopApp,
  k: string,
  modifiers: Partial<{ ctrl: boolean; shift: boolean; alt: boolean }> = {},
): void => {
  app.keydown({
    key: k,
    ctrlOrMeta: modifiers.ctrl ?? false,
    shift: modifiers.shift ?? false,
    alt: modifiers.alt ?? false,
  })
}

beforeEach(() => {
  document.body.innerHTML = ''
})

describe('boot', () => {
  it('renders the shell chrome and the empty state for a new project', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    expect(document.querySelector('[data-testid="project-app"]')).not.toBeNull()
    expect(document.querySelector('[data-testid="empty-state"]')).not.toBeNull()
    expect(document.querySelector('[data-testid="task-grid"]')).not.toBeNull()
    expect(document.querySelector('[data-testid="timeline"]')).not.toBeNull()
    expect(document.querySelector('[data-testid="time-axis"]')).not.toBeNull()
    expect(document.querySelector('[data-testid="file-label"]')?.textContent).toBe('Untitled')
    expect(dirtyIndicator()).toBe('false')
  })
})

describe('task creation through the renderer-core builder', () => {
  it('Insert creates, selects, focuses, and schedules the default task', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const created = rows()
    expect(created).toHaveLength(1)
    expect(rowText(created[0]!)).toContain('New Task')
    expect(created[0]!.dataset.selected).toBe('true')
    expect(created[0]!.dataset.focused).toBe('true')
    // The canonical schedule derived the start/finish cells (non-empty).
    expect(rowText(created[0]!)).toMatch(/\d{4}-\d{2}-\d{2}/)
    // … and the Gantt bar exists with the authority's geometry.
    expect(document.querySelector('[data-testid="gantt-bar"]')).not.toBeNull()
    expect(document.querySelector('[data-testid="calendar-band"]')).not.toBeNull()
  })

  it('a second task lands as the second root row', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    key(app, 'Insert')
    expect(rows()).toHaveLength(2)
    expect(rows()[1]!.dataset.taskId).toBe('t2')
  })
})

describe('cell editing through the canonical edit flow', () => {
  it('F2 → draft → Enter renames through a RenameTask command', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    key(app, 'F2')
    const editor = document.querySelector('[data-testid="cell-editor"]') as HTMLInputElement
    expect(editor).not.toBeNull()
    expect(editor.dataset.taskId).toBe('t1')
    editor.value = 'Pour foundations'
    editor.dispatchEvent(new Event('input', { bubbles: true }))
    key(app, 'Enter')
    expect(document.querySelector('[data-testid="cell-editor"]')).toBeNull()
    expect(rowText(rows()[0]!)).toContain('Pour foundations')
    expect(app.dirty).toBe(true)
  })

  it('Escape cancels without a command', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const before = app.state.session
    key(app, 'F2')
    const editor = document.querySelector('[data-testid="cell-editor"]') as HTMLInputElement
    editor.value = 'Discarded name'
    editor.dispatchEvent(new Event('input', { bubbles: true }))
    key(app, 'Escape')
    expect(app.state.session).toBe(before)
    expect(rowText(rows()[0]!)).toContain('New Task')
  })

  it('a duration edit moves the derived finish (the real scheduler)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const finishBefore = rowText(rows()[0]!)
    // Direct field edit through the intent vocabulary (the dblclick path
    // maps to the same intent).
    app.execute({
      kind: 'intent',
      intent: { type: 'beginTaskEdit', taskId: asTaskId('t1'), field: 'duration' },
    })
    const editor = document.querySelector('[data-testid="cell-editor"]') as HTMLInputElement
    expect(editor.dataset.field).toBe('duration')
    editor.value = '1920'
    editor.dispatchEvent(new Event('input', { bubbles: true }))
    key(app, 'Enter')
    const finishAfter = rowText(rows()[0]!)
    expect(finishAfter).not.toBe(finishBefore)
    // 1920 working minutes = 4 working days at 480/day: the finish moves.
    expect(
      plainMinutes(
        app.state.session.schedule?.taskSchedules[asTaskId('t1')]?.duration ??
          app.state.session.document.tasks[0]!.duration,
      ),
    ).toBe(1920)
  })

  it('an invalid duration draft surfaces the canonical reason and ends the editor', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    app.execute({
      kind: 'intent',
      intent: { type: 'beginTaskEdit', taskId: asTaskId('t1'), field: 'duration' },
    })
    const editor = document.querySelector('[data-testid="cell-editor"]') as HTMLInputElement
    editor.value = 'two days'
    editor.dispatchEvent(new Event('input', { bubbles: true }))
    key(app, 'Enter')
    expect(statusText()).toContain('Invalid edit')
    expect(document.querySelector('[data-testid="cell-editor"]')).toBeNull()
    expect(plainMinutes(app.state.session.document.tasks[0]!.duration)).toBe(480)
  })
})

describe('history', () => {
  it('Ctrl+Z undoes; Ctrl+Shift+Z redoes (document + schedule + view)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    // The pristine template reference — captured before any command.
    const initialDocument = app.state.session.document
    key(app, 'Insert')
    key(app, 'F2')
    const editor = document.querySelector('[data-testid="cell-editor"]') as HTMLInputElement
    editor.value = 'Named'
    editor.dispatchEvent(new Event('input', { bubbles: true }))
    key(app, 'Enter')
    expect(rows()).toHaveLength(1)
    // Undo #1 restores the pre-RENAME document (the task exists, unnamed).
    key(app, 'z', { ctrl: true })
    expect(rows()).toHaveLength(1)
    expect(rowText(rows()[0]!)).toContain('New Task')
    // Undo #2 restores the pre-CREATE document — the EXACT initial reference.
    key(app, 'z', { ctrl: true })
    expect(rows()).toHaveLength(0)
    expect(app.state.session.document).toBe(initialDocument)
    expect(app.dirty).toBe(false)
    // Redo restores the CREATE (the rename is still in the future).
    key(app, 'z', { ctrl: true })
    key(app, 'Z', { ctrl: true, shift: true })
    expect(rows()).toHaveLength(1)
    expect(rowText(rows()[0]!)).toContain('New Task')
  })

  it('undo with empty history is an honest no-op', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'z', { ctrl: true })
    expect(statusText()).toContain('Nothing to undo')
  })
})

describe('dirty tracking by document identity', () => {
  it('save persists through the canonical adapter and clears the dirty flag', async () => {
    const { bridge, files, calls } = fakeBridge({ pickSavePath: '/tmp/save-target.gproj' })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    expect(dirtyIndicator()).toBe('true')
    await app.execute({ kind: 'file', action: 'save' })
    expect(calls.pickSaveFile).toBe(1)
    expect(calls.writeFile).toEqual(['/tmp/save-target.gproj'])
    expect(files.get('/tmp/save-target.gproj')).toBeDefined()
    expect(dirtyIndicator()).toBe('false')
    expect(document.querySelector('[data-testid="file-label"]')?.textContent).toBe(
      'save-target.gproj',
    )
  })

  it('a save-as to an .xml path switches the canonical format', async () => {
    const { bridge, calls } = fakeBridge({ pickSavePath: '/tmp/interchange.xml' })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    await app.execute({ kind: 'file', action: 'saveAs' })
    expect(calls.writeFile).toEqual(['/tmp/interchange.xml'])
    expect(app.state.format).toBe('mspdi')
    expect(
      (document.querySelector('[data-testid="format-label"]') as HTMLElement).textContent,
    ).toBe('MSPDI')
  })

  it('undo back to the saved document clears dirty (identity, not revision)', async () => {
    const { bridge } = fakeBridge({ pickSavePath: '/tmp/identity.gproj' })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    await app.execute({ kind: 'file', action: 'save' })
    expect(app.dirty).toBe(false)
    key(app, 'Insert')
    expect(app.dirty).toBe(true)
    key(app, 'z', { ctrl: true })
    // Restored the EXACT saved document reference → not dirty.
    expect(app.dirty).toBe(false)
  })
})

describe('open through the bridge + canonical adapter', () => {
  it('loads a .gproj file and renders its rows + schedule', async () => {
    const source = newProjectDocument('Opened Fixture')
    const exported = exportDocumentBytes(source, 'gproj')
    const { bridge, calls } = fakeBridge({
      pickOpenFile: async () => ({
        path: '/tmp/opened.gproj',
        read: { ok: true, bytes: exported.bytes },
      }),
    })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    await app.execute({ kind: 'file', action: 'open' })
    expect(calls.pickOpenFile).toBe(1)
    expect(document.querySelector('[data-testid="file-label"]')?.textContent).toBe('opened.gproj')
    expect(app.state.format).toBe('gproj')
    expect(app.state.filePath).toBe('/tmp/opened.gproj')
  })

  it('a failed import keeps the current document and surfaces the error', async () => {
    const { bridge } = fakeBridge({
      pickOpenFile: async () => ({
        path: '/tmp/broken.gproj',
        read: { ok: true, bytes: new TextEncoder().encode('not a project') },
      }),
      discardChoice: 'discard',
    })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const documentBefore = app.state.session.document
    await app.execute({ kind: 'file', action: 'open' })
    expect(app.state.session.document).toBe(documentBefore)
    expect(statusText()).toContain('Open failed')
  })

  it('the argv open path (onOpenRequested) loads through readFile', async () => {
    const source = newProjectDocument('Argv Doc')
    const exported = exportDocumentBytes(source, 'gproj')
    const { bridge, files, calls } = fakeBridge()
    files.set('/tmp/argv.gproj', exported.bytes)
    let openHandler: ((path: string) => void) | undefined
    const app = createProjectDesktopApp({
      bridge: {
        ...bridge,
        onOpenRequested: (handler) => {
          openHandler = handler
        },
      },
      root: mount(),
    })
    app.start()
    expect(openHandler).toBeDefined()
    await openHandler!('/tmp/argv.gproj')
    expect(calls.readFile).toEqual(['/tmp/argv.gproj'])
    expect(app.state.filePath).toBe('/tmp/argv.gproj')
    expect(document.querySelector('[data-testid="file-label"]')?.textContent).toBe('argv.gproj')
  })

  it('an oversized picker read surfaces the transport cap error and keeps the document', async () => {
    // The picker selected a file over the transport cap: the bridge returns
    // the bounded-read ERROR (no bytes ever crossed), and the current
    // document survives by reference.
    const { bridge, calls } = fakeBridge({
      pickOpenFile: async () => ({
        path: '/tmp/oversized.gproj',
        read: { ok: false, error: 'File exceeds the 104857600 byte limit' },
      }),
      discardChoice: 'discard',
    })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const documentBefore = app.state.session.document
    await app.execute({ kind: 'file', action: 'open' })
    expect(calls.pickOpenFile).toBe(1)
    expect(statusText()).toContain('Open failed')
    expect(statusText()).toContain('File exceeds the 104857600 byte limit')
    // The renderer never received content: the session is untouched.
    expect(app.state.session.document).toBe(documentBefore)
    expect(rows()).toHaveLength(1)
    expect(app.state.filePath).toBe(null)
  })

  it('an oversized argv read (onOpenRequested) surfaces the cap error, session untouched', async () => {
    const { bridge, calls } = fakeBridge({
      readFile: () => ({ ok: false, error: 'File exceeds the 104857600 byte limit' }),
      discardChoice: 'discard',
    })
    let openHandler: ((path: string) => void) | undefined
    const app = createProjectDesktopApp({
      bridge: {
        ...bridge,
        onOpenRequested: (handler) => {
          openHandler = handler
        },
      },
      root: mount(),
    })
    app.start()
    key(app, 'Insert')
    const documentBefore = app.state.session.document
    await openHandler!('/tmp/oversized-argv.gproj')
    expect(calls.readFile).toEqual(['/tmp/oversized-argv.gproj'])
    expect(statusText()).toContain('Open failed')
    expect(statusText()).toContain('File exceeds the 104857600 byte limit')
    expect(app.state.session.document).toBe(documentBefore)
    expect(rows()).toHaveLength(1)
  })

  it('a missing argv path surfaces the read error (ENOENT) without touching the session', async () => {
    const { bridge, calls } = fakeBridge()
    let openHandler: ((path: string) => void) | undefined
    const app = createProjectDesktopApp({
      bridge: {
        ...bridge,
        onOpenRequested: (handler) => {
          openHandler = handler
        },
      },
      root: mount(),
    })
    app.start()
    const documentBefore = app.state.session.document
    await openHandler!('/tmp/definitely-missing.gproj')
    expect(calls.readFile).toEqual(['/tmp/definitely-missing.gproj'])
    expect(statusText()).toContain('Open failed')
    expect(statusText()).toContain('ENOENT')
    expect(app.state.session.document).toBe(documentBefore)
    expect(document.querySelector('[data-testid="empty-state"]')).not.toBeNull()
  })
})

describe('the close guard handshake', () => {
  it('a clean document approves the close immediately', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    await app.handleCloseRequested()
    expect(calls.approveClose).toBe(1)
    expect(calls.confirmDiscard).toHaveLength(0)
  })

  it('a dirty document asks; Save persists then approves', async () => {
    const { bridge, calls } = fakeBridge({
      discardChoice: 'save',
      pickSavePath: '/tmp/close-save.gproj',
    })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    await app.handleCloseRequested()
    expect(calls.confirmDiscard).toHaveLength(1)
    expect(calls.writeFile).toEqual(['/tmp/close-save.gproj'])
    expect(calls.approveClose).toBe(1)
  })

  it('Cancel refuses the close (no save, no approval)', async () => {
    const { bridge, calls } = fakeBridge({ discardChoice: 'cancel' })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    await app.handleCloseRequested()
    expect(calls.writeFile).toEqual([])
    expect(calls.approveClose).toBe(0)
  })
})

describe('menu command dispatch', () => {
  it('menu commands drive the same actions as the keyboard (file.new)', async () => {
    const { bridge } = fakeBridge({ discardChoice: 'discard' })
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    app.menuCommand('file.new' as MenuCommandId)
    await flush()
    // The dirty discard path ran: the untitled template replaced the doc.
    expect(rows()).toHaveLength(0)
    expect(document.querySelector('[data-testid="file-label"]')?.textContent).toBe('Untitled')
  })

  it('view.fit refits the viewport through the canonical fitViewport', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    app.menuCommand('view.zoomOut' as MenuCommandId)
    app.menuCommand('view.fit' as MenuCommandId)
    // The reducer's fit intent result equals the canonical fitViewport()
    // output for this document + schedule — the authority's own answer.
    expect(app.state.viewState.viewport).toEqual(
      fitViewport(app.state.session.document, app.state.session.schedule),
    )
  })
})

describe('keyboard navigation + selection', () => {
  it('arrow keys walk the visible rows and shift extends', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    key(app, 'Insert')
    key(app, 'ArrowDown')
    expect(app.state.viewState.tasks.focusId).toBe(asTaskId('t2'))
    key(app, 'ArrowUp', { shift: true })
    // Extend from the anchor: both rows selected.
    expect([...app.state.viewState.tasks.taskIds].sort()).toEqual(['t1', 't2'])
  })

  it('Delete removes the selected tasks through the builder commands', () => {
    const { bridge } = fakeBridge()
    const app = createProjectDesktopApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    key(app, 'Insert')
    key(app, 'ArrowDown')
    key(app, 'Delete')
    // The focused row (t2) was selected → deleted; t1 remains.
    expect(rows()).toHaveLength(1)
    expect(rows()[0]!.dataset.taskId).toBe('t1')
    key(app, 'ArrowUp')
    key(app, 'Delete')
    expect(rows()).toHaveLength(0)
  })
})
