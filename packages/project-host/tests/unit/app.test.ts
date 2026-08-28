/**
 * The shared host controller battery (jsdom + the in-memory bridge fake +
 * the REAL engine/scheduler/adapters through the app's own bindings;
 * established as the PROJECT-027 desktop suite, moved with the controller
 * to `@genoffice/project-host` at PROJECT-028).
 *
 * Proves the host binding end-to-end at the unit layer: boot render,
 * task creation through the renderer-core builder, cell editing through the
 * canonical edit flow with the SCHEDULE moving as the authority derives it,
 * undo/redo, dirty tracking by document identity, open/save through the
 * canonical adapters over the bridge, the close-guard handshake, menu
 * command dispatch, keyboard navigation — and, since PROJECT-030, the
 * SHARED dialogs: the unsaved-changes dialog rendered by the host binding
 * itself (no bridge surface — the tests click its real DOM buttons) and
 * the Task Information dialog committing through the canonical command
 * pipeline.
 */
import { beforeEach, describe, expect, it } from 'vitest'
import { asTaskId, plainMinutes } from '@genoffice/project-contracts'
import { fitViewport } from '@genoffice/project-renderer-core'
import type {
  MenuCommandId,
  NativeReadResult,
  OpenFileSelection,
  ProjectHostBridge,
} from '../../src/bridge.js'
import { createProjectApp } from '../../src/app.js'
import type { ProjectHostApp } from '../../src/app.js'
import { exportDocumentBytes, newProjectDocument } from '../../src/document.js'

const flush = async (): Promise<void> => {
  await new Promise((resolve) => setTimeout(resolve, 0))
}

/** The in-memory bridge fake: config-driven behavior, observable calls.
 * The read surfaces mirror the REAL contract: `NativeReadResult` values
 * (errors are values, never throws — the real main process returns the
 * bounded-read outcome for every read, so the fake's `readFile` never
 * rejects and never fabricates uncapped bytes). Since PROJECT-030 the
 * fake carries NO dialog surface — the unsaved-changes dialog is shared
 * presentation the tests drive through its real DOM buttons. */
interface BridgeConfig {
  readonly pickOpenFile?: () => Promise<OpenFileSelection | null>
  readonly pickSavePath?: string | null
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
    approveClose: 0,
  }
  const bridge: ProjectHostBridge = {
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

/** Clicks a button of the SHARED unsaved-changes dialog (the real DOM —
 * the dialog opens synchronously inside the awaited flow's prefix). */
const clickDiscard = (testid: string): void => {
  const button = document.querySelector(`[data-testid="${testid}"]`) as HTMLElement | null
  expect(button, `the shared discard dialog must be open (${testid})`).not.toBeNull()
  button!.click()
}

const discardDialogOpen = (): boolean =>
  document.querySelector('[data-testid="discard-dialog"]') !== null

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
  app: ProjectHostApp,
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'z', { ctrl: true })
    expect(statusText()).toContain('Nothing to undo')
  })
})

describe('dirty tracking by document identity', () => {
  it('save persists through the canonical adapter and clears the dirty flag', async () => {
    const { bridge, files, calls } = fakeBridge({ pickSavePath: '/tmp/save-target.gproj' })
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    })
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const documentBefore = app.state.session.document
    const pending = app.execute({ kind: 'file', action: 'open' })
    expect(discardDialogOpen()).toBe(true)
    clickDiscard('discard-dont-save')
    await pending
    expect(app.state.session.document).toBe(documentBefore)
    expect(statusText()).toContain('Open failed')
  })

  it('the argv open path (onOpenRequested) loads through readFile', async () => {
    const source = newProjectDocument('Argv Doc')
    const exported = exportDocumentBytes(source, 'gproj')
    const { bridge, files, calls } = fakeBridge()
    files.set('/tmp/argv.gproj', exported.bytes)
    let openHandler: ((path: string) => void) | undefined
    const app = createProjectApp({
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
    })
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const documentBefore = app.state.session.document
    const pending = app.execute({ kind: 'file', action: 'open' })
    clickDiscard('discard-dont-save')
    await pending
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
    })
    let openHandler: ((path: string) => void) | undefined
    const app = createProjectApp({
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
    const pending = openHandler!('/tmp/oversized-argv.gproj') as unknown as Promise<void>
    clickDiscard('discard-dont-save')
    await pending
    expect(calls.readFile).toEqual(['/tmp/oversized-argv.gproj'])
    expect(statusText()).toContain('Open failed')
    expect(statusText()).toContain('File exceeds the 104857600 byte limit')
    expect(app.state.session.document).toBe(documentBefore)
    expect(rows()).toHaveLength(1)
  })

  it('a missing argv path surfaces the read error (ENOENT) without touching the session', async () => {
    const { bridge, calls } = fakeBridge()
    let openHandler: ((path: string) => void) | undefined
    const app = createProjectApp({
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

describe('the close guard handshake (over the SHARED unsaved-changes dialog)', () => {
  it('a clean document approves the close immediately (no dialog)', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    await app.handleCloseRequested()
    expect(calls.approveClose).toBe(1)
    expect(discardDialogOpen()).toBe(false)
  })

  it('a dirty document asks through the shared dialog; Save persists then approves', async () => {
    const { bridge, calls } = fakeBridge({ pickSavePath: '/tmp/close-save.gproj' })
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const closeRequest = app.handleCloseRequested()
    expect(discardDialogOpen()).toBe(true)
    clickDiscard('discard-save')
    await closeRequest
    expect(calls.writeFile).toEqual(['/tmp/close-save.gproj'])
    expect(calls.approveClose).toBe(1)
    expect(discardDialogOpen()).toBe(false)
  })

  it('“Don’t Save” approves without writing', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const closeRequest = app.handleCloseRequested()
    clickDiscard('discard-dont-save')
    await closeRequest
    expect(calls.writeFile).toEqual([])
    expect(calls.approveClose).toBe(1)
  })

  it('Cancel refuses the close (no save, no approval)', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const closeRequest = app.handleCloseRequested()
    clickDiscard('discard-cancel')
    await closeRequest
    expect(calls.writeFile).toEqual([])
    expect(calls.approveClose).toBe(0)
    expect(discardDialogOpen()).toBe(false)
  })

  it('Escape cancels the shared dialog (the native dialog keyboard behavior)', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const closeRequest = app.handleCloseRequested()
    document.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
    )
    await closeRequest
    expect(calls.approveClose).toBe(0)
    expect(discardDialogOpen()).toBe(false)
  })
})

describe('menu command dispatch', () => {
  it('menu commands drive the same actions as the keyboard (file.new)', async () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    app.menuCommand('file.new' as MenuCommandId)
    // The dirty discard path consulted the SHARED dialog; Don't Save loads
    // the untitled template.
    clickDiscard('discard-dont-save')
    await flush()
    expect(rows()).toHaveLength(0)
    expect(document.querySelector('[data-testid="file-label"]')?.textContent).toBe('Untitled')
  })

  it('view.fit refits the viewport through the canonical fitViewport', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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
    const app = createProjectApp({ bridge, root: mount() })
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

describe('the Task Information dialog (PROJECT-030 — operates on commands)', () => {
  const taskDialog = (): HTMLElement | null =>
    document.querySelector('[data-testid="task-info-dialog"]')
  const taskDialogOpen = (): boolean => taskDialog() !== null
  const nameInput = (): HTMLInputElement =>
    document.querySelector('[data-testid="task-info-name"]') as HTMLInputElement
  const durationInput = (): HTMLInputElement =>
    document.querySelector('[data-testid="task-info-duration"]') as HTMLInputElement
  const errorText = (): string =>
    (document.querySelector('[data-testid="task-info-error"]') as HTMLElement | null)
      ?.textContent ?? ''
  const clickTaskButton = (testid: string): void => {
    const button = document.querySelector(`[data-testid="${testid}"]`) as HTMLElement | null
    expect(button, `the task dialog must be open (${testid})`).not.toBeNull()
    button!.click()
  }
  const openDialog = (app: ProjectHostApp): void => {
    app.menuCommand('task.information')
    expect(taskDialogOpen()).toBe(true)
  }

  it('no selection → the honest status answer, no dialog', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    app.menuCommand('task.information')
    expect(taskDialogOpen()).toBe(false)
    expect(statusText()).toContain('No task selected')
  })

  it('opens with the focused task’s DISPLAYED values (drafts + read-only schedule)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    expect(taskDialog()!.dataset.taskId).toBe('t1')
    expect(nameInput().value).toBe('New Task')
    expect(durationInput().value).toBe('480')
    expect(durationInput().disabled).toBe(false)
    // The displayed schedule instants (sliced ISO — the shared convention).
    const start = (document.querySelector('[data-testid="task-info-start"]') as HTMLElement)
      .textContent
    const finish = (document.querySelector('[data-testid="task-info-finish"]') as HTMLElement)
      .textContent
    expect(start).toMatch(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/)
    expect(finish).toMatch(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/)
    // Focus starts on the name input (the primary field).
    expect(document.activeElement).toBe(nameInput())
  })

  it('a summary task opens with the duration input disabled (the 023 rule)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    key(app, 'Insert')
    key(app, 'ArrowDown')
    key(app, 'ArrowRight', { alt: true, shift: true }) // indent t2 under t1 → t1 summary
    key(app, 'ArrowUp') // focus t1 (the summary)
    openDialog(app)
    expect(taskDialog()!.dataset.taskId).toBe('t1')
    expect(durationInput().disabled).toBe(true)
    expect(durationInput().getAttribute('aria-disabled')).toBe('true')
  })

  it('OK commits name + duration as TWO commands through the real session', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    nameInput().value = 'Pour foundations'
    durationInput().value = '1920'
    clickTaskButton('task-info-ok')
    // The dialog closed; both fields landed through the canonical flow.
    expect(taskDialogOpen()).toBe(false)
    const task = app.state.session.document.tasks[0]!
    expect(task.name).toBe('Pour foundations')
    expect(plainMinutes(task.duration)).toBe(1920)
    // The REAL scheduler derived the widened schedule.
    expect(
      plainMinutes(
        app.state.session.schedule?.taskSchedules[asTaskId('t1')]?.duration ?? task.duration,
      ),
    ).toBe(1920)
    // THREE journaled commands total: the boot Insert + the dialog's two
    // (RenameTask + SetTaskDuration) — the dialog adds no journal shape
    // of its own.
    expect(app.state.session.past).toHaveLength(3)
    expect(app.dirty).toBe(true)
  })

  it('undo twice restores the exact pre-dialog document reference', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const before = app.state.session.document
    openDialog(app)
    nameInput().value = 'Renamed'
    durationInput().value = '960'
    clickTaskButton('task-info-ok')
    key(app, 'z', { ctrl: true }) // undo the duration
    key(app, 'z', { ctrl: true }) // undo the rename
    expect(app.state.session.document).toBe(before)
  })

  it('an unparseable duration keeps the dialog open with the reason; nothing applied', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const before = app.state.session
    openDialog(app)
    durationInput().value = 'two days'
    clickTaskButton('task-info-ok')
    expect(taskDialogOpen()).toBe(true)
    expect(errorText()).toContain('Invalid edit')
    // No command crossed the session (the name was unchanged → noChange).
    expect(app.state.session).toBe(before)
    expect(statusText()).toContain('Invalid edit')
  })

  it('the engine is the single semantic validator: a negative duration refuses with INVALID_DURATION', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    durationInput().value = '-5'
    clickTaskButton('task-info-ok')
    expect(taskDialogOpen()).toBe(true)
    expect(errorText()).toContain('INVALID_DURATION')
    expect(plainMinutes(app.state.session.document.tasks[0]!.duration)).toBe(480)
  })

  it('partial honesty: an earlier field applies; the refused one stays for the fix', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    nameInput().value = 'Applies First'
    durationInput().value = 'two days'
    clickTaskButton('task-info-ok')
    // The name (field 1) applied as a real command; the duration refused.
    expect(taskDialogOpen()).toBe(true)
    expect(errorText()).toContain('Invalid edit')
    expect(app.state.session.document.tasks[0]!.name).toBe('Applies First')
    // The boot Insert + the applied rename — the refused duration added
    // nothing.
    expect(app.state.session.past).toHaveLength(2)
    // The fix: only the duration remains (the name now matches).
    durationInput().value = '1440'
    clickTaskButton('task-info-ok')
    expect(taskDialogOpen()).toBe(false)
    expect(plainMinutes(app.state.session.document.tasks[0]!.duration)).toBe(1440)
    expect(app.state.session.past).toHaveLength(3)
  })

  it('Cancel and Escape mutate nothing (no command, no journal)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const before = app.state.session
    openDialog(app)
    nameInput().value = 'Discarded'
    durationInput().value = '9999'
    clickTaskButton('task-info-cancel')
    expect(taskDialogOpen()).toBe(false)
    expect(app.state.session).toBe(before)
    // Escape path.
    openDialog(app)
    nameInput().value = 'Also Discarded'
    document.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
    )
    expect(taskDialogOpen()).toBe(false)
    expect(app.state.session).toBe(before)
  })

  it('Enter submits (the form convention)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    nameInput().value = 'Enter Committed'
    nameInput().dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    )
    expect(taskDialogOpen()).toBe(false)
    expect(app.state.session.document.tasks[0]!.name).toBe('Enter Committed')
  })

  it('noChange: OK with unchanged displayed values dispatches NOTHING', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    const before = app.state.session
    openDialog(app)
    clickTaskButton('task-info-ok')
    expect(taskDialogOpen()).toBe(false)
    expect(app.state.session).toBe(before)
  })

  it('the modal gate: commands and keys are suspended while open; the open command refreshes', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    // Suspended: menu + keyboard dispatch never reach the session.
    app.menuCommand('task.create' as MenuCommandId)
    key(app, 'Insert')
    key(app, 'z', { ctrl: true })
    expect(rows()).toHaveLength(1)
    expect(app.state.session.past).toHaveLength(1) // only the boot Insert
    expect(taskDialogOpen()).toBe(true)
    // The dialog's OWN command re-opens (refresh from the current document).
    nameInput().value = 'Suspended Typing'
    app.menuCommand('task.information')
    expect(taskDialogOpen()).toBe(true)
    expect(nameInput().value).toBe('New Task') // refreshed — drafts discarded
    // After cancel the app resumes.
    clickTaskButton('task-info-cancel')
    app.menuCommand('task.create' as MenuCommandId)
    expect(rows()).toHaveLength(2)
  })

  it('the close handshake still runs while the dialog is open (lifecycle, not command)', async () => {
    const { bridge, calls } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    const closeRequest = app.handleCloseRequested()
    // The shared unsaved-changes dialog stacks over the task dialog.
    expect(discardDialogOpen()).toBe(true)
    clickDiscard('discard-dont-save')
    await closeRequest
    expect(calls.approveClose).toBe(1)
  })

  it('loading a document closes the dialog (the target context is gone)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    const source = newProjectDocument('Replacement')
    const exported = exportDocumentBytes(source, 'gproj')
    app.openPath('/tmp/replacement.gproj', exported.bytes)
    expect(taskDialogOpen()).toBe(false)
  })

  it('opening the dialog ends an active cell edit (the dialog supersedes it)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    key(app, 'F2')
    expect(document.querySelector('[data-testid="cell-editor"]')).not.toBeNull()
    openDialog(app)
    expect(document.querySelector('[data-testid="cell-editor"]')).toBeNull()
    expect(app.state.viewState.editing).toBeUndefined()
  })

  it('the keyboard translation still covers the full vocabulary (task.information → dialog)', () => {
    const { bridge } = fakeBridge()
    const app = createProjectApp({ bridge, root: mount() })
    app.start()
    key(app, 'Insert')
    openDialog(app)
    clickTaskButton('task-info-cancel')
    // Reaching the same action through execute (the one dispatch pipeline).
    void app.execute({ kind: 'dialog', action: 'taskInformation' })
    expect(taskDialogOpen()).toBe(true)
  })
})
