/**
 * PROJECT-028 — the web transport bridge battery (jsdom + the REAL shared
 * contract types): the ONE bounded web read (size-first rejection — the
 * browser analog of the desktop's stat-first bounded read), the
 * external-file (drag-and-drop) staging + readFile surface, the
 * three-button discard dialog, the BEFOREUNLOAD CLOSE GUARD (the
 * corrected lifecycle boundary: the unload event is purely synchronous
 * and NEVER initiates the controller's asynchronous close handshake —
 * asserted with a registered close-handler spy AND against the REAL
 * shared controller mounted on the REAL web bridge), the in-app close
 * request (the one firing path for the registered handshake), the
 * menu-command dispatch path, and the download save flow (with the
 * `URL.createObjectURL` seam jsdom lacks stubbed; the real download is
 * E2E-proven in chromium).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createWebBridge, MAX_WEB_FILE_BYTES, readCapped } from '../../src/web-bridge.js'
import type { WebBridge } from '../../src/web-bridge.js'
import type { MenuCommandId, NativeReadResult } from '@genoffice/project-host'
import { createProjectApp } from '@genoffice/project-host'

const flush = async (): Promise<void> => {
  await new Promise((resolve) => setTimeout(resolve, 0))
}

/** A File whose arrayBuffer is a spy (the no-read proof for oversized input). */
function spyFile(name: string, size: number): File & { arrayBuffer: ReturnType<typeof vi.fn> } {
  const arrayBuffer = vi.fn(async () => new ArrayBuffer(0))
  return { name, size, arrayBuffer } as unknown as File & { arrayBuffer: ReturnType<typeof vi.fn> }
}

describe('the ONE bounded web read (readCapped)', () => {
  it('reads a small file exactly', async () => {
    const file = new File([new Uint8Array([1, 2, 3, 4])], 'small.gproj')
    const result = await readCapped(file)
    expect(result).toEqual({ ok: true, bytes: new Uint8Array([1, 2, 3, 4]) })
  })

  it('reads an empty file (ok, zero bytes)', async () => {
    const file = new File([], 'empty.gproj')
    const result = await readCapped(file)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.bytes.byteLength).toBe(0)
  })

  it('accepts a file at EXACTLY the cap', async () => {
    const file = new File([new Uint8Array(MAX_WEB_FILE_BYTES)], 'exact.gproj')
    const result = await readCapped(file)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.bytes.byteLength).toBe(MAX_WEB_FILE_BYTES)
  })

  it('rejects cap+1 on the SIZE, with NO read and NO bytes in the result', async () => {
    const file = spyFile('oversized.gproj', MAX_WEB_FILE_BYTES + 1)
    const result = await readCapped(file)
    // The size check fires BEFORE a byte is read (the stat-first analog).
    expect(file.arrayBuffer).not.toHaveBeenCalled()
    expect(result).toEqual({
      ok: false,
      error: `File exceeds the ${MAX_WEB_FILE_BYTES} byte limit`,
    })
  })

  it('returns a read error as a VALUE (never throws, no bytes)', async () => {
    const failing = spyFile('broken.gproj', 10)
    failing.arrayBuffer.mockRejectedValue(new Error('read failed'))
    const result = await readCapped(failing)
    expect(result).toEqual({ ok: false, error: 'read failed' })
  })
})

describe('the external-file (drag-and-drop) open surface', () => {
  it('an unstaged name reads as a missing-file VALUE (never throws)', async () => {
    const bridge = createWebBridge()
    const result = await bridge.readFile('never-dropped.gproj')
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('never-dropped.gproj')
  })

  it('a staged file fires onOpenRequested with its name, then reads capped through readFile', async () => {
    const bridge = createWebBridge()
    const requested: string[] = []
    bridge.onOpenRequested((path) => requested.push(path))
    const bytes = new Uint8Array([9, 8, 7])
    const file = new File([bytes], 'dropped.gproj')
    bridge.stageExternalFile(file)
    expect(requested).toEqual(['dropped.gproj'])
    const read: NativeReadResult = await bridge.readFile('dropped.gproj')
    expect(read).toEqual({ ok: true, bytes })
  })

  it('an OVERSIZED staged file is refused by readFile with the cap error (no bytes cross)', async () => {
    const bridge = createWebBridge()
    const requested: string[] = []
    bridge.onOpenRequested((path) => requested.push(path))
    const file = spyFile('huge.gproj', MAX_WEB_FILE_BYTES + 1)
    bridge.stageExternalFile(file)
    expect(requested).toEqual(['huge.gproj'])
    const read = await bridge.readFile('huge.gproj')
    expect(file.arrayBuffer).not.toHaveBeenCalled()
    expect(read).toEqual({
      ok: false,
      error: `File exceeds the ${MAX_WEB_FILE_BYTES} byte limit`,
    })
  })
})

describe('the unsaved-changes dialog', () => {
  let bridge: WebBridge
  beforeEach(() => {
    bridge = createWebBridge()
  })
  afterEach(() => {
    document.querySelector('[data-testid="discard-dialog"]')?.remove()
  })

  it('Save / Don’t Save / Cancel resolve the contract choices', async () => {
    for (const [testid, choice] of [
      ['discard-save', 'save'],
      ['discard-dont-save', 'discard'],
      ['discard-cancel', 'cancel'],
    ] as const) {
      const pending = bridge.confirmDiscard('E2E Build')
      await flush()
      const button = document.querySelector<HTMLButtonElement>(`[data-testid="${testid}"]`)
      expect(button, `the ${testid} button must render`).not.toBeNull()
      button!.click()
      await expect(pending).resolves.toBe(choice)
      expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
    }
  })

  it('the dialog names the project', async () => {
    const pending = bridge.confirmDiscard('Tower Construction')
    await flush()
    expect(document.querySelector('[data-testid="discard-dialog"]')?.textContent).toContain(
      'Tower Construction',
    )
    document.querySelector<HTMLButtonElement>('[data-testid="discard-cancel"]')!.click()
    await expect(pending).resolves.toBe('cancel')
  })

  it('Escape cancels (the native dialog keyboard behavior)', async () => {
    const pending = bridge.confirmDiscard('E2E Build')
    await flush()
    document.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
    )
    await expect(pending).resolves.toBe('cancel')
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
  })
})

describe('the beforeunload close guard (purely synchronous)', () => {
  // ONE bridge per window — exactly the production shape (the entry
  // creates a single bridge for the page's lifetime); the guard states
  // are driven through the same bridge's probe surface. THE CORRECTED
  // LIFECYCLE BOUNDARY: the unload event NEVER initiates the controller's
  // asynchronous close handshake — the close-handler spy below must stay
  // at ZERO across every unload dispatch, clean or dirty; the unload
  // decision is the dirty probe alone.
  const bridge = createWebBridge()
  let closeRequests = 0
  let dirty = false
  bridge.onCloseRequested(() => {
    closeRequests += 1
  })
  bridge.setDirtyProbe(() => dirty)

  afterEach(() => {
    dirty = false // the shared window outlives this describe: never leave a stale preventing probe
    document.querySelector('[data-testid="discard-dialog"]')?.remove()
  })

  it('a CLEAN unload allows the page to leave — and NEVER invokes the controller close handler', () => {
    dirty = false
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(closeRequests).toBe(0)
    expect(event.defaultPrevented).toBe(false)
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
  })

  it('a DIRTY unload is prevented (the native leave confirmation) — and NEVER invokes the controller close handler', () => {
    dirty = true
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(closeRequests).toBe(0)
    expect(event.defaultPrevented).toBe(true)
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
  })

  it('a DIRTY unload creates NO orphaned discard dialog — not during the event, not after the queue drains', async () => {
    // The async close handshake is never BEGUN, so no DOM dialog can
    // appear mid-unload; flush the microtask/timer queue to prove no
    // deferred dialog materializes either.
    dirty = true
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    await flush()
    await flush()
    expect(closeRequests).toBe(0)
    expect(event.defaultPrevented).toBe(true)
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
  })

  it('approveClose is a no-op on the web host — it creates NO unload bypass (a dirty document still prompts)', () => {
    // The browser owns the unload decision (the dirty probe alone): no
    // approval may silently allow a dirty unload, and no pending host
    // close exists for an approval to release.
    dirty = true
    bridge.approveClose()
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(event.defaultPrevented).toBe(true)
    expect(closeRequests).toBe(0)
  })
})

describe('the in-app close request (the native window-close button’s web analog)', () => {
  it('requestClose fires the registered controller close handshake — the in-app path, where it CAN complete', () => {
    const bridge = createWebBridge()
    let closeRequests = 0
    bridge.onCloseRequested(() => {
      closeRequests += 1
    })
    expect(closeRequests).toBe(0)
    bridge.requestClose()
    expect(closeRequests).toBe(1)
  })

  it('requestClose with NO registered handler is a harmless no-op', () => {
    const bridge = createWebBridge()
    expect(() => bridge.requestClose()).not.toThrow()
  })
})

describe('the unload lifecycle NEVER starts the controller close flow (the REAL shared controller on the REAL web bridge)', () => {
  // The integration proof of the corrected boundary: the controller the
  // entry mounts (createProjectApp + start — which registers the REAL
  // close handshake through onCloseRequested) is driven through a real
  // edit, and a beforeunload dispatch must leave it completely untouched
  // — no handshake, no approval, no dialog, the dirty document intact.
  const mount = (): HTMLElement => {
    const root = document.createElement('div')
    document.body.appendChild(root)
    return root
  }
  const insert = (app: ReturnType<typeof createProjectApp>): void => {
    app.keydown({ key: 'Insert', ctrlOrMeta: false, shift: false, alt: false })
  }
  const undoAll = (app: ReturnType<typeof createProjectApp>): void => {
    // The shared window outlives each test: leave the app CLEAN so this
    // describe's bridges never prevent a LATER test's unload dispatch.
    app.keydown({ key: 'z', ctrlOrMeta: true, shift: false, alt: false })
  }

  afterEach(() => {
    document.body.innerHTML = ''
  })

  it('a CLEAN beforeunload proceeds, with NO close handshake and NO dialog', async () => {
    const root = mount()
    const bridge = createWebBridge()
    const approvals: string[] = []
    bridge.approveClose = () => approvals.push('approved')
    const app = createProjectApp({ bridge, root })
    bridge.setDirtyProbe(() => app.dirty)
    app.start()
    expect(app.dirty).toBe(false)

    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)

    expect(event.defaultPrevented).toBe(false)
    await flush()
    expect(approvals).toEqual([]) // NOT even the clean-close approval ran
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
  })

  it('a DIRTY beforeunload is prevented, with NO close handshake, NO approval, and NO orphaned discard dialog', async () => {
    const root = mount()
    const bridge = createWebBridge()
    const approvals: string[] = []
    bridge.approveClose = () => approvals.push('approved')
    const app = createProjectApp({ bridge, root })
    bridge.setDirtyProbe(() => app.dirty)
    app.start()
    insert(app)
    expect(app.dirty).toBe(true)

    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)

    expect(event.defaultPrevented).toBe(true)
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
    await flush()
    await flush()
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
    expect(approvals).toEqual([]) // the handshake never ran
    expect(app.dirty).toBe(true) // the document survives untouched
    undoAll(app)
    expect(app.dirty).toBe(false)
  })

  it('the IN-APP close request runs the REAL handshake to completion (the flow that CAN complete)', async () => {
    const root = mount()
    const bridge = createWebBridge()
    const approvals: string[] = []
    bridge.approveClose = () => approvals.push('approved')
    const app = createProjectApp({ bridge, root })
    bridge.setDirtyProbe(() => app.dirty)
    app.start()
    insert(app)
    expect(app.dirty).toBe(true)

    // In-app close → the shared controller's Save/Don't-Save/Cancel dialog.
    bridge.requestClose()
    await flush()
    expect(document.querySelector('[data-testid="discard-dialog"]')).not.toBeNull()

    // Don't Save → the handshake completes with the approval.
    document.querySelector<HTMLButtonElement>('[data-testid="discard-dont-save"]')!.click()
    await flush()
    expect(approvals).toEqual(['approved'])
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()

    // The page does NOT unload (the browser owns that): a subsequent
    // unload attempt still consults the dirty probe synchronously.
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(event.defaultPrevented).toBe(true)
    undoAll(app)
    expect(app.dirty).toBe(false)
  })
})

describe('the menu-command dispatch path', () => {
  it('the registered controller handler receives the dispatched id', () => {
    const bridge = createWebBridge()
    const received: MenuCommandId[] = []
    bridge.onMenuCommand((command) => received.push(command))
    bridge.dispatchMenuCommand('file.open')
    bridge.dispatchMenuCommand('task.create')
    expect(received).toEqual(['file.open', 'task.create'])
  })
})

describe('the download save flow', () => {
  afterEach(() => {
    // jsdom lacks URL.createObjectURL; tests may stub it.
    ;(URL as { createObjectURL?: unknown }).createObjectURL = undefined
    ;(URL as { revokeObjectURL?: unknown }).revokeObjectURL = undefined
  })

  it('downloads the bytes under the file name (ok value)', async () => {
    // The seam jsdom lacks, provided by the test (the real browser flow is
    // E2E-proven through Playwright's download event in chromium).
    const objects: string[] = []
    ;(URL as { createObjectURL?: (blob: Blob) => string }).createObjectURL = () => {
      const url = `blob:mock-${objects.length}`
      objects.push(url)
      return url
    }
    ;(URL as { revokeObjectURL?: (url: string) => void }).revokeObjectURL = (url: string) => {
      const index = objects.indexOf(url)
      if (index >= 0) objects.splice(index, 1)
    }

    const bridge = createWebBridge()
    const clicks: string[] = []
    document.addEventListener('click', (event) => {
      const target = event.target as HTMLElement
      if (target?.dataset?.testid === 'save-download') {
        clicks.push(target.getAttribute('download') ?? '')
      }
    })
    const outcome = await bridge.writeFile('saved-as.gproj', new Uint8Array([5, 6]))
    expect(outcome).toEqual({ ok: true })
    expect(clicks).toEqual(['saved-as.gproj'])
  })

  it('a write failure is an error VALUE (never a throw)', async () => {
    // No createObjectURL stub → the browser API is missing → the honest
    // error-value path (the controller's status surface shows it).
    const bridge = createWebBridge()
    const outcome = await bridge.writeFile('x.gproj', new Uint8Array([1]))
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toBeTruthy()
  })

  it('the picker returns the default save name (the download flow is the save dialog)', async () => {
    const bridge = createWebBridge()
    await expect(bridge.pickSaveFile('Project1.gproj')).resolves.toBe('Project1.gproj')
  })
})
