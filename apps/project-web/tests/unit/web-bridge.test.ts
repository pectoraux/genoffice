/**
 * PROJECT-028 — the web transport bridge battery (jsdom + the REAL shared
 * contract types): the ONE bounded web read (size-first rejection — the
 * browser analog of the desktop's stat-first bounded read), the
 * external-file (drag-and-drop) staging + readFile surface, the
 * three-button discard dialog, the beforeunload close guard, the
 * menu-command dispatch path, and the download save flow (with the
 * `URL.createObjectURL` seam jsdom lacks stubbed; the real download is
 * E2E-proven in chromium).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createWebBridge, MAX_WEB_FILE_BYTES, readCapped } from '../../src/web-bridge.js'
import type { WebBridge } from '../../src/web-bridge.js'
import type { MenuCommandId, NativeReadResult } from '@genoffice/project-host'

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

describe('the beforeunload close guard', () => {
  // ONE bridge per window — exactly the production shape (the entry creates
  // a single bridge for the page's lifetime); the guard states are driven
  // through the same bridge's probe/approval surfaces.
  const bridge = createWebBridge()
  let closed = 0
  let dirty = false
  bridge.onCloseRequested(() => {
    closed += 1
  })
  bridge.setDirtyProbe(() => dirty)

  it('a clean document unloads without the native confirmation', () => {
    dirty = false
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(closed).toBe(1)
    expect(event.defaultPrevented).toBe(false)
  })

  it('a DIRTY document triggers the native leave confirmation', () => {
    dirty = true
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(event.defaultPrevented).toBe(true)
  })

  it('an approved close unloads without the confirmation', () => {
    dirty = true
    bridge.approveClose()
    const event = new Event('beforeunload', { cancelable: true })
    window.dispatchEvent(event)
    expect(event.defaultPrevented).toBe(false)
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
