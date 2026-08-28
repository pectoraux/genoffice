/**
 * The shared dialog layer battery (PROJECT-030).
 *
 * The dialog modules directly (`confirmUnsavedChanges`,
 * `createTaskInformationDialog` on jsdom): structure, keyboard behavior,
 * the once-only settle discipline, listener hygiene, build determinism —
 * the presentation contract BOTH hosts render (the controller integration
 * lives in app.test.ts; the shell batteries drive the real apps in E2E).
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { confirmUnsavedChanges, createTaskInformationDialog } from '../../src/dialogs.js'
import type { TaskInformationInput } from '../../src/dialogs.js'

const flush = async (): Promise<void> => {
  await new Promise((resolve) => setTimeout(resolve, 0))
}

const SAMPLE_INPUT: TaskInformationInput = {
  taskId: 't1',
  name: 'Pour foundations',
  duration: '480',
  durationEditable: true,
  start: '2026-01-05T09:00:00.000Z',
  finish: '2026-01-06T17:00:00.000Z',
}

const keydown = (key: string): void => {
  document.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
}

beforeEach(() => {
  document.body.innerHTML = ''
})

afterEach(() => {
  document.body.innerHTML = ''
})

describe('the shared unsaved-changes dialog', () => {
  it('Save / Don’t Save / Cancel resolve the contract choices', async () => {
    for (const [testid, expected] of [
      ['discard-save', 'save'],
      ['discard-dont-save', 'discard'],
      ['discard-cancel', 'cancel'],
    ] as const) {
      const pending = confirmUnsavedChanges('E2E Build')
      await flush()
      const button = document.querySelector(`[data-testid="${testid}"]`) as HTMLElement
      expect(button).not.toBeNull()
      button.click()
      await expect(pending).resolves.toBe(expected)
      // Settled once: the overlay is gone.
      expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
    }
  })

  it('the dialog names the project (the alertdialog title)', async () => {
    const pending = confirmUnsavedChanges('Riverside Refit')
    await flush()
    const title = document.querySelector('[data-testid="discard-dialog"] .gp-dialog-title')
    expect(title?.textContent).toContain('Riverside Refit')
    ;(document.querySelector('[data-testid="discard-cancel"]') as HTMLElement).click()
    await expect(pending).resolves.toBe('cancel')
  })

  it('Escape cancels (the native dialog keyboard behavior)', async () => {
    const pending = confirmUnsavedChanges('Escape Case')
    await flush()
    keydown('Escape')
    await expect(pending).resolves.toBe('cancel')
    expect(document.querySelector('[data-testid="discard-dialog"]')).toBeNull()
  })

  it('settles EXACTLY ONCE — later clicks are inert', async () => {
    const pending = confirmUnsavedChanges('Once Only')
    await flush()
    const save = document.querySelector('[data-testid="discard-save"]') as HTMLElement
    const cancel = document.querySelector('[data-testid="discard-cancel"]') as HTMLElement
    save.click()
    save.click()
    cancel.click()
    await expect(pending).resolves.toBe('save')
  })

  it('focus starts on Save and the key listener detaches on settle (no leaks)', async () => {
    const pending = confirmUnsavedChanges('Focus Case')
    await flush()
    const save = document.querySelector('[data-testid="discard-save"]') as HTMLElement
    expect(document.activeElement).toBe(save)
    save.click()
    await pending
    // A LATER Escape must find no listener (the dialog is gone) — and
    // nothing throws.
    expect(() => keydown('Escape')).not.toThrow()
  })

  it('determinism: two sequential consultations render identical dialog chrome', async () => {
    const first = confirmUnsavedChanges('Determinism')
    await flush()
    const one = (document.querySelector('[data-testid="discard-dialog"]') as HTMLElement).outerHTML
    ;(document.querySelector('[data-testid="discard-cancel"]') as HTMLElement).click()
    await first
    const second = confirmUnsavedChanges('Determinism')
    await flush()
    const two = (document.querySelector('[data-testid="discard-dialog"]') as HTMLElement).outerHTML
    ;(document.querySelector('[data-testid="discard-cancel"]') as HTMLElement).click()
    await second
    expect(one).toBe(two)
  })
})

describe('the shared Task Information dialog', () => {
  it('renders the fields from the controller-computed input', () => {
    const dialog = createTaskInformationDialog()
    let cancelled = false
    dialog.open(SAMPLE_INPUT, {
      onCommit: () => ({ ok: true }),
      onCancelled: () => (cancelled = true),
    })
    const overlay = document.querySelector('[data-testid="task-info-dialog"]') as HTMLElement
    expect(overlay).not.toBeNull()
    expect(overlay.dataset.taskId).toBe('t1')
    expect(
      (document.querySelector('[data-testid="task-info-name"]') as HTMLInputElement).value,
    ).toBe('Pour foundations')
    expect(
      (document.querySelector('[data-testid="task-info-duration"]') as HTMLInputElement).value,
    ).toBe('480')
    expect(
      (document.querySelector('[data-testid="task-info-duration"]') as HTMLInputElement).disabled,
    ).toBe(false)
    // The displayed instants are sliced ISO (the shared convention).
    expect(document.querySelector('[data-testid="task-info-start"]')?.textContent).toBe(
      '2026-01-05 09:00',
    )
    expect(document.querySelector('[data-testid="task-info-finish"]')?.textContent).toBe(
      '2026-01-06 17:00',
    )
    // ARIA dialog semantics + focus on the name input.
    const box = overlay.querySelector('.gp-dialog')
    expect(box?.getAttribute('role')).toBe('dialog')
    expect(box?.getAttribute('aria-modal')).toBe('true')
    expect(box?.getAttribute('aria-label')).toBe('Task Information')
    expect(document.activeElement).toBe(document.querySelector('[data-testid="task-info-name"]'))
    expect(cancelled).toBe(false)
  })

  it('the summary rule: duration renders disabled with the honest hint', () => {
    const dialog = createTaskInformationDialog()
    dialog.open(
      { ...SAMPLE_INPUT, durationEditable: false },
      { onCommit: () => ({ ok: true }), onCancelled: () => {} },
    )
    const duration = document.querySelector(
      '[data-testid="task-info-duration"]',
    ) as HTMLInputElement
    expect(duration.disabled).toBe(true)
    expect(duration.getAttribute('aria-disabled')).toBe('true')
    expect(duration.title).toContain('Summary roll-up')
  })

  it('OK hands the collected drafts to onCommit and closes on ok', () => {
    const dialog = createTaskInformationDialog()
    const requests: Array<{ name: string; duration: string }> = []
    dialog.open(SAMPLE_INPUT, {
      onCommit: (request) => {
        requests.push(request)
        return { ok: true }
      },
      onCancelled: () => {},
    })
    ;(document.querySelector('[data-testid="task-info-name"]') as HTMLInputElement).value =
      'Renamed'
    ;(document.querySelector('[data-testid="task-info-duration"]') as HTMLInputElement).value =
      '960'
    ;(document.querySelector('[data-testid="task-info-ok"]') as HTMLElement).click()
    expect(requests).toEqual([{ name: 'Renamed', duration: '960' }])
    expect(document.querySelector('[data-testid="task-info-dialog"]')).toBeNull()
  })

  it('a refused commit keeps the dialog open and surfaces the reason (role=alert)', () => {
    const dialog = createTaskInformationDialog()
    let calls = 0
    dialog.open(SAMPLE_INPUT, {
      onCommit: () => {
        calls += 1
        return calls === 1
          ? { ok: false, reason: 'INVALID_DURATION: negative duration' }
          : { ok: true }
      },
      onCancelled: () => {},
    })
    ;(document.querySelector('[data-testid="task-info-ok"]') as HTMLElement).click()
    const overlay = document.querySelector('[data-testid="task-info-dialog"]')
    expect(overlay).not.toBeNull()
    const error = document.querySelector('[data-testid="task-info-error"]') as HTMLElement
    expect(error.getAttribute('role')).toBe('alert')
    expect(error.textContent).toContain('INVALID_DURATION')
    // The fix commits and closes.
    ;(document.querySelector('[data-testid="task-info-ok"]') as HTMLElement).click()
    expect(document.querySelector('[data-testid="task-info-dialog"]')).toBeNull()
  })

  it('Cancel and Escape fire onCancelled; no commit ever runs', () => {
    const dialog = createTaskInformationDialog()
    let commits = 0
    let cancels = 0
    dialog.open(SAMPLE_INPUT, {
      onCommit: () => {
        commits += 1
        return { ok: true }
      },
      onCancelled: () => (cancels += 1),
    })
    ;(document.querySelector('[data-testid="task-info-cancel"]') as HTMLElement).click()
    expect(cancels).toBe(1)
    expect(document.querySelector('[data-testid="task-info-dialog"]')).toBeNull()
    dialog.open(SAMPLE_INPUT, {
      onCommit: () => {
        commits += 1
        return { ok: true }
      },
      onCancelled: () => (cancels += 1),
    })
    keydown('Escape')
    expect(cancels).toBe(2)
    expect(document.querySelector('[data-testid="task-info-dialog"]')).toBeNull()
    expect(commits).toBe(0)
  })

  it('Enter inside the dialog submits; the settled dialog is inert', () => {
    const dialog = createTaskInformationDialog()
    const requests: Array<{ name: string; duration: string }> = []
    dialog.open(SAMPLE_INPUT, {
      onCommit: (request) => {
        requests.push(request)
        return { ok: true }
      },
      onCancelled: () => {},
    })
    const name = document.querySelector('[data-testid="task-info-name"]') as HTMLInputElement
    name.value = 'Enter Key'
    name.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    )
    expect(requests).toEqual([{ name: 'Enter Key', duration: '480' }])
    expect(document.querySelector('[data-testid="task-info-dialog"]')).toBeNull()
    // A later Enter finds no listener (settled once — nothing throws, no
    // second request).
    expect(() =>
      name.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })),
    ).not.toThrow()
    expect(requests).toHaveLength(1)
  })

  it('re-open replaces the instance and refreshes from the new input', () => {
    const dialog = createTaskInformationDialog()
    dialog.open(SAMPLE_INPUT, { onCommit: () => ({ ok: true }), onCancelled: () => {} })
    dialog.open(
      { ...SAMPLE_INPUT, taskId: 't2', name: 'Second Task', duration: '120' },
      { onCommit: () => ({ ok: true }), onCancelled: () => {} },
    )
    const overlays = document.querySelectorAll('[data-testid="task-info-dialog"]')
    expect(overlays).toHaveLength(1)
    expect((overlays[0] as HTMLElement).dataset.taskId).toBe('t2')
    expect(
      (document.querySelector('[data-testid="task-info-name"]') as HTMLInputElement).value,
    ).toBe('Second Task')
  })

  it('close() detaches without callbacks (the context-vanished path)', () => {
    const dialog = createTaskInformationDialog()
    let cancels = 0
    dialog.open(SAMPLE_INPUT, { onCommit: () => ({ ok: true }), onCancelled: () => (cancels += 1) })
    expect(dialog.isOpen()).toBe(true)
    dialog.close()
    expect(dialog.isOpen()).toBe(false)
    expect(document.querySelector('[data-testid="task-info-dialog"]')).toBeNull()
    expect(cancels).toBe(0) // a context close is not a user cancel
    // A later Escape finds no listener.
    expect(() => keydown('Escape')).not.toThrow()
  })

  it('determinism: two fresh instances over the same input render identical chrome', () => {
    const first = createTaskInformationDialog()
    first.open(SAMPLE_INPUT, { onCommit: () => ({ ok: true }), onCancelled: () => {} })
    const one = (document.querySelector('[data-testid="task-info-dialog"]') as HTMLElement)
      .outerHTML
    first.close()
    const second = createTaskInformationDialog()
    second.open(SAMPLE_INPUT, { onCommit: () => ({ ok: true }), onCancelled: () => {} })
    const two = (document.querySelector('[data-testid="task-info-dialog"]') as HTMLElement)
      .outerHTML
    second.close()
    expect(one).toBe(two)
  })

  it('the topmost rule: a stacked dialog below stays inert on Escape', async () => {
    const bottom = createTaskInformationDialog()
    let bottomCancels = 0
    bottom.open(SAMPLE_INPUT, {
      onCommit: () => ({ ok: true }),
      onCancelled: () => (bottomCancels += 1),
    })
    // The unsaved-changes dialog stacks on top (the rare drag-drop-while-
    // open case): only IT answers Escape.
    const pending = confirmUnsavedChanges('Stacked')
    await flush()
    keydown('Escape')
    await expect(pending).resolves.toBe('cancel')
    expect(bottomCancels).toBe(0)
    expect(document.querySelector('[data-testid="task-info-dialog"]')).not.toBeNull()
    bottom.close()
  })
})
