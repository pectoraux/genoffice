/**
 * PROJECT-030 — the shared Project dialog layer.
 *
 * The modal dialogs BOTH hosts render — implemented ONCE in the shared
 * host binding (`@genoffice/project-host`), mounted on `document.body`
 * (the modal-overlay convention), identical on desktop (Electron
 * renderer) and web (browser) by construction: the shells add NO dialog
 * code of their own (the discipline suites pin it).
 *
 * Two dialogs, one discipline:
 *
 * - **The unsaved-changes dialog** (`confirmUnsavedChanges`) — the
 *   three-button Save / Don't Save / Cancel decision the controller's
 *   unsaved gate and close-guard handshake await. This ONE dialog
 *   replaces the two bespoke per-host implementations the shells carried
 *   before PROJECT-030 (the desktop native `dialog.showMessageBox` and
 *   the web bridge's private DOM dialog): a dialog is shared
 *   PRESENTATION, not host transport, so it left the bridge contract.
 * - **The Task Information dialog** (`createTaskInformationDialog`) —
 *   the semantic dialog that OPERATES ON COMMANDS: it collects the
 *   user's field drafts and hands them to the controller, which runs
 *   each changed field through the CANONICAL one-call commit flow
 *   (`commitTaskEditThroughSession`: begin → draft → semantic command →
 *   journal + reconcile + the derived-timing refresh) — exactly the
 *   pipeline the cell editor runs. The dialog module itself builds NO
 *   command, reads NO canonical state, and owns no Project semantics:
 *   its input is plain presentation data the controller computes; its
 *   output is plain draft strings.
 *
 * Presentation rules shared by both dialogs (the pre-030 web dialog's
 * accepted behavior, now canonical):
 *
 * - Escape cancels; focus starts on the primary input/button; the
 *   promise/callback settles EXACTLY ONCE; the overlay is removed and
 *   the document key listener detached on settle.
 * - Only the TOPMOST dialog answers keyboard input (a stacked dialog
 *   below stays inert) — deterministic under the rare stack (an
 *   external open request arriving while a dialog is open).
 * - Determinism: the DOM is a pure function of the input — no wall
 *   clock, no randomness, no date computation (display strings are
 *   passed in or sliced from ISO text).
 *
 * This module imports NOTHING (the strictest shape in the package —
 * pinned by the discipline suite): it is pure shared presentation.
 */

/** The unsaved-changes dialog answer (moved out of the bridge contract
 * at PROJECT-030 — the dialog is shared presentation, not transport). */
export type DiscardChoice = 'save' | 'discard' | 'cancel'

/** The Task Information dialog's pure input — presentation data the
 * controller computes (canonical draft texts + the 023 editability
 * fact + the displayed start/finish instants). */
export interface TaskInformationInput {
  /** The target task (carried for the controller's callbacks and the
   * E2E assertions; never rendered as semantics). */
  readonly taskId: string
  /** The name draft (the DISPLAYED task name, the cell editor's
   * canonical draft protocol). */
  readonly name: string
  /** The duration draft (decimal working minutes, the displayed value). */
  readonly duration: string
  /** Whether duration is editable (leaf rows only — a summary's
  scheduling values are derived roll-ups, the 023 rule). */
  readonly durationEditable: boolean
  /** The displayed start instant (ISO text, '' when none). */
  readonly start: string
  /** The displayed finish instant (ISO text, '' when none). */
  readonly finish: string
}

/** The dialog's collected drafts on OK (plain strings — the controller
 * runs the canonical translation; the dialog parses nothing). */
export interface TaskInformationRequest {
  readonly name: string
  readonly duration: string
}

/** The commit outcome the controller reports back synchronously (the
 * controller pipeline is synchronous): ok closes the dialog; a reason
 * keeps it open with the reason surfaced. */
export type TaskInformationResult =
  { readonly ok: true } | { readonly ok: false; readonly reason: string }

export interface TaskInformationCallbacks {
  /** OK: run the commit sequence over the drafts. Synchronous by
   * construction (the controller's commit pipeline never awaits). */
  onCommit(request: TaskInformationRequest): TaskInformationResult
  /** Cancel / Escape: nothing runs (no command, no mutation). */
  onCancelled(): void
}

export interface TaskInformationDialog {
  /** Opens (or re-opens — refreshing from the CURRENT document) the
   * dialog for the input; replaces any open instance. */
  open(input: TaskInformationInput, callbacks: TaskInformationCallbacks): void
  /** Closes the dialog if open (no callbacks fire — used when the
   * context vanished, e.g. a new document was loaded). */
  close(): void
  /** Whether the dialog is currently open (the controller's modal gate). */
  isOpen(): boolean
}

const el = (tag: string, className?: string): HTMLElement => {
  const node = document.createElement(tag)
  if (className !== undefined) node.className = className
  return node
}

/** Presentation-only instant label: `YYYY-MM-DD HH:mm` (UTC, by slicing
 * — the shared UI layer's convention, never date computation). */
const formatInstantLabel = (iso: string): string => `${iso.slice(0, 10)} ${iso.slice(11, 16)}`

/** Whether `overlay` is the topmost dialog overlay in the document (the
 * keyboard-answer rule: stacked dialogs below stay inert). */
const isTopmostOverlay = (overlay: HTMLElement): boolean => {
  if (!overlay.isConnected) return false
  const overlays = document.querySelectorAll('.gp-dialog-overlay')
  return overlays.length === 0 || overlays[overlays.length - 1] === overlay
}

/** Detaches a settled dialog's global key listener (the once-only
 * settle discipline — no listener ever leaks past its dialog). */
function detachKeyListener(onKey: (event: KeyboardEvent) => void): void {
  document.removeEventListener('keydown', onKey, { capture: true })
}

/**
 * The shared unsaved-changes dialog: Save / Don't Save / Cancel, named
 * after the project, Escape = Cancel, focus on Save. Resolves exactly
 * once; the overlay is removed on settle.
 */
export function confirmUnsavedChanges(projectName: string): Promise<DiscardChoice> {
  return new Promise((resolve) => {
    const overlay = el('div', 'gp-dialog-overlay')
    overlay.dataset.testid = 'discard-dialog'

    const dialog = el('div', 'gp-dialog')
    dialog.setAttribute('role', 'alertdialog')
    dialog.setAttribute('aria-modal', 'true')
    dialog.setAttribute('aria-label', 'Unsaved changes')

    const title = el('div', 'gp-dialog-title')
    title.textContent = `Save changes to '${projectName}'?`

    const detail = el('div', 'gp-dialog-detail')
    detail.textContent = 'Unsaved changes will be lost if you don’t save them.'

    const buttons = el('div', 'gp-dialog-buttons')

    let settled = false
    const answer = (choice: DiscardChoice): void => {
      if (settled) return
      settled = true
      overlay.remove()
      detachKeyListener(onKey)
      resolve(choice)
    }

    const addButton = (label: string, choice: DiscardChoice, testid: string): void => {
      const button = document.createElement('button')
      button.type = 'button'
      button.dataset.testid = testid
      button.textContent = label
      button.addEventListener('click', () => answer(choice))
      buttons.appendChild(button)
    }
    addButton('Save', 'save', 'discard-save')
    addButton("Don't Save", 'discard', 'discard-dont-save')
    addButton('Cancel', 'cancel', 'discard-cancel')

    const onKey = (event: KeyboardEvent): void => {
      if (settled || !overlay.isConnected) {
        detachKeyListener(onKey)
        return
      }
      if (event.key !== 'Escape' || !isTopmostOverlay(overlay)) return
      event.stopPropagation()
      answer('cancel')
    }
    document.addEventListener('keydown', onKey, { capture: true })

    dialog.append(title, detail, buttons)
    overlay.appendChild(dialog)
    document.body.appendChild(overlay)
    overlay.querySelector<HTMLButtonElement>('[data-testid="discard-save"]')?.focus()
  })
}

/**
 * Creates the shared Task Information dialog surface (one instance per
 * controller; `open` re-uses it). The dialog is presentation only: it
 * renders the controller-computed input, collects drafts, and reports
 * them through the callbacks — the controller runs the canonical
 * command pipeline (the dialog operates ON commands, never beside
 * them).
 */
export function createTaskInformationDialog(): TaskInformationDialog {
  let overlay: HTMLElement | null = null
  let onKey: ((event: KeyboardEvent) => void) | null = null

  const detach = (): void => {
    if (onKey !== null) {
      detachKeyListener(onKey)
      onKey = null
    }
    overlay?.remove()
    overlay = null
  }

  function open(input: TaskInformationInput, callbacks: TaskInformationCallbacks): void {
    // Re-open refreshes from the current document (the prior drafts are
    // the user's uncommitted input — a refresh discards them, exactly
    // like closing and reopening any modal).
    detach()

    const surface = el('div', 'gp-dialog-overlay')
    surface.dataset.testid = 'task-info-dialog'
    surface.dataset.taskId = input.taskId
    overlay = surface

    const dialog = el('div', 'gp-dialog')
    dialog.setAttribute('role', 'dialog')
    dialog.setAttribute('aria-modal', 'true')
    dialog.setAttribute('aria-label', 'Task Information')

    const title = el('div', 'gp-dialog-title')
    title.textContent = 'Task Information'

    const fields = el('div', 'gp-dialog-fields')

    const nameField = el('div', 'gp-dialog-field')
    const nameLabel = el('label', 'gp-dialog-label')
    nameLabel.textContent = 'Name'
    const nameInput = document.createElement('input')
    nameInput.type = 'text'
    nameInput.className = 'gp-dialog-input'
    nameInput.dataset.testid = 'task-info-name'
    nameInput.setAttribute('aria-label', 'Task name')
    nameInput.spellcheck = false
    nameInput.value = input.name
    nameField.append(nameLabel, nameInput)

    const durationField = el('div', 'gp-dialog-field')
    const durationLabel = el('label', 'gp-dialog-label')
    durationLabel.textContent = 'Duration (min)'
    const durationInput = document.createElement('input')
    durationInput.type = 'text'
    durationInput.className = 'gp-dialog-input'
    durationInput.dataset.testid = 'task-info-duration'
    durationInput.setAttribute('aria-label', 'Duration in working minutes')
    durationInput.spellcheck = false
    durationInput.value = input.duration
    // A summary's duration is a derived roll-up (the 023 editability
    // rule): the input is disabled and says why.
    durationInput.disabled = !input.durationEditable
    durationInput.setAttribute('aria-disabled', String(!input.durationEditable))
    if (!input.durationEditable) durationInput.title = 'Summary roll-up — not editable'
    durationField.append(durationLabel, durationInput)

    const startField = el('div', 'gp-dialog-field')
    const startLabel = el('label', 'gp-dialog-label')
    startLabel.textContent = 'Start'
    const startValue = el('span', 'gp-dialog-value')
    startValue.dataset.testid = 'task-info-start'
    startValue.textContent = input.start === '' ? '' : formatInstantLabel(input.start)
    startField.append(startLabel, startValue)

    const finishField = el('div', 'gp-dialog-field')
    const finishLabel = el('label', 'gp-dialog-label')
    finishLabel.textContent = 'Finish'
    const finishValue = el('span', 'gp-dialog-value')
    finishValue.dataset.testid = 'task-info-finish'
    finishValue.textContent = input.finish === '' ? '' : formatInstantLabel(input.finish)
    finishField.append(finishLabel, finishValue)

    fields.append(nameField, durationField, startField, finishField)

    const error = el('div', 'gp-dialog-error')
    error.dataset.testid = 'task-info-error'
    error.setAttribute('role', 'alert')
    error.textContent = ''

    const buttons = el('div', 'gp-dialog-buttons')
    const okButton = document.createElement('button')
    okButton.type = 'button'
    okButton.dataset.testid = 'task-info-ok'
    okButton.textContent = 'OK'
    const cancelButton = document.createElement('button')
    cancelButton.type = 'button'
    cancelButton.dataset.testid = 'task-info-cancel'
    cancelButton.textContent = 'Cancel'
    buttons.append(okButton, cancelButton)

    const cancel = (): void => {
      detach()
      callbacks.onCancelled()
    }

    const submit = (): void => {
      const outcome = callbacks.onCommit({
        name: nameInput.value,
        duration: durationInput.value,
      })
      if (outcome.ok) {
        detach()
        return
      }
      // The commit sequence refused: surface the controller's reason and
      // keep the dialog open (already-applied fields stay applied — they
      // were real commands through the controller; the user fixes the
      // refused field or cancels).
      error.textContent = outcome.reason
      nameInput.focus()
    }

    okButton.addEventListener('click', submit)
    cancelButton.addEventListener('click', cancel)

    onKey = (event: KeyboardEvent): void => {
      if (overlay === null) {
        detachKeyListener(onKey!)
        return
      }
      if (!overlay.isConnected || !isTopmostOverlay(overlay)) return
      if (event.key === 'Escape') {
        event.stopPropagation()
        cancel()
        return
      }
      if (
        event.key === 'Enter' &&
        dialog.contains(event.target instanceof Node ? event.target : null)
      ) {
        event.preventDefault()
        submit()
      }
    }
    document.addEventListener('keydown', onKey, { capture: true })

    dialog.append(title, fields, error, buttons)
    overlay.appendChild(dialog)
    document.body.appendChild(overlay)
    nameInput.focus()
  }

  return {
    open,
    close: detach,
    isOpen: () => overlay !== null && overlay.isConnected,
  }
}
