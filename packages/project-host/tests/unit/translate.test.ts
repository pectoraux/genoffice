/**
 * PROJECT-027 — the host translation tables (pure unit battery).
 *
 * Proves the keyboard/menu translation is exactly the accepted vocabulary:
 * navigation keys → moveTaskFocus intents (with the shift-extend rule),
 * activation keys → the edit flow, modifiers → file/history/document
 * actions, and the MENU ids converge on the SAME actions as the keys that
 * display them in the native menu (the single-translation-path discipline).
 */
import { describe, expect, it } from 'vitest'
import {
  translateKeyDown,
  translateMenuCommand,
  ZOOM_IN_FACTOR,
  ZOOM_OUT_FACTOR,
} from '../../src/translate.js'
import type { KeyInput } from '../../src/translate.js'
import { MENU_COMMAND_IDS } from '../../src/bridge.js'
import { menuAcceleratorFor } from '../../src/menu-presentation.js'

const key = (partial: Partial<KeyInput>): KeyInput => ({
  key: '',
  ctrlOrMeta: false,
  shift: false,
  alt: false,
  ...partial,
})

describe('keyboard translation — navigation', () => {
  it('arrow keys move the task focus', () => {
    expect(translateKeyDown(key({ key: 'ArrowUp' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveTaskFocus', direction: 'up' },
    })
    expect(translateKeyDown(key({ key: 'ArrowDown' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveTaskFocus', direction: 'down' },
    })
  })

  it('home/end jump to the first/last visible row', () => {
    expect(translateKeyDown(key({ key: 'Home' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveTaskFocus', direction: 'first' },
    })
    expect(translateKeyDown(key({ key: 'End' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveTaskFocus', direction: 'last' },
    })
  })

  it('shift extends the selection while moving', () => {
    expect(translateKeyDown(key({ key: 'ArrowDown', shift: true }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveTaskFocus', direction: 'down', extend: true },
    })
  })

  it('the focused-cell column keys (PROJECT-031): arrows and Tab walk the cells', () => {
    expect(translateKeyDown(key({ key: 'ArrowRight' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveCellFocus', direction: 'next' },
    })
    expect(translateKeyDown(key({ key: 'ArrowLeft' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveCellFocus', direction: 'previous' },
    })
    expect(translateKeyDown(key({ key: 'Tab' }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveCellFocus', direction: 'next' },
    })
    expect(translateKeyDown(key({ key: 'Tab', shift: true }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'moveCellFocus', direction: 'previous' },
    })
    // While editing, the column keys pass through to the input (the
    // pinned 023 editor rule — Tab is the browser's focus key there).
    expect(translateKeyDown(key({ key: 'ArrowRight' }), { editing: true })).toEqual({
      kind: 'none',
    })
    expect(translateKeyDown(key({ key: 'Tab' }), { editing: true })).toEqual({ kind: 'none' })
  })

  it('the outline gestures still own the ALT arrows (indent/outdent, not cell moves)', () => {
    expect(translateKeyDown(key({ key: 'ArrowLeft', alt: true }), { editing: false })).toEqual({
      kind: 'document',
      action: 'outdentSelection',
    })
    expect(
      translateKeyDown(key({ key: 'ArrowRight', alt: true, shift: true }), { editing: false }),
    ).toEqual({ kind: 'document', action: 'indentSelection' })
  })

  it('unmapped keys are none (the host invents no bindings)', () => {
    expect(translateKeyDown(key({ key: 'q' }), { editing: false })).toEqual({ kind: 'none' })
    expect(translateKeyDown(key({ key: 'F9' }), { editing: false })).toEqual({ kind: 'none' })
  })
})

describe('keyboard translation — editing', () => {
  it('enter and F2 begin the edit of the FOCUSED CELL (taskName when no field is focused)', () => {
    expect(translateKeyDown(key({ key: 'Enter' }), { editing: false })).toEqual({
      kind: 'edit',
      action: 'beginEditFocusedCell',
    })
    expect(translateKeyDown(key({ key: 'F2' }), { editing: false })).toEqual({
      kind: 'edit',
      action: 'beginEditFocusedCell',
    })
  })

  it('while editing, only Enter/Escape belong to the editor', () => {
    expect(translateKeyDown(key({ key: 'Enter' }), { editing: true })).toEqual({
      kind: 'edit',
      action: 'commit',
    })
    expect(translateKeyDown(key({ key: 'Escape' }), { editing: true })).toEqual({
      kind: 'edit',
      action: 'cancel',
    })
    // Everything else — including Delete and arrows — passes through to the
    // input (native caret movement + text editing), never to document
    // commands: the mid-edit menu-fire hazard is structurally impossible.
    expect(translateKeyDown(key({ key: 'Delete' }), { editing: true })).toEqual({ kind: 'none' })
    expect(translateKeyDown(key({ key: 'ArrowUp' }), { editing: true })).toEqual({ kind: 'none' })
    expect(translateKeyDown(key({ key: 'z', ctrlOrMeta: true }), { editing: true })).toEqual({
      kind: 'none',
    })
  })
})

describe('keyboard translation — modifier shortcuts', () => {
  it('undo/redo (the menu accelerators, unregistered → handled here)', () => {
    expect(translateKeyDown(key({ key: 'z', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'history',
      action: 'undo',
    })
    expect(
      translateKeyDown(key({ key: 'Z', ctrlOrMeta: true, shift: true }), { editing: false }),
    ).toEqual({ kind: 'history', action: 'redo' })
    expect(translateKeyDown(key({ key: 'y', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'history',
      action: 'redo',
    })
  })

  it('file shortcuts', () => {
    expect(translateKeyDown(key({ key: 'n', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'file',
      action: 'new',
    })
    expect(translateKeyDown(key({ key: 'o', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'file',
      action: 'open',
    })
    expect(translateKeyDown(key({ key: 's', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'file',
      action: 'save',
    })
    expect(
      translateKeyDown(key({ key: 'S', ctrlOrMeta: true, shift: true }), { editing: false }),
    ).toEqual({ kind: 'file', action: 'saveAs' })
  })

  it('task + outline shortcuts', () => {
    expect(translateKeyDown(key({ key: 'Insert' }), { editing: false })).toEqual({
      kind: 'document',
      action: 'createTask',
    })
    expect(translateKeyDown(key({ key: 't', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'document',
      action: 'createTask',
    })
    expect(translateKeyDown(key({ key: 'Delete' }), { editing: false })).toEqual({
      kind: 'document',
      action: 'deleteSelection',
    })
    expect(
      translateKeyDown(key({ key: 'ArrowRight', alt: true, shift: true }), { editing: false }),
    ).toEqual({ kind: 'document', action: 'indentSelection' })
    expect(
      translateKeyDown(key({ key: 'ArrowLeft', alt: true, shift: true }), { editing: false }),
    ).toEqual({ kind: 'document', action: 'outdentSelection' })
  })

  it('zoom + fit + collapse', () => {
    expect(translateKeyDown(key({ key: '=', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'scaleViewport', factor: ZOOM_IN_FACTOR },
    })
    expect(translateKeyDown(key({ key: '-', ctrlOrMeta: true }), { editing: false })).toEqual({
      kind: 'intent',
      intent: { type: 'scaleViewport', factor: ZOOM_OUT_FACTOR },
    })
    expect(
      translateKeyDown(key({ key: 'F', ctrlOrMeta: true, shift: true }), { editing: false }),
    ).toEqual({ kind: 'intent', intent: { type: 'fitViewport' } })
    expect(translateKeyDown(key({ key: '-', alt: true, shift: true }), { editing: false })).toEqual(
      { kind: 'view', action: 'collapseSelection' },
    )
    expect(translateKeyDown(key({ key: '+', alt: true, shift: true }), { editing: false })).toEqual(
      { kind: 'view', action: 'expandSelection' },
    )
  })
})

describe('menu translation converges with the keyboard table', () => {
  it('the menu command ids are the complete, unique vocabulary', () => {
    expect(new Set(MENU_COMMAND_IDS).size).toBe(MENU_COMMAND_IDS.length)
    expect([...MENU_COMMAND_IDS].sort()).toEqual(
      [
        'file.new',
        'file.open',
        'file.save',
        'file.saveAs',
        'edit.undo',
        'edit.redo',
        'edit.deleteTask',
        'task.create',
        'task.information',
        'task.indent',
        'task.outdent',
        'view.zoomIn',
        'view.zoomOut',
        'view.fit',
        'view.collapse',
        'view.expand',
      ].sort(),
    )
  })

  it('every menu command maps to a non-none action', () => {
    for (const id of MENU_COMMAND_IDS) {
      expect(translateMenuCommand(id).kind).not.toBe('none')
    }
  })

  it('task.information opens the shared dialog (the PROJECT-030 command)', () => {
    // The dialog action — deliberately WITHOUT a keyboard binding (the
    // shared keyboard table invents no binding for it; menu/ribbon
    // activation are the firing surfaces).
    expect(translateMenuCommand('task.information')).toEqual({
      kind: 'dialog',
      action: 'taskInformation',
    })
    for (const input of [
      key({ key: 'i', ctrlOrMeta: true }),
      key({ key: 'F4' }),
      key({ key: 'Enter' }),
    ]) {
      expect(translateKeyDown(input, { editing: false })).not.toEqual(
        translateMenuCommand('task.information'),
      )
    }
  })

  it('the displayed accelerators and the menu ids agree with the keyboard table', () => {
    // Ctrl+N ↔ file.new, Ctrl+Z ↔ edit.undo, Ctrl+= ↔ view.zoomIn …
    expect(translateKeyDown(key({ key: 'n', ctrlOrMeta: true }), { editing: false })).toEqual(
      translateMenuCommand('file.new'),
    )
    expect(translateKeyDown(key({ key: 'z', ctrlOrMeta: true }), { editing: false })).toEqual(
      translateMenuCommand('edit.undo'),
    )
    expect(translateKeyDown(key({ key: 'o', ctrlOrMeta: true }), { editing: false })).toEqual(
      translateMenuCommand('file.open'),
    )
    expect(translateKeyDown(key({ key: 's', ctrlOrMeta: true }), { editing: false })).toEqual(
      translateMenuCommand('file.save'),
    )
    expect(translateKeyDown(key({ key: 'Insert' }), { editing: false })).toEqual(
      translateMenuCommand('task.create'),
    )
    expect(translateKeyDown(key({ key: 'Delete' }), { editing: false })).toEqual(
      translateMenuCommand('edit.deleteTask'),
    )
    expect(translateKeyDown(key({ key: '=', ctrlOrMeta: true }), { editing: false })).toEqual(
      translateMenuCommand('view.zoomIn'),
    )
    expect(translateKeyDown(key({ key: '-', ctrlOrMeta: true }), { editing: false })).toEqual(
      translateMenuCommand('view.zoomOut'),
    )
    expect(
      translateKeyDown(key({ key: 'ArrowRight', alt: true, shift: true }), { editing: false }),
    ).toEqual(translateMenuCommand('task.indent'))
  })

  it('EVERY displayed accelerator is a real keyboard-table binding of its command (PROJECT-031 presentation parity)', () => {
    // The one-to-one map from the shared display form to the KeyInput
    // facts the table binds — the proof that no menu DISPLAYS a binding
    // the shared translation table does not own (and that every bound
    // command's display agrees with its menu action).
    const inputByAccelerator: Record<string, KeyInput> = {
      'Ctrl+N': key({ key: 'n', ctrlOrMeta: true }),
      'Ctrl+O': key({ key: 'o', ctrlOrMeta: true }),
      'Ctrl+S': key({ key: 's', ctrlOrMeta: true }),
      'Ctrl+Shift+S': key({ key: 'S', ctrlOrMeta: true, shift: true }),
      'Ctrl+Z': key({ key: 'z', ctrlOrMeta: true }),
      'Ctrl+Y': key({ key: 'y', ctrlOrMeta: true }),
      Delete: key({ key: 'Delete' }),
      Insert: key({ key: 'Insert' }),
      'Alt+Shift+Right': key({ key: 'ArrowRight', alt: true, shift: true }),
      'Alt+Shift+Left': key({ key: 'ArrowLeft', alt: true, shift: true }),
      'Ctrl+=': key({ key: '=', ctrlOrMeta: true }),
      'Ctrl+-': key({ key: '-', ctrlOrMeta: true }),
      'Ctrl+Shift+F': key({ key: 'F', ctrlOrMeta: true, shift: true }),
      'Alt+Shift+Minus': key({ key: '-', alt: true, shift: true }),
      'Alt+Shift+Plus': key({ key: '+', alt: true, shift: true }),
    }
    for (const command of MENU_COMMAND_IDS) {
      const accelerator = menuAcceleratorFor(command)
      if (accelerator === undefined) {
        // No display ⇔ no binding: the one command without an accelerator
        // (task.information) must not be keyboard-reachable either.
        expect(command).toBe('task.information')
        continue
      }
      const input = inputByAccelerator[accelerator]
      expect(input, `the display "${accelerator}" must be a known binding form`).toBeDefined()
      expect(translateKeyDown(input!, { editing: false })).toEqual(translateMenuCommand(command))
    }
  })
})
