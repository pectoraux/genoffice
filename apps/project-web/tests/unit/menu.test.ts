/**
 * PROJECT-028 — the web menu bar battery (jsdom): the bar's vocabulary,
 * the dropdown open/close behavior, the activation dispatch, and the
 * lockstep with the shared menu contract. PROJECT-031 adds the
 * rendered-presentation parity pin: the bar renders the SHARED menu
 * presentation table (labels + accelerator display strings), so the web
 * menu and the desktop native menu cannot drift.
 */
import { beforeEach, describe, expect, it } from 'vitest'
import { MENU_COMMAND_IDS, MENU_PRESENTATION } from '@genoffice/project-host'
import { createMenuBar, WEB_MENU_COMMAND_IDS } from '../../src/menu.js'
import type { MenuCommandId } from '@genoffice/project-host'

const mount = (): { container: HTMLElement; dispatched: MenuCommandId[] } => {
  document.body.innerHTML = ''
  const container = document.createElement('div')
  document.body.appendChild(container)
  const dispatched: MenuCommandId[] = []
  createMenuBar(container, (command) => dispatched.push(command))
  return { container, dispatched }
}

const click = (element: HTMLElement): void => {
  element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
}

describe('the web menu bar', () => {
  let mounted: ReturnType<typeof mount>
  beforeEach(() => {
    mounted = mount()
  })

  it('renders the four menu tops', () => {
    const tops = [...document.querySelectorAll('[data-menu-top]')]
    expect(tops.map((top) => top.getAttribute('data-menu-top'))).toEqual([
      'file',
      'edit',
      'task',
      'view',
    ])
  })

  it('the web menu vocabulary is EXACTLY the shared contract vocabulary', () => {
    expect([...WEB_MENU_COMMAND_IDS].sort()).toEqual([...MENU_COMMAND_IDS].sort())
  })

  it('a top click opens its dropdown; a second click closes it', () => {
    const file = document.querySelector<HTMLButtonElement>('[data-menu-top="file"]')!
    click(file)
    let items = [...document.querySelectorAll('[data-menu-id]')]
    expect(items.map((item) => item.getAttribute('data-menu-id'))).toEqual([
      'file.new',
      'file.open',
      'file.save',
      'file.saveAs',
    ])
    expect(file.getAttribute('aria-expanded')).toBe('true')
    click(file)
    items = [...document.querySelectorAll('[data-menu-id]')]
    expect(items).toHaveLength(0)
    expect(file.getAttribute('aria-expanded')).toBe('false')
  })

  it('the Task dropdown carries the shared dialog command (PROJECT-030)', () => {
    const task = document.querySelector<HTMLButtonElement>('[data-menu-top="task"]')!
    click(task)
    const items = [...document.querySelectorAll('[data-menu-id]')]
    expect(items.map((item) => item.getAttribute('data-menu-id'))).toEqual([
      'task.create',
      'task.information',
      'task.indent',
      'task.outdent',
    ])
    // The dialog item displays NO accelerator (the keyboard table binds no
    // key to it — menu/ribbon activation are the firing surfaces).
    const information = items[1]!
    expect(information.textContent).toContain('Task Information')
    expect(information.querySelector('.gp-web-menu-accelerator')).toBeNull()
    click(task)
  })

  it('every menu id is present in its dropdown (the complete vocabulary)', () => {
    const seen: string[] = []
    for (const label of ['file', 'edit', 'task', 'view']) {
      const top = document.querySelector<HTMLButtonElement>(`[data-menu-top="${label}"]`)!
      click(top)
      seen.push(
        ...[...document.querySelectorAll('[data-menu-id]')].map((item) =>
          item.getAttribute('data-menu-id')!,
        ),
      )
      click(top)
    }
    expect(seen.sort()).toEqual([...MENU_COMMAND_IDS].sort())
  })

  it('an item click dispatches the id and closes the dropdown', () => {
    const task = document.querySelector<HTMLButtonElement>('[data-menu-top="task"]')!
    click(task)
    click(document.querySelector<HTMLElement>('[data-menu-id="task.create"]')!)
    expect(mounted.dispatched).toEqual(['task.create'])
    expect(document.querySelectorAll('[data-menu-id]')).toHaveLength(0)
  })

  it('Escape closes the open dropdown', () => {
    const view = document.querySelector<HTMLButtonElement>('[data-menu-top="view"]')!
    click(view)
    expect(document.querySelectorAll('[data-menu-id]')).not.toHaveLength(0)
    document.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
    )
    expect(document.querySelectorAll('[data-menu-id]')).toHaveLength(0)
  })

  it('an outside click closes the open dropdown', () => {
    const edit = document.querySelector<HTMLButtonElement>('[data-menu-top="edit"]')!
    click(edit)
    expect(document.querySelectorAll('[data-menu-id]')).not.toHaveLength(0)
    click(document.body)
    expect(document.querySelectorAll('[data-menu-id]')).toHaveLength(0)
  })

  it('items display their accelerators (displayed, never executed here)', () => {
    const file = document.querySelector<HTMLButtonElement>('[data-menu-top="file"]')!
    click(file)
    const open = document.querySelector<HTMLElement>('[data-menu-id="file.open"]')!
    expect(open.textContent).toContain('Open Project')
    expect(open.textContent).toContain('Ctrl+O')
  })

  it('renders EXACTLY the shared presentation table — sections, labels, and accelerator displays (PROJECT-031)', () => {
    for (const section of MENU_PRESENTATION) {
      const top = document.querySelector<HTMLButtonElement>(
        `[data-menu-top="${section.label.toLowerCase()}"]`,
      )!
      expect(top, `the ${section.label} top exists`).not.toBeNull()
      click(top)
      const items = [...document.querySelectorAll('[data-menu-id]')]
      expect(items.map((item) => item.getAttribute('data-menu-id'))).toEqual(
        section.items.map((item) => item.id),
      )
      for (const [index, item] of section.items.entries()) {
        const rendered = items[index]!
        expect(rendered.textContent).toContain(item.label)
        const accelerator = rendered.querySelector('.gp-web-menu-accelerator')
        if (item.accelerator === undefined) {
          expect(accelerator).toBeNull()
        } else {
          expect(accelerator?.textContent).toBe(item.accelerator)
        }
      }
      click(top)
    }
  })
})
