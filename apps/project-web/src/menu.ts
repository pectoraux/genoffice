/**
 * PROJECT-028 — the Project web host menu bar (PROJECT-031: the presentation
 * lockstep — the bar renders the SHARED menu presentation table).
 *
 * The DOM analog of the desktop's native application menu: transport
 * chrome ONLY. Every item carries a `MenuCommandId` from the shared
 * contract vocabulary and forwards ACTIVATION to the entry-provided sink
 * (which dispatches through the bridge's menu-command path — the same
 * surface the native menu's IPC channel feeds on desktop). No item
 * interprets Project state; labels are static.
 *
 * PROJECT-031 — presentation parity: the bar's structure, labels, and
 * accelerator display strings come from the ONE shared presentation table
 * (`MENU_PRESENTATION` in `@genoffice/project-host`), the same table the
 * desktop native menu mirrors and the shared ribbon's tooltips display.
 * This module owns NO label or accelerator literal — the cross-host
 * presentation lockstep is structural, not conventional.
 *
 * Accelerators are DISPLAYED but execute nothing here — the web entry's
 * keyboard listener + the shared translation table are the single
 * execution path for accelerator keys (the exact desktop discipline), so
 * an active cell editor keeps its own keys.
 */
import { MENU_PRESENTATION } from '@genoffice/project-host'
import type { MenuCommandId } from '@genoffice/project-host'

/** The shared menu presentation — the ONE table (PROJECT-031); this module
 * renders it and owns no presentation literal of its own. */
const MENUS = MENU_PRESENTATION

/** The menu command ids in bar order (used by the discipline suite). */
export const WEB_MENU_COMMAND_IDS: readonly MenuCommandId[] = MENUS.flatMap((menu) =>
  menu.items.map((item) => item.id),
)

/**
 * Builds the menu bar into `container`. `dispatch` is the activation sink
 * (the web entry forwards it through the bridge's menu-command path).
 * Returns the menu's open state — while a dropdown is open the menu owns
 * the keyboard (native-menu parity), so the entry's app keydown handler
 * consults `isOpen()` and swallows the key.
 */
export function createMenuBar(
  container: HTMLElement,
  dispatch: (command: MenuCommandId) => void,
): { readonly isOpen: () => boolean } {
  const bar = document.createElement('div')
  bar.dataset.testid = 'menu-bar'
  bar.className = 'gp-web-menu-bar'
  bar.setAttribute('role', 'menubar')
  bar.setAttribute('aria-label', 'Project menu')

  let openDropdown: HTMLElement | null = null
  let openTop: HTMLButtonElement | null = null

  const closeMenu = (): void => {
    openDropdown?.remove()
    openDropdown = null
    if (openTop !== null) {
      openTop.setAttribute('aria-expanded', 'false')
      openTop = null
    }
  }

  const openMenu = (top: HTMLButtonElement, menu: (typeof MENUS)[number]): void => {
    closeMenu()
    const dropdown = document.createElement('div')
    dropdown.dataset.testid = 'menu-dropdown'
    dropdown.className = 'gp-web-menu-dropdown'
    dropdown.setAttribute('role', 'menu')
    dropdown.setAttribute('aria-label', `${menu.label} menu`)
    for (const item of menu.items) {
      const entry = document.createElement('button')
      entry.type = 'button'
      entry.dataset.menuId = item.id
      entry.className = 'gp-web-menu-item'
      entry.setAttribute('role', 'menuitem')
      const label = document.createElement('span')
      label.textContent = item.label
      entry.appendChild(label)
      if (item.accelerator !== undefined) {
        const accelerator = document.createElement('span')
        accelerator.className = 'gp-web-menu-accelerator'
        accelerator.textContent = item.accelerator
        entry.appendChild(accelerator)
      }
      // Menu clicks close the menu FIRST, then dispatch — the dropdown
      // never overlaps the click's re-render.
      entry.addEventListener('click', () => {
        closeMenu()
        dispatch(item.id)
      })
      dropdown.appendChild(entry)
    }
    bar.appendChild(dropdown)
    // Position deterministically under the top button (the dropdown is
    // fixed-positioned; explicit offsets avoid flow-position ambiguity).
    const topRect = top.getBoundingClientRect()
    dropdown.style.left = `${topRect.left}px`
    dropdown.style.top = `${topRect.bottom}px`
    openDropdown = dropdown
    openTop = top
    top.setAttribute('aria-expanded', 'true')
  }

  for (const menu of MENUS) {
    const top = document.createElement('button')
    top.type = 'button'
    top.dataset.menuTop = menu.label.toLowerCase()
    top.className = 'gp-web-menu-top'
    top.setAttribute('role', 'menuitem')
    top.setAttribute('aria-haspopup', 'menu')
    top.setAttribute('aria-expanded', 'false')
    top.textContent = menu.label
    top.addEventListener('click', () => {
      if (openTop === top) {
        closeMenu()
        return
      }
      openMenu(top, menu)
    })
    top.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown') {
        event.preventDefault()
        openMenu(top, menu)
      }
    })
    bar.appendChild(top)
  }

  // Outside clicks close the open dropdown; Escape closes it (the Escape
  // never reaches the app — the menu owns it while open).
  document.addEventListener('click', (event) => {
    if (openDropdown === null) return
    if (event.target instanceof Node && bar.contains(event.target)) return
    closeMenu()
  })
  document.addEventListener('keydown', (event) => {
    if (openDropdown === null) return
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      closeMenu()
    }
  })

  container.appendChild(bar)
  return { isOpen: () => openDropdown !== null }
}
