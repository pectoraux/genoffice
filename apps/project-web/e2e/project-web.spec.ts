/**
 * PROJECT-028 — the Project web shell E2E battery.
 *
 * REAL browser verification: every test drives the built production bundle
 * (vite build → vite preview) in real chromium — the web host mounting the
 * SAME shared renderer binding the desktop shell runs, with the REAL
 * engine/scheduler/file adapters executing in the browser. The browser
 * transports are driven through their REAL surfaces (the file picker via
 * Playwright's filechooser, the drag-and-drop open via real DataTransfer
 * events, the save flow via real download events) — nothing is stubbed.
 *
 * Assertions are grounded in the canonical authority's own output for the
 * fixture document (the SAME constants the desktop battery asserts —
 * cross-host parity of the shared experience, by construction).
 */
import { mkdtemp } from 'node:fs/promises'
import { readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import type { Page } from '@playwright/test'
import { gprojFileAdapter } from '@genoffice/project-file'
import { e2eFixtureBytes, writeE2EFixture } from './fixtures'

/** The canonical fixture values (derived by the real scheduler — the same
 * constants the desktop battery asserts: cross-host parity). */
const T1_START = '2026-01-05 09:00'
const T1_FINISH_960 = '2026-01-06 17:00'
const T1_FINISH_1920 = '2026-01-08 17:00'
const T2_START = '2026-01-07 09:00'

/** The web transport cap (the web bridge's own constant, mirrored here for
 * the assertion messages). */
const MAX_WEB_FILE_BYTES = 104_857_600
const CAP_ERROR = `File exceeds the ${MAX_WEB_FILE_BYTES} byte limit`

const rows = (page: Page) => page.locator('[data-testid="task-row"]')
const rowOf = (page: Page, taskId: string) =>
  page.locator(`[data-testid="task-row"][data-task-id="${taskId}"]`)
const cellOf = (page: Page, taskId: string, column: string) =>
  page.locator(`[data-testid="task-row"][data-task-id="${taskId}"] [data-column="${column}"]`)
const statusText = (page: Page) => page.locator('[data-testid="status-text"]')

/** The full menu vocabulary (the shared contract's). */
const MENU_IDS = [
  'file.new',
  'file.open',
  'file.save',
  'file.saveAs',
  'edit.undo',
  'edit.redo',
  'edit.deleteTask',
  'task.create',
  'task.indent',
  'task.outdent',
  'view.zoomIn',
  'view.zoomOut',
  'view.fit',
  'view.collapse',
  'view.expand',
]

async function boot(page: Page): Promise<void> {
  await page.goto('/')
  await expect(page.locator('[data-testid="project-app"]')).toBeVisible()
}

/** Drops a file onto the window (the web's external-open path) with the
 * bytes built IN the page (the drop payload is a real DataTransfer File). */
async function dropBytes(
  page: Page,
  name: string,
  build: { base: number[]; total?: number },
): Promise<void> {
  await page.evaluate(
    async ({ name: fileName, base, total }) => {
      const buffer = new Uint8Array(total ?? base.length)
      buffer.set(base)
      if (total !== undefined && total > base.length) buffer.fill(0x20, base.length)
      const file = new File([buffer], fileName)
      const transfer = new DataTransfer()
      transfer.items.add(file)
      window.dispatchEvent(
        new DragEvent('drop', { dataTransfer: transfer, bubbles: true, cancelable: true }),
      )
    },
    { name, base: build.base, total: build.total },
  )
}

/** Drops the canonical fixture document. */
async function dropFixture(page: Page): Promise<void> {
  await dropBytes(page, 'e2e-build.gproj', { base: Array.from(e2eFixtureBytes()) })
}

test.describe('W01 — application boot and the DOM menu bar', () => {
  test('boots to an untitled empty project with the complete menu vocabulary', async ({ page }) => {
    await boot(page)
    await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
    await expect(statusText(page)).toHaveText('Ready')
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'false',
    )
    await expect(page.locator('[data-testid="time-axis"]')).toBeVisible()

    // The DOM menu bar carries the complete shared command vocabulary with
    // displayed accelerators (the native menu's web analog).
    const seen: string[] = []
    for (const top of ['file', 'edit', 'task', 'view']) {
      await page.click(`[data-menu-top="${top}"]`)
      seen.push(
        ...(await page
          .locator('[data-menu-id]')
          .evaluateAll((items) => items.map((item) => (item as HTMLElement).dataset.menuId ?? ''))),
      )
      await page.keyboard.press('Escape')
    }
    expect(seen.sort()).toEqual([...MENU_IDS].sort())
    // Accelerators are DISPLAYED (the native menu's displayed-but-
    // unregistered discipline — the shared keyboard table owns execution).
    await page.click('[data-menu-top="file"]')
    await expect(page.locator('[data-menu-id="file.open"]')).toContainText('Ctrl+O')
    await page.keyboard.press('Escape')
  })

  test('the page reports zero console errors over a full boot', async ({ page }) => {
    const errors: string[] = []
    page.on('pageerror', (error) => errors.push(String(error)))
    page.on('console', (message) => {
      if (message.type() === 'error') errors.push(message.text())
    })
    await boot(page)
    await dropFixture(page)
    await expect(rows(page)).toHaveCount(4)
    expect(errors).toEqual([])
  })
})

test.describe('W02 — open through drag-and-drop and render the canonical projection', () => {
  test('renders rows, derived dates, bars, links, calendar, criticality, and resources', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
    await expect(rows(page)).toHaveCount(4)
    await expect(cellOf(page, 't1', 'taskName')).toContainText('Design')
    // The scheduler authority's derived dates (not host-computed values) —
    // the SAME values the desktop battery asserts (parity).
    await expect(cellOf(page, 't1', 'start')).toHaveText(T1_START)
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_960)
    await expect(cellOf(page, 't2', 'start')).toHaveText(T2_START)
    // Gantt geometry: 3 bars + 1 milestone diamond + 2 dependency links.
    await expect(page.locator('[data-testid="gantt-bar"]')).toHaveCount(3)
    await expect(page.locator('[data-testid="gantt-milestone"]')).toHaveCount(1)
    await expect(
      page.locator('[data-testid="gantt-milestone"][data-task-id="t3"]'),
    ).toHaveAttribute('data-critical', 'true')
    await expect(page.locator('[data-testid="dependency-link"]')).toHaveCount(2)
    // The critical chain carries the authority's flags; the float task not.
    await expect(page.locator('[data-testid="gantt-bar"][data-task-id="t1"]')).toHaveAttribute(
      'data-critical',
      'true',
    )
    await expect(page.locator('[data-testid="gantt-bar"][data-task-id="t4"]')).toHaveAttribute(
      'data-critical',
      'false',
    )
    // t4's float → the PROJECT-026 slack bar (totalSlack 2400).
    await expect(page.locator('[data-testid="slack-bar"][data-task-id="t4"]')).toHaveAttribute(
      'data-total-slack',
      '2400',
    )
    // Calendar working-time bands (working and non-working both present).
    await expect(
      page.locator('[data-testid="calendar-band"][data-working="true"]'),
    ).not.toHaveCount(0)
    await expect(
      page.locator('[data-testid="calendar-band"][data-working="false"]'),
    ).not.toHaveCount(0)
    // Resource utilization: the Crew row with an over-allocated band.
    await expect(page.locator('[data-testid="resource-row"][data-resource-id="r1"]')).toContainText(
      'Crew',
    )
    const over = page.locator('[data-testid="resource-band"][data-overallocated="true"]')
    await expect(over).not.toHaveCount(0)
    await expect(over.first()).toHaveAttribute('data-demand-units', '2')
    await expect(over.first()).toHaveAttribute('data-capacity-units', '1')
  })
})

test.describe('W03 — task creation + cell editing through the command pipeline', () => {
  test('Insert creates a scheduled task; the name editor renames through RenameTask', async ({
    page,
  }) => {
    await boot(page)
    await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
    await page.keyboard.press('Insert')
    await expect(rows(page)).toHaveCount(1)
    await expect(cellOf(page, 't1', 'taskName')).toContainText('New Task')
    await expect(cellOf(page, 't1', 'start')).toHaveText('2026-01-05 09:00')
    await expect(page.locator('[data-testid="gantt-bar"][data-task-id="t1"]')).toBeVisible()
    await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="taskName"]`)
    const editor = page.locator('[data-testid="cell-editor"]')
    await expect(editor).toBeVisible()
    await editor.fill('Pour foundations')
    await page.keyboard.press('Enter')
    await expect(editor).toBeHidden()
    await expect(cellOf(page, 't1', 'taskName')).toContainText('Pour foundations')
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'true',
    )
  })
})

test.describe('W04 — a duration edit moves the derived finish (real scheduling)', () => {
  test('SetTaskDuration through the session re-schedules the chain', async ({ page }) => {
    await boot(page)
    await dropFixture(page)
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_960)
    await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="duration"]`)
    const editor = page.locator('[data-testid="cell-editor"]')
    await expect(editor).toHaveAttribute('data-field', 'duration')
    await editor.fill('1920')
    await page.keyboard.press('Enter')
    // 4 working days from Mon 2026-01-05 → Thu 2026-01-08 17:00.
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
    // The FS successor moves with it (the dependency semantics).
    await expect(cellOf(page, 't2', 'start')).toHaveText('2026-01-09 09:00')
  })
})

test.describe('W05 — undo / redo', () => {
  test('Ctrl+Z restores the prior document AND schedule; redo reapplies', async ({ page }) => {
    await boot(page)
    await dropFixture(page)
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_960)
    await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="duration"]`)
    const editor = page.locator('[data-testid="cell-editor"]')
    await editor.fill('1920')
    await page.keyboard.press('Enter')
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
    await expect(cellOf(page, 't2', 'start')).toHaveText('2026-01-09 09:00')
    // Undo restores the exact prior document — and the schedule re-derives.
    await page.keyboard.press('Control+z')
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_960)
    await expect(cellOf(page, 't2', 'start')).toHaveText(T2_START)
    // Redo reapplies.
    await page.keyboard.press('Control+y')
    await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
    await expect(cellOf(page, 't2', 'start')).toHaveText('2026-01-09 09:00')
  })
})

test.describe('W06 — save through the browser download transport', () => {
  test('Ctrl+S downloads the canonical bytes; a dropped reopen restores the edited state', async ({
    page,
  }) => {
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-project-web-save-'))
    await boot(page)
    await dropFixture(page)
    // Edit two cells (duration + name).
    await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="duration"]`)
    const editor = page.locator('[data-testid="cell-editor"]')
    await editor.fill('1920')
    await page.keyboard.press('Enter')
    await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="taskName"]`)
    const nameEditor = page.locator('[data-testid="cell-editor"]')
    await nameEditor.fill('Design Extended')
    await page.keyboard.press('Enter')
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'true',
    )
    // Ctrl+S → the canonical adapter → the browser download transport.
    const downloadPromise = page.waitForEvent('download')
    await page.keyboard.press('Control+s')
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('e2e-build.gproj')
    const savedPath = join(outDir, download.suggestedFilename())
    await download.saveAs(savedPath)
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'false',
    )
    await expect(statusText(page)).toContainText('Saved')
    // The downloaded bytes are the adapter's own output (readable back).
    const roundTrip = gprojFileAdapter.import(new Uint8Array(readFileSync(savedPath)))
    expect(roundTrip.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
    expect(roundTrip.document.tasks[0]!.name).toBe('Design Extended')
    // A fresh page load + drop of the SAVED file restores the edited state
    // (the web analog of the desktop reopen).
    const fresh = await page.context().newPage()
    await boot(fresh)
    await dropBytes(fresh, 'e2e-build.gproj', {
      base: Array.from(new Uint8Array(readFileSync(savedPath))),
    })
    await expect(cellOf(fresh, 't1', 'duration')).toHaveText('1920')
    await expect(cellOf(fresh, 't1', 'finish')).toHaveText(T1_FINISH_1920)
    await expect(cellOf(fresh, 't1', 'taskName')).toContainText('Design Extended')
    await fresh.close()
  })
})

test.describe('W07 — Save As through the menu', () => {
  test('Save As… downloads under the shared default file name', async ({ page }) => {
    await boot(page)
    await dropFixture(page)
    await expect(rows(page)).toHaveCount(4)
    const downloadPromise = page.waitForEvent('download')
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.saveAs"]')
    const download = await downloadPromise
    // The shared document flow's default name (the project name).
    expect(download.suggestedFilename()).toBe('E2E Build.gproj')
    await expect(statusText(page)).toContainText('Saved')
  })
})

test.describe('W08 — open through the real file picker transport', () => {
  test('the menu click → real file input → picker → bounded read → adapter → renderer pipeline', async ({
    page,
  }) => {
    const fixture = await writeE2EFixture()
    await boot(page)
    const chooserPromise = page.waitForEvent('filechooser')
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.open"]')
    const chooser = await chooserPromise
    await chooser.setFiles(fixture)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
    await expect(rows(page)).toHaveCount(4)
    await expect(cellOf(page, 't2', 'start')).toHaveText(T2_START)
  })

  test('a failed open (corrupt file) keeps the current document and surfaces the error', async ({
    page,
  }) => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-project-web-corrupt-'))
    const corrupt = join(dir, 'corrupt.gproj')
    const { writeFile } = await import('node:fs/promises')
    await writeFile(corrupt, new Uint8Array([0x7b, 0x7f, 0xff, 0xfe, 0x00, 0x7d])) // {!…}
    await boot(page)
    await dropFixture(page)
    await expect(rows(page)).toHaveCount(4)
    // Open the corrupt file through the real picker.
    const chooserPromise = page.waitForEvent('filechooser')
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.open"]')
    const chooser = await chooserPromise
    await chooser.setFiles(corrupt)
    await expect(statusText(page)).toContainText('Open failed')
    // The current document survives untouched.
    await expect(rows(page)).toHaveCount(4)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
  })
})

test.describe('W09 — keyboard navigation + selection', () => {
  test('arrows walk the rows; shift extends; grid click selects; bar click hit-tests', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await rowOf(page, 't1').click()
    await expect(rowOf(page, 't1')).toHaveAttribute('data-selected', 'true')
    await page.keyboard.press('ArrowDown')
    await expect(rowOf(page, 't2')).toHaveAttribute('data-focused', 'true')
    await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
    await page.keyboard.press('Shift+ArrowUp')
    await expect(rowOf(page, 't1')).toHaveAttribute('data-selected', 'true')
    await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
    await page.keyboard.press('Home')
    await expect(rowOf(page, 't1')).toHaveAttribute('data-focused', 'true')
    // Clicking a Gantt BAR selects the row through the canonical hit-test
    // inverse (hitTestGantt).
    await page.locator('[data-testid="gantt-bar"][data-task-id="t2"]').click()
    await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
  })
})

test.describe('W10 — viewport zoom + fit', () => {
  test('zoom out widens the axis; fit through the menu restores the project span', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    const bands = page.locator('[data-testid="time-axis"] .gp-axis-band')
    const initial = await bands.count()
    await page.keyboard.press('Control+-')
    const zoomedOut = await bands.count()
    expect(zoomedOut).toBeGreaterThan(initial)
    // Fit through the REAL menu path restores the span-derived axis.
    await page.click('[data-menu-top="view"]')
    await page.click('[data-menu-id="view.fit"]')
    const fitted = await bands.count()
    expect(fitted).toBeLessThan(zoomedOut)
  })
})

test.describe('W11 — the unsaved-changes guards', () => {
  test('a clean reload raises no beforeunload confirmation', async ({ page }) => {
    await boot(page)
    let dialogs = 0
    page.on('dialog', async (dialog) => {
      dialogs += 1
      await dialog.dismiss()
    })
    await page.reload()
    await expect(page.locator('[data-testid="project-app"]')).toBeVisible()
    await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
    expect(dialogs).toBe(0)
  })

  test('a DIRTY reload raises the native beforeunload confirmation', async ({ page }) => {
    await boot(page)
    await dropFixture(page)
    await page.keyboard.press('Insert')
    await expect(rows(page)).toHaveCount(5)
    // An auto-accepting listener: page.reload() blocks until the dialog is
    // handled, so the accept must run from the listener (not awaited
    // afterwards — that would deadlock the reload).
    const dialogs: string[] = []
    page.on('dialog', async (dialog) => {
      dialogs.push(dialog.type())
      await dialog.accept()
    })
    await page.reload()
    expect(dialogs).toEqual(['beforeunload'])
    // After accepting the leave, the fresh boot is clean/untitled.
    await expect(page.locator('[data-testid="project-app"]')).toBeVisible()
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
  })

  test('File→New over unsaved changes: Cancel keeps the document; Don’t Save discards', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await page.keyboard.press('Insert')
    await expect(rows(page)).toHaveCount(5)
    // Cancel: the dialog refuses, the document survives.
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.new"]')
    await expect(page.locator('[data-testid="discard-dialog"]')).toBeVisible()
    await page.click('[data-testid="discard-cancel"]')
    await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    await expect(rows(page)).toHaveCount(5)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
    // Don't Save: the new untitled project loads.
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.new"]')
    await page.click('[data-testid="discard-dont-save"]')
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
    await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
  })

  test('File→New over unsaved changes: Save downloads first, then loads the new project', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await page.keyboard.press('Insert')
    await expect(rows(page)).toHaveCount(5)
    const downloadPromise = page.waitForEvent('download')
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.new"]')
    await page.click('[data-testid="discard-save"]')
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('e2e-build.gproj')
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
    await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
  })
})

test.describe('W12 — dependency editing through the PROJECT-024 flow', () => {
  test('link click selects; a real double-click edits the lag through ChangeLag', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await page.locator('[data-testid="dependency-link"][data-dependency-id="d1"]').first().click()
    await expect(
      page.locator('[data-testid="dependency-link"][data-dependency-id="d1"]').first(),
    ).toHaveAttribute('data-selected', 'true')
    await page
      .locator('[data-testid="dependency-link"][data-dependency-id="d1"]')
      .first()
      .dblclick()
    const editor = page.locator('[data-testid="cell-editor"]')
    await expect(editor).toHaveAttribute('data-dependency-id', 'd1')
    await editor.fill('120')
    await page.keyboard.press('Enter')
    // The grid's predecessor cell of t2 shows the canonical link text.
    await expect(cellOf(page, 't2', 'predecessors')).toHaveText('1FS+120')
    // The lag edit moved the schedule: +120 working minutes after the
    // predecessor's finish (Tue 17:00) → Wed 11:00.
    await expect(cellOf(page, 't2', 'start')).toHaveText('2026-01-07 11:00')
  })
})

test.describe('W13 — determinism of the rendered app', () => {
  test('two fresh page loads of the same document render identical canonical DOM', async ({
    page,
  }) => {
    const content: string[] = []
    for (let load = 0; load < 2; load += 1) {
      await boot(page)
      await dropFixture(page)
      await expect(rows(page)).toHaveCount(4)
      content.push(await canonicalDom(page))
    }
    expect(content[1]).toBe(content[0])

    async function canonicalDom(page: Page): Promise<string> {
      return page.evaluate(() => {
        const app = document.querySelector('[data-testid="project-app"]')
        if (app === null) throw new Error('app root missing')
        const clone = app.cloneNode(true) as HTMLElement
        // Drop the purely presentational style attributes; keep the
        // canonical data-* surface (ids, dates, flags, counts).
        for (const element of [...clone.querySelectorAll('[style]')]) {
          element.removeAttribute('style')
        }
        return clone.innerHTML
      })
    }
  })
})

test.describe('W14 — bounded web reads (the single web transport cap)', () => {
  test('an OVERSIZED dropped document is refused: the cap error surfaces, nothing loads', async ({
    page,
  }) => {
    await boot(page)
    // The oversized payload is a VALID document padded with trailing JSON
    // whitespace (the desktop battery's construction — byte-size is the
    // only variable, so a refused load proves the transport refused).
    const base = Array.from(e2eFixtureBytes())
    await dropBytes(page, 'oversized.gproj', { base, total: MAX_WEB_FILE_BYTES + 1 })
    await expect(statusText(page)).toContainText('Open failed')
    await expect(statusText(page)).toContainText(CAP_ERROR)
    await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
    await expect(rows(page)).toHaveCount(0)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
  })

  test('a read at EXACTLY the cap succeeds and delivers exactly the cap', async ({ page }) => {
    await boot(page)
    // The boundary READ through the REAL web transport (the staged File →
    // the bounded readFile): ok with EXACTLY the cap's bytes delivered
    // into the page. Driving the exposed bridge (`window.projectWeb`, the
    // desktop battery's `window.projectDesktop` precedent) and detaching
    // the controller's open flow keeps this proof about the TRANSPORT
    // layer only — the .gproj adapter's per-byte UTF-8 decode cannot parse
    // a 100 MiB document in bounded memory (the pre-existing project-file
    // characteristic the desktop exact-cap test documents identically).
    const base = Array.from(e2eFixtureBytes())
    const read = await page.evaluate(
      async ({ base: baseBytes }) => {
        const bridge = (
          window as unknown as {
            projectWeb: {
              onOpenRequested(handler: (path: string) => void): void
              stageExternalFile(file: File): void
              readFile(path: string): Promise<{ ok: boolean; bytes?: Uint8Array; error?: string }>
            }
          }
        ).projectWeb
        bridge.onOpenRequested(() => {})
        const buffer = new Uint8Array(104_857_600)
        buffer.set(baseBytes)
        buffer.fill(0x20, baseBytes.length)
        const file = new File([buffer], 'boundary.gproj')
        bridge.stageExternalFile(file)
        const result = await bridge.readFile('boundary.gproj')
        return result.ok
          ? { ok: true, len: result.bytes?.byteLength ?? -1 }
          : { ok: false, error: result.error }
      },
      { base },
    )
    expect(read).toEqual({ ok: true, len: MAX_WEB_FILE_BYTES })
    // The app itself stays healthy after the boundary transfer.
    await expect(statusText(page)).toHaveText('Ready')
  })

  test('an OVERSIZED picker selection is refused; the current document survives', async ({
    page,
  }) => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-project-web-padded-'))
    const oversized = join(dir, 'oversized.gproj')
    const { writeFile } = await import('node:fs/promises')
    const base = e2eFixtureBytes()
    const padded = new Uint8Array(MAX_WEB_FILE_BYTES + 1)
    padded.set(base)
    padded.fill(0x20, base.byteLength)
    await writeFile(oversized, padded)
    await boot(page)
    await dropFixture(page)
    await expect(rows(page)).toHaveCount(4)
    const chooserPromise = page.waitForEvent('filechooser')
    await page.click('[data-menu-top="file"]')
    await page.click('[data-menu-id="file.open"]')
    const chooser = await chooserPromise
    await chooser.setFiles(oversized)
    await expect(statusText(page)).toContainText('Open failed')
    await expect(statusText(page)).toContainText(CAP_ERROR)
    // The current document survives untouched (no bytes crossed).
    await expect(rows(page)).toHaveCount(4)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
  })
})

test.describe('W15 — the corrected unload lifecycle boundary', () => {
  // The beforeunload guard is PURELY SYNCHRONOUS: it consults the dirty
  // probe and (when dirty) prevents the unload so the browser asks the
  // user natively — it NEVER initiates the controller's asynchronous
  // Save/Don't-Save/Cancel close handshake, which cannot complete inside
  // the browser's synchronous unload lifecycle. The in-app close request
  // (`window.projectWeb.requestClose`, the native window-close button's
  // web analog) is the one firing path for that handshake.

  test('a DIRTY beforeunload dispatch stays purely synchronous: no controller dialog, no orphaned DOM dialog', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await page.keyboard.press('Insert')
    await expect(rows(page)).toHaveCount(5)
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'true',
    )

    // A synthetic beforeunload dispatch runs the real guard listener; a
    // synthetic event raises NO native leave prompt (only real
    // navigations do), so any dialog observed here would be the
    // controller's — the defect this correction removes.
    const dialogs: string[] = []
    page.on('dialog', async (dialog) => {
      dialogs.push(dialog.type())
      await dialog.dismiss()
    })

    const prevented = await page.evaluate(() => {
      const event = new Event('beforeunload', { cancelable: true })
      window.dispatchEvent(event)
      return event.defaultPrevented
    })
    expect(prevented).toBe(true)

    // The async close handshake was never begun: no DOM discard dialog,
    // now or after the queue drains; the dirty document is untouched.
    await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    await page.waitForTimeout(150)
    await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    await expect(rows(page)).toHaveCount(5)
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'true',
    )
    expect(dialogs).toEqual([])
  })

  test('the IN-APP close request runs the shared Save/Don’t-Save/Cancel flow to completion — never during unload', async ({
    page,
  }) => {
    await boot(page)
    await dropFixture(page)
    await page.keyboard.press('Insert')
    await expect(rows(page)).toHaveCount(5)
    await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'true',
    )

    // The in-app path never touches the unload lifecycle: no native
    // beforeunload dialog may appear at any point in this test.
    const dialogs: string[] = []
    page.on('dialog', async (dialog) => {
      dialogs.push(dialog.type())
      await dialog.dismiss()
    })

    // In-app close → the shared controller's three-button dialog.
    await page.evaluate(() => {
      ;(window as unknown as { projectWeb: { requestClose(): void } }).projectWeb.requestClose()
    })
    await expect(page.locator('[data-testid="discard-dialog"]')).toBeVisible()

    // Cancel refuses: the dialog settles, the document survives.
    await page.click('[data-testid="discard-cancel"]')
    await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    await expect(rows(page)).toHaveCount(5)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')

    // In-app close again → Don't Save completes the handshake. The page
    // does NOT unload (the browser owns that): the still-open document
    // remains loaded — the in-app close settles the guard question only.
    await page.evaluate(() => {
      ;(window as unknown as { projectWeb: { requestClose(): void } }).projectWeb.requestClose()
    })
    await expect(page.locator('[data-testid="discard-dialog"]')).toBeVisible()
    await page.click('[data-testid="discard-dont-save"]')
    await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    await expect(rows(page)).toHaveCount(5)
    await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')

    // The unload lifecycle remains independent of the settled handshake:
    // a dirty-document unload attempt still prompts natively (here: the
    // synthetic dispatch is still prevented, still with no DOM dialog).
    const prevented = await page.evaluate(() => {
      const event = new Event('beforeunload', { cancelable: true })
      window.dispatchEvent(event)
      return event.defaultPrevented
    })
    expect(prevented).toBe(true)
    await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    expect(dialogs).toEqual([])
  })
})

test.describe('W16 — the shared ribbon (PROJECT-029)', () => {
  test('boots with the complete ribbon vocabulary and honest boot echoes', async ({ page }) => {
    await boot(page)
    await expect(page.locator('[data-testid="ribbon"]')).toBeVisible()
    // The three tabs; the Task tab is active by default; tab switching works.
    await expect(page.locator('[data-testid="ribbon-tab"][data-tab="task"]')).toHaveAttribute(
      'aria-selected',
      'true',
    )
    await expect(page.locator('[data-testid="ribbon-panel"][data-tab="view"]')).toBeHidden()
    await page.click('[data-testid="ribbon-tab"][data-tab="view"]')
    await expect(page.locator('[data-testid="ribbon-panel"][data-tab="view"]')).toBeVisible()
    await expect(page.locator('[data-testid="ribbon-panel"][data-tab="task"]')).toBeHidden()
    await page.click('[data-testid="ribbon-tab"][data-tab="task"]')
    // The complete shared command vocabulary, exactly once per command —
    // the SAME 15 ids the DOM menu bar carries (one transport vocabulary).
    const commands = await page
      .locator('[data-testid="ribbon-button"]')
      .evaluateAll((buttons) => buttons.map((button) => button.dataset.command))
    expect(new Set(commands).size).toBe(15)
    expect([...commands].sort()).toEqual([...MENU_IDS].sort())
    // Boot echoes: empty journal, no selection, clean document.
    await expect(
      page.locator('[data-testid="ribbon-button"][data-command="edit.undo"]'),
    ).toBeDisabled()
    await expect(
      page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
    ).toBeDisabled()
    await expect(page.locator('[data-testid="ribbon-dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'false',
    )
  })

  test('the semantic controls drive the REAL engine paths through the ribbon (in the browser)', async ({
    page,
  }) => {
    await boot(page)
    // New Task through the ribbon: builder → engine → scheduler, in-page.
    await page.click('[data-testid="ribbon-button"][data-command="task.create"]')
    await expect(rows(page)).toHaveCount(1)
    await expect(rowOf(page, 't1')).toHaveAttribute('data-selected', 'true')
    // The scheduler derived the new task's cells (the authority's dates).
    await expect(cellOf(page, 't1', 'start')).toContainText('2026-')
    // The selection echoes flip on.
    await expect(
      page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
    ).toBeEnabled()

    // Undo/Redo through the ribbon buttons (the session journal). The undo
    // prunes the deleted-from-document selection (the honest view-state
    // rule), so reselect through the real grid pointer path before the
    // delete — the selection echo follows the real state.
    await page.click('[data-testid="ribbon-button"][data-command="task.create"]')
    await expect(rows(page)).toHaveCount(2)
    const undo = page.locator('[data-testid="ribbon-button"][data-command="edit.undo"]')
    await undo.click()
    await expect(rows(page)).toHaveCount(1)
    await expect(
      page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
    ).toBeDisabled()
    await page.click('[data-testid="ribbon-button"][data-command="edit.redo"]')
    await expect(rows(page)).toHaveCount(2)

    // Delete through the ribbon: the real engine path removes the task and
    // the selection echo follows (the pruned selection disables the control).
    await rowOf(page, 't2').click()
    await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
    await expect(
      page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
    ).toBeEnabled()
    await page.click('[data-testid="ribbon-button"][data-command="edit.deleteTask"]')
    await expect(rows(page)).toHaveCount(1)
    await expect(
      page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
    ).toBeDisabled()
  })

  test('Save through the ribbon downloads the canonical bytes; View controls zoom and fit', async ({
    page,
  }) => {
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-project-web-ribbon-'))
    await boot(page)
    await dropFixture(page)
    // A real edit through the canonical cell-editor flow.
    await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="taskName"]`)
    const editor = page.locator('[data-testid="cell-editor"]')
    await editor.fill('Ribbon Renamed')
    await page.keyboard.press('Enter')
    await expect(page.locator('[data-testid="ribbon-dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'true',
    )
    // Save through the ribbon (the File tab) → the browser download
    // transport → the adapter's own bytes.
    await page.click('[data-testid="ribbon-tab"][data-tab="file"]')
    await expect(page.locator('[data-testid="ribbon-panel"][data-tab="file"]')).toBeVisible()
    const downloadPromise = page.waitForEvent('download')
    await page.click('[data-testid="ribbon-button"][data-command="file.save"]')
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('e2e-build.gproj')
    const savedPath = join(outDir, download.suggestedFilename())
    await download.saveAs(savedPath)
    await expect(page.locator('[data-testid="ribbon-dirty-indicator"]')).toHaveAttribute(
      'data-dirty',
      'false',
    )
    // The downloaded bytes are the adapter's own output (readable back).
    const roundTrip = gprojFileAdapter.import(new Uint8Array(readFileSync(savedPath)))
    expect(roundTrip.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
    expect(roundTrip.document.tasks[0]!.name).toBe('Ribbon Renamed')

    // The View tab: zoom out widens the axis; fit restores the span
    // (the same assertions the W10 keyboard/menu battery makes).
    await page.click('[data-testid="ribbon-tab"][data-tab="view"]')
    await expect(page.locator('[data-testid="ribbon-panel"][data-tab="view"]')).toBeVisible()
    const bands = page.locator('[data-testid="time-axis"] .gp-axis-band')
    const initial = await bands.count()
    await page.click('[data-testid="ribbon-button"][data-command="view.zoomOut"]')
    const zoomedOut = await bands.count()
    expect(zoomedOut).toBeGreaterThan(initial)
    await page.click('[data-testid="ribbon-button"][data-command="view.fit"]')
    const fitted = await bands.count()
    expect(fitted).toBeLessThan(zoomedOut)
  })
})
