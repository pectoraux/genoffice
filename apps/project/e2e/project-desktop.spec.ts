/**
 * PROJECT-027 — the Project desktop E2E battery.
 *
 * REAL desktop verification: every test boots the actual built Electron app
 * (main process + preload bridge + renderer binding the shared renderer
 * core, the real engine/scheduler/adapters). Native file dialogs are the
 * only stubbed surface (Playwright cannot drive native dialogs — the
 * established repo E2E pattern); the dialog STUB still crosses the real
 * IPC → dialog API → fs read/write → adapter → renderer pipeline.
 *
 * Assertions are grounded in the canonical authority's own output for the
 * fixture document (see `fixtures.ts`): the schedule dates, criticality,
 * slack, and over-allocation asserted here were derived by the real
 * scheduler, not chosen by hand.
 */
import { existsSync, readFileSync, statSync } from 'node:fs'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import type { ElectronApplication, Page } from '@playwright/test'
import { closeApp, launchProjectApp, spawnSecondInstance } from './launch'
import { writeE2EFixture, writePaddedE2EFixture } from './fixtures'
import { gprojFileAdapter } from '@genoffice/project-file'
import { MAX_FILE_BYTES } from '../src/main/bounded-read.js'

/** The canonical fixture values (derived by the real scheduler). */
const T1_START = '2026-01-05 09:00'
const T1_FINISH_960 = '2026-01-06 17:00'
const T1_FINISH_1920 = '2026-01-08 17:00'
const T2_START = '2026-01-07 09:00'

const rows = (page: Page) => page.locator('[data-testid="task-row"]')
const rowOf = (page: Page, taskId: string) =>
  page.locator(`[data-testid="task-row"][data-task-id="${taskId}"]`)
const cellOf = (page: Page, taskId: string, column: string) =>
  page.locator(`[data-testid="task-row"][data-task-id="${taskId}"] [data-column="${column}"]`)
const statusText = (page: Page) => page.locator('[data-testid="status-text"]')

/** The transport cap message the bounded helper emits (asserted verbatim
 * wherever a read is refused). */
const CAP_ERROR = `File exceeds the ${MAX_FILE_BYTES} byte limit`

/** Padded fixtures are expensive (size-exact real documents); build each
 * size once per suite run and reuse the on-disk file (read-only usage). */
const paddedFixtures = new Map<number, Promise<string>>()
function paddedFixture(totalBytes: number, name: string): Promise<string> {
  let cached = paddedFixtures.get(totalBytes)
  if (cached === undefined) {
    cached = writePaddedE2EFixture(totalBytes, name)
    paddedFixtures.set(totalBytes, cached)
  }
  return cached
}

/** Dispatches a native menu item click through the REAL application menu. */
async function clickMenuItem(app: ElectronApplication, id: string): Promise<void> {
  await app.evaluate(({ Menu, BrowserWindow }, itemId) => {
    const item = Menu.getApplicationMenu()?.getMenuItemById(itemId)
    if (item === undefined || item === null) throw new Error(`No menu item ${itemId}`)
    item.click(item, BrowserWindow.getAllWindows()[0] ?? undefined, {} as never)
  }, id)
}

/** Stubs the native open dialog to "choose" the given path. */
async function stubOpenDialog(app: ElectronApplication, path: string): Promise<void> {
  await app.evaluate(({ dialog }, chosen) => {
    dialog.showOpenDialog = (async () => ({
      canceled: false,
      filePaths: [chosen],
    })) as typeof dialog.showOpenDialog
  }, path)
}

test.describe('E01 — application boot', () => {
  test('boots to an untitled empty project with the native menu installed', async () => {
    const { app, page } = await launchProjectApp()
    try {
      await expect(page.locator('[data-testid="project-app"]')).toBeVisible()
      await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
      await expect(statusText(page)).toHaveText('Ready')
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'false',
      )
      await expect(page.locator('[data-testid="time-axis"]')).toBeVisible()
      // The real native menu carries the PROJECT-027 command vocabulary.
      const menu = await app.evaluate(({ Menu }) => {
        const items: string[] = []
        for (const top of Menu.getApplicationMenu()?.items ?? []) {
          for (const sub of top.submenu?.items ?? []) items.push(`${top.label}:${sub.id}`)
        }
        return items
      })
      expect(menu).toContain('File:file.new')
      expect(menu).toContain('Edit:edit.undo')
      expect(menu).toContain('Task:task.create')
      expect(menu).toContain('View:view.fit')
      // Accelerators are displayed but NOT registered (single keyboard path).
      const accelerators = await app.evaluate(
        ({ Menu }) =>
          (Menu.getApplicationMenu()?.getMenuItemById('edit.undo') ?? { accelerator: null })
            .accelerator,
      )
      expect(accelerators).toContain('Z')
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E02 — open from argv and render the canonical projection', () => {
  test('renders rows, derived dates, bars, links, calendar, criticality, and resources', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
      await expect(rows(page)).toHaveCount(4)
      await expect(cellOf(page, 't1', 'taskName')).toContainText('Design')
      // The scheduler authority's derived dates (not host-computed values).
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
      // Resource utilization: the Crew row with an over-allocated band
      // (demand 2 over capacity 1 on the shared first day).
      await expect(
        page.locator('[data-testid="resource-row"][data-resource-id="r1"]'),
      ).toContainText('Crew')
      const over = page.locator('[data-testid="resource-band"][data-overallocated="true"]')
      await expect(over).not.toHaveCount(0)
      await expect(over.first()).toHaveAttribute('data-demand-units', '2')
      await expect(over.first()).toHaveAttribute('data-capacity-units', '1')
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E03 — task creation + cell editing through the command pipeline', () => {
  test('Insert creates a scheduled task; the name editor renames through RenameTask', async () => {
    const { app, page } = await launchProjectApp()
    try {
      await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
      // Create via the keyboard translation path.
      await page.keyboard.press('Insert')
      await expect(rows(page)).toHaveCount(1)
      await expect(cellOf(page, 't1', 'taskName')).toContainText('New Task')
      await expect(cellOf(page, 't1', 'start')).toHaveText('2026-01-05 09:00')
      await expect(page.locator('[data-testid="gantt-bar"][data-task-id="t1"]')).toBeVisible()
      // Rename through the real cell editor.
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
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E04 — a duration edit moves the derived finish (real scheduling)', () => {
  test('SetTaskDuration through the session re-schedules the chain', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
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
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E05 — undo / redo', () => {
  test('Ctrl+Z restores the prior document AND schedule; redo reapplies', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="duration"]`)
      const editor = page.locator('[data-testid="cell-editor"]')
      await editor.fill('1920')
      await page.keyboard.press('Enter')
      await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
      // Undo restores the exact prior snapshot (finish + successor both).
      await page.keyboard.press('Control+z')
      await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_960)
      await expect(cellOf(page, 't2', 'start')).toHaveText(T2_START)
      // Redo reapplies.
      await page.keyboard.press('Control+Shift+z')
      await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
      // The history indicators track availability.
      await expect(page.locator('[data-testid="history-label"]')).toHaveAttribute(
        'data-can-undo',
        'true',
      )
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E06 — save and reopen through the canonical file adapter', () => {
  test('Ctrl+S persists to the open path; a fresh launch reopens the edited state', async () => {
    const fixture = await writeE2EFixture()
    const first = await launchProjectApp({ openFile: fixture })
    try {
      await page0EditAndSave(first.page)
    } finally {
      await closeApp(first.app)
    }
    // Reopen: the persisted document round-trips through the adapter.
    const second = await launchProjectApp({ openFile: fixture })
    try {
      await expect(cellOf(second.page, 't1', 'duration')).toHaveText('1920')
      await expect(cellOf(second.page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
      await expect(cellOf(second.page, 't1', 'taskName')).toContainText('Design Extended')
    } finally {
      await closeApp(second.app)
    }

    async function page0EditAndSave(page: Page): Promise<void> {
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
      // Ctrl+S → the canonical adapter → the real filesystem write.
      await page.keyboard.press('Control+s')
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'false',
      )
      await expect(statusText(page)).toContainText('Saved')
      // The bytes on disk are the adapter's own output (readable back).
      const bytes = readFileSync(fixture)
      const roundTrip = gprojFileAdapter.import(bytes)
      expect(roundTrip.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
      expect(roundTrip.document.tasks[0]!.name).toBe('Design Extended')
    }
  })

  test('Save As… with the native save dialog stub writes a new canonical file', async () => {
    const fixture = await writeE2EFixture()
    const launched = await launchProjectApp({ openFile: fixture })
    const { app, page } = launched
    const saveAsPath = join(launched.userDataDir, 'saved-as.gproj')
    try {
      // The fixture document must be loaded before the save is driven.
      await expect(rows(page)).toHaveCount(4)
      await app.evaluate(({ dialog }, path) => {
        dialog.showSaveDialog = (async () => ({
          canceled: false,
          filePath: path,
        })) as typeof dialog.showSaveDialog
      }, saveAsPath)
      await clickMenuItem(app, 'file.saveAs')
      await expect(statusText(page)).toContainText('Saved')
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('saved-as.gproj')
      expect(existsSync(saveAsPath)).toBe(true)
      const roundTrip = gprojFileAdapter.import(readFileSync(saveAsPath))
      expect(roundTrip.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
      expect(roundTrip.document.tasks).toHaveLength(4)
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E07 — open through the native menu + dialog transport', () => {
  test('the real menu click → IPC → dialog stub → fs read → adapter → renderer pipeline', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp()
    try {
      await stubOpenDialog(app, fixture)
      // The REAL menu item click (activation forwarding through the main
      // process), not a synthetic renderer event.
      await clickMenuItem(app, 'file.open')
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
      await expect(rows(page)).toHaveCount(4)
      await expect(cellOf(page, 't2', 'start')).toHaveText(T2_START)
    } finally {
      await closeApp(app)
    }
  })

  test('a failed open (corrupt file) keeps the current document and surfaces the error', async () => {
    const { app, page } = await launchProjectApp()
    try {
      await page.keyboard.press('Insert')
      await expect(rows(page)).toHaveCount(1)
      const corruptPath = join('/tmp', 'corrupt-e2e.gproj')
      await writeFile(corruptPath, new TextEncoder().encode('definitely not a project'))
      await stubOpenDialog(app, corruptPath)
      await clickMenuItem(app, 'file.open')
      // The dirty document: the open flow consults the SHARED unsaved-
      // changes dialog first (PROJECT-030 — the native stub is gone);
      // "Don't Save" discards and continues to the (corrupt) open.
      await expect(page.locator('[data-testid="discard-dialog"]')).toBeVisible()
      await page.click('[data-testid="discard-dont-save"]')
      // The corrupt open failed: the current document survives untouched
      // (the controller never replaces the session on a failed import).
      await expect(statusText(page)).toContainText('Open failed')
      await expect(rows(page)).toHaveCount(1)
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E08 — keyboard navigation + selection', () => {
  test('arrows walk the rows; shift extends; grid click selects', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      // Click the first grid row (the DOM pointer path).
      await rowOf(page, 't1').click()
      await expect(rowOf(page, 't1')).toHaveAttribute('data-selected', 'true')
      // Arrow down moves the focus (and selection) to t2.
      await page.keyboard.press('ArrowDown')
      await expect(rowOf(page, 't2')).toHaveAttribute('data-focused', 'true')
      await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
      // Shift+ArrowUp extends back to t1: both selected.
      await page.keyboard.press('Shift+ArrowUp')
      await expect(rowOf(page, 't1')).toHaveAttribute('data-selected', 'true')
      await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
      // Home jumps to the first row.
      await page.keyboard.press('Home')
      await expect(rowOf(page, 't1')).toHaveAttribute('data-focused', 'true')
      // Clicking a Gantt BAR selects the row through the canonical
      // hit-test inverse (hitTestGantt).
      await page.locator('[data-testid="gantt-bar"][data-task-id="t2"]').click()
      await expect(rowOf(page, 't2')).toHaveAttribute('data-selected', 'true')
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E09 — viewport zoom + fit', () => {
  test('zoom out widens the axis; fit restores the project span', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      const bands = page.locator('[data-testid="time-axis"] .gp-axis-band')
      const initial = await bands.count()
      await page.keyboard.press('Control+-')
      const zoomedOut = await bands.count()
      expect(zoomedOut).toBeGreaterThan(initial)
      // Fit through the REAL menu path restores the span-derived axis.
      await clickMenuItem(app, 'view.fit')
      const fitted = await bands.count()
      expect(fitted).toBeLessThan(zoomedOut)
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E10 — the unsaved-changes close guard (the SHARED dialog)', () => {
  // Since PROJECT-030 the unsaved-changes dialog is the shared host
  // binding's own DOM dialog (the native message box is gone): the close
  // handshake is unchanged (main prevents the close once and forwards to
  // the renderer, which consults the SHARED dialog and approves), so the
  // tests drive the REAL dialog — no stubbed native surface anymore.

  test('a clean window closes immediately (no dialog consultation)', async () => {
    const { app } = await launchProjectApp()
    try {
      await app.evaluate(({ BrowserWindow }) => {
        BrowserWindow.getAllWindows()[0]?.close()
      })
      await Promise.race([
        app.close(),
        new Promise((_resolve, reject) =>
          setTimeout(
            () => reject(new Error('the clean close was blocked (guard consulted?)')),
            8000,
          ),
        ),
      ])
    } finally {
      await closeApp(app).catch(() => undefined)
    }
  })

  test('a dirty window consults the shared dialog; “Don’t Save” approves the close', async () => {
    const { app, page } = await launchProjectApp()
    try {
      await page.keyboard.press('Insert')
      await expect(rows(page)).toHaveCount(1)
      await app.evaluate(({ BrowserWindow }) => {
        BrowserWindow.getAllWindows()[0]?.close()
      })
      // The SHARED dialog renders in the renderer (the close was
      // prevented; the page is alive and interactive).
      const dialog = page.locator('[data-testid="discard-dialog"]')
      await expect(dialog).toBeVisible()
      await expect(dialog).toContainText('Save changes to')
      await page.click('[data-testid="discard-dont-save"]')
      await Promise.race([
        app.close(),
        new Promise((_resolve, reject) =>
          setTimeout(() => reject(new Error('the dirty close never completed')), 8000),
        ),
      ])
    } finally {
      await closeApp(app).catch(() => undefined)
    }
  })

  test('Cancel keeps the dirty window open', async () => {
    const { app, page } = await launchProjectApp()
    try {
      await page.keyboard.press('Insert')
      await expect(rows(page)).toHaveCount(1)
      await app.evaluate(({ BrowserWindow }) => {
        BrowserWindow.getAllWindows()[0]?.close()
      })
      await expect(page.locator('[data-testid="discard-dialog"]')).toBeVisible()
      await page.click('[data-testid="discard-cancel"]')
      // The close was refused: the window (and the dirty document) survive.
      await page.waitForTimeout(500)
      await expect(rows(page)).toHaveCount(1)
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'true',
      )
      await expect(page.locator('[data-testid="discard-dialog"]')).toHaveCount(0)
    } finally {
      await closeApp(app)
    }
  })

  test('Save persists the dirty document, then approves the close', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      // A real edit, then a close: Save must write the adapter's bytes to
      // the OPENED path (no picker — the file has a path) before the
      // approval releases the close.
      await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="taskName"]`)
      const editor = page.locator('[data-testid="cell-editor"]')
      await editor.fill('Close-Save Renamed')
      await page.keyboard.press('Enter')
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'true',
      )
      await app.evaluate(({ BrowserWindow }) => {
        BrowserWindow.getAllWindows()[0]?.close()
      })
      await expect(page.locator('[data-testid="discard-dialog"]')).toBeVisible()
      await page.click('[data-testid="discard-save"]')
      await Promise.race([
        app.close(),
        new Promise((_resolve, reject) =>
          setTimeout(() => reject(new Error('the saving close never completed')), 8000),
        ),
      ])
      // The bytes on disk are the adapter's own output (the save ran
      // before the approval).
      const roundTrip = gprojFileAdapter.import(readFileSync(fixture))
      expect(roundTrip.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
      expect(roundTrip.document.tasks[0]!.name).toBe('Close-Save Renamed')
    } finally {
      await closeApp(app).catch(() => undefined)
    }
  })
})

test.describe('E11 — dependency editing through the PROJECT-024 flow', () => {
  test('link click selects; a real double-click edits the lag through ChangeLag', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      // Select the link through the SVG hit surface (real click).
      await page.locator('[data-testid="dependency-link"][data-dependency-id="d1"]').first().click()
      await expect(
        page.locator('[data-testid="dependency-link"][data-dependency-id="d1"]').first(),
      ).toHaveAttribute('data-selected', 'true')
      // A real double-click on the link activates the lag editor (the host's
      // pointer double-press translation — robust to the selection re-render).
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
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E12 — determinism of the rendered app', () => {
  test('two fresh boots of the same document render identical canonical DOM', async () => {
    const fixture = await writeE2EFixture()
    const first = await launchProjectApp({ openFile: fixture })
    await expect(rows(first.page)).toHaveCount(4)
    const content1 = await canonicalDom(first.page)
    await closeApp(first.app)
    const second = await launchProjectApp({ openFile: fixture })
    await expect(rows(second.page)).toHaveCount(4)
    const content2 = await canonicalDom(second.page)
    await closeApp(second.app)
    expect(content2).toBe(content1)

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

test.describe('E13 — bounded native reads (the single transport cap)', () => {
  test('an OVERSIZED argv document is refused: the cap error surfaces, nothing loads', async () => {
    // The oversized file is a VALID document (padded JSON): if its bytes had
    // crossed the IPC boundary it would have loaded. It must not — the
    // renderer can never receive uncapped file contents.
    const oversized = await paddedFixture(MAX_FILE_BYTES + 1, 'oversized.gproj')
    expect(statSync(oversized).size).toBe(MAX_FILE_BYTES + 1)
    const { app, page } = await launchProjectApp({ openFile: oversized })
    try {
      await expect(statusText(page)).toContainText('Open failed')
      await expect(statusText(page)).toContainText(CAP_ERROR)
      await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
      await expect(rows(page)).toHaveCount(0)
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
    } finally {
      await closeApp(app)
    }
  })

  test('a read at EXACTLY the cap succeeds and delivers exactly the cap', async () => {
    const boundary = await paddedFixture(MAX_FILE_BYTES, 'boundary.gproj')
    expect(statSync(boundary).size).toBe(MAX_FILE_BYTES)
    const { app, page } = await launchProjectApp()
    try {
      // The boundary READ through the REAL transport (the preload-exposed
      // bridge → IPC → the main-process canonical bounded read): ok with
      // EXACTLY the cap's bytes delivered into the renderer. Driving the
      // bridge keeps this proof about the TRANSPORT layer: the .gproj
      // adapter's pure-TS UTF-8 decode is per-byte and cannot parse a
      // 100 MiB document in bounded memory (a pre-existing project-file
      // characteristic, out of this increment's scope and unchanged by
      // the correction — the adapter's own input cap admits the size, so
      // a future decoder fast-path would make the full load work).
      const read = await page.evaluate(async (path) => {
        const bridge = (
          window as unknown as {
            projectDesktop: {
              readFile(
                filePath: string,
              ): Promise<{ ok: boolean; bytes?: Uint8Array; error?: string }>
            }
          }
        ).projectDesktop
        const result = await bridge.readFile(path)
        return result.ok
          ? { ok: true, len: result.bytes!.byteLength }
          : { ok: false, error: result.error }
      }, boundary)
      expect(read).toEqual({ ok: true, len: MAX_FILE_BYTES })
      // The app itself stays healthy after the boundary transfer.
      await expect(statusText(page)).toHaveText('Ready')
    } finally {
      await closeApp(app)
    }
  })

  test('an OVERSIZED picker selection is refused; the current document survives', async () => {
    const fixture = await writeE2EFixture()
    const oversized = await paddedFixture(MAX_FILE_BYTES + 1, 'oversized.gproj')
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      await expect(rows(page)).toHaveCount(4)
      await stubOpenDialog(app, oversized)
      await clickMenuItem(app, 'file.open')
      await expect(statusText(page)).toContainText('Open failed')
      await expect(statusText(page)).toContainText(CAP_ERROR)
      // The current document survives untouched (no bytes crossed).
      await expect(rows(page)).toHaveCount(4)
      await expect(page.locator('[data-testid="file-label"]')).toHaveText('e2e-build.gproj')
    } finally {
      await closeApp(app)
    }
  })

  test('a missing argv path surfaces the read error (ENOENT), nothing loads', async () => {
    const missing = join(tmpdir(), 'genoffice-project-e2e-missing', 'nope.gproj')
    const { app, page } = await launchProjectApp({ openFile: missing })
    try {
      await expect(statusText(page)).toContainText('Open failed')
      await expect(statusText(page)).toContainText('ENOENT')
      await expect(page.locator('[data-testid="empty-state"]')).toBeVisible()
      await expect(rows(page)).toHaveCount(0)
    } finally {
      await closeApp(app)
    }
  })

  test('a directory argv path is refused as unreadable (EISDIR), never loaded', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-project-e2e-dir-'))
    const directoryPath = join(dir, 'unreadable.gproj')
    await mkdir(directoryPath)
    const { app, page } = await launchProjectApp({ openFile: directoryPath })
    try {
      await expect(statusText(page)).toContainText('Open failed')
      await expect(statusText(page)).toContainText('EISDIR')
      await expect(rows(page)).toHaveCount(0)
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E14 — single-instance lock isolation', () => {
  test('a second instance defers (no window) and forwards its argv document to the first', async () => {
    // ONE shared scratch userData dir: both instances key their lock on it.
    const shared = await mkdtemp(join(tmpdir(), 'genoffice-project-e2e-shared-'))
    const doc = await writeE2EFixture('forwarded.gproj')
    const first = await launchProjectApp({ userDataDir: shared })
    try {
      await expect(first.page.locator('[data-testid="empty-state"]')).toBeVisible()
      // Instance B: same dir, a document in argv. It must exit (it deferred
      // to A's lock — and never created a window of its own).
      const code = await spawnSecondInstance({ userDataDir: shared, openFile: doc })
      expect(code).toBe(0)
      // B's document arrived in A through the real second-instance → IPC →
      // bounded-read → adapter → renderer pipeline.
      await expect(rows(first.page)).toHaveCount(4)
      await expect(first.page.locator('[data-testid="file-label"]')).toHaveText('forwarded.gproj')
      await expect(cellOf(first.page, 't1', 'finish')).toHaveText(T1_FINISH_960)
      // A kept EXACTLY one window (B never opened one).
      const windows = await first.app.evaluate(
        ({ BrowserWindow }) => BrowserWindow.getAllWindows().length,
      )
      expect(windows).toBe(1)
    } finally {
      await closeApp(first.app)
    }
  })

  test('a second instance forwards an OVERSIZED document: the cap error surfaces in the first', async () => {
    const shared = await mkdtemp(join(tmpdir(), 'genoffice-project-e2e-shared-'))
    const oversized = await paddedFixture(MAX_FILE_BYTES + 1, 'oversized.gproj')
    const first = await launchProjectApp({ userDataDir: shared })
    try {
      await expect(first.page.locator('[data-testid="empty-state"]')).toBeVisible()
      const code = await spawnSecondInstance({ userDataDir: shared, openFile: oversized })
      expect(code).toBe(0)
      // The forwarded read crossed the SAME canonical bounded helper: the
      // cap error surfaces in A and the empty project survives.
      await expect(statusText(first.page)).toContainText('Open failed')
      await expect(statusText(first.page)).toContainText(CAP_ERROR)
      await expect(rows(first.page)).toHaveCount(0)
      await expect(first.page.locator('[data-testid="file-label"]')).toHaveText('Untitled')
    } finally {
      await closeApp(first.app)
    }
  })

  test('two launches with SEPARATE scratch dirs run independently (isolated locks)', async () => {
    // The discriminating scenario for the ordering correction: with the
    // userData path installed BEFORE the lock request, each scratch dir
    // carries its OWN lock — both instances stay alive simultaneously.
    // (With the path installed after the lock, the second launch would have
    // keyed its lock on the REAL profile, deferred, and quit.)
    const docA = await writeE2EFixture('independent-a.gproj')
    const docB = await writeE2EFixture('independent-b.gproj')
    const first = await launchProjectApp({ openFile: docA })
    const second = await launchProjectApp({ openFile: docB })
    try {
      await expect(first.page.locator('[data-testid="file-label"]')).toHaveText(
        'independent-a.gproj',
      )
      await expect(second.page.locator('[data-testid="file-label"]')).toHaveText(
        'independent-b.gproj',
      )
      await expect(rows(first.page)).toHaveCount(4)
      await expect(rows(second.page)).toHaveCount(4)
      // Each instance owns exactly one window; neither deferred nor died.
      expect(
        await first.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length),
      ).toBe(1)
      expect(
        await second.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length),
      ).toBe(1)
    } finally {
      await closeApp(second.app)
      await closeApp(first.app)
    }
  })
})

test.describe('E15 — the shared ribbon (PROJECT-029)', () => {
  test('boots with the complete ribbon vocabulary and honest boot echoes', async () => {
    const { app, page } = await launchProjectApp()
    try {
      const ribbon = page.locator('[data-testid="ribbon"]')
      await expect(ribbon).toBeVisible()
      // The three tabs; the Task tab is active by default.
      await expect(page.locator('[data-testid="ribbon-tab"][data-tab="task"]')).toHaveAttribute(
        'aria-selected',
        'true',
      )
      await expect(page.locator('[data-testid="ribbon-tab"][data-tab="view"]')).toHaveAttribute(
        'aria-selected',
        'false',
      )
      await expect(page.locator('[data-testid="ribbon-panel"][data-tab="task"]')).toBeVisible()
      await expect(page.locator('[data-testid="ribbon-panel"][data-tab="view"]')).toBeHidden()
      // The complete shared command vocabulary, exactly once per command.
      const commands = await page
        .locator('[data-testid="ribbon-button"]')
        .evaluateAll((buttons) => buttons.map((button) => button.dataset.command))
      expect(new Set(commands).size).toBe(16)
      expect([...commands].sort()).toEqual(
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
      // Boot echoes: empty journal, no selection, clean document.
      await expect(
        page.locator('[data-testid="ribbon-button"][data-command="edit.undo"]'),
      ).toBeDisabled()
      await expect(
        page.locator('[data-testid="ribbon-button"][data-command="edit.redo"]'),
      ).toBeDisabled()
      await expect(
        page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
      ).toBeDisabled()
      await expect(page.locator('[data-testid="ribbon-dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'false',
      )
    } finally {
      await closeApp(app)
    }
  })

  test('the semantic controls drive the REAL engine paths through the ribbon', async () => {
    const { app, page } = await launchProjectApp()
    try {
      // New Task through the ribbon: the builder → engine → scheduler chain.
      await page.click('[data-testid="ribbon-button"][data-command="task.create"]')
      await expect(rows(page)).toHaveCount(1)
      const created = page.locator('[data-testid="task-row"][data-task-id="t1"]')
      await expect(created).toHaveAttribute('data-selected', 'true')
      // The scheduler derived the new task's cells (the authority's dates).
      await expect(cellOf(page, 't1', 'start')).toContainText('2026-')
      // The selection echoes flip on (delete/indent/outdent enabled).
      await expect(
        page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
      ).toBeEnabled()
      await expect(
        page.locator('[data-testid="ribbon-button"][data-command="task.indent"]'),
      ).toBeEnabled()

      // A second task, then Undo/Redo through the ribbon buttons (the
      // session journal — the disabled states follow the honest flags).
      await page.click('[data-testid="ribbon-button"][data-command="task.create"]')
      await expect(rows(page)).toHaveCount(2)
      const undo = page.locator('[data-testid="ribbon-button"][data-command="edit.undo"]')
      const redo = page.locator('[data-testid="ribbon-button"][data-command="edit.redo"]')
      await undo.click()
      await expect(rows(page)).toHaveCount(1)
      await redo.click()
      await expect(rows(page)).toHaveCount(2)
      await undo.click()
      await undo.click()
      await expect(rows(page)).toHaveCount(0)
      await expect(undo).toBeDisabled()

      // Delete through the ribbon: the real engine path removes the task.
      await page.click('[data-testid="ribbon-button"][data-command="task.create"]')
      await expect(rows(page)).toHaveCount(1)
      await page.click('[data-testid="ribbon-button"][data-command="edit.deleteTask"]')
      await expect(rows(page)).toHaveCount(0)
      await expect(
        page.locator('[data-testid="ribbon-button"][data-command="edit.deleteTask"]'),
      ).toBeDisabled()
    } finally {
      await closeApp(app)
    }
  })

  test('Save through the ribbon persists the canonical bytes; View controls zoom and fit', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      // A real edit through the canonical cell-editor flow.
      await page.dblclick(`[data-testid="task-row"][data-task-id="t1"] [data-column="taskName"]`)
      const editor = page.locator('[data-testid="cell-editor"]')
      await editor.fill('Ribbon Renamed')
      await page.keyboard.press('Enter')
      await expect(page.locator('[data-testid="ribbon-dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'true',
      )
      // Save through the ribbon (the File tab): the adapter export → the
      // real fs write.
      await page.click('[data-testid="ribbon-tab"][data-tab="file"]')
      await expect(page.locator('[data-testid="ribbon-panel"][data-tab="file"]')).toBeVisible()
      await page.click('[data-testid="ribbon-button"][data-command="file.save"]')
      await expect(page.locator('[data-testid="ribbon-dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'false',
      )
      await expect(statusText(page)).toContainText('Saved')
      // The bytes on disk are the adapter's own output (readable back).
      const roundTrip = gprojFileAdapter.import(readFileSync(fixture))
      expect(roundTrip.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
      expect(roundTrip.document.tasks[0]!.name).toBe('Ribbon Renamed')

      // The View tab: zoom out widens the axis; fit restores the span.
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
    } finally {
      await closeApp(app)
    }
  })
})

test.describe('E16 — the Task Information dialog (PROJECT-030 — operates on commands)', () => {
  test('opens from the ribbon with the displayed values; commits through the real engine', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      await expect(rows(page)).toHaveCount(4)
      await rowOf(page, 't1').click()
      // The ribbon's Task Information control (Task tab is the default).
      await page.click('[data-testid="ribbon-button"][data-command="task.information"]')
      const dialog = page.locator('[data-testid="task-info-dialog"]')
      await expect(dialog).toBeVisible()
      await expect(dialog).toHaveAttribute('data-task-id', 't1')
      // The DISPLAYED values — the scheduler's own derived dates for the
      // fixture (the canonical constants the battery pins everywhere).
      await expect(page.locator('[data-testid="task-info-name"]')).toHaveValue('Design')
      await expect(page.locator('[data-testid="task-info-duration"]')).toHaveValue('960')
      await expect(page.locator('[data-testid="task-info-duration"]')).toBeEnabled()
      // The displayed schedule (sliced ISO — the shared convention).
      await expect(page.locator('[data-testid="task-info-start"]')).toHaveText(T1_START)
      await expect(page.locator('[data-testid="task-info-finish"]')).toHaveText(T1_FINISH_960)

      // Name + duration through the REAL command pipeline.
      await page.fill('[data-testid="task-info-name"]', 'Dialog Renamed')
      await page.fill('[data-testid="task-info-duration"]', '1920')
      await page.click('[data-testid="task-info-ok"]')
      await expect(page.locator('[data-testid="task-info-dialog"]')).toHaveCount(0)
      // The grid reflects both fields; the scheduler derived the wider span.
      await expect(cellOf(page, 't1', 'taskName')).toContainText('Dialog Renamed')
      await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'true',
      )
      // The dialog's two commands walk back in order under Ctrl+Z
      // (duration first, then the rename).
      await page.keyboard.press('Control+z')
      await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_960)
      await expect(cellOf(page, 't1', 'taskName')).toContainText('Dialog Renamed')
      await page.keyboard.press('Control+z')
      await expect(cellOf(page, 't1', 'taskName')).toContainText('Design')
    } finally {
      await closeApp(app)
    }
  })

  test('an unparseable duration keeps the dialog open with the reason; the fix commits', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      await expect(rows(page)).toHaveCount(4)
      await rowOf(page, 't1').click()
      await page.click('[data-testid="ribbon-button"][data-command="task.information"]')
      await page.fill('[data-testid="task-info-duration"]', 'two days')
      await page.click('[data-testid="task-info-ok"]')
      // The dialog refused the duration: the reason surfaced, dialog open.
      await expect(page.locator('[data-testid="task-info-dialog"]')).toBeVisible()
      await expect(page.locator('[data-testid="task-info-error"]')).toContainText('Invalid edit')
      await expect(cellOf(page, 't1', 'duration')).toHaveText('960')
      // The fix commits through the same pipeline.
      await page.fill('[data-testid="task-info-duration"]', '1920')
      await page.click('[data-testid="task-info-ok"]')
      await expect(page.locator('[data-testid="task-info-dialog"]')).toHaveCount(0)
      await expect(cellOf(page, 't1', 'finish')).toHaveText(T1_FINISH_1920)
    } finally {
      await closeApp(app)
    }
  })

  test('Cancel and Escape mutate nothing; the summary rule disables duration', async () => {
    const fixture = await writeE2EFixture()
    const { app, page } = await launchProjectApp({ openFile: fixture })
    try {
      await expect(rows(page)).toHaveCount(4)
      // A CLEAN loaded document: Cancel must keep it exactly clean.
      await rowOf(page, 't1').click()
      await page.click('[data-testid="ribbon-button"][data-command="task.information"]')
      await page.fill('[data-testid="task-info-name"]', 'Discarded')
      await page.click('[data-testid="task-info-cancel"]')
      await expect(page.locator('[data-testid="task-info-dialog"]')).toHaveCount(0)
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'false',
      )
      await expect(cellOf(page, 't1', 'taskName')).toContainText('Design')
      // Escape: the same honest nothing.
      await page.click('[data-testid="ribbon-button"][data-command="task.information"]')
      await page.fill('[data-testid="task-info-name"]', 'Also Discarded')
      await page.keyboard.press('Escape')
      await expect(page.locator('[data-testid="task-info-dialog"]')).toHaveCount(0)
      await expect(page.locator('[data-testid="dirty-indicator"]')).toHaveAttribute(
        'data-dirty',
        'false',
      )
      await expect(cellOf(page, 't1', 'taskName')).toContainText('Design')

      // (The fixture's tasks carry dependencies — an indent under a linked
      // predecessor is an honest engine rejection, so the summary scenario
      // runs on a fresh document.)
    } finally {
      await closeApp(app)
    }
    const fresh = await launchProjectApp()
    try {
      const page2 = fresh.page
      await page2.keyboard.press('Insert')
      await expect(rows(page2)).toHaveCount(1)
      await page2.keyboard.press('Insert')
      await expect(rows(page2)).toHaveCount(2)
      // Indent t2 under t1 → t1 becomes a summary (a derived roll-up).
      await page2.keyboard.press('ArrowDown')
      await page2.keyboard.press('Alt+Shift+ArrowRight')
      await expect(page2.locator('[data-testid="task-row"][data-task-id="t1"]')).toHaveAttribute(
        'data-summary',
        'true',
      )
      await page2.keyboard.press('ArrowUp')
      await page2.click('[data-testid="ribbon-button"][data-command="task.information"]')
      await expect(page2.locator('[data-testid="task-info-dialog"]')).toHaveAttribute(
        'data-task-id',
        't1',
      )
      await expect(page2.locator('[data-testid="task-info-duration"]')).toBeDisabled()
      await page2.click('[data-testid="task-info-cancel"]')
    } finally {
      await closeApp(fresh.app)
    }
  })

  test('the modal gate: the keyboard is suspended while the dialog is open', async () => {
    const { app, page } = await launchProjectApp()
    try {
      await page.keyboard.press('Insert')
      await expect(rows(page)).toHaveCount(1)
      await page.click('[data-testid="ribbon-button"][data-command="task.information"]')
      await expect(page.locator('[data-testid="task-info-dialog"]')).toBeVisible()
      // Keys reach the dialog's input only — the app's translation path
      // is suspended (Insert would create a second task).
      await page.keyboard.press('Insert')
      await page.waitForTimeout(300)
      await expect(rows(page)).toHaveCount(1)
      // Cancel resumes the app.
      await page.click('[data-testid="task-info-cancel"]')
      await page.keyboard.press('Insert')
      await expect(rows(page)).toHaveCount(2)
    } finally {
      await closeApp(app)
    }
  })
})
