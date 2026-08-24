/**
 * Shared helpers for the Playwright browser E2E tests.
 *
 * Covers the app's real auth flow (demo login → tenant selection) and the
 * editors' download capture. Everything crosses the REAL HTTP boundary —
 * no routeOffice() calls, no engine imports in the browser path.
 */
import type { Page, Download } from '@playwright/test'
import { expect } from '@playwright/test'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

/**
 * Log in through the real UI (demo Owner account) and select the demo
 * tenant. The API host runs in demo mode (in-memory PGlite with seeded
 * demo users), so this exercises the real password/session/cookie flow.
 */
export async function loginAsDemoOwner(page: Page): Promise<void> {
  await page.goto('/')
  // Login screen — demo quick-login row.
  await page.getByRole('button', { name: 'Owner', exact: true }).click()
  // Tenant selection screen — the demo org membership.
  await page.getByRole('button', { name: /Enter →/ }).click()
  // AppShell header appears once the tenant is selected.
  await expect(page.getByRole('button', { name: 'Office' }).first()).toBeVisible()
}

/** Navigate to a hash route (the app uses hash-based routing). */
export async function gotoHashRoute(page: Page, route: string): Promise<void> {
  await page.evaluate((r) => {
    window.location.hash = r
  }, route)
}

/**
 * Wait for a genuinely-sized grid canvas inside the Univer container.
 *
 * With SheetsCorePreset header:false + formulaBar:false (create-browser-
 * univer.ts), Univer does NOT mount the SpreadsheetHeader plugin. The
 * container still holds multiple canvases (the grid canvas + a 0×0
 * cell-editor overlay `univer-doc-main-canvas` that is only sized while a
 * cell is being edited). A plain `page.waitForSelector('#... canvas')`
 * latches onto the 0×0 overlay and times out. This waits for ANY canvas
 * with real area — which is the grid.
 */
export async function waitForGridCanvas(page: Page): Promise<void> {
  await page.waitForFunction(
    () => {
      const cs = Array.from(document.querySelectorAll('#genoffice-web-excel canvas')) as HTMLCanvasElement[]
      return cs.some((c) => c.getBoundingClientRect().width > 200 && c.getBoundingClientRect().height > 100)
    },
    { timeout: 30_000 },
  )
}

/**
 * Upload a file through an editor's hidden file input and wait for the
 * editor to report that the open completed (the status line flips from
 * "Opening…" to "Opened <name>").
 */
export async function uploadFixture(
  page: Page,
  inputSelector: string,
  filePath: string,
  fileName: string,
): Promise<void> {
  await page.setInputFiles(inputSelector, filePath)
  await expect(page.getByText(`Opened ${fileName}`)).toBeVisible({ timeout: 30_000 })
}

/**
 * Click the Save button and capture the file the editor downloads (the
 * editors trigger a blob-URL download of the bytes returned by the API).
 * Returns the downloaded bytes.
 */
export async function clickSaveAndCaptureDownload(
  page: Page,
  saveButtonName: string,
): Promise<Buffer> {
  const downloadPromise = page.waitForEvent('download', { timeout: 30_000 })
  await page.getByRole('button', { name: saveButtonName, exact: true }).click()
  const download: Download = await downloadPromise
  // Read the bytes directly from the download stream.
  const stream = await download.createReadStream()
  const chunks: Buffer[] = []
  for await (const chunk of stream) {
    chunks.push(chunk as Buffer)
  }
  return Buffer.concat(chunks)
}

/** Persist a buffer under tests/e2e/.artifacts for inspection/debugging. */
export function saveArtifact(name: string, bytes: Buffer): string {
  const dir = join(__dirname, '.artifacts')
  mkdirSync(dir, { recursive: true })
  const path = join(dir, name)
  writeFileSync(path, bytes)
  return path
}
