/**
 * PROJECT-027 — the Project desktop app launcher for E2E tests.
 *
 * Boots the real built app (`apps/project/out`) against a scratch userData
 * dir (via GENOFFICE_PROJECT_USER_DATA) so runs never touch real settings
 * and never collide with a running install's single-instance lock — the
 * established repo shell-launcher pattern (`e2e/helpers.ts`).
 *
 * Build first: `npm run build -w @genoffice/project-desktop`.
 */
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'

export const APP_DIR = resolve(fileURLToPath(new URL('..', import.meta.url)))
const APP_MAIN = join(APP_DIR, 'out/main/index.js')

export interface LaunchOptions {
  /** Absolute document path passed as argv — opened by the app at launch. */
  openFile?: string
}

export interface LaunchedProjectApp {
  app: ElectronApplication
  page: Page
  userDataDir: string
}

export async function launchProjectApp(options: LaunchOptions = {}): Promise<LaunchedProjectApp> {
  if (!existsSync(APP_MAIN)) {
    throw new Error(
      `Missing build output at ${APP_MAIN} — run \`npm run build -w @genoffice/project-desktop\` first`,
    )
  }
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-project-e2e-'))
  const require = createRequire(join(APP_DIR, 'package.json'))
  const executablePath = require('electron') as unknown as string
  // ELECTRON_RUN_AS_NODE (set by VS Code/CI hosts) would boot Electron as
  // plain Node with no windows — strip it so the app always starts as an app.
  const { ELECTRON_RUN_AS_NODE: _electronRunAsNode, ...hostEnv } = process.env
  // Linux CI runners restrict unprivileged user namespaces (no usable SUID
  // sandbox) and run under xvfb without GPU — without these the window
  // opens but the renderer never loads (the repo-wide e2e job pattern).
  // Switches go BEFORE the app path so Chromium consumes them and they
  // never leak into the argv the app parses for documents to open.
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  args.push(APP_DIR)
  if (options.openFile) args.push(options.openFile)
  const app = await electron.launch({
    executablePath,
    args,
    env: {
      ...hostEnv,
      GENOFFICE_PROJECT_USER_DATA: userDataDir,
      ...(process.platform === 'linux' ? { ELECTRON_DISABLE_SANDBOX: '1' } : {}),
    },
  })
  const page = await app.firstWindow()
  await waitForDocumentReady(app, page)
  return { app, page, userDataDir }
}

/**
 * Playwright can attach to the Electron window mid-navigation and miss the
 * load lifecycle events entirely (Linux timing) — poll through evaluate
 * instead (the repo pattern). Readiness requires the APP ITSELF to be
 * mounted (`[data-testid="project-app"]`), not merely the document parsed:
 * the renderer's module script (which installs the keyboard listener and
 * the bridge handlers) runs between 'interactive' and 'complete', and a
 * keystroke sent in that window would be lost.
 */
async function waitForDocumentReady(
  app: ElectronApplication,
  page: Page,
  timeoutMs = 30_000,
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    const ready = await page
      .evaluate(() =>
        document.readyState !== 'loading' &&
        window.location.href !== 'about:blank' &&
        document.querySelector('[data-testid="project-app"]') !== null
          ? document.readyState
          : null,
      )
      .catch(() => null)
    if (ready) return
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  const diag = await app
    .evaluate(({ app: electronApp, BrowserWindow }) => ({
      processes: electronApp.getAppMetrics().map((m) => m.type),
      contents: BrowserWindow.getAllWindows().map((w) => w.webContents.getURL()),
    }))
    .catch((e) => String(e))
  throw new Error(`Project window never loaded (url: ${page.url()}, diag: ${JSON.stringify(diag)})`)
}

/** Closes the app with escalating force so the suite never wedges and never
 * leaks an instance (a leaked instance would hold the single-instance lock
 * and kill every later launch). */
export async function closeApp(app: ElectronApplication): Promise<void> {
  const process = app.process()
  let killTimer: NodeJS.Timeout | undefined
  await Promise.race([
    app.close(),
    new Promise<void>((resolvePromise) => {
      killTimer = setTimeout(() => {
        process.kill('SIGTERM')
        resolvePromise()
      }, 10_000)
    }),
  ])
  if (killTimer) clearTimeout(killTimer)
  // Hard guarantee: no lingering main process may survive the suite.
  for (let attempt = 0; attempt < 20 && process.exitCode === null; attempt += 1) {
    if (attempt === 5) process.kill('SIGKILL')
    await new Promise((resolve) => setTimeout(resolve, 200))
  }
}
