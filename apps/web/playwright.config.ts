/**
 * Playwright config — REAL browser E2E for the GenOffice web office editors.
 *
 * Boots the full production-shaped stack:
 *   1. The Node API host (packages/web-host/src/dev-server.ts) on :5179.
 *      This is the SAME vercel-handler used by the deployed Vercel function —
 *      /api/office/* requests cross the real HTTP boundary into
 *      routeOffice() → @genoffice/xlsx-gateway / @genoffice/docx-engine.
 *      It runs in demo mode (in-memory PGlite, no DATABASE_URL) so the
 *      browser can log in via the demo accounts.
 *   2. The Vite dev server (apps/web) on :5178, which proxies /api to :5179
 *      exactly like the local dev workflow.
 *
 * The tests then drive the real browser:
 *   browser → Vite proxy → HTTP → vercel-handler → routeOffice → engines → bytes → browser
 *
 * Run from apps/web:  npx playwright test
 * (or from repo root: npx playwright test -c apps/web/playwright.config.ts)
 */
import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

const API_PORT = 5179
const WEB_PORT = 5178
const API_URL = `http://localhost:${API_PORT}`
const WEB_URL = `http://localhost:${WEB_PORT}`
const APP_DIR = fileURLToPath(new URL('.', import.meta.url))

// Dev-only secrets for the local API host (session cookie signing + magic
// link config). These are throwaway values — the host runs in demo mode with
// an in-memory PGlite database that dies with the process.
const API_ENV = {
  CG_WEB_PORT: String(API_PORT),
  CG_SESSION_SECRET: 'playwright-e2e-session-secret-0123456789abcdef',
  CG_MAGIC_LINK_SECRET: 'playwright-e2e-magic-link-secret-0123456789abcdef',
  // No DATABASE_URL → demo mode (in-memory PGlite + seeded demo accounts).
}

export default defineConfig({
  testDir: './tests/e2e',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL: WEB_URL,
    headless: true,
    viewport: { width: 1440, height: 900 },
    actionTimeout: 20_000,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  outputDir: './tests/e2e/.results',
  webServer: [
    {
      command: 'npx tsx ../../packages/web-host/src/dev-server.ts',
      url: `${API_URL}/api/auth/dev-mode`,
      env: API_ENV,
      reuseExistingServer: false,
      timeout: 120_000,
      stdout: 'ignore',
      stderr: 'pipe',
      cwd: APP_DIR,
    },
    {
      command: 'npx vite --port 5178 --strictPort',
      url: WEB_URL,
      reuseExistingServer: false,
      timeout: 120_000,
      stdout: 'ignore',
      stderr: 'pipe',
      cwd: APP_DIR,
    },
  ],
})
