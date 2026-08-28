import { defineConfig } from '@playwright/test'
import { fileURLToPath } from 'node:url'

/**
 * Playwright config — REAL browser E2E for the GenOffice Project web shell
 * (PROJECT-028).
 *
 * Boots the production-shaped stack: `vite build` (run by the test:e2e
 * script) produces the real browser bundle, and `vite preview` serves it
 * on :5190. The tests then drive the real browser:
 *
 *   chromium → preview server → the built bundle → the shared
 *   @genoffice/project-host binding → the REAL engine / scheduler / file
 *   adapters (all in the browser — the structural "no Node/Electron in the
 *   browser-side implementation" proof, additionally scanned over dist/ in
 *   CI before this suite runs).
 *
 * Run from apps/project-web:  npm run test:e2e
 */
const PORT = 5190
const BASE_URL = `http://localhost:${PORT}`
const APP_DIR = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig({
  testDir: '.',
  outputDir: './test-results',
  timeout: 90_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL: BASE_URL,
    headless: true,
    viewport: { width: 1440, height: 900 },
    actionTimeout: 20_000,
    acceptDownloads: true,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: {
    command: 'npx vite preview --port 5190 --strictPort',
    url: BASE_URL,
    reuseExistingServer: false,
    timeout: 120_000,
    stdout: 'ignore',
    stderr: 'pipe',
    cwd: APP_DIR,
  },
})
