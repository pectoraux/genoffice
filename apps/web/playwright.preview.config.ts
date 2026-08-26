/**
 * Playwright config override for PRODUCTION-PREVIEW runs (EXCEL-020
 * pattern): the API host (:5179) and the vite preview server (:5178) are
 * started EXTERNALLY by scripts/run-table-vs-preview.sh; this config only
 * points the suite at them and reuses the running servers.
 */
import { defineConfig } from '@playwright/test'

const API_PORT = 5179
const WEB_PORT = 5178

export default defineConfig({
  testDir: './tests/e2e',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL: `http://localhost:${WEB_PORT}`,
    headless: true,
    viewport: { width: 1440, height: 900 },
    actionTimeout: 20_000,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  outputDir: './tests/e2e/.results-preview',
  webServer: [
    {
      command: 'npx tsx ../../packages/web-host/src/dev-server.ts',
      url: `http://localhost:${API_PORT}/api/auth/dev-mode`,
      reuseExistingServer: true,
      timeout: 120_000,
      stdout: 'ignore',
      stderr: 'pipe',
    },
    {
      command: 'npx vite preview --port 5178 --strictPort',
      url: `http://localhost:${WEB_PORT}`,
      reuseExistingServer: true,
      timeout: 120_000,
      stdout: 'ignore',
      stderr: 'pipe',
    },
  ],
})
