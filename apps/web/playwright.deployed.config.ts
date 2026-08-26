/**
 * Playwright config — verify a DEPLOYED GenOffice instance.
 *
 * Runs the real browser E2E specs (excel-browser / word-browser /
 * word-fidelity) against a live deployment instead of the local dev stack.
 * This is the post-deploy acceptance gate: the browser → deployed HTTPS →
 * serverless function → office engines → bytes path.
 *
 * Usage:
 *   DEPLOYED_BASE_URL=https://genoffice.vercel.app npx playwright test -c playwright.deployed.config.ts
 *
 * Notes:
 *   - No webServer: the deployment under test must already be live.
 *   - The nested-runs spec is EXCLUDED: it imports app source modules via
 *     Vite dev-server URLs (/src/...), which only exist in the local dev
 *     stack, not in a built deployment.
 */
import { defineConfig } from '@playwright/test'

const baseURL = process.env.DEPLOYED_BASE_URL
if (!baseURL) {
  console.error('DEPLOYED_BASE_URL is required (e.g. https://genoffice.vercel.app)')
  process.exit(2)
}

export default defineConfig({
  testDir: './tests/e2e',
  testIgnore: '**/nested-runs.spec.ts',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL,
    headless: true,
    viewport: { width: 1440, height: 900 },
    actionTimeout: 20_000,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  outputDir: './tests/e2e/.results-deployed',
})
