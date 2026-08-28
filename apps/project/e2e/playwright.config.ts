import { defineConfig } from '@playwright/test'

/**
 * PROJECT-027 — desktop E2E config for the GenOffice Project Electron app.
 *
 * Tests launch the real built app (`electron.launch` against
 * `apps/project/out`), so they run serially — parallel Electron instances
 * fight over the GPU cache. Run with (Linux CI):
 * `xvfb-run --auto-servernum -- npm run test:e2e -w @genoffice/project-desktop`
 * after `npm run build -w @genoffice/project-desktop`.
 */
export default defineConfig({
  testDir: '.',
  outputDir: './test-results',
  timeout: 90_000,
  fullyParallel: false,
  workers: 1,
  retries: 0,
  expect: { timeout: 15_000 },
  reporter: [['list']],
})
