import { defineConfig } from 'vitest/config'

/**
 * The shared host binding suite (PROJECT-028). jsdom exercises the real DOM
 * rendering layer; the transport bridge is faked in-memory (the same fake
 * discipline both hosts' controllers are tested with). The per-host E2E
 * suites live in the host apps with their own configs.
 */
export default defineConfig({
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
  },
})
