import { defineConfig } from 'vitest/config'

/**
 * The Project web host unit suite (PROJECT-028). jsdom exercises the real
 * web transport bridge DOM (the file input, the discard dialog, the
 * beforeunload guard) and the menu bar. The Playwright web E2E suite lives
 * under e2e/ with its own config and is deliberately outside this include
 * set.
 */
export default defineConfig({
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
  },
})
