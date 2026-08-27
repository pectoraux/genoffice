import { defineConfig } from 'vitest/config'

/**
 * The Project desktop host unit suite (PROJECT-027). jsdom exercises the
 * real DOM rendering layer; the desktop bridge is faked in-memory. The
 * Playwright desktop E2E suite lives under e2e/ with its own config and is
 * deliberately outside this include set.
 */
export default defineConfig({
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
  },
})
