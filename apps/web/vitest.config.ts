import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    // Architecture tests scan source files; round-trip tests import Node-only
    // office engines. Run in node.
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    testTimeout: 30_000,
    hookTimeout: 30_000,
  },
})
