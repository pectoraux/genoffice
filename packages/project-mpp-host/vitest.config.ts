import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    environment: 'node',
    // PROJECT-020: test FILES run sequentially. Every file in this package
    // spawns real JVM sidecar conversions; running files in parallel on a
    // shared CI runner multiplies JVM contention and was the root cause of
    // the documented transient per-test timeouts (the PROJECT-019 CI round
    // and the first PROJECT-020 round). Within a file, tests already run
    // sequentially by default — so at most one conversion pipeline runs at
    // any time on the runner.
    fileParallelism: false,
  },
})
