/** Vitest/vite module declarations for the typechecked test suite. */
declare module '*?raw' {
  const content: string
  export default content
}

interface ImportMeta {
  /** Vite's glob import (raw module sources for the architecture scans). */
  glob(
    pattern: string,
    options?: { query?: string; import?: string; eager?: boolean },
  ): Record<string, unknown>
}
