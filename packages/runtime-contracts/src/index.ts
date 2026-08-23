/**
 * @genoffice/runtime-contracts — barrel export.
 *
 * Platform-neutral runtime contracts: the RuntimeContext interface, the
 * getRuntime()/setRuntime() bootstrap mechanism, the domain service
 * interfaces, the shell coordinator interface, and the runtime-independent
 * domain types.
 *
 * BOUNDARY CORRECTION (2026-08-21, contract direction):
 *   This package defines its OWN types (in src/types/docs.ts). It does NOT
 *   import from @genoffice/docs-shared or any other @genoffice/*-shared
 *   alias (which point to apps <star> /src/shared). The runtime-independent
 *   layer sits UNDERNEATH the legacy application, not above it.
 *
 * Zero app imports. Zero @genoffice/*-shared imports.
 */
export * from './runtime.js'
export * from './services/docs.js'
export * from './services/sheets.js'
export * from './services/slides.js'
export * from './services/pdf.js'
export * from './services/markdown.js'
export * from './services/project.js'
export * from './services/spreadsheet-engine.js'
export * from './services/spreadsheet-pdf-renderer.js'
export * from './services/pivot-definition.js'
export * from './types/docs.js'
