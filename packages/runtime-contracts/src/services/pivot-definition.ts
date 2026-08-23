/**
 * WorkbookPivotDefinition — runtime-independent contract for the parsed
 * definition of an XLSX pivot table (the structured data a renderer needs
 * to recompute the pivot from source data, refresh on edit, or render
 * the pivot UI).
 *
 * WHY THIS FILE EXISTS (Increment 15A):
 *   Previously `SpreadsheetService.readPivotDefinition()` returned
 *   `Promise<unknown>`, and the engine exposed a generic
 *   `readArchiveEntry(handle, entryPath): Promise<string>` escape-hatch
 *   that let the service pluck arbitrary OOXML parts out of the workbook
 *   zip. The user-driven design review (Phase 2 Increment 15A) closed
 *   both gaps:
 *
 *     1. The engine contract now exposes a Sheets-specific
 *        `readPivotDefinition(handle, pivotTablePath, cacheDefinitionPath)`
 *        that returns a typed `WorkbookPivotDefinition` — no generic ZIP
 *        entry API on the contract surface. The engine implementation
 *        still reads the two XML parts from the on-disk temp file, but
 *        that is private to the adapter.
 *
 *     2. The service and coordinator pass the typed
 *        `WorkbookPivotDefinition` through unchanged. The migrated IPC
 *        handler runs the value through `workbookPivotDefinitionSchema.parse()`
 *        as a frozen-IPC sanity check before returning it to the renderer.
 *
 * STRUCTURAL MIRROR:
 *   The shape of this contract mirrors `PivotDefinition` in
 *   `@genoffice/xlsx-gateway` (the canonical parser). The two types are
 *   structurally assignable in both directions, but they live in separate
 *   packages because runtime-contracts MUST NOT depend on xlsx-gateway
 *   (runtime-contracts is Layer 1, xlsx-gateway is a peer of
 *   platform-electron at Layer 4a). The engine implementation is the
 *   single translation point: it calls the canonical parser and returns
 *   its result typed as `WorkbookPivotDefinition`.
 *
 * FORBIDDEN in this file (and all runtime-contracts):
 *   Electron, node:*, sidecar, Rust, stdio, child_process, snapshotPath,
 *   BrowserWindow, WebContents, wcId, filesystem paths, archive entry paths.
 *   This is a pure data contract.
 */

// ── Shared items (cache field values) ─────────────────────────────────

/**
 * A single value stored in a pivot cache field's `sharedItems` array.
 *
 * Mirrors `PivotSharedItem` in `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`.
 */
export type WorkbookPivotSharedItem = string | number | boolean | null

// ── Field item (axis placement) ───────────────────────────────────────

/**
 * A single axis item in a cache field's `fieldItems` array.
 *
 * `x` is the sharedItems index of the value placed at this axis position
 * (null for subtotal / data-field-position markers).
 *
 * Mirrors `PivotFieldItem` in `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`.
 */
export interface WorkbookPivotFieldItem {
  readonly x: number | null
  readonly hidden: boolean
}

// ── Field grouping (date/numeric ranges) ─────────────────────────────

/**
 * The date unit for a date-based grouped field (year / quarter / month).
 *
 * Mirrors `PivotDateUnit` in
 * `@genoffice/xlsx-gateway/src/domain/pivot-grouping.ts`.
 */
export type WorkbookPivotDateUnit = 'year' | 'quarter' | 'month'

/**
 * Grouping rule for a pivot cache field (date buckets or numeric ranges).
 *
 * When `grouping` is present, `sharedItems` are group labels (not raw
 * source values) and refresh re-buckets raw source values before matching.
 *
 * Mirrors `PivotFieldGrouping` in
 * `@genoffice/xlsx-gateway/src/domain/pivot-grouping.ts`.
 */
export type WorkbookPivotFieldGrouping =
  | { readonly kind: 'date'; readonly dateUnit: WorkbookPivotDateUnit }
  | {
      readonly kind: 'range'
      /** Interval step (one bucket per rangeStep); must be positive. */
      readonly rangeStep: number
      /** Interval start (default 0): bucket boundaries are rangeStart + k*rangeStep. */
      readonly rangeStart?: number | undefined
    }

// ── Cache field (a column in the source data) ───────────────────────

/**
 * A cache field — one column of the pivot's source data range.
 *
 * `sharedItems` are the distinct values in that column (group labels
 * when `grouping` is present).
 *
 * `formula` is set for calculated fields (cacheField@formula) — these
 * take no cache records; their value is computed by the formula over
 * grouped aggregates of other fields.
 *
 * Mirrors `PivotCacheField` in
 * `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`.
 */
export interface WorkbookPivotCacheField {
  readonly name: string
  readonly sharedItems: readonly WorkbookPivotSharedItem[]
  readonly grouping?: WorkbookPivotFieldGrouping | undefined
  readonly formula?: string | undefined
}

// ── Layout line (a row or column in the rendered pivot) ──────────────

/**
 * One row or column in the rendered pivot grid.
 *
 * `t` is the line type: 'data' rows/cols carry members; 'default' is a
 * subtotal, 'grand' is a grand total.
 *
 * `members` is the sharedItems index per axis level that this line fixes
 * (already propagated across the OOXML `r`-attribute repeats).
 *
 * `depth` is how many leading axis levels the line fixes (subtotal depth).
 *
 * `dataField` is the data-field index this line renders (multi-value pivots).
 *
 * Mirrors `PivotLayoutLine` in
 * `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`.
 */
export interface WorkbookPivotLayoutLine {
  readonly t: string
  readonly members: readonly (number | null)[]
  readonly depth: number
  readonly dataField: number
}

// ── Data field (a measure being aggregated) ──────────────────────────

/**
 * "Show values as" modes that the parser carries over (subset of the
 * ECMA-376 `dataField@showDataAs` enum).
 *
 * Mirrors `PivotShowDataAs` in
 * `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`.
 */
export type WorkbookPivotShowDataAs =
  | 'percentOfTotal'
  | 'percentOfRow'
  | 'percentOfCol'

/**
 * A measure being aggregated in the pivot (e.g. SUM of Sales).
 *
 * `formula` is set when `field` points at a cacheField with a formula
 * (calculated data field): each group first SUMs the referenced source
 * fields, then evaluates the formula.
 *
 * Mirrors `PivotDataField` in
 * `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`.
 */
export interface WorkbookPivotDataField {
  readonly name: string
  readonly field: number
  readonly subtotal: string
  readonly showDataAs?: WorkbookPivotShowDataAs | undefined
  readonly formula?: string | undefined
}

// ── Filters (label / value) ──────────────────────────────────────────

/**
 * Label filter: keeps members whose label matches the operation
 * (case-insensitive).
 *
 * Mirrors `PivotLabelFilter` in
 * `@genoffice/xlsx-gateway/src/domain/pivot-filters.ts`.
 */
export interface WorkbookPivotLabelFilter {
  readonly kind: 'label'
  /** The filtered row/column dimension field (cache-field index). */
  readonly field: number
  readonly op: 'equal' | 'contains' | 'beginsWith'
  readonly value: string
}

/**
 * Value filter: keeps members by their aggregate on a data field.
 * Aggregate first, then filter.
 *
 * Mirrors `PivotValueFilter` in
 * `@genoffice/xlsx-gateway/src/domain/pivot-filters.ts`.
 */
export interface WorkbookPivotValueFilter {
  readonly kind: 'value'
  /** The filtered row/column dimension field (cache-field index). */
  readonly field: number
  /** The data field it measures (dataFields index, OOXML iMeasureFld). */
  readonly dataField: number
  readonly op: 'top' | 'greaterThan' | 'between'
  /** top: keep the count members with the largest aggregates. */
  readonly count?: number | undefined
  /** greaterThan: aggregate > from; between: from ≤ aggregate ≤ to. */
  readonly from?: number | undefined
  readonly to?: number | undefined
}

/**
 * Discriminated union of label / value filters.
 *
 * Mirrors `PivotFilterDef` in
 * `@genoffice/xlsx-gateway/src/domain/pivot-filters.ts`.
 */
export type WorkbookPivotFilterDef =
  | WorkbookPivotLabelFilter
  | WorkbookPivotValueFilter

// ── Top-level contract ───────────────────────────────────────────────

/**
 * The parsed definition of an XLSX pivot table — runtime-independent.
 *
 * Contains the structured data a renderer needs to recompute the pivot
 * from source data, refresh on edit, or render the pivot UI. Carries
 * NO engine-specific archive type, NO filesystem paths, NO Electron /
 * Node references.
 *
 * The shape mirrors `PivotDefinition` in
 * `@genoffice/xlsx-gateway/src/gateway/xlsx-pivot.ts`. The engine
 * implementation calls the canonical `parsePivotDefinition()` parser
 * and returns its result typed as `WorkbookPivotDefinition` — the two
 * types are structurally assignable.
 */
export interface WorkbookPivotDefinition {
  /** The output range (e.g. 'Sheet1!A1:C5') the pivot table occupies. */
  readonly outputRef: string
  /** First data row (1-indexed, per OOXML location@firstDataRow). */
  readonly firstDataRow: number
  /** First data column (1-indexed, per OOXML location@firstDataCol). */
  readonly firstDataCol: number
  /** Cache fields (one per source column). */
  readonly fields: readonly WorkbookPivotCacheField[]
  /** Per cache field: axis items (empty for pure data fields). */
  readonly fieldItems: readonly (readonly WorkbookPivotFieldItem[])[]
  /** Cache-field indexes for the row axis; -2 marks the data-field position. */
  readonly rowFields: readonly number[]
  /** Cache-field indexes for the column axis; -2 marks the data-field position. */
  readonly colFields: readonly number[]
  /** Rendered row lines (one per row in the pivot output). */
  readonly rowLines: readonly WorkbookPivotLayoutLine[]
  /** Rendered column lines (one per column in the pivot output). */
  readonly colLines: readonly WorkbookPivotLayoutLine[]
  /** Data fields (measures being aggregated). */
  readonly dataFields: readonly WorkbookPivotDataField[]
  /** Page (report-filter) selections: sharedItems index or null for "all". */
  readonly pageFields: readonly { readonly field: number; readonly item: number | null }[]
  /** Value/label filters (pivotFilters). */
  readonly filters: readonly WorkbookPivotFilterDef[]
  /** Source sheet name. */
  readonly sourceSheet: string
  /** Source range (A1 notation on sourceSheet). */
  readonly sourceRef: string
  /** Reasons this pivot cannot be recomputed; empty = refresh supported. */
  readonly unsupported: readonly string[]
}
