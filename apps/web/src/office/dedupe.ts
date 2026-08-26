/**
 * Remove Duplicates — browser-side pure dedupe (EXCEL-018).
 *
 * Two pure exports:
 *   1. `dedupeRows(rows, hasHeader)` — kept for backward compatibility with
 *      the existing unit suite; returns the kept-row MATRIX (no indices).
 *      Still mirrors the frozen desktop reference's value-level algorithm.
 *   2. `dedupeRowIndices(rows, hasHeader)` — the EXCEL-018 canonical entry
 *      point. Returns the duplicate row INDICES (0-based, relative to the
 *      input matrix) so the runtime can issue `worksheet.deleteRows(...)`
 *      for each duplicate in descending order. The browser NEVER rewrites
 *      moved rows via `setValues` — that path destroys formulas (a moved
 *      `=B6` becomes the literal computed value `30`). Structural row
 *      deletion preserves the entire cell record (formulas, styles,
 *      merges, numfmt, hyperlinks, notes, DV) atomically: the cell
 *      travels untouched inside its `<c>` element while the gateway's
 *      `remove-rows` op renumbers `r=` attributes and rewrites formula
 *      references via `transformFormulas` + `shiftFormulaText` (the same
 *      canonical path `excel-structural.spec.ts` already proves for
 *      Insert/Delete Rows).
 *
 * Both functions are PURE — no Univer, no journal, no save plan, no host
 * API. The caller (`useExcelRuntime.removeDuplicates`) reads the live
 * Univer range values, hands them to `dedupeRowIndices`, and issues
 * `ws.deleteRows(startRow + offset, 1)` per duplicate (in DESCENDING
 * offset order so earlier deletes don't shift later indices). Each call
 * fires `sheet.mutation.remove-rows`, journaled by ExcelEditor's existing
 * `STRUCTURAL_MUTATION_IDS` subscription as a `remove-rows` structural
 * op in the save plan. The gateway's `applyStructuralOps` applies the
 * ops atomically — `transformSheetRows` renumbers rows, `transformFormulas`
 * rewrites formula references (relative + absolute + mixed), and
 * `transformRangedFeatures` shifts merges / autoFilter / hyperlink sqref
 * / dataValidation sqref / conditionalFormatting sqref.
 *
 * Browser purity: this module imports NOTHING. No `electron`, no `node:*`,
 * no `jszip`, no raw OOXML. The architecture test in
 * `apps/web/tests/architecture.test.ts` enforces that the dedupe surface
 * stays canonical AND that the runtime wires through the structural
 * `remove-rows` path (NOT a value-rewrite path that would destroy
 * formulas on moved rows).
 *
 * Equality semantics (matching the frozen desktop reference verbatim):
 *   - Text values are compared CASE-INSENSITIVELY via `toLowerCase()`.
 *     Trailing whitespace is preserved in storage but participates in the
 *     key (so `"Apple "` and `"Apple"` are distinct rows).
 *   - Numbers, booleans, and null are compared by `===` value.
 *   - TYPES are strict: `1` (number) and `"1"` (string) are NOT duplicates.
 *   - Empty string (`""`) and `null` are distinct (a blank cell is `null`,
 *     an explicitly empty-string cell is `""`).
 *   - Formula cells: the caller passes the COMPUTED result (Univer's
 *     `FRange.getValues()` returns computed results, not formula text),
 *     so two rows with the same result but different formulas ARE
 *     duplicates — matching the desktop's behavior exactly.
 *
 * Header handling:
 *   - When `hasHeader === true`, the first row of `rows` is kept verbatim
 *     and is excluded from both the seen-set and the removal check. The
 *     header is never treated as a duplicate even if a later data row
 *     matches it. This matches the desktop's `dedupeRows(rows, true)`
 *     branch exactly.
 *
 * Output shape:
 *   - `dedupeRows` returns `{ rows: keptRows, removed }` (kept rows only
 *     — no padding) for backward compatibility.
 *   - `dedupeRowIndices` returns `{ keptIndices, duplicateIndices, removed }`
 *     where `duplicateIndices` is the list of 0-based offsets to DELETE
 *     (in ascending order; the caller issues the deletes in DESCENDING
 *     order so earlier deletes don't shift later indices).
 */
export type DedupeValue = string | number | boolean | null

/**
 * Compute the dedupe of a value matrix (value-level, kept for backward
 * compatibility with the unit suite). The runtime no longer uses this
 * — it calls `dedupeRowIndices` instead and issues `ws.deleteRows(...)`.
 *
 * @param rows 2D matrix of cell values (the result of
 *   `FRange.getValues()` mapped through `(v) => v ?? null`).
 * @param hasHeader When true, the first row is preserved verbatim and is
 *   excluded from the seen-set / removal check.
 * @returns `{ rows: kept rows (no padding), removed: count of removed rows }`.
 */
export function dedupeRows(
  rows: ReadonlyArray<ReadonlyArray<DedupeValue>>,
  hasHeader: boolean,
): { rows: DedupeValue[][]; removed: number } {
  const keptIndices = dedupeRowIndices(rows, hasHeader)
  return {
    rows: keptIndices.keptIndices.map((i) => [...rows[i]!]),
    removed: keptIndices.removed,
  }
}

/**
 * Compute the dedupe of a value matrix and return the duplicate INDICES
 * (the canonical EXCEL-018 entry point).
 *
 * @param rows 2D matrix of cell values (the result of
 *   `FRange.getValues()` mapped through `(v) => v ?? null`).
 * @param hasHeader When true, the first row is preserved verbatim and is
 *   excluded from the seen-set / removal check.
 * @returns `{ keptIndices, duplicateIndices, removed }`:
 *   - `keptIndices`: 0-based offsets of the surviving rows (ascending).
 *   - `duplicateIndices`: 0-based offsets of the rows to DELETE
 *     (ascending — the caller issues the deletes in DESCENDING order so
 *     earlier deletes don't shift later indices).
 *   - `removed`: count of duplicate rows (= `duplicateIndices.length`).
 */
export function dedupeRowIndices(
  rows: ReadonlyArray<ReadonlyArray<DedupeValue>>,
  hasHeader: boolean,
): {
  keptIndices: number[]
  duplicateIndices: number[]
  removed: number
} {
  const keptIndices: number[] = []
  const duplicateIndices: number[] = []
  const seen = new Set<string>()
  for (const [index, row] of rows.entries()) {
    if (hasHeader && index === 0) {
      // Header is preserved verbatim AND never matched against. The
      // seen-set is not populated with the header, so a data row that
      // happens to equal the header is treated as a duplicate OF THE
      // HEADER'S VALUE only if a prior data row already produced the
      // same key — the header itself never matches.
      keptIndices.push(index)
      continue
    }
    // Mirror the desktop's JSON.stringify key: text is lowercased,
    // numbers/booleans/null ride through as-is. JSON.stringify
    // distinguishes `1` (number) from `"1"` (string) by emitting `1`
    // vs `"1"`, so type-strict equality holds. `null` serializes as
    // `null` and is distinct from `""` (which serializes as `""`).
    const key = JSON.stringify(
      row.map((value) => (typeof value === 'string' ? value.toLowerCase() : value)),
    )
    if (seen.has(key)) {
      duplicateIndices.push(index)
      continue
    }
    seen.add(key)
    keptIndices.push(index)
  }
  return { keptIndices, duplicateIndices, removed: duplicateIndices.length }
}
