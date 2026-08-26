/**
 * Remove Duplicates — browser-side pure dedupe (EXCEL-018).
 *
 * This module is the web parity of the frozen desktop reference at
 * `apps/sheets/src/renderer/dedupe.ts`. It is a PURE function: it does not
 * touch Univer, the journal, the save plan, the gateway, or any host API.
 * The caller (`useExcelRuntime.removeDuplicates`) reads the live Univer
 * range values, hands them to `dedupeRows`, and writes the result back
 * through `FWorksheet.getRange(...).setValues(...)` — which fires the
 * existing `sheet.mutation.set-range-values` channel and is journaled
 * by `cellEditFromMutation` into the canonical `CellEdit[]` save plan.
 *
 * Browser purity: this module imports NOTHING. No `electron`, no `node:*`,
 * no `jszip`, no raw OOXML. The architecture test in
 * `apps/web/tests/architecture.test.ts` enforces that the dedupe surface
 * stays canonical.
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
 * Output shape: the returned `rows` array is SHRUNK to the deduped count
 * (kept rows only — no padding). The caller pads with `null` rows back to
 * the original selection height so the in-place rewrite leaves blank rows
 * at the bottom of the selection (where duplicates were removed). This
 * matches the desktop's `while (rows.length < values.length) rows.push(null)`
 * padding step.
 */
export type DedupeValue = string | number | boolean | null

/**
 * Compute the dedupe of a value matrix.
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
  const kept: DedupeValue[][] = []
  const seen = new Set<string>()
  let removed = 0
  for (const [index, row] of rows.entries()) {
    if (hasHeader && index === 0) {
      // Header is preserved verbatim AND never matched against. The
      // seen-set is not populated with the header, so a data row that
      // happens to equal the header is treated as a duplicate OF THE
      // HEADER'S VALUE only if a prior data row already produced the
      // same key — the header itself never matches.
      kept.push([...row])
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
      removed += 1
      continue
    }
    seen.add(key)
    kept.push([...row])
  }
  return { rows: kept, removed }
}
