/**
 * Unit tests for the pure dedupe algorithm (EXCEL-018).
 *
 * Covers BOTH exports of `apps/web/src/office/dedupe.ts`:
 *   - `dedupeRows(rows, hasHeader)` — the legacy value-level entry
 *     (kept for backward compat; mirrors the desktop reference
 *     algorithm at apps/sheets/tests/dedupe.test.ts verbatim).
 *   - `dedupeRowIndices(rows, hasHeader)` — the canonical EXCEL-018
 *     entry point. Returns the SET of duplicate row INDICES (0-based
 *     offsets relative to the input matrix). The runtime issues one
 *     `ws.deleteRows(startRow + offset, 1)` call per duplicate in
 *     DESCENDING offset order so earlier deletes don't shift later
 *     indices. Each call fires `sheet.mutation.remove-rows`, journaled
 *     by ExcelEditor's existing `STRUCTURAL_MUTATION_IDS` subscription
 *     as a `remove-rows` structural op in the save plan. The gateway's
 *     `applyStructuralOps` applies each op atomically —
 *     `transformSheetRows` renumbers `<row>` r= and inner `<c>` r=
 *     (cell contents travel UNTOUCHED inside their `<c>` elements),
 *     `transformFormulas` rewrites `<f>` bodies via `shiftFormulaText`
 *     (relative + absolute + mixed references all track the moved
 *     cells — the `$` markers are preserved by `shiftReferenceToken`'s
 *     colDollar/rowDollar capture groups), and `transformRangedFeatures`
 *     shifts merges, autoFilter, hyperlink sqref, dataValidation sqref,
 *     conditionalFormatting sqref. This is the EXACT canonical path
 *     `excel-structural.spec.ts` already proves for Insert/Delete Rows —
 *     no value-rewrite, no formula loss.
 */
import { describe, expect, it } from 'vitest'

import { dedupeRows, dedupeRowIndices } from '../src/office/dedupe'

describe('dedupeRows', () => {
  it('keeps the first occurrence and reports the removed count', () => {
    const { rows, removed } = dedupeRows(
      [
        ['a', 1],
        ['b', 2],
        ['a', 1],
        ['a', 2],
        ['b', 2],
      ],
      false,
    )
    expect(rows).toEqual([
      ['a', 1],
      ['b', 2],
      ['a', 2],
    ])
    expect(removed).toBe(2)
  })

  it('compares text case-insensitively', () => {
    const { rows, removed } = dedupeRows([['Apple'], ['APPLE'], ['apple ']], false)
    expect(rows).toEqual([['Apple'], ['apple ']])
    expect(removed).toBe(1)
  })

  it('never removes the header row and never matches against it', () => {
    const { rows, removed } = dedupeRows([['Name'], ['Name'], ['name']], true)
    expect(rows).toEqual([['Name'], ['Name']])
    expect(removed).toBe(1)
  })

  it('does not confuse types or nulls with equal string forms', () => {
    const { rows, removed } = dedupeRows(
      [[1, null], ['1', null], [null, null], [null, null], [true], ['true']],
      false,
    )
    expect(removed).toBe(1)
    expect(rows).toHaveLength(5)
  })

  // ── Additional coverage for the EXCEL-018 verification matrix ──────

  it('handles a basic two-column duplicate-row scenario', () => {
    // Example from the work-item brief:
    //   A  B
    //   A  X
    //   A  X    ← duplicate of row 0
    //   B  Y
    const { rows, removed } = dedupeRows(
      [
        ['A', 'X'],
        ['A', 'X'],
        ['B', 'Y'],
      ],
      false,
    )
    expect(rows).toEqual([
      ['A', 'X'],
      ['B', 'Y'],
    ])
    expect(removed).toBe(1)
  })

  it('treats the header as data when hasHeader is false', () => {
    // Same fixture as the header case, but hasHeader=false → the header
    // row IS in the seen-set, so a later row matching it gets removed.
    const { rows, removed } = dedupeRows([['Name'], ['Name'], ['name']], false)
    expect(rows).toEqual([['Name']])
    expect(removed).toBe(2)
  })

  it('uses ALL selected columns as the comparison key', () => {
    // All rows are identical EXCEPT one selected comparison column:
    //   row 0: A, 1
    //   row 1: A, 1   ← duplicate of row 0 (both columns match)
    //   row 2: A, 2   ← NOT a duplicate (column 1 differs)
    //   row 3: A, 1   ← duplicate of row 0
    const { rows, removed } = dedupeRows(
      [
        ['A', 1],
        ['A', 1],
        ['A', 2],
        ['A', 1],
      ],
      false,
    )
    expect(rows).toEqual([
      ['A', 1],
      ['A', 2],
    ])
    expect(removed).toBe(2)
  })

  it('keeps duplicate rows adjacent to non-duplicate rows correctly', () => {
    //   row 0: A, 1
    //   row 1: B, 2   ← unique
    //   row 2: A, 1   ← duplicate of row 0 (NOT adjacent)
    //   row 3: C, 3   ← unique
    //   row 4: B, 2   ← duplicate of row 1 (NOT adjacent)
    const { rows, removed } = dedupeRows(
      [
        ['A', 1],
        ['B', 2],
        ['A', 1],
        ['C', 3],
        ['B', 2],
      ],
      false,
    )
    expect(rows).toEqual([
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ])
    expect(removed).toBe(2)
  })

  it('distinguishes blank cells from empty-string cells', () => {
    // null (blank cell) and '' (empty-string cell) are DISTINCT under
    // the desktop's equality rule — JSON.stringify distinguishes them
    // (null vs "") and the dedupe preserves both as separate keys.
    const { rows, removed } = dedupeRows(
      [
        [null],
        [''],
        [null], // duplicate of row 0
        [''], // duplicate of row 1
      ],
      false,
    )
    expect(rows).toEqual([[null], ['']])
    expect(removed).toBe(2)
  })

  it('distinguishes number from text with the same printed form', () => {
    // 1 (number) and '1' (string) are NOT duplicates — the desktop's
    // equality is type-strict.
    const { rows, removed } = dedupeRows([[1], ['1'], [1]], false)
    expect(rows).toEqual([[1], ['1']])
    expect(removed).toBe(1)
  })

  it('distinguishes boolean from text with the same printed form', () => {
    // true (boolean) and 'true' (string) are NOT duplicates.
    const { rows, removed } = dedupeRows([[true], ['true'], [true]], false)
    expect(rows).toEqual([[true], ['true']])
    expect(removed).toBe(1)
  })

  it('returns an empty matrix when all rows are duplicates of the first', () => {
    const { rows, removed } = dedupeRows(
      [
        ['A', 1],
        ['A', 1],
        ['A', 1],
      ],
      false,
    )
    expect(rows).toEqual([['A', 1]])
    expect(removed).toBe(2)
  })

  it('preserves a 1-row input with hasHeader=true (header-only, no data)', () => {
    // Edge case: a single header row and no data rows. The dedupe must
    // keep the header and report zero removed.
    const { rows, removed } = dedupeRows([['Header']], true)
    expect(rows).toEqual([['Header']])
    expect(removed).toBe(0)
  })

  it('does not mutate the input rows', () => {
    // The function must not mutate the caller's input. The caller
    // (useExcelRuntime) reads `range.getValues()` into a fresh array,
    // but the dedupe must still be pure so a future caller that passes
    // a shared matrix is safe.
    const input: ReadonlyArray<ReadonlyArray<string | number | null>> = [
      ['a', 1],
      ['a', 1],
    ]
    dedupeRows(input, false)
    expect(input).toEqual([
      ['a', 1],
      ['a', 1],
    ])
  })
})

// ──────────────────────────────────────────────────────────────────────────
// dedupeRowIndices — the canonical EXCEL-018 entry point.
//
// Returns the SET of duplicate row INDICES (0-based offsets relative
// to the input matrix). The runtime issues `ws.deleteRows(startRow +
// offset, 1)` per duplicate in DESCENDING offset order so earlier
// deletes don't shift later indices. This is the structural `remove-rows`
// path — formulas/styles/merges/etc. travel atomically with the row.
// ──────────────────────────────────────────────────────────────────────────

describe('dedupeRowIndices', () => {
  it('returns the indices of duplicate rows (0-based, ascending) and the count', () => {
    //   row 0: A, 1   ← first occurrence, kept
    //   row 1: B, 2   ← first occurrence, kept
    //   row 2: A, 1   ← duplicate of row 0 — DELETE
    //   row 3: A, 2   ← first occurrence, kept
    //   row 4: B, 2   ← duplicate of row 1 — DELETE
    const { keptIndices, duplicateIndices, removed } = dedupeRowIndices(
      [
        ['a', 1],
        ['b', 2],
        ['a', 1],
        ['a', 2],
        ['b', 2],
      ],
      false,
    )
    expect(keptIndices, 'kept indices (ascending)').toEqual([0, 1, 3])
    expect(duplicateIndices, 'duplicate indices (ascending)').toEqual([2, 4])
    expect(removed).toBe(2)
  })

  it('never marks the header row as a duplicate', () => {
    //   row 0: Name (header)  ← preserved verbatim AND never matched
    //   row 1: Name           ← NOT a duplicate of the header
    //   row 2: name           ← duplicate of row 1 (case-insensitive)
    const { keptIndices, duplicateIndices, removed } = dedupeRowIndices(
      [['Name'], ['Name'], ['name']],
      true,
    )
    expect(keptIndices).toEqual([0, 1])
    expect(duplicateIndices).toEqual([2])
    expect(removed).toBe(1)
  })

  it('treats the header as data when hasHeader is false', () => {
    // Same fixture as above, but hasHeader=false → row 0 IS in the
    // seen-set, so a later row matching it gets removed.
    const { keptIndices, duplicateIndices, removed } = dedupeRowIndices(
      [['Name'], ['Name'], ['name']],
      false,
    )
    expect(keptIndices).toEqual([0])
    expect(duplicateIndices).toEqual([1, 2])
    expect(removed).toBe(2)
  })

  it("handles the architect's mandatory regression case (Apple/Banana/Cherry)", () => {
    // The architect's mandatory EXCEL-018 regression fixture:
    //   row 0: Name, Qty        (header — kept)
    //   row 1: Apple, 10        (first — kept)
    //   row 2: Apple, 10        (DUP of row 1 — DELETE)
    //   row 3: Banana, 20       (first — kept)
    //   row 4: Apple, 10        (DUP of row 1 — DELETE)
    //   row 5: Cherry, 30       (first — kept)
    //   row 6: Apple, 30 (=B6)  (first — kept; result 30 differs from row 1's 10)
    //
    // The runtime will issue ws.deleteRows(startRow + 4, 1) then
    // ws.deleteRows(startRow + 2, 1) — descending order so earlier
    // deletes don't shift later indices. The survivor at row 6 (the
    // =B6 formula row) compacts to row 4 after the two deletes; the
    // gateway's transformFormulas rewrites =B6 → =B4 (Cherry/30
    // shifted from row 6 to row 4).
    const { keptIndices, duplicateIndices, removed } = dedupeRowIndices(
      [
        ['Name', 'Qty'],
        ['Apple', 10],
        ['Apple', 10],
        ['Banana', 20],
        ['Apple', 10],
        ['Cherry', 30],
        ['Apple', 30],
      ],
      true,
    )
    expect(keptIndices, 'header + 4 unique data rows kept').toEqual([0, 1, 3, 5, 6])
    expect(duplicateIndices, 'two duplicates at rows 2 and 4').toEqual([2, 4])
    expect(removed).toBe(2)
  })

  it('uses ALL selected columns as the comparison key', () => {
    // All rows are identical EXCEPT one selected comparison column:
    //   row 0: A, 1
    //   row 1: A, 1   ← duplicate of row 0 (both columns match)
    //   row 2: A, 2   ← NOT a duplicate (column 1 differs)
    //   row 3: A, 1   ← duplicate of row 0
    const { duplicateIndices, removed } = dedupeRowIndices(
      [
        ['A', 1],
        ['A', 1],
        ['A', 2],
        ['A', 1],
      ],
      false,
    )
    expect(duplicateIndices).toEqual([1, 3])
    expect(removed).toBe(2)
  })

  it('returns an empty duplicateIndices array when there are no duplicates (no-op)', () => {
    // The runtime MUST NOT fire any ws.deleteRows calls when the dedupe
    // finds no duplicates — this is the fail-closed no-op path that
    // surfaces the "No duplicate rows found" status message.
    const { keptIndices, duplicateIndices, removed } = dedupeRowIndices(
      [
        ['A', 1],
        ['B', 2],
        ['C', 3],
      ],
      false,
    )
    expect(keptIndices).toEqual([0, 1, 2])
    expect(duplicateIndices).toEqual([])
    expect(removed).toBe(0)
  })

  it('distinguishes blank cells from empty-string cells (null vs "")', () => {
    const { duplicateIndices, removed } = dedupeRowIndices(
      [
        [null],
        [''],
        [null], // duplicate of row 0
        [''], // duplicate of row 1
      ],
      false,
    )
    expect(duplicateIndices).toEqual([2, 3])
    expect(removed).toBe(2)
  })

  it('distinguishes number from text with the same printed form', () => {
    // 1 (number) and '1' (string) are NOT duplicates — type-strict.
    const { duplicateIndices, removed } = dedupeRowIndices([[1], ['1'], [1]], false)
    expect(duplicateIndices).toEqual([2])
    expect(removed).toBe(1)
  })

  it('distinguishes boolean from text with the same printed form', () => {
    const { duplicateIndices, removed } = dedupeRowIndices([[true], ['true'], [true]], false)
    expect(duplicateIndices).toEqual([2])
    expect(removed).toBe(1)
  })

  it('preserves a 1-row input with hasHeader=true (header-only, no data)', () => {
    const { keptIndices, duplicateIndices, removed } = dedupeRowIndices([['Header']], true)
    expect(keptIndices).toEqual([0])
    expect(duplicateIndices).toEqual([])
    expect(removed).toBe(0)
  })

  it('does not mutate the input rows', () => {
    const input: ReadonlyArray<ReadonlyArray<string | number | null>> = [
      ['a', 1],
      ['a', 1],
    ]
    dedupeRowIndices(input, false)
    expect(input).toEqual([
      ['a', 1],
      ['a', 1],
    ])
  })

  it('returned indices are stable for descending-order deletion', () => {
    // The contract: duplicateIndices is ASCENDING. The runtime iterates
    // the array in REVERSE (descending) so earlier deletes don't shift
    // later indices. This test verifies the ascending contract holds
    // for a fixture where the descending order is essential.
    const { duplicateIndices } = dedupeRowIndices(
      [
        ['A'],
        ['B'],
        ['A'], // dup of row 0
        ['C'],
        ['B'], // dup of row 1
        ['D'],
        ['A'], // dup of row 0
      ],
      false,
    )
    expect(duplicateIndices).toEqual([2, 4, 6])
    // Descending iteration: 6, 4, 2 — each delete doesn't shift the
    // remaining indices (the remaining dups are at lower row indices).
    const descending = [...duplicateIndices].sort((a, b) => b - a)
    expect(descending).toEqual([6, 4, 2])
  })

  it('dedupeRowIndices and dedupeRows agree on the kept rows (parity)', () => {
    // The legacy dedupeRows returns the kept-row MATRIX; the canonical
    // dedupeRowIndices returns the kept-row INDICES. The two must agree:
    //   dedupeRows(...).rows must equal dedupeRowIndices(...).keptIndices
    //   mapped through the input matrix.
    const input: ReadonlyArray<ReadonlyArray<string | number | null>> = [
      ['Name', 'Qty'],
      ['Apple', 10],
      ['Apple', 10],
      ['Banana', 20],
      ['Apple', 10],
      ['Cherry', 30],
      ['Apple', 30],
    ]
    const legacy = dedupeRows(input, true)
    const canonical = dedupeRowIndices(input, true)
    expect(legacy.rows).toEqual(canonical.keptIndices.map((i) => [...input[i]!]))
    expect(legacy.removed).toBe(canonical.removed)
  })
})
