/**
 * Unit tests for the pure dedupe algorithm (EXCEL-018).
 *
 * Mirrors the desktop reference at apps/sheets/tests/dedupe.test.ts
 * verbatim — the same fixtures and assertions — so the web parity is
 * provable from source: the same algorithm runs in both hosts.
 *
 * The dedupe function lives in apps/web/src/office/dedupe.ts and is a
 * PURE function — no Univer, no journal, no save plan. The
 * useExcelRuntime hook wires it through the live Univer facade (read
 * values, dedupe, write back per-row via FWorksheet.getRange(...).
 * setValues(...) which fires sheet.mutation.set-range-values — the
 * SAME canonical mutation channel Sort uses). The end-to-end save/
 * reopen path is exercised by the E2E suite at
 * apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts.
 */
import { describe, expect, it } from 'vitest'

import { dedupeRows } from '../src/office/dedupe'

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
      [
        [1, null],
        ['1', null],
        [null, null],
        [null, null],
        [true],
        ['true'],
      ],
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
