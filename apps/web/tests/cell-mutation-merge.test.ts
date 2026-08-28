/**
 * Unit tests — cell-mutation merge formula ordering (Increment 12 hardening).
 *
 * These tests prove that mutation ordering can never silently convert a
 * formula back into a literal. They exercise the PURE merge logic
 * (cellEditFromMutation + mergeCellEdit) that the Excel editor folds Univer's
 * set-range-values mutations through before sending the save plan.
 *
 * The four formula-related payload shapes (Phase A forensic audit):
 *   ① formula edit:        { f: "=SUM(A1:A2)*2", v: null }
 *   ② recalc echo:         { v: 20, t: 2 }           (f ABSENT)
 *   ③ explicit formula clear: { f: null, v: 30 }      (f EXPLICITLY null)
 *   ④ style-only:          { s: { bl: 1 } }
 *
 * The six ordering sequences required by Increment 12 §4 are exercised
 * verbatim. The engine-level cached-value semantics (unchanged formula byte
 * preservation, edited formula drops cached <v>, formula→literal no <f>,
 * formula→blank no <f>/<v>) are covered by office-excel-formula.test.ts;
 * these tests assert the canonical CellEdit the merge produces — the exact
 * payload the engine then turns into XLSX XML.
 */
import { describe, expect, it } from 'vitest'
import {
  cellEditFromMutation,
  mergeCellEdit,
  ooxmlTextRotationToUniver,
} from '../src/office/cell-mutation-merge'
import type { CellEdit } from '@genoffice/xlsx-gateway'

/** Fold a sequence of raw Univer cell payloads for one cell into the final
 *  canonical CellEdit the dirty map would hold (and the save plan send). */
function runSequence(payloads: unknown[]): CellEdit | undefined {
  let current: CellEdit | undefined
  for (const cell of payloads) {
    const parsed = cellEditFromMutation('Sheet1', 2, 0, cell)
    if (!parsed) continue
    current = mergeCellEdit(current, parsed)
  }
  return current
}

describe('Cell-mutation merge — formula mutation ordering', () => {
  it('① formula → ② recalc echo → ④ style: formula + style survive', () => {
    const final = runSequence([
      { f: '=SUM(A1:A2)*2', v: null }, // ① formula edit
      { v: 20, t: 2 }, // ② recalc echo (f ABSENT) — must NOT overwrite formula
      { s: { bl: 1 } }, // ④ style-only — must merge without destroying formula
    ])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    expect(final!.cell.formula).toBe('SUM(A1:A2)*2')
    expect(final!.cell.value).toBe('')
    expect(final!.style?.bold).toBe(true)
    // No literal value leaked in from the recalc echo.
    expect(final!.cell.value).not.toBe(20)
  })

  it('① formula → ④ style → ② recalc echo: formula + style survive', () => {
    const final = runSequence([
      { f: '=SUM(A1:A2)*2', v: null }, // ①
      { s: { bl: 1 } }, // ④ style first
      { v: 20, t: 2 }, // ② recalc echo after — must NOT overwrite formula
    ])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    expect(final!.cell.formula).toBe('SUM(A1:A2)*2')
    expect(final!.cell.value).toBe('')
    expect(final!.style?.bold).toBe(true)
    expect(final!.cell.value).not.toBe(20)
  })

  it('① formula → ③ literal → ④ style: literal replaces formula; style merges', () => {
    const final = runSequence([
      { f: '=SUM(A1:A2)*2', v: null }, // ① formula edit
      { f: null, v: 30 }, // ③ explicit formula clear → literal REPLACES formula
      { s: { bl: 1 } }, // ④ style merges onto the literal
    ])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    // The formula is GONE — the literal 30 replaced it.
    expect(final!.cell.formula).toBeUndefined()
    expect(final!.cell.value).toBe(30)
    expect(final!.style?.bold).toBe(true)
  })

  it('① formula → blank: both formula and value removed', () => {
    const final = runSequence([
      { f: '=SUM(A1:A2)*2', v: null }, // ① formula edit
      { v: null, f: null }, // explicit blank clear
    ])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    expect(final!.cell.formula).toBeUndefined()
    expect(final!.cell.value).toBeNull()
  })

  it('literal → ④ style → ② recalc echo: value updates; style preserved; no formula', () => {
    const final = runSequence([
      { v: 30 }, // plain literal on a non-formula cell (f absent + non-null v)
      { s: { bl: 1 } }, // ④ style
      { v: 20, t: 2 }, // ② echo-shaped value update — updates literal, keeps style
    ])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    // No formula ever appeared (an echo can never synthesize one).
    expect(final!.cell.formula).toBeUndefined()
    expect(final!.cell.value).toBe(20)
    expect(final!.style?.bold).toBe(true)
  })

  it('④ style → ① formula → ② recalc echo: formula + style survive', () => {
    const final = runSequence([
      { s: { bl: 1 } }, // ④ style-only first (no value yet)
      { f: '=SUM(A1:A2)', v: null }, // ① formula edit wins cell state
      { v: 20, t: 2 }, // ② recalc echo — must NOT overwrite formula
    ])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    expect(final!.cell.formula).toBe('SUM(A1:A2)')
    expect(final!.cell.value).toBe('')
    expect(final!.style?.bold).toBe(true)
    expect(final!.cell.value).not.toBe(20)
  })
})

describe('Cell-mutation merge — blank shapes all clear a journaled formula', () => {
  // The Delete key produces several indistinguishable blank shapes; each
  // must clear a journaled formula (the latent Increment 12 bug — the
  // pre-hardening merge dropped a blank in favor of a journaled formula).
  const blankShapes: Array<{ name: string; payload: unknown }> = [
    { name: 'null cell', payload: null },
    { name: 'undefined cell', payload: undefined },
    { name: '{ v: null }', payload: { v: null } },
    { name: '{ v: null, f: null }', payload: { v: null, f: null } },
  ]
  for (const { name, payload } of blankShapes) {
    it(`① formula → blank (${name}): formula + value removed`, () => {
      const final = runSequence([
        { f: '=SUM(A1:A2)*2', v: null }, // ① formula journaled
        payload, // blank — must clear the formula
      ])
      expect(final).toBeDefined()
      expect(final!.writeValue).toBe(true)
      expect(final!.cell.formula).toBeUndefined()
      expect(final!.cell.value).toBeNull()
    })
  }
})

describe('Cell-mutation merge — recalc echo never overwrites a journaled formula', () => {
  it('a bare recalc echo with no prior formula becomes the literal value', () => {
    // On a cell with NO journaled formula, the echo-shaped payload {v, t} is
    // just a value edit — it wins (there is no formula to protect).
    const final = runSequence([{ v: 42, t: 2 }])
    expect(final).toBeDefined()
    expect(final!.writeValue).toBe(true)
    expect(final!.cell.formula).toBeUndefined()
    expect(final!.cell.value).toBe(42)
  })

  it('a recalc echo repeated never replaces a journaled formula', () => {
    const final = runSequence([
      { f: '=A1+B1', v: null },
      { v: 7, t: 2 },
      { v: 8, t: 2 },
      { v: 9, t: 2 },
    ])
    expect(final).toBeDefined()
    expect(final!.cell.formula).toBe('A1+B1')
    expect(final!.cell.value).toBe('')
  })

  it('style-only mutations accumulate without touching value/formula', () => {
    const final = runSequence([
      { f: '=A1+B1', v: null },
      { s: { bl: 1 } }, // bold
      { s: { it: 1 } }, // italic
      { s: { bg: { rgb: '#FFF2CC' } } }, // fill
    ])
    expect(final).toBeDefined()
    expect(final!.cell.formula).toBe('A1+B1')
    expect(final!.cell.value).toBe('')
    expect(final!.style?.bold).toBe(true)
    expect(final!.style?.italic).toBe(true)
    expect(final!.style?.fillColor).toBe('#FFF2CC')
  })
})

describe('Cell-mutation merge — explicit formula clear (③) over a journaled literal', () => {
  it('literal → ③ explicit clear: clear replaces literal too', () => {
    // A journaled literal (no formula) cleared via {f:null, v:null} → blank.
    const final = runSequence([
      { v: 30 }, // literal
      { f: null, v: null }, // explicit clear
    ])
    expect(final).toBeDefined()
    expect(final!.cell.formula).toBeUndefined()
    expect(final!.cell.value).toBeNull()
  })

  it('f:null cleanup echo (no v, no style) is ignored', () => {
    // Univer's f:null cleanup echo (si:null, no v, no s) carries no editable
    // state — it must be ignored so it cannot spuriously clear a formula.
    const final = runSequence([
      { f: '=A1+B1', v: null },
      { f: null }, // cleanup echo — ignored
    ])
    expect(final).toBeDefined()
    expect(final!.cell.formula).toBe('A1+B1')
    expect(final!.cell.value).toBe('')
  })
})

describe('Cell-mutation merge — formula-bar display contract', () => {
  // The formula bar canvas reads the cell's canonical `f` field. These tests
  // prove that field is correct after each merge (the canvas paints exactly
  // this). The browser cannot assert DOM text (Univer renders on canvas) —
  // the exposed runtime's snapshot mirrors this `f` (see excel-formula.spec.ts
  // `formulaBarDisplay`); the real editing flow is exercised there.
  it('a freshly typed formula surfaces a canonical `f` (leading = stripped)', () => {
    const final = runSequence([{ f: '=SUM(A1:A2)*2', v: null }])
    expect(final!.cell.formula).toBe('SUM(A1:A2)*2')
  })

  it('a formula without a leading = is accepted verbatim', () => {
    const final = runSequence([{ f: 'SUM(A1:A2)*2', v: null }])
    expect(final!.cell.formula).toBe('SUM(A1:A2)*2')
  })
})

describe('styleDeltaFromUniver — EXCEL-027 advanced formatting journal mapping', () => {
  it('maps a border mutation with all four edges to canonical border deltas', () => {
    const parsed = cellEditFromMutation('Sheet1', 0, 0, {
      s: {
        bd: {
          t: { s: 1, cl: { rgb: '#FF0000' } },
          b: { s: 8 },
          l: { s: 7, cl: { rgb: '#00B050' } },
          r: { s: 13 },
        },
      },
    })
    expect(parsed?.edit.style).toEqual({
      borderTop: { style: 'thin', color: '#FF0000' },
      borderBottom: { style: 'medium' },
      borderLeft: { style: 'double', color: '#00B050' },
      borderRight: { style: 'thick' },
    })
    expect(parsed?.edit.writeValue).toBe(false)
  })

  it('maps the engine clearBorder emission (all-null edges) to null clears', () => {
    const parsed = cellEditFromMutation('Sheet1', 0, 0, {
      s: { bd: { t: null, b: null, l: null, r: null } },
    })
    expect(parsed?.edit.style).toEqual({
      borderTop: null,
      borderBottom: null,
      borderLeft: null,
      borderRight: null,
    })
  })

  it('preserves untouched edges — a one-edge mutation journals only that edge', () => {
    const parsed = cellEditFromMutation('Sheet1', 0, 0, {
      s: { bd: { l: { s: 4, cl: { rgb: '#123456' } } } },
    })
    expect(parsed?.edit.style).toEqual({ borderLeft: { style: 'dashed', color: '#123456' } })
  })

  it('ignores diagonal border keys and unknown style numbers (fail closed)', () => {
    const parsed = cellEditFromMutation('Sheet1', 0, 0, {
      s: { bd: { tl_br: { s: 1 }, bl_tr: { s: 1 }, t: { s: 99 } } },
    })
    // The unknown style 99 is not one of the 13 canonical values — nothing
    // is journaled and the file's own border XML is preserved.
    expect(parsed).toBeNull()
  })

  it('maps text rotation in all four engine forms', () => {
    // Positive angle (counterclockwise) → OOXML 1..90.
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { tr: { a: 45 } } })?.edit.style).toEqual({
      textRotation: 45,
    })
    // Negative angle (clockwise) → OOXML 91..180 (90 - a).
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { tr: { a: -45 } } })?.edit.style).toEqual({
      textRotation: 135,
    })
    // Stacked text → OOXML 255.
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { tr: { a: 0, v: 1 } } })?.edit.style).toEqual(
      { textRotation: 255 },
    )
    // Clear → the writer's 0 sentinel.
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { tr: null } })?.edit.style).toEqual({
      textRotation: 0,
    })
    // A 0-degree angle is "no rotation" — the sentinel, not an angle.
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { tr: { a: 0 } } })?.edit.style).toEqual({
      textRotation: 0,
    })
  })

  it('maps indent padding to OOXML steps and clears with null', () => {
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { pd: { l: 16 } } })?.edit.style).toEqual({
      indent: 2,
    })
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { pd: { l: 0 } } })?.edit.style).toEqual({
      indent: 0,
    })
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { pd: null } })?.edit.style).toEqual({
      indent: 0,
    })
    // Non-integer padding rounds to the nearest step; clamped to the wire
    // bound 0..250.
    expect(cellEditFromMutation('Sheet1', 0, 0, { s: { pd: { l: 4000 } } })?.edit.style).toEqual({
      indent: 250,
    })
  })

  it('a combined border + rotation + bold mutation journals every family at once', () => {
    const parsed = cellEditFromMutation('Sheet1', 0, 0, {
      s: { bl: 1, tr: { a: 90 }, bd: { t: { s: 3, cl: { rgb: '#ABCDEF' } } } },
    })
    expect(parsed?.edit.style).toEqual({
      bold: true,
      textRotation: 90,
      borderTop: { style: 'dotted', color: '#ABCDEF' },
    })
  })

  it('style-only advanced-format mutations merge with a prior value edit', () => {
    const final = runSequence([{ v: 42 }, { s: { bd: { t: { s: 1 } }, tr: { a: 45 } } }])
    // runSequence journals at its fixed cell (row 2, column 0).
    expect(final).toEqual({
      sheetName: 'Sheet1',
      row: 2,
      column: 0,
      writeValue: true,
      cell: { value: 42 },
      style: { borderTop: { style: 'thin' }, textRotation: 45 },
    })
  })
})

describe('ooxmlTextRotationToUniver — EXCEL-027 import conversion', () => {
  it('converts every OOXML rotation form to the engine value', () => {
    expect(ooxmlTextRotationToUniver(45)).toEqual({ a: 45 })
    expect(ooxmlTextRotationToUniver(135)).toEqual({ a: -45 })
    expect(ooxmlTextRotationToUniver(255)).toEqual({ a: 0, v: 1 })
    expect(ooxmlTextRotationToUniver(90)).toEqual({ a: 90 })
    expect(ooxmlTextRotationToUniver(0)).toBeNull()
    // 181..254 is outside the OOXML domain — no engine value.
    expect(ooxmlTextRotationToUniver(200)).toBeNull()
  })
})
