/**
 * Cell-mutation merge logic — the browser's formula-fidelity core.
 *
 * Univer emits `sheet.mutation.set-range-values` events carrying a sparse
 * `{ [row]: { [col]: ICellData } }` matrix. Each cell payload may carry:
 *   - `v`: the cell value (string | number | boolean | null)
 *   - `f`: the formula string (with or without a leading `=`), OR explicit
 *          `null` to clear a formula, OR absent
 *   - `s`: a style delta
 *   - `t`: the value type (recalc echo marker)
 *
 * Three distinct formula-related payload shapes were identified by the
 * Phase A forensic audit:
 *   ① formula edit:  `{ f: "=SUM(A1:A2)*2", v: null }`  (leading `=` present)
 *   ② recalc echo:   `{ v: 20, t: 2 }` — Univer's formula engine writes the
 *      recalculated cached value as a SEPARATE mutation, with NO `f` field.
 *      This must NOT overwrite a journaled formula: the engine drops the
 *      cached `<v>` on formula writes anyway and Excel recalculates on open.
 *   ③ formula clear: `{ f: null, v: 30 }` — an explicit `f: null` REPLACES the
 *      formula with the literal value.
 *
 * The merge folds these into one canonical `CellEdit` per cell so a value
 * edit and a style edit on the same cell compose into a single patch. The
 * formula-priority rule (Increment 12 hardening) guarantees mutation
 * ordering can never silently convert a formula back into a literal:
 *
 *   - An incoming formula edit always wins the cell state.
 *   - An incoming explicit formula clear / blank (replacesFormula) wins the
 *     cell state over a journaled formula — the user intentionally replaced
 *     the formula with a literal or emptied the cell.
 *   - The recalc echo (replacesFormula === false, no formula, value present)
 *     must NOT overwrite a journaled formula.
 *   - A value edit on a cell with NO journaled formula wins (plain literal).
 *   - A style-only edit never overwrites the cell state but merges style.
 *
 * The dirty map stores plain `CellEdit` values; `replacesFormula` is a
 * TRANSIENT merge-time signal on `ParsedMutation` only — it never reaches
 * the wire save plan (the route's `expectCellEdit` validates the unchanged
 * `CellEdit` shape: sheetName/row/column/writeValue/cell/style/rich/styleReset).
 */
import type { CellEdit, CellState } from '@genoffice/xlsx-gateway'

/**
 * A parsed Univer cell mutation: the canonical `CellEdit` it produces plus
 * the transient `replacesFormula` flag consumed by the merge rule.
 */
export interface ParsedMutation {
  readonly edit: CellEdit
  /**
   * True when this mutation explicitly replaces (clears) any journaled
   * formula — either an explicit `f: null` clear, a blank/clear that empties
   * the cell, or a fresh formula edit. The recalc echo (`f` absent + non-null
   * value) is FALSE so a journaled formula survives it.
   */
  readonly replacesFormula: boolean
}

/**
 * Extract a `ParsedMutation` from a Univer mutation cell payload.
 *
 * Returns `null` when the cell carries no recognizable value, formula, or
 * style (e.g. an `f: null` cleanup echo with no `si`, no `v`, and no `s`).
 */
export function cellEditFromMutation(
  sheetName: string,
  row: number,
  column: number,
  cell: unknown,
): ParsedMutation | null {
  if (cell === null || cell === undefined) {
    // Clearing the cell entirely → blank; replaces any journaled formula.
    return {
      edit: { sheetName, row, column, writeValue: true, cell: { value: null } },
      replacesFormula: true,
    }
  }
  if (typeof cell !== 'object') return null
  const data = cell as { v?: unknown; f?: unknown; s?: unknown }
  const style = styleDeltaFromUniver(data.s)
  const formulaRaw = data.f
  // Explicit `f: null` (the key present and null) is the formula-clear signal
  // — it distinguishes ③ (clear → literal replaces formula) from ② (recalc
  // echo, `f` absent). An empty-string `f` is not a clear (Univer treats it
  // as no formula); only an explicit null clears.
  const fExplicitlyNull = 'f' in data && data.f === null
  if (typeof formulaRaw === 'string' && formulaRaw.length > 0) {
    // ① formula edit — a fresh formula always wins the cell state.
    const formula = formulaRaw.startsWith('=') ? formulaRaw.slice(1) : formulaRaw
    return {
      edit: {
        sheetName,
        row,
        column,
        writeValue: true,
        cell: { value: '', formula },
        ...(style ? { style } : {}),
      },
      replacesFormula: true,
    }
  }
  const v = data.v
  if (v === undefined) {
    // No formula and no value in the payload: a style-only mutation, or an
    // `f: null` cleanup echo with no editable state. A cleanup echo with no
    // style carries no editable state — ignore it.
    if (!style) return null
    return {
      edit: { sheetName, row, column, writeValue: false, cell: { value: null }, style },
      replacesFormula: false,
    }
  }
  if (v === null || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
    // ③ explicit formula clear (`f: null` + `v: 30`) → literal replaces
    //    formula (replacesFormula = true).
    // Blank (`v: null`, `f` absent or null) → clears formula too.
    // ② recalc echo (`f` absent + non-null `v`) → does NOT replace a journaled
    //    formula (replacesFormula = false); the merge rule preserves it.
    const replacesFormula = fExplicitlyNull || v === null
    const cellState: CellState = { value: v }
    return {
      edit: {
        sheetName,
        row,
        column,
        writeValue: true,
        cell: cellState,
        ...(style ? { style } : {}),
      },
      replacesFormula,
    }
  }
  return null
}

/**
 * Merge an incoming parsed mutation into the existing journaled `CellEdit`
 * (if any) for the same cell. The result is a plain `CellEdit` for the dirty
 * map — `replacesFormula` is consumed here and never persisted.
 *
 * Merge rules (formula-priority, Increment 12 hardening):
 *   1. An incoming formula edit always wins the cell state (formula + value).
 *   2. An incoming explicit formula clear / blank (replacesFormula) wins the
 *      cell state over a journaled formula.
 *   3. A value edit on a cell with NO journaled formula wins (plain literal).
 *   4. The recalc echo (no formula, value present, not a clear) never
 *      overwrites a journaled formula.
 *   5. A style-only edit never overwrites the cell state but always merges
 *      its style fields (later deltas win per field).
 */
export function mergeCellEdit(existing: CellEdit | undefined, incoming: ParsedMutation): CellEdit {
  const edit = incoming.edit
  if (!existing) return edit
  const existingHasFormula = !!existing.cell.formula
  const incomingHasFormula = !!edit.cell.formula
  const incomingIsValueEdit = edit.writeValue
  const cellWins =
    incomingHasFormula || incoming.replacesFormula || (!existingHasFormula && incomingIsValueEdit)
  const mergedStyle =
    edit.style || existing.style ? { ...(existing.style ?? {}), ...(edit.style ?? {}) } : undefined
  return {
    sheetName: edit.sheetName,
    row: edit.row,
    column: edit.column,
    writeValue: cellWins ? edit.writeValue : existing.writeValue,
    cell: cellWins ? edit.cell : existing.cell,
    ...(mergedStyle ? { style: mergedStyle } : {}),
  }
}

/**
 * Map a Univer `IStyleData` (the `s` payload of a set-range-values mutation)
 * to a canonical `WorkbookStyleEdit` delta. Only PRESENT keys map — Univer's
 * formatting commands send partial deltas (`{ bl: 1 }` for bold-on,
 * `{ bl: 0 }` for bold-off), and the engine applies deltas on top of each
 * cell's current `cellXfs` entry, so absent keys leave the file's own
 * properties alone.
 *
 * Colors convert to the `#RRGGBB` convention of `WorkbookStyleEdit`.
 */
export function styleDeltaFromUniver(s: unknown): CellEdit['style'] | undefined {
  if (typeof s !== 'object' || s === null) return undefined
  const d = s as Record<string, unknown>
  const out: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    underlineStyle?: 'single' | 'double'
    strikethrough?: boolean
    fontFamily?: string
    fontSize?: number
    fontColor?: string | null
    fillColor?: string | null
    horizontalAlignment?: 'left' | 'center' | 'right'
    verticalAlignment?: 'top' | 'center' | 'bottom'
    wrapText?: boolean
  } = {}
  if ('bl' in d) out.bold = d.bl === 1
  if ('it' in d) out.italic = d.it === 1
  if ('ul' in d) {
    const ul = d.ul
    const on = typeof ul === 'object' && ul !== null ? (ul as { s?: unknown }).s === 1 : ul === 1
    out.underline = on
    if (on) out.underlineStyle = 'single'
  }
  if ('st' in d) {
    const st = d.st
    out.strikethrough =
      typeof st === 'object' && st !== null ? (st as { s?: unknown }).s === 1 : st === 1
  }
  if (typeof d.ff === 'string' && d.ff !== '') out.fontFamily = d.ff
  if (typeof d.fs === 'number' && Number.isFinite(d.fs) && d.fs > 0) out.fontSize = d.fs
  if ('cl' in d) {
    const rgb = univerColorToHex((d.cl as { rgb?: unknown } | null)?.rgb)
    out.fontColor = rgb === undefined && d.cl === null ? null : rgb
  }
  if ('bg' in d) {
    const rgb = univerColorToHex((d.bg as { rgb?: unknown } | null)?.rgb)
    out.fillColor = rgb === undefined && d.bg === null ? null : rgb
  }
  if (typeof d.ht === 'number') {
    if (d.ht === 1) out.horizontalAlignment = 'left'
    else if (d.ht === 2) out.horizontalAlignment = 'center'
    else if (d.ht === 3) out.horizontalAlignment = 'right'
  }
  if (typeof d.vt === 'number') {
    if (d.vt === 1) out.verticalAlignment = 'top'
    else if (d.vt === 2) out.verticalAlignment = 'center'
    else if (d.vt === 3) out.verticalAlignment = 'bottom'
  }
  if ('tb' in d) {
    // WrapStrategy enum (Univer): UNSPECIFIED=0, OVERFLOW=1, CLIP=2, WRAP=3.
    // wrapText=true only when the strategy is WRAP (3). The previous
    // implementation checked `d.tb === 1` (OVERFLOW) which was inverted —
    // a wrap-ON toggle produced wrapText:false in the save plan, and a
    // wrap-OFF produced wrapText:true.
    out.wrapText = d.tb === 3
  }
  return Object.keys(out).length > 0 ? out : undefined
}

/** Univer color ('#RRGGBB') → WorkbookStyleEdit '#RRGGBB'; non-hex → undefined. */
export function univerColorToHex(rgb: unknown): string | undefined {
  if (typeof rgb !== 'string' || rgb === '') return undefined
  if (/^#[0-9A-Fa-f]{6}$/.test(rgb)) return rgb.toUpperCase()
  if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return `#${rgb.toUpperCase()}`
  return undefined
}

// ── Number-format journaling (Phase 4 Increment 3 — Objective 1) ───────────
//
// Univer's numfmt facade mixin (sheets-numfmt preset) fires the
// `sheet.mutation.set.numfmt` mutation, whose params carry:
//   {
//     values: { [id]: { ranges: IRange[] } },
//     refMap: { [id]: { pattern: string } },
//     unitId, subUnitId,
//   }
// The browser's existing `set-range-values` journal subscription does NOT
// capture this mutation (it filters by mutation ID), so without this
// helper a number-format change is in-session only — it appears on the
// grid but does NOT survive save/reopen. ExcelEditor's expanded
// subscription now also handles `sheet.mutation.set.numfmt`: for each
// (pattern, range) pair it emits one style-only CellEdit per cell with
// `style.numberFormat = pattern`, which the existing
// `WorkbookStyleEdit.numberFormat` field persists through the canonical
// `applyCellEditsToXlsx` save path (xlsx-styles.ts writes numFmtId on
// the cellXfs entry). The canonical `WorkbookStyleEdit` already exposes
// `numberFormat` — the wire is unchanged.

/**
 * Build style-only CellEdits for a numfmt mutation. Each cell in every
 * range receives a CellEdit with `writeValue: false` (the cell's stored
 * content stays untouched) and `style: { numberFormat: pattern }`.
 *
 * Returns an empty array when the params are malformed (the journal
 * ignores the mutation rather than crashing).
 */
export function numfmtEditsFromMutation(
  sheetName: string,
  params: unknown,
): ReadonlyArray<{
  readonly row: number
  readonly column: number
  readonly edit: CellEdit
}> {
  if (typeof params !== 'object' || params === null) return []
  const p = params as {
    values?: unknown
    refMap?: unknown
  }
  if (typeof p.values !== 'object' || p.values === null || typeof p.refMap !== 'object' || p.refMap === null) {
    return []
  }
  const values = p.values as Record<string, { ranges?: unknown }>
  const refMap = p.refMap as Record<string, { pattern?: unknown }>
  const out: Array<{ row: number; column: number; edit: CellEdit }> = []
  for (const [id, entry] of Object.entries(values)) {
    if (!entry || typeof entry !== 'object') continue
    const ranges = (entry as { ranges?: unknown }).ranges
    if (!Array.isArray(ranges)) continue
    const ref = refMap[id]
    if (!ref || typeof ref.pattern !== 'string') continue
    const pattern = ref.pattern
    for (const range of ranges) {
      if (typeof range !== 'object' || range === null) continue
      const r = range as { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number }
      const startRow = Number.isInteger(r.startRow) ? (r.startRow as number) : -1
      const endRow = Number.isInteger(r.endRow) ? (r.endRow as number) : -1
      const startColumn = Number.isInteger(r.startColumn) ? (r.startColumn as number) : -1
      const endColumn = Number.isInteger(r.endColumn) ? (r.endColumn as number) : -1
      if (startRow < 0 || endRow < 0 || startColumn < 0 || endColumn < 0) continue
      for (let row = startRow; row <= endRow; row++) {
        for (let column = startColumn; column <= endColumn; column++) {
          out.push({
            row,
            column,
            edit: {
              sheetName,
              row,
              column,
              writeValue: false,
              cell: { value: null },
              style: { numberFormat: pattern },
            },
          })
        }
      }
    }
  }
  return out
}

// ── Sort journaling (Phase 4 Increment 3 — Objective 3) ────────────────────
//
// Univer's sort command (sheet.command.sort-range) fires the
// `sheet.mutation.reorder-range` mutation, which writes directly into the
// worksheet's cellDataMatrix in-memory — it does NOT dispatch a separate
// `sheet.mutation.set-range-values`. The existing journal's
// set-range-values subscription therefore misses sort. ExcelEditor's
// expanded subscription now also handles `sheet.mutation.reorder-range`:
// after the mutation lands, the post-sort cell values are read straight
// from the worksheet model and journaled as plain value CellEdits. On
// save, applyCellEditsToXlsx writes them back into the XLSX in the
// sorted order — the canonical write path, no bespoke sort mutation
// family on the wire.

/**
 * Build value CellEdits for a reorder-range (sort) mutation. Reads the
 * post-mutation cell values directly from the worksheet model and emits
 * one `writeValue: true` CellEdit per cell in the sorted range, so the
 * save plan writes the new row order into the XLSX.
 *
 * `readCell` is a thin indirection so the pure merge module never imports
 * Univer — the caller (ExcelEditor) supplies a closure over the live
 * worksheet. Returns an empty array when the params are malformed.
 */
export function reorderEditsFromMutation(
  sheetName: string,
  params: unknown,
  readCell: (row: number, column: number) =>
    | { value: string | number | boolean | null; formula?: string }
    | null,
): ReadonlyArray<{
  readonly row: number
  readonly column: number
  readonly edit: CellEdit
}> {
  if (typeof params !== 'object' || params === null) return []
  const p = params as { range?: unknown }
  if (typeof p.range !== 'object' || p.range === null) return []
  const r = p.range as { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number }
  const startRow = Number.isInteger(r.startRow) ? (r.startRow as number) : -1
  const endRow = Number.isInteger(r.endRow) ? (r.endRow as number) : -1
  const startColumn = Number.isInteger(r.startColumn) ? (r.startColumn as number) : -1
  const endColumn = Number.isInteger(r.endColumn) ? (r.endColumn as number) : -1
  if (startRow < 0 || endRow < 0 || startColumn < 0 || endColumn < 0) return []
  const out: Array<{ row: number; column: number; edit: CellEdit }> = []
  for (let row = startRow; row <= endRow; row++) {
    for (let column = startColumn; column <= endColumn; column++) {
      const cell = readCell(row, column)
      if (cell === null) continue
      // Preserve any formula on the post-sort cell — sort moves the whole
      // cell (value + formula). A formula cell keeps its formula; a plain
      // literal keeps its literal. An empty cell clears the destination.
      out.push({
        row,
        column,
        edit: {
          sheetName,
          row,
          column,
          writeValue: true,
          cell: cell.formula
            ? { value: cell.value, formula: cell.formula }
            : { value: cell.value },
        },
      })
    }
  }
  return out
}
