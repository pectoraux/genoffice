/**
 * Browser-local A1 address helpers.
 *
 * Mirrors the canonical `parseAddress` / `parseRange` in
 * `@genoffice/xlsx-gateway`, but kept here in apps/web so the browser
 * bundle never pulls in the engine's runtime — only the type-only
 * imports from `@genoffice/xlsx-gateway` reach the browser bundle, and
 * those are erased at compile time.
 *
 * The engine remains the source of truth: the wire shape (WorkbookSnapshot
 * addresses) is what we parse here. If the canonical parser changes,
 * these helpers must change to match — that's an acceptable cost given
 * how small the A1 grammar is.
 */

export interface CellCoordinates {
  readonly row: number
  readonly column: number
}

export interface RangeBounds {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

/** Parse "A1" → { row: 0, column: 0 }. Throws on invalid input. */
export function parseAddress(address: string): CellCoordinates {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(address)
  if (!match?.[1] || !match[2]) throw new Error(`Invalid cell address: ${address}`)
  let column = 0
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return { row: Number(match[2]) - 1, column: column - 1 }
}

/** Parse "A1:C10" (or "B2") → inclusive row/column bounds. */
export function parseRange(range: string): RangeBounds {
  const parts = range.split(':')
  if (parts.length > 2 || !parts[0]) throw new Error(`Invalid range: ${range}`)
  const first = parseAddress(parts[0])
  const second = parts[1] ? parseAddress(parts[1]) : first
  return {
    startRow: Math.min(first.row, second.row),
    startColumn: Math.min(first.column, second.column),
    endRow: Math.max(first.row, second.row),
    endColumn: Math.max(first.column, second.column),
  }
}

/** Column label "A" → 0, "Z" → 25, "AA" → 26. */
export function columnIndex(label: string): number {
  if (!/^[A-Z]+$/.test(label)) throw new Error(`Invalid column label: ${label}`)
  let column = 0
  for (const character of label) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}
