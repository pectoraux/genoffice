/**
 * PROJECT-022 — row virtualization for the synchronized Gantt surface.
 *
 * The ONE shared row window both panes (task grid and timeline) render from:
 * the same visible-row slice of the projection, computed once by
 * `buildGanttView` and consumed by every view builder — grid rows, timeline
 * rows, bars, milestones, and dependency links all address rows by their
 * ABSOLUTE index into the projection's visible-row sequence, so the two
 * panes cannot drift apart (the "synchronized" acceptance contract).
 *
 * Scroll position is HOST state, not renderer view state: hosts convert
 * pixel scroll offsets into the logical `firstRow` (and know their visible
 * row capacity in rows) and pass both as plain arguments. The core never
 * sees a pixel (the PROJECT-021 fraction-space discipline); nothing here is
 * persisted in `ProjectViewState`.
 *
 * Pure integer arithmetic, fully deterministic.
 */

/** An inclusive [firstIndex, lastIndex] window over the visible rows.
 * `lastIndex < firstIndex` is the canonical EMPTY window (rowCount 0 or
 * visibleRows 0). */
export interface ProjectRowWindow {
  readonly firstIndex: number
  readonly lastIndex: number
}

/** The host-provided layout inputs (logical rows, never pixels). */
export interface RowWindowInput {
  /** The first row the host is scrolled to (0-based). Clamped to the row
   * range; negative values behave like 0. */
  readonly firstRow: number
  /** How many rows the host can show. Clamped to [0, rowCount]. */
  readonly visibleRows: number
  /** Extra rows rendered on each side of the window (virtualization
   * overscan, default 0). Clamped to >= 0. */
  readonly overscan?: number
}

/** Builds the clamped inclusive row window for a visible-row count. */
export function buildRowWindow(rowCount: number, input: RowWindowInput): ProjectRowWindow {
  if (
    !Number.isFinite(rowCount) ||
    rowCount <= 0 ||
    !Number.isFinite(input.firstRow) ||
    !Number.isFinite(input.visibleRows) ||
    input.visibleRows <= 0
  ) {
    return { firstIndex: 0, lastIndex: -1 }
  }
  const overscan = Number.isFinite(input.overscan)
    ? Math.max(0, Math.floor(input.overscan ?? 0))
    : 0
  const firstRow = clampInt(input.firstRow, 0, rowCount - 1)
  const visibleRows = clampInt(input.visibleRows, 0, rowCount)
  const firstIndex = Math.max(0, firstRow - overscan)
  const lastIndex = Math.min(rowCount - 1, firstRow + visibleRows - 1 + overscan)
  return { firstIndex, lastIndex }
}

/** True when the window covers no rows at all. */
export function rowWindowIsEmpty(window: ProjectRowWindow): boolean {
  return window.lastIndex < window.firstIndex
}

const clampInt = (value: number, lo: number, hi: number): number =>
  Math.max(lo, Math.min(hi, Math.floor(value)))
