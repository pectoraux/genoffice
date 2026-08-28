import type {
  BorderPatch,
  CellFormatPatch,
  LayoutOperation,
  StructuralOperation,
} from './workbook-dsl'
import type { SheetVisual } from './chart-visual'
import type { SheetFilterState } from '../gateway/xlsx-filter'
import type { DvWireRule } from '../gateway/xlsx-dv'
import type { CfWireRule } from '../gateway/xlsx-cf'
import type { SheetNote } from '../gateway/xlsx-notes'
import type { SheetTableInfo } from '../gateway/xlsx-table-read'
import type { SheetChartInfo } from '../gateway/xlsx-chart-read'
import type { SheetImageInfo } from '../gateway/xlsx-image-read'
import type { DefinedNameEntry } from '../gateway/xlsx-defined-names'

export type CellScalar = string | number | boolean | null

export interface CellState {
  readonly value: CellScalar
  readonly formula?: string | undefined
}

/** resolved per-cell formatting; unlike CellFormatPatch, never holds nulls */
export interface CellFormatState {
  readonly bold?: boolean | undefined
  readonly italic?: boolean | undefined
  readonly underline?: boolean | undefined
  readonly strikethrough?: boolean | undefined
  readonly fontFamily?: string | undefined
  readonly fontSize?: number | undefined
  readonly fontColor?: string | undefined
  readonly fillColor?: string | undefined
  readonly numberFormat?: string | undefined
  readonly horizontalAlign?: 'left' | 'center' | 'right' | undefined
  readonly verticalAlign?: 'top' | 'center' | 'bottom' | undefined
  readonly wrapText?: boolean | undefined
  readonly textRotation?: number | 'vertical' | undefined
  readonly indent?: number | undefined
  readonly border?: BorderPatch | undefined
}

export interface WorksheetState {
  readonly id: string
  readonly name: string
  readonly cells: Readonly<Record<string, CellState>>
  /** demo-mode formatting, keyed by address; absent means "no explicit format" */
  readonly styles?: Readonly<Record<string, CellFormatState>> | undefined
  /** demo-mode merged ranges ("B2:D2"), replayed on rebuild */
  readonly merges?: readonly string[] | undefined
  /** demo-mode row heights in points, keyed by 1-based row number */
  readonly rowHeights?: Readonly<Record<string, number>> | undefined
  /** demo-mode column widths in px, keyed by column label */
  readonly colWidths?: Readonly<Record<string, number>> | undefined
  /** demo-mode charts added by AI add_chart, replayed on rebuild */
  readonly visuals?: readonly SheetVisual[] | undefined
  /**
   * Frozen-pane state parsed from the worksheet's <sheetView><pane>.
   * Absent (undefined) means "no pane / frozen rows and columns are 0";
   * present with `frozenRows` and `frozenColumns` both 0 also means no
   * freeze. The web Sheets shell (View → Freeze Panes) journals freeze
   * changes through the canonical `SheetPageSetupState` save family, and
   * reads it back via this field on reopen so freeze survives round-trip.
   */
  readonly freeze?:
    Readonly<{ readonly frozenRows: number; readonly frozenColumns: number }> | undefined
  /**
   * View-state flags parsed from the worksheet's first <sheetView> (EXCEL-026).
   * Each is exposed ONLY when the file explicitly sets the NON-DEFAULT value —
   * absent means the schema default the canonical writer also restores by
   * dropping the attribute (showGridLines / showRowColHeaders default true,
   * showFormulas defaults false). Malformed boolean values are ignored for
   * modeling (the raw attribute stays byte-preserved on no-op saves because
   * such a sheet never enters the pageSetupStates save family; a user toggle
   * overwrites it with a definite canonical value). The web shell (View →
   * Show) journals changes through the canonical `SheetPageSetupState` save
   * family and reads them back via these fields on reopen.
   */
  readonly view?:
    | Readonly<{
        /** sheetView@showGridLines="0" — gridlines hidden. */
        readonly showGridlines: false
        /** sheetView@showFormulas="1" — the sheet renders formulas, not values. */
        readonly showFormulas: true
        /** sheetView@showRowColHeaders="0" — row/column heading strips hidden. */
        readonly showHeadings: false
      }>
    | undefined
  /**
   * AutoFilter state parsed from the worksheet's <autoFilter> element plus
   * the row-visibility it implies. Absent means no parseable filter —
   * including when the element carries criteria the canonical model cannot
   * represent (top10, dynamicFilter, iconFilter, dateGroup, colorFilters),
   * which fail closed: the browser never renders such a filter and a no-op
   * save preserves the file's XML byte-for-byte. The web Sheets shell
   * (Data → Filter) journals filter changes through the canonical
   * `SheetFilterState` save family and reads it back via this field on
   * reopen, so filter range + criteria + hidden rows survive round-trip.
   */
  readonly filterState?: Readonly<SheetFilterState> | undefined
  /**
   * Data-validation rules parsed from the worksheet's <dataValidations>
   * section. Absent means no parseable rules — including when the section
   * carries constructs the canonical model cannot represent (x14 extensions,
   * unknown types/operators/error styles, malformed sqref), which fail
   * closed: the browser never renders such rules and a no-op save preserves
   * the file's XML byte-for-byte. The web Sheets shell (Data → Data
   * Validation) journals changes through the canonical SheetDvState save
   * family and reads it back via this field on reopen, so validation rules
   * survive round-trip.
   */
  readonly dvRules?: readonly DvWireRule[] | undefined
  /**
   * Legacy notes (cell comments) parsed from the worksheet's comments part
   * (resolved through the worksheet rels). Absent means no parseable notes —
   * including when the part carries constructs the canonical model cannot
   * represent (unreadable refs, missing text), which fail closed: the
   * browser never renders such notes and a no-op save preserves the file's
   * parts byte-for-byte. The web Sheets shell (Review → New Comment)
   * journals note changes through the canonical SheetNoteState save family
   * and reads it back via this field on reopen.
   */
  readonly notes?: readonly SheetNote[] | undefined
  /**
   * Excel tables (ListObjects) parsed from the worksheet's <tableParts>
   * (resolved through the worksheet rels). Absent means the sheet carries
   * no representable tables — including when a part holds constructs the
   * model cannot represent (no readable ref), which are skipped per table,
   * or when the table wiring is unreadable, which fails closed PER SHEET:
   * the browser never renders such tables and a no-op save preserves the
   * file's parts byte-for-byte. The web Sheets shell (Insert → Table)
   * journals table creations through the canonical `tableAdditions` save
   * family and reads the state back via this field on reopen, so tables
   * survive round-trip; banding colors are pre-resolved (theme accents +
   * HSL tints + custom tableStyle dxfs) so the browser paints without
   * OOXML knowledge.
   */
  readonly tables?: readonly SheetTableInfo[] | undefined
  /**
   * Worksheet pictures (EXCEL-022) parsed from the sheet's drawing part
   * (resolved through the worksheet rels). Absent means the sheet carries
   * no representable pictures — including when the drawing wiring is
   * unreadable (fail closed PER SHEET) or when individual pictures are
   * unsupported (media type, size, missing part — skipped per picture,
   * anchors still counted for drawingIndex parity). The web Sheets shell
   * renders these through Univer's over-grid image model and journals
   * move/resize/delete through the canonical visualEdits save family,
   * keyed by the (drawingPath, drawingIndex) locator. NOTE: this is a
   * DEDICATED image surface — the chart-oriented `visuals` field above is
   * the desktop demo/AI chart replay state and is deliberately NOT
   * repurposed (EXCEL-022 architecture decision, worklog).
   */
  readonly images?: readonly SheetImageInfo[] | undefined
  /**
   * Worksheet charts (EXCEL-023) parsed from the sheet's drawing part
   * (resolved through the worksheet rels → drawing rels → chart parts).
   * Each entry carries the canonical (drawingPath, drawingIndex) anchor
   * locator for move/resize/delete through the visualEdits family, the
   * chartPath part locator for semantic edits through the chartEdits
   * family, and the full canonical ChartVisualState (the shared domain
   * model the desktop also renders from). Absent means the sheet carries
   * no representable charts — including when the drawing wiring is
   * unreadable (fail closed PER SHEET) or when individual charts are
   * unsupported (3-D plots, bubble/stock/surface, chartEx extensions,
   * non-canonical multi-plot combinations, absolute anchors — skipped per
   * chart, anchors still counted for drawingIndex parity). The web Sheets
   * shell renders charts as its own SVG visual surface floated over the
   * Univer grid (the desktop's exact rendering architecture — Univer
   * 0.25.1 ships no chart plugin) and journals semantic changes through
   * the canonical chartEdits save family, keyed by chartPath.
   */
  readonly charts?: readonly SheetChartInfo[] | undefined
  /**
   * Conditional-formatting rules (EXCEL-024) parsed from the worksheet's
   * `<conditionalFormatting>` sections. Each entry is the canonical
   * CfWireRule — the Univer conditional-formatting model shape with the
   * dxf style PRE-RESOLVED into rule.style (the browser installs them via
   * the add-conditional-rule mutation without touching style XML). Rules
   * come back priority-ascending (the install order — Univer applies
   * rules in insertion order, and lower xlsx priority = higher
   * precedence). Absent means the sheet carries no representable rules —
   * including when the sections hold constructs the canonical model
   * cannot represent (x14 extensions, time periods, unknown rule types,
   * malformed sqref, unresolvable dxf styling), which fail closed PER
   * SHEET with `cfLocked: true`: the browser never renders such a
   * surface, refuses CF edits on the sheet (so a rewrite that would
   * silently drop the unrepresentable rules can never be requested), and
   * a no-op save preserves the file's XML byte-for-byte. The web Sheets
   * shell (Home → Conditional Formatting) journals rule changes through
   * the canonical `cfStates` save family and reads the state back via
   * this field on reopen.
   */
  readonly cfRules?: readonly CfWireRule[] | undefined
  /**
   * EXCEL-024 fail-closed marker: the worksheet carries conditional
   * formatting the canonical typed model cannot represent (see cfRules).
   * The browser must refuse conditional-formatting mutations on this
   * sheet — the canonical writer would rewrite every section of a
   * CF-dirty sheet from the live-model snapshot, so an edit here would
   * silently drop the unrepresentable rules. A no-op save (no CF edits)
   * leaves the sheet's CF XML untouched.
   */
  readonly cfLocked?: boolean | undefined
  /**
   * Sheet protection state parsed from the worksheet's <sheetProtection>
   * element. Absent means the worksheet carries NO element (not
   * protected); present with `protected: false` means an element exists
   * but protection is disabled. `hasPassword` marks either password form
   * (legacy hash or modern algorithmName/hashValue) — unprotecting such a
   * sheet fails closed, so the browser must refuse the toggle up front.
   * The web Sheets shell (Review → Protect Sheet) journals changes
   * through the canonical SheetProtectionState save family and reads the
   * state back via this field on reopen, so protection survives
   * round-trip.
   */
  readonly sheetProtection?:
    Readonly<{ readonly protected: boolean; readonly hasPassword: boolean }> | undefined
}

export interface WorkbookSnapshot {
  readonly revision: number
  readonly sheets: readonly WorksheetState[]
  /**
   * Defined names (EXCEL-025) parsed from workbook.xml's
   * `<definedNames>` — the canonical `DefinedNamesState` split: `names`
   * are the entries the editor can safely model (valid per the writer's
   * rules, non-hidden, non-`_xlnm`, scoped inside the workbook, the
   * ranked winner of any duplicate group), `preserveNames` are the rest
   * (invalid names, out-of-range scopes, empty bodies, duplicate
   * losers) — entries the declarative rewrite must retain byte-verbatim
   * because it would otherwise drop them. Hidden names and `_xlnm.*`
   * built-ins (Print_Area/Print_Titles, owned by the page-setup family)
   * never appear here: the writer's keep-rules preserve them without a
   * preserve-list entry. The browser installs `names` into the real
   * Univer defined-name model (desktop applyDefinedNames rank parity),
   * unions its own install rejects into the preserve list, and journals
   * name edits through the canonical `definedNamesState` save family.
   * Absent means the workbook carries no names at all.
   */
  readonly definedNames?:
    | Readonly<{
        readonly names: readonly DefinedNameEntry[]
        readonly preserveNames: readonly string[]
      }>
    | undefined
  /**
   * EXCEL-025 fail-closed marker: workbook.xml carries a
   * `<definedNames>` section the reader could not structurally parse.
   * The browser must refuse name mutations on this workbook — the
   * declarative save would rewrite the section from a model that never
   * saw every entry, silently dropping the unparseable ones. A no-op
   * save (no name edits) leaves the section's XML byte-for-byte.
   */
  readonly namesLocked?: boolean | undefined
  /**
   * Workbook structure protection parsed from workbook.xml's
   * <workbookProtection> element. Absent means no element. The web shell
   * (Review → Protect Workbook) journals changes through the canonical
   * workbook-protection save family and reads the state back on reopen.
   */
  readonly workbookProtection?:
    Readonly<{ readonly lockStructure: boolean; readonly hasPassword: boolean }> | undefined
}

export interface CellChange {
  readonly sheetId: string
  readonly address: string
  readonly before: CellState
  readonly after: CellState
}

export interface SheetRename {
  readonly sheetId: string
  readonly before: string
  readonly after: string
}

export interface StructuralChange {
  readonly op: StructuralOperation | LayoutOperation
  readonly label: string
}

export interface FormatChange {
  readonly sheetId: string
  readonly range: string
  readonly format: CellFormatPatch
  readonly label: string
}

export interface ChangePlan {
  readonly transactionId: string
  readonly baseRevision: number
  readonly cellChanges: readonly CellChange[]
  readonly sheetRenames: readonly SheetRename[]
  readonly structuralChanges: readonly StructuralChange[]
  readonly formatChanges: readonly FormatChange[]
  readonly warnings: readonly string[]
}

/** result of applying a proposed plan to the live workbook */
export interface ApplyOutcome {
  readonly ok: boolean
  readonly reason?: string
  /** the failure hit mid-batch: earlier operations were already committed */
  readonly partiallyApplied?: boolean
}

export interface CommitReceipt {
  readonly transactionId: string
  readonly previousRevision: number
  readonly revision: number
}

export interface WorkbookAdapter {
  getSnapshot(): WorkbookSnapshot
  plan(input: unknown): ChangePlan
  apply(plan: ChangePlan): CommitReceipt
  undo(): CommitReceipt
}
