import type {
  BorderPatch,
  CellFormatPatch,
  LayoutOperation,
  StructuralOperation,
} from './workbook-dsl'
import type { SheetVisual } from './chart-visual'
import type { SheetFilterState } from '../gateway/xlsx-filter'
import type { DvWireRule } from '../gateway/xlsx-dv'
import type { SheetNote } from '../gateway/xlsx-notes'
import type { SheetTableInfo } from '../gateway/xlsx-table-read'

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
