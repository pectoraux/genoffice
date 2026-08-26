/**
 * @genoffice/xlsx-gateway — canonical PURE XLSX planning implementation.
 *
 * This package owns the pure SavePlan → archive-patch translation logic
 * used by both apps/sheets and packages/platform-electron.
 *
 * The authoritative entry point is `planCellEditsToXlsx` (in src/gateway/
 * xlsx-gateway.ts), which accepts an EntrySource + mutation families and
 * produces a MutationPlan (replaced/added/addedBinary/removedEntries/
 * addedEntries/touchedEntries).
 *
 * PURITY (Increment 3F):
 *   ZERO Electron imports. ZERO node:* imports. ZERO apps/sheets imports.
 *   The package operates only on EntrySource (abstract archive reader),
 *   Buffers (in-memory), and pure XML/XLSX helpers. Runtime filesystem
 *   I/O (syncFileBestEffort, writeXlsxAtomically, mutateXlsxFile, sha256)
 *   and sidecar archive interaction (saveWorkbookViaSidecar, ArchiveClient)
 *   live in packages/platform-electron/src/capabilities/.
 */

// ── Gateway planning (pure) ──
export {
  planCellEditsToXlsx,
  applyPlanToXlsx,
  applyCellEditsToXlsx,
  createBufferEntrySource,
  assembleWithJsZip,
  readBasicWorkbook,
  inventoryXlsx,
  toA1Address,
  assertOnlyTouchedEntriesChanged,
  type CellEdit,
  type EntrySource,
  type MutationPlan,
  type XlsxMutation,
  type PackageEntry,
  type ImportedXlsx,
  type SheetStructuralOps,
  type SheetHyperlinkEdits,
  type SheetCfState,
  type SheetDvState,
  type SheetProtectionState,
  type SheetProtectedRangesState,
  type SheetFormulaValues,
  type SheetNoteState,
  type SheetVisualAddition,
  type SheetTableAddition,
  type SheetPivotAddition,
  type SheetSparklineAddition,
  type PivotRefreshUpdate,
} from './gateway/xlsx-gateway.js'

// ── Gateway types (previously from apps/sheets/src/shared/desktop-api) ──
export type {
  HexColor,
  EditableBorderStyle,
  StyleEditBorder,
  WorkbookStyleEdit,
  WorkbookRichRun,
  DrawingAnchor,
  ChartSeriesSetEntry,
  ChartSeriesEdit,
  ChartAxisTitles,
  ChartValueAxis,
  WorkbookChartEdit,
  WorkbookVisualEdit,
} from './types.js'

// ── Sheet edit plan ──
export type { SheetEditPlan, SheetAllocation } from './gateway/xlsx-sheets.js'
export { SheetEditError, validateSheetName } from './gateway/xlsx-sheets.js'

// ── Structural ops ──
export type { StructuralOp, AxisAttributeOp } from './gateway/xlsx-structure.js'

// ── Filter state ──
export type { SheetFilterState, FilterColumnState } from './gateway/xlsx-filter.js'
export { parseAutoFilter, FilterReadError, FilterEditError } from './gateway/xlsx-filter.js'

// ── Data validation ──
export type { DvWireRule, DvCellArea } from './gateway/xlsx-dv.js'
export { parseDataValidations, DvReadError, DvEditError } from './gateway/xlsx-dv.js'

// ── Notes (legacy cell comments) ──
export type { SheetNote } from './gateway/xlsx-notes.js'
export { parseCommentsPart, NoteReadError, NoteEditError } from './gateway/xlsx-notes.js'

// ── Defined names ──
export type { DefinedNamesState } from './gateway/xlsx-defined-names.js'
export { DefinedNameError } from './gateway/xlsx-defined-names.js'

// ── Page setup ──
export type { SheetPageSetupState } from './gateway/xlsx-page-setup.js'

// ── Theme ──
export type { WorkbookThemeState } from './gateway/xlsx-theme.js'

// ── Domain types (used by the gateway) ──
export type {
  CellScalar,
  CellState,
  CellFormatState,
  WorksheetState,
  WorkbookSnapshot,
  CellChange,
  SheetRename,
  StructuralChange,
  FormatChange,
  ChangePlan,
  ApplyOutcome,
  CommitReceipt,
  WorkbookAdapter,
} from './domain/workbook.types.js'

// ── Cell address utilities ──
export {
  columnIndex,
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
  rangeCellCount,
} from './domain/cell-address.js'

// ── Shape types ──
export { ADDABLE_SHAPE_TYPES, type AddableShapeType } from './shared/shape-types.js'

// ── Pure SHA-256 (for entry inventory hashing — no node:crypto) ──
export { sha256Hex } from './sha256.js'
