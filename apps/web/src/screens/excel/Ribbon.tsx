/**
 * GenOffice web Sheets — Ribbon.
 *
 * Seven top-level tabs matching the Electron Sheets ribbon (Home, Insert,
 * Page Layout, Formulas, Data, Review, View). Each tab carries commands
 * that have a CLEAR, CANONICAL save path through the wire
 * (BrowserWorkbookSavePlan → routeOffice → applyCellEditsToXlsx). Commands
 * whose save path is NOT yet wired through the canonical contract are
 * visibly DISABLED with a `title` tooltip that documents the exact
 * architectural reason — never faked (per spec: "a disabled button is
 * preferable to a fake feature").
 *
 * Functional commands (Phase 4 Increment 3):
 *   Home: undo/redo, font family/size, bold/italic/underline, font/fill
 *         color, alignment, wrap, merge, number format — ALL persist via
 *         style.numberFormat / WorkbookStyleEdit on CellEdit.
 *   Data: Sort ascending / Sort descending — FRange.sort() fires
 *         sheet.command.sort-range → ReorderRangeMutation, journaled by
 *         ExcelEditor's expanded subscription as per-cell writeValue edits.
 *   Formulas: ƒx Insert Function + AutoSum — write through commitFormula
 *         (=SUM(...)) → set-range-values → existing journal. No second
 *         formula engine runs.
 *   View: Gridlines toggle, Zoom, Freeze Panes — freeze fires
 *         sheet.command.set-frozen → journaled as a per-sheet
 *         BrowserSheetPageSetupState, persisted by the canonical
 *         applyPageSetupState (<pane> element).
 *
 * Disabled commands (with reason):
 *   Insert → Picture/Chart — the wire save plan does not expose
 *            visualAdditions / chartEdits families (applyCellEditsToXlsx
 *            accepts them but routeOffice does not pass them through).
 *   Formulas → Show formulas / Name Manager — needs pageSetupStates.
 *             showFormulas (a page-setup field) and definedNamesState,
 *             neither wired through the web save plan.
 *   Page Layout → all commands — pageSetupStates exists in the wire but
 *            only frozenRows/frozenColumns are wired by the web shell;
 *            orientation/margins/print area remain disabled until the
 *            web shell emits them.
 *
 * EXCEL-021 (Insert → Tables): ENABLED. Table journals the canonical
 * tableAdditions family over the active range (the browser sanitizes the
 * column names from the header row, desktop applyAiTableAdd parity);
 * Delete Table is convert-to-range for session tables (journal splice —
 * nothing reaches the file) and refuses file-native tables with the
 * desktop's exact message. File tables import through
 * WorksheetState.tables (banding painted into the cell matrix; Univer
 * registration with a muted plain theme), and a sheet whose filter
 * belongs to a table refuses filter edits (BeforeCommandExecute gate).
 *
 * EXCEL-020 (Review → Protection): ENABLED. Protect Sheet / Unprotect
 * Sheet journals the canonical sheetProtections family; Protect Workbook /
 * Unprotect Workbook journals workbookProtectionState; Lock Cell / Unlock
 * Cell journal WorkbookStyleEdit.protectionLocked style-only CellEdits
 * (the desktop's neutral-delta path — Univer's OSS presets carry no
 * cell-protection model). Password-protected sheets/structures are refused
 * up front (fail-closed — the gateway rejects them too). The editor itself
 * does not enforce protection (desktop parity: enforcement lives in the
 * saved file, for Excel and other readers).
 *
 * EXCEL-018 (Data → Remove Duplicates): ENABLED. The command opens a
 * small inline dialog with a "My data has headers" checkbox (default
 * checked — Excel's own default), then calls
 * `api.removeDuplicates(hasHeader)`. The dedupe runs through the pure
 * `apps/web/src/office/dedupe.ts` algorithm (mirrors the frozen desktop
 * reference at apps/sheets/src/renderer/dedupe.ts verbatim), writes the
 * result back per-row via `FWorksheet.getRange(...).setValues(...)`,
 * which fires `sheet.mutation.set-range-values` — the SAME canonical
 * channel Sort uses, journaled by ExcelEditor's existing subscription
 * as CellEdits and persisted by `applyCellEditsToXlsx`. No new save-plan
 * family, no wire change, no gateway change.
 *
 * The desktop's ExcelShell.tsx ribbon (3286 lines) is a frozen surface and
 * is NOT imported; this is a fresh web implementation using the desktop
 * only as the visual/interaction reference.
 */
import { useState, type ReactNode } from 'react'
import type { ExcelRuntimeApi, ExcelRuntimeState, RemoveDuplicatesResult } from './useExcelRuntime'
import {
  UndoIcon,
  RedoIcon,
  BoldIcon,
  ItalicIcon,
  UnderlineIcon,
  FontColorIcon,
  FillColorIcon,
  AlignLeftIcon,
  AlignCenterIcon,
  AlignRightIcon,
  AlignTopIcon,
  AlignMiddleIcon,
  AlignBottomIcon,
  WrapIcon,
  MergeIcon,
  TableIcon,
  ChartIcon,
  ImageIcon,
  GridlinesIcon,
  FreezePaneIcon,
  FunctionIcon,
  CfIcon,
} from './RibbonIcons'

type TabId = 'home' | 'insert' | 'page' | 'formulas' | 'data' | 'review' | 'view'

const TABS: ReadonlyArray<{ id: TabId; label: string }> = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'page', label: 'Page Layout' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
]

const FONT_FAMILIES = [
  'Calibri',
  'Cambria',
  'Arial',
  'Comic Sans MS',
  'Courier New',
  'Georgia',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
]
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]
const NUMBER_FORMATS: ReadonlyArray<{ label: string; pattern: string }> = [
  { label: 'General', pattern: 'General' },
  { label: 'Number', pattern: '#,##0' },
  { label: 'Currency', pattern: '$#,##0' },
  { label: 'Percentage', pattern: '0%' },
  { label: 'Date', pattern: 'yyyy-mm-dd' },
  { label: 'Time', pattern: 'hh:mm' },
  { label: 'Text', pattern: '@' },
]

function Group({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="excel-ribbon-group">
      <div className="excel-ribbon-group-controls">{children}</div>
      <div className="excel-ribbon-group-label">{label}</div>
    </div>
  )
}

function RibbonButton({
  label,
  title,
  icon,
  active,
  disabled,
  onClick,
}: {
  label?: string
  title: string
  icon?: ReactNode
  active?: boolean
  disabled?: boolean
  onClick?: () => void
}) {
  return (
    <button
      type="button"
      className={`rb-btn${active ? ' active' : ''}`}
      title={title}
      aria-label={title}
      aria-pressed={active ? true : undefined}
      data-tip={title}
      disabled={disabled}
      onClick={onClick}
    >
      {icon}
      {label ? <span>{label}</span> : null}
    </button>
  )
}

/** Fallback state shown before the runtime is ready, so the shell layout is
 * stable from first paint (the grid gets its final height before Univer
 * measures it — otherwise the canvas mounts at 0×0). */
const NULL_STATE: ExcelRuntimeState = {
  ready: false,
  activeCellA1: '',
  selectionFormat: {
    bold: false,
    italic: false,
    underline: false,
    fontFamily: null,
    fontSize: null,
    fontColor: null,
    fillColor: null,
    hAlign: null,
    vAlign: null,
    wrap: false,
  },
  zoomPercent: 100,
  gridlinesVisible: true,
  canUndo: false,
  canRedo: false,
  seq: 0,
}

/**
 * EXCEL-020 protection surface (Review → Protection group). The shell
 * (ExcelEditor) owns the journal + file state and passes it down here;
 * the buttons only reflect the echo and call back.
 */
export interface RibbonProtectionProps {
  /** Effective protection of the active sheet; null = no live sheet. */
  readonly sheetProtected: boolean | null
  /** The active sheet's file state carries a password (fail-closed). */
  readonly sheetHasPassword: boolean
  /** Effective structure lock; null = no file open. */
  readonly workbookLocked: boolean | null
  /** The workbook structure carries a password (fail-closed). */
  readonly workbookHasPassword: boolean
  readonly onToggleSheetProtection: () => void
  readonly onToggleWorkbookProtection: () => void
  /** Set the selection's protectionLocked flag (Lock Cell / Unlock Cell). */
  readonly onSetCellsLocked: (locked: boolean) => void
}

/**
 * EXCEL-021 tables surface (Insert → Tables group). The shell
 * (ExcelEditor) owns the file-state refs + the session journal and passes
 * the two commands down here; the buttons only call back.
 */
export interface RibbonTablesProps {
  /** Insert → Table: journal a table over the active range. */
  readonly onInsertTable: () => void
  /** Insert → Delete Table: convert-to-range for a session table. */
  readonly onDeleteTable: () => void
}

export interface RibbonImagesProps {
  /** Insert → Picture: open the browser file picker for an image embed. */
  readonly onInsertPicture: () => void
}

/**
 * EXCEL-024 conditional-formatting surface (Home → Styles group). The
 * shell (ExcelEditor) owns the runtime and passes the command down here;
 * the button only calls back and invokes Univer's own panel — the web
 * shell invents no CF UI of its own.
 */
export interface RibbonCfProps {
  /** Home → Conditional Formatting: open Univer's manage-rules panel. */
  readonly onOpenConditionalFormatting: () => void
}

/**
 * EXCEL-023 charts surface (Insert → Charts group). The shell
 * (ExcelEditor) owns the chart store + session journal and passes the
 * command down here; the button only calls back (the Chart Design pane
 * handles type selection over the active selection).
 */
export interface RibbonChartsProps {
  /** Insert → Chart: open the Chart Design pane in create mode. */
  readonly onInsertChart: () => void
}

export function Ribbon({
  api,
  protection,
  tables,
  images,
  charts,
  cf,
}: {
  api: ExcelRuntimeApi | null
  protection?: RibbonProtectionProps | null
  tables?: RibbonTablesProps | null
  images?: RibbonImagesProps | null
  charts?: RibbonChartsProps | null
  cf?: RibbonCfProps | null
}) {
  const [tab, setTab] = useState<TabId>('home')
  // EXCEL-018 — Remove Duplicates dialog state. Mirrors the desktop's
  // `setShowDedupeDialog(true)` flow: the Data → Remove Duplicates
  // button opens a small inline dialog with a "My data has headers"
  // checkbox (Excel's own default is checked). OK runs the dedupe and
  // surfaces the result as a transient message below the ribbon.
  const [dedupeOpen, setDedupeOpen] = useState(false)
  const [dedupeHasHeader, setDedupeHasHeader] = useState(true)
  const [dedupeMessage, setDedupeMessage] = useState<string | null>(null)
  const s = api?.state ?? NULL_STATE
  const disabled = !s.ready

  const runRemoveDuplicates = (hasHeader: boolean) => {
    if (!api) return
    const result: RemoveDuplicatesResult = api.removeDuplicates(hasHeader)
    let msg: string
    switch (result.kind) {
      case 'done':
        msg = `Removed ${result.removed} duplicate row(s).`
        break
      case 'noop':
        msg = result.message
        break
      case 'select':
        msg = 'Select the rows to check for duplicates first.'
        break
      case 'error':
        msg = result.message
        break
    }
    setDedupeMessage(msg)
    // Auto-clear the message after 5s — matches the desktop's transient
    // status bar behavior.
    window.setTimeout(() => {
      setDedupeMessage((cur) => (cur === msg ? null : cur))
    }, 5000)
  }

  return (
    <div className="excel-ribbon" data-testid="excel-ribbon">
      <div className="excel-ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
        {TABS.map((t) => (
          <button
            key={t.id}
            role="tab"
            aria-selected={tab === t.id}
            className={`excel-ribbon-tab${tab === t.id ? ' active' : ''}`}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className="excel-ribbon-panel">
        {tab === 'home' && (
          <>
            <Group label="Undo">
              <RibbonButton
                title="Undo"
                icon={<UndoIcon />}
                disabled={!s.canUndo}
                onClick={() => api?.undo()}
              />
              <RibbonButton
                title="Redo"
                icon={<RedoIcon />}
                disabled={!s.canRedo}
                onClick={() => api?.redo()}
              />
            </Group>
            <Group label="Font">
              <select
                className="rb-select"
                aria-label="Font family"
                value={s.selectionFormat.fontFamily ?? 'Calibri'}
                disabled={disabled}
                onChange={(e) => api?.setFontFamily(e.target.value)}
              >
                {FONT_FAMILIES.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
              <select
                className="rb-select"
                aria-label="Font size"
                style={{ maxWidth: 64 }}
                value={s.selectionFormat.fontSize ?? 11}
                disabled={disabled}
                onChange={(e) => api?.setFontSize(Number(e.target.value))}
              >
                {FONT_SIZES.map((n) => (
                  <option key={n} value={n}>
                    {n}
                  </option>
                ))}
              </select>
              <RibbonButton
                title="Bold"
                icon={<BoldIcon />}
                active={s.selectionFormat.bold}
                disabled={disabled}
                onClick={() => api?.toggleBold()}
              />
              <RibbonButton
                title="Italic"
                icon={<ItalicIcon />}
                active={s.selectionFormat.italic}
                disabled={disabled}
                onClick={() => api?.toggleItalic()}
              />
              <RibbonButton
                title="Underline"
                icon={<UnderlineIcon />}
                active={s.selectionFormat.underline}
                disabled={disabled}
                onClick={() => api?.toggleUnderline()}
              />
              <label
                className="rb-btn"
                title="Font color"
                aria-label="Font color"
                style={{ position: 'relative', padding: 0 }}
              >
                <FontColorIcon />
                <input
                  type="color"
                  aria-label="Font color picker"
                  className="rb-color"
                  style={{
                    position: 'absolute',
                    inset: 0,
                    opacity: 0,
                    width: '100%',
                    height: '100%',
                    padding: 0,
                    border: 'none',
                    background: 'transparent',
                    cursor: 'pointer',
                  }}
                  value={s.selectionFormat.fontColor ?? '#000000'}
                  disabled={disabled}
                  onChange={(e) => api?.setFontColor(e.target.value)}
                />
              </label>
              <label
                className="rb-btn"
                title="Fill color"
                aria-label="Fill color"
                style={{ position: 'relative', padding: 0 }}
              >
                <FillColorIcon />
                <input
                  type="color"
                  aria-label="Fill color picker"
                  className="rb-color"
                  style={{
                    position: 'absolute',
                    inset: 0,
                    opacity: 0,
                    width: '100%',
                    height: '100%',
                    padding: 0,
                    border: 'none',
                    background: 'transparent',
                    cursor: 'pointer',
                  }}
                  value={s.selectionFormat.fillColor ?? '#ffffff'}
                  disabled={disabled}
                  onChange={(e) => api?.setFillColor(e.target.value)}
                />
              </label>
            </Group>
            <Group label="Alignment">
              <RibbonButton
                title="Align left"
                icon={<AlignLeftIcon />}
                active={s.selectionFormat.hAlign === 'left'}
                disabled={disabled}
                onClick={() => api?.setHAlign(1)}
              />
              <RibbonButton
                title="Align center"
                icon={<AlignCenterIcon />}
                active={s.selectionFormat.hAlign === 'center'}
                disabled={disabled}
                onClick={() => api?.setHAlign(2)}
              />
              <RibbonButton
                title="Align right"
                icon={<AlignRightIcon />}
                active={s.selectionFormat.hAlign === 'right'}
                disabled={disabled}
                onClick={() => api?.setHAlign(3)}
              />
              <RibbonButton
                title="Align top"
                icon={<AlignTopIcon />}
                active={s.selectionFormat.vAlign === 'top'}
                disabled={disabled}
                onClick={() => api?.setVAlign(1)}
              />
              <RibbonButton
                title="Align middle"
                icon={<AlignMiddleIcon />}
                active={s.selectionFormat.vAlign === 'middle'}
                disabled={disabled}
                onClick={() => api?.setVAlign(2)}
              />
              <RibbonButton
                title="Align bottom"
                icon={<AlignBottomIcon />}
                active={s.selectionFormat.vAlign === 'bottom'}
                disabled={disabled}
                onClick={() => api?.setVAlign(3)}
              />
              <RibbonButton
                title="Wrap text"
                icon={<WrapIcon />}
                active={s.selectionFormat.wrap}
                disabled={disabled}
                onClick={() => api?.toggleWrap()}
              />
              <RibbonButton
                title="Merge & center"
                icon={<MergeIcon />}
                disabled={disabled}
                onClick={() => api?.toggleMerge()}
              />
            </Group>
            <Group label="Number">
              <select
                className="rb-select"
                aria-label="Number format"
                defaultValue="General"
                disabled={disabled}
                onChange={(e) => api?.setNumberFormat(e.target.value)}
              >
                {NUMBER_FORMATS.map((n) => (
                  <option key={n.label} value={n.pattern}>
                    {n.label}
                  </option>
                ))}
              </select>
            </Group>
            <Group label="Styles">
              {/* EXCEL-024 — Home → Conditional Formatting. Opens Univer's
                  own conditional-formatting panel (the preset the runtime
                  already registers); the shell invents no CF UI. Rule
                  changes journal through the canonical cfStates family —
                  the command gate refuses edits on sheets whose file CF
                  cannot be represented (fail-closed). */}
              <RibbonButton
                label="Conditional Formatting"
                title="Conditional Formatting — manage rules for the current selection (highlight, color scales, data bars, icon sets)"
                icon={<CfIcon />}
                disabled={disabled || !cf}
                onClick={() => cf?.onOpenConditionalFormatting()}
              />
            </Group>
          </>
        )}

        {tab === 'insert' && (
          <>
            <Group label="Tables">
              {/* EXCEL-021 — ENABLED. Insert → Table journals the canonical
                  tableAdditions family (BrowserWorkbookSavePlan →
                  routeOffice strict validation → the gateway's table-add
                  writer creates the table part with its worksheet wiring,
                  relationship, and content-type override). Delete Table is
                  convert-to-range for session tables (journal splice —
                  nothing reaches the file); file-native tables refuse with
                  the desktop's message. */}
              <RibbonButton
                label="Table"
                title="Table — create a table from the selected range (Insert → Table)"
                icon={<TableIcon />}
                disabled={disabled}
                onClick={() => tables?.onInsertTable()}
              />
              <RibbonButton
                label="Delete Table"
                title="Delete Table — convert a table created this session back to a range"
                icon={<TableIcon />}
                disabled={disabled}
                onClick={() => tables?.onDeleteTable()}
              />
            </Group>
            <Group label="Charts">
              {/* EXCEL-023: Insert → Chart — the Chart Design pane parses
                  the active selection (shared chartDataFromValues) and
                  journals a session chart persisted through the canonical
                  visualAdditions.chart family on save. */}
              <RibbonButton
                label="Chart"
                title="Chart — insert a chart from the selected data range (column, bar, line, area, pie, doughnut, scatter, radar, combo)"
                icon={<ChartIcon />}
                disabled={disabled || !charts}
                onClick={() => charts?.onInsertChart()}
              />
            </Group>
            <Group label="Illustrations">
              {/* EXCEL-022: Insert → Picture — a browser File/Blob upload
                  persisted through the canonical visualAdditions family
                  (PNG / JPEG / GIF only; the wire validates strictly). */}
              <RibbonButton
                label="Picture"
                title="Picture — insert an image from your device (PNG, JPEG, GIF)"
                icon={<ImageIcon />}
                disabled={disabled || !images}
                onClick={() => images?.onInsertPicture()}
              />
            </Group>
          </>
        )}

        {tab === 'page' && (
          <>
            <Group label="Page Setup">
              {/* Disabled: pageSetupStates is exposed on the wire but only
                 frozenRows/frozenColumns are wired by the web shell today.
                 Orientation / paperSize / margins / print area remain
                 disabled until the web shell emits them. */}
              <RibbonButton
                label="Margins"
                title="Margins — disabled: pageSetupStates is wired only for freeze panes today"
                disabled
              />
              <RibbonButton
                label="Orientation"
                title="Orientation — disabled: pageSetupStates is wired only for freeze panes today"
                disabled
              />
              <RibbonButton
                label="Size"
                title="Size — disabled: pageSetupStates is wired only for freeze panes today"
                disabled
              />
            </Group>
            <Group label="Scale">
              <RibbonButton
                label="Width"
                title="Width — disabled: pageSetupStates is wired only for freeze panes today"
                disabled
              />
              <RibbonButton
                label="Height"
                title="Height — disabled: pageSetupStates is wired only for freeze panes today"
                disabled
              />
            </Group>
          </>
        )}

        {tab === 'formulas' && (
          <>
            <Group label="Function Library">
              <RibbonButton
                label="ƒx"
                title="Insert function — writes =SUM() into the active cell via the formula bar"
                icon={<FunctionIcon />}
                disabled={disabled}
                onClick={() => api?.insertFunction('SUM()')}
              />
              <RibbonButton
                label="AutoSum"
                title="AutoSum — inserts =SUM(<range above>) into the active cell"
                disabled={disabled}
                onClick={() => api?.insertFunction('SUM()')}
              />
            </Group>
            <Group label="Defined Names">
              {/* Disabled: the wire save plan does not expose
                 definedNamesState. */}
              <RibbonButton
                label="Name Manager"
                title="Name Manager — disabled: the web save plan does not yet expose definedNamesState"
                disabled
              />
            </Group>
          </>
        )}

        {tab === 'data' && (
          <>
            <Group label="Sort & Filter">
              <RibbonButton
                label="Sort Asc"
                title="Sort ascending by the first column of the active range"
                disabled={disabled}
                onClick={() => api?.sortRange(true)}
              />
              <RibbonButton
                label="Sort Desc"
                title="Sort descending by the first column of the active range"
                disabled={disabled}
                onClick={() => api?.sortRange(false)}
              />
              {/* Data → Filter: toggles the AutoFilter through the REAL
                  Univer command (smart-toggle-filter). Persistence chain
                  (Phase 4 Increment 4): filter mutations mark the sheet
                  filter-dirty → on save, the live filter model is snapshotted
                  as canonical SheetFilterState → /api/office/workbooks/save
                  → xlsx-gateway applyFilterState writes the autoFilter
                  element + row visibility → reopen parses it back
                  (ribbon-filter.spec.ts proves the round-trip). Value
                  filters, blank filtering, and the six supported custom
                  operators round-trip; color filters fail closed at save
                  with an explicit error. */}
              <RibbonButton
                label="Filter"
                title="Toggle the AutoFilter on the active sheet (Data → Filter)"
                disabled={disabled}
                onClick={() => api?.toggleFilter()}
              />
            </Group>
            <Group label="Data Tools">
              {/* EXCEL-018 — Remove Duplicates (ENABLED). Opens a small
                  inline dialog with a "My data has headers" checkbox,
                  then runs `api.removeDuplicates(hasHeader)`. The dedupe
                  result is written back per-row through
                  FWorksheet.getRange().setValues(...) — the SAME
                  canonical mutation channel Sort uses — journaled by
                  ExcelEditor's existing subscription as CellEdits, and
                  persisted by the canonical applyCellEditsToXlsx on save.
                  The dedupe algorithm mirrors the frozen desktop
                  reference at apps/sheets/src/renderer/dedupe.ts
                  verbatim (case-insensitive text, type-strict, header
                  preserved, unchanged rows untouched so formulas/styles
                  survive). */}
              <RibbonButton
                label="Remove Duplicates"
                title="Remove duplicate rows from the selected range (Data → Remove Duplicates)"
                disabled={disabled}
                onClick={() => {
                  setDedupeHasHeader(true)
                  setDedupeMessage(null)
                  setDedupeOpen(true)
                }}
              />
              <RibbonButton
                label="Data Validation"
                title="Open the Data Validation panel for the active sheet (Data → Data Validation)"
                disabled={disabled}
                onClick={() => api?.openDataValidation()}
              />
            </Group>
          </>
        )}

        {tab === 'review' && (
          <>
            <Group label="Comments">
              {/* Disabled: the wire save plan does not expose noteStates. */}
              <RibbonButton
                label="New Comment"
                title="Open the note editor for the selected cell (Review → New Comment)"
                disabled={disabled}
                onClick={() => api?.addNote()}
              />
            </Group>
            <Group label="Protection">
              {/* EXCEL-020 — Protect Sheet / Unprotect Sheet. Journal-only
                  semantics (desktop parity): the toggle records the desired
                  state, the save writes the canonical sheetProtections
                  family, and the label flips with the effective echo. The
                  title IS the accessible name (exact) — the rich feedback
                  (no-password note, enforcement note) lives in the status
                  messages, exactly like the desktop's message strings. */}
              <RibbonButton
                label={protection?.sheetProtected ? 'Unprotect Sheet' : 'Protect Sheet'}
                title={protection?.sheetProtected ? 'Unprotect Sheet' : 'Protect Sheet'}
                active={protection?.sheetProtected === true}
                disabled={disabled}
                onClick={() => protection?.onToggleSheetProtection()}
              />
              {/* EXCEL-020 — Protect Workbook / Unprotect Workbook: the
                  structure lock, written through the canonical
                  workbookProtectionState family. */}
              <RibbonButton
                label={protection?.workbookLocked ? 'Unprotect Workbook' : 'Protect Workbook'}
                title={protection?.workbookLocked ? 'Unprotect Workbook' : 'Protect Workbook'}
                active={protection?.workbookLocked === true}
                disabled={disabled}
                onClick={() => protection?.onToggleWorkbookProtection()}
              />
              {/* EXCEL-020 — Lock Cell / Unlock Cell: journal canonical
                  WorkbookStyleEdit.protectionLocked deltas for the
                  selection (the desktop's Format Cells → Protection
                  neutral-delta path). Together with a protected sheet this
                  is Excel's editable-vs-locked semantics in the file:
                  unlocked cells stay editable, everything else is
                  read-only for readers that enforce protection. */}
              <RibbonButton
                label="Lock Cell"
                title="Lock Cell"
                disabled={disabled}
                onClick={() => protection?.onSetCellsLocked(true)}
              />
              <RibbonButton
                label="Unlock Cell"
                title="Unlock Cell"
                disabled={disabled}
                onClick={() => protection?.onSetCellsLocked(false)}
              />
            </Group>
          </>
        )}

        {tab === 'view' && (
          <>
            <Group label="Show">
              <RibbonButton
                label="Gridlines"
                title="Toggle gridlines (in-session)"
                icon={<GridlinesIcon />}
                active={s.gridlinesVisible}
                disabled={disabled}
                onClick={() => api?.toggleGridlines()}
              />
            </Group>
            <Group label="Zoom">
              <RibbonButton
                title="Zoom out"
                label="−"
                disabled={disabled}
                onClick={() => api?.zoomOut()}
              />
              <input
                className="rb-select"
                type="range"
                min={50}
                max={400}
                step={10}
                aria-label="Zoom slider"
                value={Math.min(400, Math.max(50, s.zoomPercent))}
                disabled={disabled}
                onChange={(e) => api?.setZoom(Number(e.target.value) / 100)}
                style={{ width: 100, accentColor: 'var(--accent)' }}
              />
              <RibbonButton
                title="Zoom in"
                label="+"
                disabled={disabled}
                onClick={() => api?.zoomIn()}
              />
              <span
                style={{
                  fontSize: 11,
                  color: 'var(--text-secondary)',
                  minWidth: 38,
                  textAlign: 'right',
                }}
              >
                {s.zoomPercent}%
              </span>
            </Group>
            <Group label="Window">
              <RibbonButton
                label="Freeze Panes"
                title="Freeze panes at the active cell — toggle (persists on save/reopen)"
                icon={<FreezePaneIcon />}
                disabled={disabled}
                onClick={() => {
                  const failure = api?.toggleFreezePanes() ?? null
                  if (failure) {
                    /* best-effort: surface failure in title attribute */
                    void failure
                  }
                }}
              />
            </Group>
          </>
        )}
      </div>
      {/* EXCEL-018 — Remove Duplicates inline dialog. Rendered at the
          ribbon level (not in the Data tab panel) so it stays mounted
          regardless of which tab the user switched to after opening it.
          Mirrors the desktop's `setShowDedupeDialog(true)` modal: a
          checkbox for "My data has headers" (default checked — Excel's
          own default) plus OK/Cancel. OK runs `api.removeDuplicates(...)`
          and surfaces the result as a transient message. */}
      {dedupeOpen && (
        <div
          className="rb-dialog-overlay"
          role="dialog"
          aria-modal="true"
          aria-labelledby="dedupe-title"
          data-testid="dedupe-dialog"
          onClick={(e) => {
            if (e.target === e.currentTarget) setDedupeOpen(false)
          }}
        >
          <div className="rb-dialog">
            <h3 id="dedupe-title" className="rb-dialog-title">
              Remove Duplicates
            </h3>
            <p className="rb-dialog-body">
              Remove duplicate rows from the selected range. The first occurrence of each unique row
              is kept; later matches are cleared. Rows outside the selection are untouched.
            </p>
            <label className="rb-dialog-check">
              <input
                type="checkbox"
                checked={dedupeHasHeader}
                onChange={(e) => setDedupeHasHeader(e.target.checked)}
              />
              <span>My data has headers</span>
            </label>
            <div className="rb-dialog-actions">
              <button
                type="button"
                className="rb-dialog-cancel"
                onClick={() => setDedupeOpen(false)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="rb-dialog-ok"
                data-testid="dedupe-ok"
                onClick={() => {
                  runRemoveDuplicates(dedupeHasHeader)
                  setDedupeOpen(false)
                }}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Transient status message — surfaces the dedupe result (or
          "select rows first" / "no duplicates found" / error) for a
          few seconds. Mirrors the desktop's `appDuplicatesRemoved` /
          `appNoDuplicates` / `appDedupeSelectRows` status strings. */}
      {dedupeMessage && (
        <div className="rb-toast" role="status" data-testid="dedupe-message">
          {dedupeMessage}
        </div>
      )}
    </div>
  )
}
