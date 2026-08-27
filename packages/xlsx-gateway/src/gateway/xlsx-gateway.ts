import JSZip from 'jszip'

import { sha256Hex } from '../sha256.js'

import type {
  CellFormatState,
  CellState,
  ChangePlan,
  WorkbookSnapshot,
  WorksheetState,
} from '../domain/workbook.types'
import type {
  WorkbookChartEdit,
  WorkbookRichRun,
  WorkbookStyleEdit,
  WorkbookVisualEdit,
} from '../types.js'
import { applyChartEdit } from './xlsx-chart'
import { applyVisualEdits } from './xlsx-drawing-edit'
import {
  applyVisualAdditions,
  type ChartAdd,
  type DrawingAnchor,
  type ImageAdd,
  relsPathFor,
  resolveRelTarget,
  type ShapeAdd,
} from './xlsx-drawing-add'
import { ImageReadError, parseSheetImages, type SheetImageInfo } from './xlsx-image-read'
import { ChartReadError, parseSheetCharts, type SheetChartInfo } from './xlsx-chart-read'
import { applyTableAdditions, type TableArea } from './xlsx-table-add'
import type { PivotFilterDef } from '../domain/pivot-filters'
import {
  applyPivotAdditions,
  type PivotAddGrouping,
  type PivotAddRowLine,
  type PivotValueSpec,
} from './xlsx-pivot-add'
import type { SheetFilterState } from './xlsx-filter'
import { applyFilterState, FilterReadError, parseAutoFilter } from './xlsx-filter'
import type { SheetAllocation, SheetEditPlan, SheetElement } from './xlsx-sheets'
import {
  addWorksheetOverride,
  addWorksheetRelationship,
  applySheetPlanToWorkbookXml,
  assertNoSheetScopedDefinedNames,
  buildWorksheetPartXml,
  chartReferencesSheet,
  classifyRemovedSheetRels,
  definedNamesReferenceSheet,
  definedNamesUseToken,
  maxRelationshipId,
  maxSheetIdInWorkbook,
  parseRelationships,
  parseSheetElements,
  partPathForRels,
  pivotCacheReadsFromSheet,
  prepareClonedSheetRels,
  removePartOverride,
  removeRelationshipById,
  tableDisplayName,
  renameSheetInPivotCacheSource,
  renameSheetReferencesInChart,
  renameSheetReferencesInDefinedNames,
  renameSheetReferencesInWorksheet,
  sanitizeClonedWorksheetXml,
  SheetEditError,
  stripPageSetupRelIds,
  validateSheetName,
  worksheetReferencesSheet,
} from './xlsx-sheets'
import type { StructuralOp } from './xlsx-structure'
import { applyCfRules, CfReadError, parseConditionalFormatting, type CfWireRule } from './xlsx-cf'
import {
  applyDefinedNamesState,
  DefinedNameError,
  type DefinedNamesState,
} from './xlsx-defined-names'
import { applyDvRules, DvReadError, parseDataValidations, type DvWireRule } from './xlsx-dv'
import { applyPageSetupState, applyPrintAreas, type SheetPageSetupState } from './xlsx-page-setup'
import {
  applyProtectedRanges,
  applySheetProtection,
  applyWorkbookProtection,
  parseSheetProtectionState,
  parseWorkbookProtectionState,
  type ProtectedRangeState,
} from './xlsx-protection'
import { applyThemeState, type WorkbookThemeState } from './xlsx-theme'
import { applySheetNotes, NoteReadError, parseCommentsPart, type SheetNote } from './xlsx-notes'
import {
  parseSheetTables,
  readCustomTableStyles,
  readTableThemePalette,
  TableReadError,
  type SheetTableInfo,
} from './xlsx-table-read'

const COMMENTS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
import {
  applySparklineAdditions,
  type SheetSparklineAddition,
  type SparklineGroupAdd,
} from './xlsx-sparkline'
import { setPivotRefreshOnLoad } from './xlsx-pivot'
import { applyPivotLayoutExpansions, type PivotRefreshUpdate } from './xlsx-pivot-expand'
import {
  applyHyperlinkEdits,
  ensureRelationshipNamespace,
  type HyperlinkEdit,
} from './xlsx-hyperlinks'
import {
  applyStructuralOps,
  isShiftingOp,
  shiftChartReferences,
  shiftCrossSheetFormulas,
  shiftDefinedNames,
  shiftDrawingAnchors,
  shiftTablePart,
  StructuralShiftError,
} from './xlsx-structure'
import { StylesheetEditor, StylesheetReader } from './xlsx-styles'

const MAX_ENTRY_COUNT = 10_000
const MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024

export interface PackageEntry {
  readonly path: string
  readonly size: number
  readonly sha256: string
}

export interface XlsxMutation {
  readonly buffer: Buffer
  readonly touchedEntries: readonly string[]
  /// Entries intentionally dropped from the package (e.g. a stale calcChain).
  readonly removedEntries: readonly string[]
  /// Entries intentionally created (e.g. an added worksheet part).
  readonly addedEntries: readonly string[]
  readonly beforeEntries: readonly PackageEntry[]
  readonly afterEntries: readonly PackageEntry[]
  /// EXCEL-022: locators of visuals persisted by this save's
  /// visualAdditions — the caller merges them into its live image state
  /// so a later edit targets the exact appended anchor.
  readonly addedVisuals?: readonly AddedVisualLocator[]
}

/** Locator of a visual persisted by applyVisualAdditions. */
export interface AddedVisualLocator {
  readonly worksheetPath: string
  readonly drawingPath: string
  readonly drawingIndex: number
  /** EXCEL-023: the allocated chart part for chart additions — later
   * chartEdits target this exact path. */
  readonly chartPath?: string | undefined
}

export interface SheetStructuralOps {
  readonly sheetName: string
  readonly ops: readonly StructuralOp[]
}

export interface SheetHyperlinkEdits {
  readonly sheetName: string
  readonly edits: readonly HyperlinkEdit[]
}

export interface SheetCfState {
  readonly sheetName: string
  readonly rules: readonly CfWireRule[]
}

export interface SheetDvState {
  readonly sheetName: string
  readonly rules: readonly DvWireRule[]
}

export interface SheetProtectionState {
  readonly sheetName: string
  readonly protected: boolean
}

/// Full allow-edit-range snapshot for one sheet ([] removes the element).
export interface SheetProtectedRangesState {
  readonly sheetName: string
  readonly ranges: readonly ProtectedRangeState[]
}

/// Recalculated cached values for formula cells: the engine already
/// computed them for the screen, and the save writes them into <v> so the file's
/// inputs and outputs agree even for readers without a formula engine
/// (openpyxl data_only, pandas, preview services).
export interface SheetFormulaValues {
  readonly sheetName: string
  readonly cells: readonly {
    readonly row: number
    readonly column: number
    readonly value: string | number | boolean | null
  }[]
}

export interface ImportedXlsx {
  readonly snapshot: WorkbookSnapshot
  readonly sheetNamesById: Readonly<Record<string, string>>
}

export interface CellEdit {
  readonly sheetName: string
  readonly row: number
  readonly column: number
  /// false = style-only edit; the cell's stored content stays untouched.
  readonly writeValue: boolean
  readonly cell: CellState
  readonly style?: WorkbookStyleEdit | undefined
  /// Per-run styling for a rich-text string value.
  readonly rich?: readonly WorkbookRichRun[] | undefined
  /// Reset the cell to the default style (xf 0) before applying `style`.
  readonly styleReset?: boolean | undefined
}

/// Read access to the entries of a source package, independent of whether
/// the bytes live in an in-memory JSZip buffer or behind the sidecar.
export interface EntrySource {
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  /// False when the entry is too large to load for patching (readText would
  /// fail). Absent means everything is patchable (in-memory sources).
  canPatch?(path: string): Promise<boolean>
  /// Whether the entry's decoded XML text contains `needle`; consulted only
  /// for entries that cannot be patched, to decide skip vs fail-closed.
  containsText?(path: string, needle: string): Promise<boolean>
}

/// The entry-level outcome of patch planning: what an assembler (in-memory
/// JSZip or the sidecar streaming writer) must replace, add, and drop.
export interface MutationPlan {
  readonly replaced: ReadonlyMap<string, string>
  readonly added: ReadonlyMap<string, string>
  /// Non-text entries (media bytes); disjoint from `added`.
  readonly addedBinary: ReadonlyMap<string, Uint8Array>
  readonly removedEntries: readonly string[]
  readonly addedEntries: readonly string[]
  readonly touchedEntries: readonly string[]
  /// EXCEL-022: locators of appended visual anchors, in addition order.
  readonly addedVisuals?: readonly AddedVisualLocator[]
}

/// Overlay of pending edits on top of a read-only source package. Planning
/// code reads through the overlay so later stages see earlier rewrites.
class PackageEditor {
  private readonly overlay = new Map<string, string>()
  private readonly binaryOverlay = new Map<string, Uint8Array>()
  private readonly removed = new Set<string>()
  private readonly addedPaths = new Set<string>()

  constructor(private readonly source: EntrySource) {}

  async paths(): Promise<readonly string[]> {
    const base = (await this.source.paths()).filter((path) => !this.removed.has(path))
    return [...base, ...this.addedPaths, ...this.binaryOverlay.keys()]
  }

  async has(path: string): Promise<boolean> {
    if (this.removed.has(path)) return false
    if (this.overlay.has(path) || this.binaryOverlay.has(path)) return true
    return this.source.has(path)
  }

  async readText(path: string): Promise<string> {
    if (!this.removed.has(path)) {
      const pending = this.overlay.get(path)
      if (pending !== undefined) return pending
      if (await this.source.has(path)) return this.source.readText(path)
    }
    throw new Error(`Workbook is missing ${path}.`)
  }

  write(path: string, content: string): void {
    if (this.removed.has(path)) throw new Error(`Cannot write removed entry ${path}.`)
    this.overlay.set(path, content)
  }

  add(path: string, content: string): void {
    this.addedPaths.add(path)
    this.overlay.set(path, content)
  }

  addBinary(path: string, bytes: Uint8Array): void {
    if (this.removed.has(path)) throw new Error(`Cannot write removed entry ${path}.`)
    this.binaryOverlay.set(path, bytes)
  }

  remove(path: string): void {
    this.overlay.delete(path)
    if (!this.addedPaths.delete(path)) this.removed.add(path)
  }

  async canPatch(path: string): Promise<boolean> {
    if (this.overlay.has(path)) return true
    return this.source.canPatch?.(path) ?? true
  }

  async containsText(path: string, needle: string): Promise<boolean> {
    const pending = this.overlay.get(path)
    if (pending !== undefined) return pending.includes(needle)
    if (!this.source.containsText) {
      throw new Error(`Cannot scan ${path} for references.`)
    }
    return this.source.containsText(path, needle)
  }

  toPlan(touchedEntries: ReadonlySet<string>): MutationPlan {
    const replaced = new Map<string, string>()
    const added = new Map<string, string>()
    for (const [path, content] of this.overlay) {
      if (this.addedPaths.has(path)) added.set(path, content)
      else replaced.set(path, content)
    }
    return {
      replaced,
      added,
      addedBinary: new Map(this.binaryOverlay),
      removedEntries: [...this.removed].sort(),
      addedEntries: [...this.addedPaths, ...this.binaryOverlay.keys()].sort(),
      touchedEntries: [...touchedEntries].sort(),
    }
  }
}

const DEFAULT_STYLESHEET_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>'

const STYLES_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
const STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'

async function addDefaultStylesheet(
  pkg: PackageEditor,
  touchedEntries: Set<string>,
): Promise<void> {
  const stylesPath = 'xl/styles.xml'
  pkg.add(stylesPath, DEFAULT_STYLESHEET_XML)
  touchedEntries.add(stylesPath)

  const relationshipsPath = 'xl/_rels/workbook.xml.rels'
  const relationships = await pkg.readText(relationshipsPath)
  if (!relationships.includes(`Type="${STYLES_REL_TYPE}"`)) {
    const relationship =
      `<Relationship Id="rId${maxRelationshipId(relationships) + 1}" ` +
      `Type="${STYLES_REL_TYPE}" Target="styles.xml"/>`
    pkg.write(
      relationshipsPath,
      relationships.replace('</Relationships>', `${relationship}</Relationships>`),
    )
    touchedEntries.add(relationshipsPath)
  }

  const contentTypesPath = '[Content_Types].xml'
  const contentTypes = await pkg.readText(contentTypesPath)
  if (!contentTypes.includes('PartName="/xl/styles.xml"')) {
    const override = `<Override PartName="/xl/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/>`
    pkg.write(contentTypesPath, contentTypes.replace('</Types>', `${override}</Types>`))
    touchedEntries.add(contentTypesPath)
  }
}

export async function createBufferEntrySource(buffer: Buffer): Promise<{
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  /// EXCEL-022: raw media bytes for the image reader. Only the in-memory
  /// buffer source can serve this — the platform archive adapter never
  /// reads images, so EntrySource itself stays unchanged.
  readBinary(path: string): Promise<Uint8Array>
}> {
  const zip = await loadSafeZip(buffer)
  return {
    paths: async () =>
      Object.entries(zip.files)
        .filter(([, file]) => !file.dir)
        .map(([path]) => path),
    has: async (path) => zip.file(path) !== null,
    readText: (path) => readTextEntry(zip, path),
    readBinary: async (path) => {
      const entry = zip.file(path)
      if (entry === null) throw new Error(`Missing ZIP entry ${path}.`)
      return new Uint8Array(await entry.async('uint8array'))
    },
  }
}

/// In-memory assembler: applies a plan onto the source buffer with JSZip.
/// Every entry is recompressed, so it stays subject to the whole-package
/// decompression limit — the sidecar streaming assembler is the large-file
/// path.
export async function assembleWithJsZip(source: Buffer, plan: MutationPlan): Promise<XlsxMutation> {
  const zip = await loadSafeZip(source)
  const beforeEntries = await inventoryXlsx(source)
  for (const path of plan.removedEntries) zip.remove(path)
  for (const [path, content] of plan.replaced) zip.file(path, content, { createFolders: false })
  for (const [path, content] of plan.added) zip.file(path, content, { createFolders: false })
  for (const [path, bytes] of plan.addedBinary) zip.file(path, bytes, { createFolders: false })
  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  return {
    buffer,
    touchedEntries: plan.touchedEntries,
    removedEntries: plan.removedEntries,
    addedEntries: plan.addedEntries,
    beforeEntries,
    afterEntries: await inventoryXlsx(buffer),
    ...(plan.addedVisuals !== undefined ? { addedVisuals: plan.addedVisuals } : {}),
  }
}

export async function readBasicWorkbook(buffer: Buffer): Promise<ImportedXlsx> {
  const zip = await createBufferEntrySource(buffer)
  const workbookXml = await zip.readText('xl/workbook.xml')
  const sharedStrings = await readSharedStrings(zip)
  // Table style context (EXCEL-021): the theme palette + custom
  // <tableStyle> dxfs are parsed ONCE for the whole workbook — every
  // sheet's banding colors resolve against the same accents.
  const tableTheme = await readTableThemePalette(zip)
  const customTableStyles = await readCustomTableStyles(zip, tableTheme)
  // Presentation reader: resolves cellXfs indexes to the editable format
  // subset so the browser grid renders existing styling (bold, fills,
  // alignment, …). Unmodeled properties stay in the file's own XML.
  const stylesXml = (await zip.has('xl/styles.xml')) ? await zip.readText('xl/styles.xml') : null
  const styleReader = stylesXml !== null ? new StylesheetReader(stylesXml) : null
  const sheets: WorksheetState[] = []
  const sheetNamesById: Record<string, string> = {}
  const sheetPattern = /<sheet\b([^>]*)\/?>/g
  let match: RegExpExecArray | null
  while ((match = sheetPattern.exec(workbookXml)) !== null) {
    const attributes = match[1] ?? ''
    const name = readXmlAttribute(attributes, 'name')
    const sheetNumber = readXmlAttribute(attributes, 'sheetId')
    if (!name || !sheetNumber) continue
    const decodedName = decodeXmlText(name)
    const id = `sheet-${sheetNumber}`
    const worksheetPath = await resolveWorksheetPath(zip, decodedName)
    const worksheetXml = await zip.readText(worksheetPath)
    const presentation = parseWorksheetPresentation(worksheetXml, styleReader)
    // AutoFilter read: fail closed PER FILTER — an unrepresentable
    // <autoFilter> (top10 / dynamicFilter / iconFilter / dateGroup / color
    // criteria) surfaces no filterState, so the browser never renders a
    // filter it cannot save faithfully, while the workbook itself still
    // opens and a no-op save preserves the file's XML byte-for-byte.
    let filterState: SheetFilterState | undefined
    try {
      const parsed = parseAutoFilter(worksheetXml, decodedName)
      if (parsed !== null) filterState = parsed
    } catch (error) {
      if (!(error instanceof FilterReadError)) throw error
    }
    // Data-validation read: fail closed PER SHEET — an unrepresentable
    // <dataValidations> section (x14 extensions, unknown
    // types/operators/error styles, malformed sqref) surfaces no dvRules,
    // so the browser never renders a validation it cannot save faithfully,
    // while the workbook itself still opens and a no-op save preserves the
    // file's XML byte-for-byte.
    let dvRules: readonly DvWireRule[] | undefined
    try {
      const parsed = parseDataValidations(worksheetXml)
      if (parsed.length > 0) dvRules = parsed
    } catch (error) {
      if (!(error instanceof DvReadError)) throw error
    }
    // Notes read: resolve the comments part through the worksheet rels
    // (the same two-step lookup resolveWorksheetPath uses) and parse it.
    // Fail closed PER SHEET — an unrepresentable comments part (unreadable
    // refs, missing text, oversized sets) surfaces no notes, the workbook
    // still opens, and a no-op save preserves the file's parts.
    let sheetNotes: readonly SheetNote[] | undefined
    try {
      const commentsPath = await resolveCommentsPath(zip, worksheetPath)
      if (commentsPath !== null && (await zip.has(commentsPath))) {
        const parsed = parseCommentsPart(await zip.readText(commentsPath))
        if (parsed.length > 0) sheetNotes = parsed
      }
    } catch (error) {
      if (!(error instanceof NoteReadError)) throw error
    }
    // Sheet protection read (EXCEL-020): the <sheetProtection> element is
    // a flat attribute bag — parse it directly. null (absent field) means
    // no element, so a no-op save preserves the file's XML byte-for-byte
    // while a protected sheet surfaces both its enabled flag and whether
    // a password guards it (the toggle must refuse up front — the gateway
    // fails closed on unprotecting password-bearing elements).
    const sheetProtection = parseSheetProtectionState(worksheetXml)
    // Table read (EXCEL-021): resolve <tableParts> through the worksheet
    // rels and parse each table part into the canonical SheetTableInfo
    // (metadata + pre-resolved banding colors). Fail closed PER SHEET —
    // unreadable table wiring surfaces no tables, the workbook still
    // opens, and a no-op save preserves the file's parts byte-for-byte.
    // Individual parts without a readable ref are skipped (desktop
    // read_sheet_tables parity).
    let sheetTables: readonly SheetTableInfo[] | undefined
    try {
      const parsed = await parseSheetTables(zip, worksheetPath, worksheetXml, {
        theme: tableTheme,
        customStyles: customTableStyles,
      })
      if (parsed.length > 0) sheetTables = parsed
    } catch (error) {
      if (!(error instanceof TableReadError)) throw error
    }
    // Image read (EXCEL-022): resolve the sheet's drawing relationship
    // chain into typed picture entries with inline media. Fail closed
    // PER SHEET — unreadable drawing wiring surfaces no images, the
    // workbook still opens, and a no-op save preserves the file's parts
    // byte-for-byte. Individual unsupported pictures (media type, size,
    // missing part) are skipped while their anchors still count toward
    // the drawingIndex parity with the desktop sidecar.
    let sheetImages: readonly SheetImageInfo[] | undefined
    try {
      const parsed = await parseSheetImages(zip, worksheetPath, worksheetXml)
      if (parsed.length > 0) sheetImages = parsed
    } catch (error) {
      if (!(error instanceof ImageReadError)) throw error
    }
    // Chart read (EXCEL-023): resolve the same drawing relationship
    // chain into typed chart entries carrying the canonical
    // ChartVisualState plus both locators (anchor pair for the
    // visualEdits family, chartPath for the chartEdits family). Fail
    // closed PER SHEET — unreadable drawing wiring surfaces no charts,
    // the workbook still opens, and a no-op save preserves the file's
    // parts byte-for-byte. Individual unsupported charts (3-D,
    // bubble/stock/surface, chartEx, non-canonical combos, absolute
    // anchors) are skipped while their anchors still count toward
    // drawingIndex parity with the desktop sidecar.
    let sheetCharts: readonly SheetChartInfo[] | undefined
    try {
      const parsed = await parseSheetCharts(zip, worksheetPath, worksheetXml)
      if (parsed.length > 0) sheetCharts = parsed
    } catch (error) {
      if (!(error instanceof ChartReadError)) throw error
    }
    // Conditional-formatting read (EXCEL-024): parse the sheet's
    // <conditionalFormatting> sections into the canonical CfWireRule[]
    // (the Univer wire shape, dxf style pre-resolved through the
    // StylesheetReader). Fail closed PER SHEET — unrepresentable rules
    // (x14 extensions, time periods, unknown types, malformed sqref,
    // unresolvable dxf styling) surface NO cfRules plus cfLocked: true,
    // the workbook still opens, the browser refuses CF edits on the
    // sheet (a CF-dirty rewrite would silently drop what the model
    // cannot hold), and a no-op save preserves the file's XML
    // byte-for-byte.
    let cfRules: readonly CfWireRule[] | undefined
    let cfLocked: boolean | undefined
    try {
      const parsed = parseConditionalFormatting(
        worksheetXml,
        styleReader === null ? () => undefined : (dxfId) => styleReader.dxfAt(dxfId),
      )
      if (parsed.length > 0) cfRules = parsed
    } catch (error) {
      if (!(error instanceof CfReadError)) throw error
      cfLocked = true
    }
    sheets.push({
      id,
      name: decodedName,
      cells: parseWorksheetCells(worksheetXml, sharedStrings),
      ...(presentation.styles && Object.keys(presentation.styles).length > 0
        ? { styles: presentation.styles }
        : {}),
      ...(presentation.merges.length > 0 ? { merges: presentation.merges } : {}),
      ...(presentation.rowHeights && Object.keys(presentation.rowHeights).length > 0
        ? { rowHeights: presentation.rowHeights }
        : {}),
      ...(presentation.colWidths && Object.keys(presentation.colWidths).length > 0
        ? { colWidths: presentation.colWidths }
        : {}),
      ...(presentation.freeze ? { freeze: presentation.freeze } : {}),
      ...(filterState ? { filterState } : {}),
      ...(dvRules ? { dvRules } : {}),
      ...(sheetNotes ? { notes: sheetNotes } : {}),
      ...(sheetTables ? { tables: sheetTables } : {}),
      ...(sheetImages ? { images: sheetImages } : {}),
      ...(sheetCharts ? { charts: sheetCharts } : {}),
      ...(cfRules ? { cfRules } : {}),
      ...(cfLocked ? { cfLocked } : {}),
      ...(sheetProtection ? { sheetProtection } : {}),
    })
    sheetNamesById[id] = decodedName
  }
  if (sheets.length === 0) throw new Error('Workbook contains no readable worksheets.')
  // Workbook structure protection (EXCEL-020): parsed from workbook.xml the
  // same way the per-sheet state comes from the worksheet part. Absent
  // field = no <workbookProtection> element (byte-preserving no-op saves).
  const workbookProtection = parseWorkbookProtectionState(workbookXml)
  return {
    snapshot: {
      revision: 0,
      sheets,
      ...(workbookProtection ? { workbookProtection } : {}),
    },
    sheetNamesById,
  }
}

/**
 * Presentation pass over one worksheet: per-cell resolved formats (via the
 * StylesheetReader), merged ranges, custom row heights (points, 1-based row
 * keys) and custom column widths (px, column-label keys — the OOXML
 * character width is converted with the Calibri-11 default-font metric the
 * ecosystem uses, a display approximation; the file keeps the exact width).
 */
function parseWorksheetPresentation(
  worksheetXml: string,
  styleReader: StylesheetReader | null,
): {
  styles: Readonly<Record<string, CellFormatState>>
  merges: readonly string[]
  rowHeights: Readonly<Record<string, number>>
  colWidths: Readonly<Record<string, number>>
  freeze: { frozenRows: number; frozenColumns: number } | undefined
} {
  const styles: Record<string, CellFormatState> = {}
  if (styleReader) {
    const cellPattern = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
    let cellMatch: RegExpExecArray | null
    while ((cellMatch = cellPattern.exec(worksheetXml)) !== null) {
      const attrs = cellMatch[1] ?? ''
      const address = readXmlAttribute(attrs, 'r')
      const styleIndex = readXmlAttribute(attrs, 's')
      if (!address || styleIndex === undefined) continue
      const format = styleReader.formatAt(Number(styleIndex))
      if (format) styles[address] = format
    }
  }
  const merges: string[] = []
  const mergePattern = /<mergeCell\b[^>]*\bref="([^"]+)"/g
  let mergeMatch: RegExpExecArray | null
  while ((mergeMatch = mergePattern.exec(worksheetXml)) !== null) {
    merges.push(mergeMatch[1] ?? '')
  }
  const rowHeights: Record<string, number> = {}
  const rowPattern = /<row\b([^>]*)\/?>/g
  let rowMatch: RegExpExecArray | null
  while ((rowMatch = rowPattern.exec(worksheetXml)) !== null) {
    const attrs = rowMatch[1] ?? ''
    if (readXmlAttribute(attrs, 'customHeight') !== '1') continue
    const rowNumber = readXmlAttribute(attrs, 'r')
    const height = Number(readXmlAttribute(attrs, 'ht'))
    if (rowNumber && Number.isFinite(height) && height > 0) {
      rowHeights[rowNumber] = height
    }
  }
  const colWidths: Record<string, number> = {}
  const colPattern = /<col\b([^>]*)\/?>/g
  let colMatch: RegExpExecArray | null
  while ((colMatch = colPattern.exec(worksheetXml)) !== null) {
    const attrs = colMatch[1] ?? ''
    if (readXmlAttribute(attrs, 'customWidth') !== '1') continue
    const min = Number(readXmlAttribute(attrs, 'min'))
    const max = Number(readXmlAttribute(attrs, 'max'))
    const width = Number(readXmlAttribute(attrs, 'width'))
    if (!Number.isInteger(min) || !Number.isInteger(max) || !Number.isFinite(width)) continue
    if (min < 1 || max < min || max - min > 1024) continue
    const px = Math.round(width * 7 + 5)
    for (let column = min; column <= max; column++) {
      colWidths[columnLabel(column)] = px
    }
  }
  const freeze = parseFrozenPane(worksheetXml)
  return { styles, merges, rowHeights, colWidths, freeze }
}

/**
 * Parse the frozen-pane state from a worksheet's <sheetView><pane>.
 *
 * The OOXML <pane> element carries:
 *   - xSplit: number of frozen columns (0 when absent)
 *   - ySplit: number of frozen rows    (0 when absent)
 *   - state: "frozen" | "frozenSplit" | "split" (only "frozen" counts)
 *   - topLeftCell: the first scrollable cell (e.g. "A4" for 3 frozen rows)
 *
 * Only a pane with state="frozen" represents a real freeze. The engine's
 * applyPageSetupState writes the same shape (xlsx-freeze.test.ts), so this
 * parser is the read-side counterpart. Returns undefined when no frozen
 * pane is present (so the WorksheetState stays minimal).
 */
function parseFrozenPane(
  worksheetXml: string,
): { frozenRows: number; frozenColumns: number } | undefined {
  const paneMatch = /<pane\b([^>]*)\/?>/.exec(worksheetXml)
  if (!paneMatch) return undefined
  const attrs = paneMatch[1] ?? ''
  const state = readXmlAttribute(attrs, 'state')
  if (state !== 'frozen' && state !== 'frozenSplit') return undefined
  const ySplit = Number(readXmlAttribute(attrs, 'ySplit') ?? '0')
  const xSplit = Number(readXmlAttribute(attrs, 'xSplit') ?? '0')
  const frozenRows = Number.isInteger(ySplit) && ySplit > 0 ? ySplit : 0
  const frozenColumns = Number.isInteger(xSplit) && xSplit > 0 ? xSplit : 0
  if (frozenRows === 0 && frozenColumns === 0) return undefined
  return { frozenRows, frozenColumns }
}

/** 1-based column index → A1 column label (1 → "A", 27 → "AA"). */
function columnLabel(column: number): string {
  let label = ''
  let n = column
  while (n > 0) {
    const rem = (n - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    n = Math.floor((n - 1) / 26)
  }
  return label || 'A'
}

export async function inventoryXlsx(buffer: Buffer): Promise<readonly PackageEntry[]> {
  const zip = await loadSafeZip(buffer)
  const entries: PackageEntry[] = []
  let totalSize = 0
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue
    const bytes = await file.async('nodebuffer')
    totalSize += bytes.length
    if (totalSize > MAX_UNCOMPRESSED_BYTES) {
      throw new Error('Workbook exceeds the uncompressed size limit.')
    }
    entries.push({ path, size: bytes.length, sha256: sha256Hex(bytes) })
  }
  return entries.sort((left, right) => left.path.localeCompare(right.path))
}

export async function applyPlanToXlsx(
  source: Buffer,
  plan: ChangePlan,
  sheetNamesById: Readonly<Record<string, string>>,
): Promise<XlsxMutation> {
  const pkg = new PackageEditor(await createBufferEntrySource(source))
  const touchedEntries = new Set<string>()

  for (const rename of plan.sheetRenames) {
    const workbookPath = 'xl/workbook.xml'
    const workbookXml = await pkg.readText(workbookPath)
    const currentName = sheetNamesById[rename.sheetId]
    if (!currentName) throw new Error(`Missing XLSX sheet mapping for ${rename.sheetId}.`)
    if (currentName !== rename.before)
      throw new Error(`Sheet ${rename.sheetId} no longer has the expected name.`)
    pkg.write(workbookPath, replaceSheetName(workbookXml, currentName, rename.after))
    touchedEntries.add(workbookPath)
  }

  for (const change of plan.cellChanges) {
    const sheetName = sheetNamesById[change.sheetId]
    if (!sheetName) throw new Error(`Missing XLSX sheet mapping for ${change.sheetId}.`)
    const worksheetPath = await resolveWorksheetPath(pkg, sheetName)
    const worksheetXml = await pkg.readText(worksheetPath)
    const actualCell = parseCell(worksheetXml, change.address)
    if (!cellsEqual(actualCell, change.before)) {
      throw new Error(`${sheetName}!${change.address} no longer has the expected content.`)
    }
    pkg.write(worksheetPath, patchCell(worksheetXml, change.address, change.after))
    touchedEntries.add(worksheetPath)
  }

  return assembleWithJsZip(source, pkg.toPlan(touchedEntries))
}

/// Save path for user edits on streamed workbooks. Unlike applyPlanToXlsx
/// there is no per-cell before-check: the caller gates on the whole-file
/// SHA-256 recorded at open time, which the streamed originals came from.
export async function applyCellEditsToXlsx(
  source: Buffer,
  edits: readonly CellEdit[],
  structuralOps: readonly SheetStructuralOps[] = [],
  chartEdits: readonly WorkbookChartEdit[] = [],
  sheetPlan?: SheetEditPlan,
  filterStates: readonly SheetFilterState[] = [],
  hyperlinkEdits: readonly SheetHyperlinkEdits[] = [],
  cfStates: readonly SheetCfState[] = [],
  dvStates: readonly SheetDvState[] = [],
  sheetProtections: readonly SheetProtectionState[] = [],
  definedNamesState: DefinedNamesState | null = null,
  pageSetupStates: readonly SheetPageSetupState[] = [],
  noteStates: readonly SheetNoteState[] = [],
  formulaValues: readonly SheetFormulaValues[] = [],
  workbookProtectionState: { readonly lockStructure: boolean } | null = null,
  tableAdditions: readonly SheetTableAddition[] = [],
  /// EXCEL-022: web-supplied visual families. Appended at the end so the
  /// desktop translator's positional call sites are untouched.
  visualAdditions: readonly SheetVisualAddition[] = [],
  visualEdits: readonly WorkbookVisualEdit[] = [],
): Promise<XlsxMutation> {
  const plan = await planCellEditsToXlsx(
    await createBufferEntrySource(source),
    edits,
    structuralOps,
    chartEdits,
    sheetPlan,
    filterStates,
    hyperlinkEdits,
    cfStates,
    dvStates,
    sheetProtections,
    definedNamesState,
    visualAdditions,
    pageSetupStates,
    noteStates,
    tableAdditions,
    [],
    [],
    [],
    visualEdits,
    [],
    formulaValues,
    null,
    workbookProtectionState,
  )
  return assembleWithJsZip(source, plan)
}

export interface SheetNoteState {
  readonly sheetName: string
  readonly notes: readonly SheetNote[]
}

export interface SheetVisualAddition {
  readonly sheetName: string
  readonly anchor: DrawingAnchor
  readonly chart?: ChartAdd | undefined
  readonly shape?: ShapeAdd | undefined
  readonly image?: ImageAdd | undefined
}

export interface SheetTableAddition {
  readonly sheetName: string
  readonly area: TableArea
  readonly name: string
  readonly columnNames: readonly string[]
  readonly style?: string | undefined
  readonly bandedRows: boolean
}

export interface SheetPivotAddition {
  /// Sheet that receives the pivot output.
  readonly sheetName: string
  readonly sourceSheetName: string
  readonly sourceArea: TableArea
  readonly location: TableArea
  readonly name: string
  readonly fieldNames: readonly string[]
  /// Indices into fieldNames for the row dimension levels (outer → inner).
  readonly rowFieldIndices: readonly number[]
  readonly columnFieldIndex?: number | undefined
  /// Indices into fieldNames for report-filter (page) fields.
  readonly pageFieldIndices?: readonly number[] | undefined
  /// Deduplicated members of the level-0 row field; with multiple levels this
  /// equals rowLevelItems[0].
  readonly rowItems: readonly string[]
  /// Deduplicated member lists per row level (required for multi-level rows, see
  /// PivotAddition).
  readonly rowLevelItems?: readonly (readonly string[])[] | undefined
  /// Row-by-row layout of the output data rows (required for multi-level rows,
  /// see PivotAddition).
  readonly rowLines?: readonly PivotAddRowLine[] | undefined
  readonly columnItems?: readonly string[] | undefined
  /// Multi-level columns: column dimension field indices (outer first; takes
  /// precedence over columnFieldIndex when provided).
  readonly columnFieldIndices?: readonly number[] | undefined
  /// Deduplicated member lists per column level (required with ≥2 column levels,
  /// see PivotAddition).
  readonly colLevelItems?: readonly (readonly string[])[] | undefined
  /// Column-by-column layout of the output data columns (excluding the trailing
  /// grand-total column; required with ≥2 column levels).
  readonly colLines?: readonly PivotAddRowLine[] | undefined
  /// Grouping rules for dimension fields (date/numeric ranges), see PivotAddition.
  readonly groupings?: readonly PivotAddGrouping[] | undefined
  /// Value/label filters plus the filtered-out hidden members, see PivotAddition.
  readonly filters?: readonly PivotFilterDef[] | undefined
  readonly rowHiddenItems?: readonly (readonly number[])[] | undefined
  readonly colHiddenItems?: readonly (readonly number[])[] | undefined
  readonly values: readonly PivotValueSpec[]
}

export type { PivotRefreshUpdate } from './xlsx-pivot-expand'
export { PivotExpandError } from './xlsx-pivot-expand'
export type { SheetSparklineAddition } from './xlsx-sparkline'

/// Assembler-independent planning: computes the patched entry contents for a
/// set of edits without materializing the output archive.
export async function planCellEditsToXlsx(
  source: EntrySource,
  edits: readonly CellEdit[],
  structuralOps: readonly SheetStructuralOps[] = [],
  chartEdits: readonly WorkbookChartEdit[] = [],
  sheetPlan?: SheetEditPlan,
  filterStates: readonly SheetFilterState[] = [],
  hyperlinkEdits: readonly SheetHyperlinkEdits[] = [],
  cfStates: readonly SheetCfState[] = [],
  dvStates: readonly SheetDvState[] = [],
  sheetProtections: readonly SheetProtectionState[] = [],
  definedNamesState: DefinedNamesState | null = null,
  visualAdditions: readonly SheetVisualAddition[] = [],
  pageSetupStates: readonly SheetPageSetupState[] = [],
  noteStates: readonly SheetNoteState[] = [],
  tableAdditions: readonly SheetTableAddition[] = [],
  pivotAdditions: readonly SheetPivotAddition[] = [],
  pivotCacheRefreshPaths: readonly string[] = [],
  pivotRefreshUpdates: readonly PivotRefreshUpdate[] = [],
  visualEdits: readonly WorkbookVisualEdit[] = [],
  sparklineAdditions: readonly SheetSparklineAddition[] = [],
  formulaValues: readonly SheetFormulaValues[] = [],
  themeState: WorkbookThemeState | null = null,
  workbookProtectionState: { readonly lockStructure: boolean } | null = null,
  protectedRangeStates: readonly SheetProtectedRangesState[] = [],
): Promise<MutationPlan> {
  // A pending pivot pins final coordinates for its source and output; shifts
  // on either sheet, and sheet renames (worksheetSource@sheet), would desync
  // the recorded ranges. Fail closed, mirroring the table-add guard.
  if (pivotAdditions.length > 0) {
    if (sheetPlan !== undefined) {
      throw new Error(
        'A new pivot cannot be saved together with sheet management changes — ' +
          'save the pivot first.',
      )
    }
    const pivotSheets = new Set(
      pivotAdditions.flatMap((pivot) => [pivot.sheetName, pivot.sourceSheetName]),
    )
    if (structuralOps.some((sheet) => sheet.ops.length > 0 && pivotSheets.has(sheet.sheetName))) {
      throw new Error(
        'A new pivot cannot be saved together with row/column changes on its ' +
          'sheets — save the pivot first.',
      )
    }
  }
  // A pending table add pins final coordinates; row/column shifts on the same
  // sheet would desync the recorded range. The renderer saves before allowing
  // further structural work, so this is a defensive fail-closed check.
  if (tableAdditions.length > 0) {
    const tableSheets = new Set(tableAdditions.map((table) => table.sheetName))
    if (structuralOps.some((sheet) => sheet.ops.length > 0 && tableSheets.has(sheet.sheetName))) {
      throw new Error(
        'A new table cannot be saved together with row/column changes on its ' +
          'sheet — save the table first.',
      )
    }
  }
  // The defined-names snapshot carries model coordinates and file sheet
  // indices; replaying structural or sheet operations underneath it would
  // desync both. The renderer blocks the combination too.
  if (
    definedNamesState !== null &&
    (structuralOps.some((sheet) => sheet.ops.length > 0) || sheetPlan !== undefined)
  ) {
    throw new DefinedNameError(
      'Defined-name edits cannot be saved together with row/column or sheet ' +
        'changes — save one of them first.',
    )
  }
  const pkg = new PackageEditor(source)
  const touchedEntries = new Set<string>()

  // Added sheets get their parts up front so cell edits, structural ops, and
  // the cross-sheet scan below all see them. Duplicates are seeded from the
  // source sheet's part; their journaled edits replay on top like any other.
  const additions =
    sheetPlan === undefined
      ? []
      : await allocateAddedSheets(
          pkg,
          sheetPlan.additions.map((addition) => addition.name),
        )
  const additionPaths = new Map(additions.map((addition) => [addition.name, addition.path]))
  for (const [index, addition] of additions.entries()) {
    const sourceSheetName = sheetPlan?.additions[index]?.sourceSheetName
    if (sourceSheetName === undefined) {
      pkg.add(addition.path, buildWorksheetPartXml())
      continue
    }
    const sourcePath = await resolveWorksheetPath(pkg, sourceSheetName)
    if (!(await pkg.canPatch(sourcePath))) {
      throw new SheetEditError(
        `${sourcePath} is too large to load — duplicating "${sourceSheetName}" ` +
          'cannot be saved.',
      )
    }
    assertNoSheetScopedDefinedNames(await pkg.readText('xl/workbook.xml'), sourceSheetName)
    let cloneXml = sanitizeClonedWorksheetXml(await pkg.readText(sourcePath))
    const sourceRelsPath = sourcePath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
    if (await pkg.has(sourceRelsPath)) {
      const rels = prepareClonedSheetRels(await pkg.readText(sourceRelsPath), sourceSheetName)
      if (rels.droppedPrinterSettings) cloneXml = stripPageSetupRelIds(cloneXml)
      if (rels.relsXml !== null) {
        pkg.add(
          addition.path.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels'),
          rels.relsXml,
        )
      }
    }
    pkg.add(addition.path, cloneXml)
  }

  const sheetNames = new Set([
    ...edits.map((edit) => edit.sheetName),
    ...structuralOps.map((sheet) => sheet.sheetName),
    ...filterStates.map((state) => state.sheetName),
    ...hyperlinkEdits.map((sheet) => sheet.sheetName),
    ...cfStates.map((state) => state.sheetName),
    ...dvStates.map((state) => state.sheetName),
    ...sheetProtections.map((state) => state.sheetName),
    ...pageSetupStates.map((state) => state.sheetName),
    ...protectedRangeStates.map((state) => state.sheetName),
  ])
  const worksheetXmls = new Map<string, string>()
  const worksheetPaths = new Map<string, string>()
  for (const sheetName of sheetNames) {
    const worksheetPath =
      additionPaths.get(sheetName) ?? (await resolveWorksheetPath(pkg, sheetName))
    worksheetPaths.set(sheetName, worksheetPath)
    worksheetXmls.set(sheetName, await pkg.readText(worksheetPath))
  }

  // Pivot layout expansion: conflict-check and update pivotTableDefinition
  // location refs before any cell edits are applied, so the worksheetXml
  // seen by worksheetHasContentInArea reflects the pre-edit state.
  if (pivotRefreshUpdates.length > 0) {
    // The renderer only carries the sheet name; resolve it to a part path here
    // before expanding.
    const resolvedUpdates: PivotRefreshUpdate[] = []
    for (const update of pivotRefreshUpdates) {
      const worksheetPath =
        update.worksheetPath ??
        (update.sheetName !== undefined
          ? await resolveWorksheetPath(pkg, update.sheetName)
          : undefined)
      if (worksheetPath === undefined) {
        throw new Error('A pivot refresh update needs a worksheetPath or sheetName.')
      }
      resolvedUpdates.push({ ...update, worksheetPath })
    }
    await applyPivotLayoutExpansions(pkg, resolvedUpdates, touchedEntries)
  }

  // Structural operations replay first: journaled cell edits are already in
  // the post-operation coordinate space. Qualified references from other
  // sheets, defined names, and chart series shift along with the edited sheet.
  const workbookPath = 'xl/workbook.xml'
  const originalWorkbookXml = await pkg.readText(workbookPath)
  let workbookXml = originalWorkbookXml
  for (const { sheetName, ops } of structuralOps) {
    if (ops.length === 0) continue
    worksheetXmls.set(
      sheetName,
      applyStructuralOps(worksheetXmls.get(sheetName) ?? '', ops, sheetName),
    )
    const editedPath = worksheetPaths.get(sheetName)
    if (editedPath !== undefined) {
      await shiftAnchoredSheetParts(
        pkg,
        editedPath,
        worksheetXmls.get(sheetName) ?? '',
        ops,
        touchedEntries,
      )
    }
    const nameByPath = new Map([...worksheetPaths].map(([name, path]) => [path, name]))
    for (const path of await pkg.paths()) {
      const isOtherSheet =
        path.startsWith('xl/worksheets/') && path.endsWith('.xml') && path !== editedPath
      const isChart = path.startsWith('xl/charts/') && path.endsWith('.xml')
      if (!isOtherSheet && !isChart) continue
      // A sheet with its own pending edits lives in worksheetXmls — shift that
      // copy, or the final write-back would overwrite this pass.
      const trackedName = nameByPath.get(path)
      if (trackedName === undefined && !(await pkg.canPatch(path))) {
        // Too large to rewrite: safe to leave alone unless it references the
        // shifted sheet — an unshifted qualified reference would corrupt it.
        if (await pkg.containsText(path, sheetName)) {
          throw new Error(
            `${path} references "${sheetName}" but is too large to rewrite — ` +
              'this structural change cannot be saved.',
          )
        }
        continue
      }
      const xml =
        trackedName !== undefined
          ? (worksheetXmls.get(trackedName) ?? '')
          : await pkg.readText(path)
      const shifted = isChart
        ? shiftChartReferences(xml, sheetName, ops)
        : shiftCrossSheetFormulas(xml, sheetName, ops)
      if (shifted === xml) continue
      if (trackedName !== undefined) {
        worksheetXmls.set(trackedName, shifted)
      } else {
        pkg.write(path, shifted)
        touchedEntries.add(path)
      }
    }
    const shiftedWorkbook = shiftDefinedNames(workbookXml, sheetName, ops)
    if (shiftedWorkbook !== workbookXml) {
      workbookXml = shiftedWorkbook
      pkg.write(workbookPath, workbookXml)
      touchedEntries.add(workbookPath)
    }
  }

  const editsBySheet = new Map<string, CellEdit[]>()
  for (const edit of edits) {
    const sheetEdits = editsBySheet.get(edit.sheetName) ?? []
    sheetEdits.push(edit)
    editsBySheet.set(edit.sheetName, sheetEdits)
  }
  let stylesheet: StylesheetEditor | null = null
  const stylesPath = 'xl/styles.xml'
  if (edits.some((edit) => edit.style !== undefined) || cfStates.length > 0) {
    if (!(await pkg.has(stylesPath))) await addDefaultStylesheet(pkg, touchedEntries)
    stylesheet = new StylesheetEditor(await pkg.readText(stylesPath))
  }
  for (const [sheetName, sheetEdits] of editsBySheet) {
    let worksheetXml = worksheetXmls.get(sheetName) ?? ''
    for (const edit of sheetEdits) {
      const address = toA1Address(edit.row, edit.column)
      let styleOverride: number | undefined
      if (edit.styleReset) {
        styleOverride = edit.style && stylesheet ? stylesheet.resolveStyle(0, edit.style) : 0
      } else if (edit.style && stylesheet) {
        const baseIndex = readCellStyleIndex(worksheetXml, address) ?? 0
        styleOverride = stylesheet.resolveStyle(baseIndex, edit.style)
      }
      worksheetXml = edit.writeValue
        ? patchCellKeepingStyle(worksheetXml, address, edit.cell, styleOverride, edit.rich)
        : patchCellStyleOnly(worksheetXml, address, styleOverride)
    }
    worksheetXmls.set(sheetName, expandWorksheetDimensionToCells(worksheetXml))
  }
  // Recalculated formula results: refresh each formula cell's cached
  // <v> while leaving its <f> alone. Applied after the value edits so a cell the
  // user turned into a literal keeps that literal.
  for (const sheet of formulaValues) {
    if (sheet.cells.length === 0) continue
    let worksheetXml = worksheetXmls.get(sheet.sheetName)
    if (worksheetXml === undefined) continue
    for (const cell of sheet.cells) {
      worksheetXml = patchFormulaCachedValue(
        worksheetXml,
        toA1Address(cell.row, cell.column),
        cell.value,
      )
    }
    worksheetXmls.set(sheet.sheetName, worksheetXml)
  }
  // Hyperlink edits carry final coordinates, so they apply after the
  // structural replay; the rels sibling is created or rewritten alongside.
  for (const sheet of hyperlinkEdits) {
    if (sheet.edits.length === 0) continue
    const worksheetPath = worksheetPaths.get(sheet.sheetName)
    const worksheetXml = worksheetXmls.get(sheet.sheetName)
    if (!worksheetPath || worksheetXml === undefined) continue
    const relsPath = worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
    const relsExisted = await pkg.has(relsPath)
    const relsXml = relsExisted ? await pkg.readText(relsPath) : null
    const patch = applyHyperlinkEdits(worksheetXml, relsXml, sheet.edits)
    worksheetXmls.set(sheet.sheetName, ensureRelationshipNamespace(patch.worksheetXml))
    if (patch.relsChanged && patch.relsXml !== null) {
      if (relsExisted) {
        pkg.write(relsPath, patch.relsXml)
      } else {
        pkg.add(relsPath, patch.relsXml)
      }
      touchedEntries.add(relsPath)
    }
  }

  // Conditional formatting is declarative like filters: every section on a
  // dirty sheet is rewritten from the snapshot (final coordinates).
  for (const state of cfStates) {
    const worksheetXml = worksheetXmls.get(state.sheetName)
    if (worksheetXml === undefined || stylesheet === null) continue
    worksheetXmls.set(state.sheetName, applyCfRules(worksheetXml, state.rules, stylesheet))
  }

  // Data validation follows the same declarative rewrite.
  for (const state of dvStates) {
    const worksheetXml = worksheetXmls.get(state.sheetName)
    if (worksheetXml === undefined) continue
    worksheetXmls.set(state.sheetName, applyDvRules(worksheetXml, state.rules))
  }

  for (const state of sheetProtections) {
    const worksheetXml = worksheetXmls.get(state.sheetName)
    if (worksheetXml === undefined) continue
    worksheetXmls.set(state.sheetName, applySheetProtection(worksheetXml, state.protected))
  }

  // Allow-edit ranges are declarative snapshots, like filters.
  for (const state of protectedRangeStates) {
    const worksheetXml = worksheetXmls.get(state.sheetName)
    if (worksheetXml === undefined) continue
    worksheetXmls.set(state.sheetName, applyProtectedRanges(worksheetXml, state.ranges))
  }

  // Page Layout settings merge attribute-by-attribute; untouched print
  // settings in the file stay verbatim.
  for (const state of pageSetupStates) {
    const worksheetXml = worksheetXmls.get(state.sheetName)
    if (worksheetXml === undefined) continue
    worksheetXmls.set(state.sheetName, applyPageSetupState(worksheetXml, state))
  }

  // Filter snapshots run after structural replay and cell edits, so their
  // coordinates and row set match the sheet's final content.
  for (const state of filterStates) {
    const worksheetXml = worksheetXmls.get(state.sheetName)
    if (worksheetXml === undefined) continue
    worksheetXmls.set(state.sheetName, applyFilterState(worksheetXml, state))
  }

  for (const [sheetName, worksheetXml] of worksheetXmls) {
    const worksheetPath = worksheetPaths.get(sheetName)
    if (!worksheetPath) continue
    if (additionPaths.has(sheetName)) {
      pkg.add(worksheetPath, worksheetXml)
    } else {
      pkg.write(worksheetPath, worksheetXml)
    }
    touchedEntries.add(worksheetPath)
  }
  if (stylesheet?.changed) {
    pkg.write(stylesPath, stylesheet.serialize())
    touchedEntries.add(stylesPath)
  }

  // Chart edits run after structural shifts so they patch the already-shifted
  // chart XML.
  for (const chartEdit of chartEdits) {
    const chartXml = await pkg.readText(chartEdit.chartPath)
    pkg.write(chartEdit.chartPath, applyChartEdit(chartXml, chartEdit))
    touchedEntries.add(chartEdit.chartPath)
  }

  // Edits to file visuals run before any new anchors are appended, so the
  // sidecar's drawingIndex still matches the file's document order.
  if (visualEdits.length > 0) {
    await applyVisualEdits(pkg, visualEdits, touchedEntries)
  }

  // New visuals run after the worksheet XML flush above so the drawing
  // element lands on the final sheet content.
  let addedVisuals: readonly AddedVisualLocator[] | undefined
  if (visualAdditions.length > 0) {
    const resolved = []
    for (const addition of visualAdditions) {
      resolved.push({
        worksheetPath:
          additionPaths.get(addition.sheetName) ??
          (await resolveWorksheetPath(pkg, addition.sheetName)),
        anchor: addition.anchor,
        chart: addition.chart,
        shape: addition.shape,
        image: addition.image,
      })
    }
    addedVisuals = await applyVisualAdditions(pkg, resolved, touchedEntries)
  }

  // Note snapshots replace each dirty sheet's whole comment set; they run
  // after the worksheet flush so the legacyDrawing element lands on the
  // final sheet XML.
  for (const state of noteStates) {
    const worksheetPath =
      additionPaths.get(state.sheetName) ?? (await resolveWorksheetPath(pkg, state.sheetName))
    await applySheetNotes(pkg, worksheetPath, state.notes, touchedEntries)
  }

  // Recomputed pivots: their output cells were saved as ordinary edits above;
  // flag the caches so Excel rebuilds them from the same source on open.
  for (const cachePath of pivotCacheRefreshPaths) {
    const cacheXml = await pkg.readText(cachePath)
    pkg.write(cachePath, setPivotRefreshOnLoad(cacheXml))
    touchedEntries.add(cachePath)
  }

  // New tables also run on the flushed worksheet XML: the <tableParts>
  // element and overlap checks see the final sheet content.
  if (tableAdditions.length > 0) {
    const resolvedTables = []
    for (const addition of tableAdditions) {
      resolvedTables.push({
        worksheetPath:
          additionPaths.get(addition.sheetName) ??
          (await resolveWorksheetPath(pkg, addition.sheetName)),
        area: addition.area,
        name: addition.name,
        columnNames: addition.columnNames,
        style: addition.style,
        bandedRows: addition.bandedRows,
      })
    }
    await applyTableAdditions(pkg, resolvedTables, touchedEntries)
  }

  // New sparklines also run on the flushed worksheet XML (extLst is the
  // worksheet's last element, after tableParts).
  if (sparklineAdditions.length > 0) {
    const groupsBySheet = new Map<string, SparklineGroupAdd[]>()
    for (const { sheetName, ...group } of sparklineAdditions) {
      const groups = groupsBySheet.get(sheetName) ?? []
      groups.push(group)
      groupsBySheet.set(sheetName, groups)
    }
    for (const [sheetName, groups] of groupsBySheet) {
      const worksheetPath =
        additionPaths.get(sheetName) ?? (await resolveWorksheetPath(pkg, sheetName))
      pkg.write(worksheetPath, applySparklineAdditions(await pkg.readText(worksheetPath), groups))
      touchedEntries.add(worksheetPath)
    }
  }

  // Any worksheet edit can invalidate the calculation chain — not just
  // structural shifts: overwriting a formula cell with a literal leaves a
  // calcChain entry pointing at a cell with no <f>, which Excel repairs with
  // a scary prompt. calcChain is a pure recalculation-order cache, so
  // drop it (with its content-type and relationship) whenever this save wrote
  // any worksheet part and let Excel rebuild it on open. Sheet set changes
  // are kept as an extra trigger (a removal-only save may touch no part).
  const sheetSetChanged =
    sheetPlan !== undefined &&
    (sheetPlan.additions.length > 0 ||
      sheetPlan.removals.length > 0 ||
      sheetPlan.orderChanged === true)
  const worksheetTouched = [...touchedEntries].some((path) => path.startsWith('xl/worksheets/'))
  if ((worksheetTouched || sheetSetChanged) && (await pkg.has('xl/calcChain.xml'))) {
    pkg.remove('xl/calcChain.xml')
    const contentTypesPath = '[Content_Types].xml'
    const contentTypes = await pkg.readText(contentTypesPath)
    const strippedTypes = contentTypes.replace(
      /<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/,
      '',
    )
    if (strippedTypes !== contentTypes) {
      pkg.write(contentTypesPath, strippedTypes)
      touchedEntries.add(contentTypesPath)
    }
    const workbookRelsPath = 'xl/_rels/workbook.xml.rels'
    const workbookRels = await pkg.readText(workbookRelsPath)
    const strippedRels = workbookRels.replace(
      /<Relationship\b[^>]*Target="calcChain\.xml"[^>]*\/>/,
      '',
    )
    if (strippedRels !== workbookRels) {
      pkg.write(workbookRelsPath, strippedRels)
      touchedEntries.add(workbookRelsPath)
    }
  }

  // Sheet-level surgery runs last: every worksheet and chart part is already
  // in its final content, so rename rewrites and removal guards see the
  // saved state.
  if (sheetPlan !== undefined) {
    workbookXml = await applySheetPlanToPackage(
      pkg,
      sheetPlan,
      additions,
      workbookXml,
      touchedEntries,
    )
  }

  if (definedNamesState !== null) {
    workbookXml = applyDefinedNamesState(workbookXml, definedNamesState)
  }

  if (workbookProtectionState !== null) {
    workbookXml = applyWorkbookProtection(workbookXml, workbookProtectionState.lockStructure)
  }

  if (themeState !== null) {
    const themePath = 'xl/theme/theme1.xml'
    if (!(await pkg.has(themePath))) {
      throw new Error('The workbook has no theme part — theme changes cannot be saved.')
    }
    pkg.write(themePath, applyThemeState(await pkg.readText(themePath), themeState))
    touchedEntries.add(themePath)
  }

  // Print areas / title rows are sheet-scoped _xlnm names; they apply to the
  // final workbook.xml (post sheet-plan, post defined-names rewrite — which
  // keeps _xlnm entries verbatim).
  const printAreas = pageSetupStates
    .filter((state) => state.printArea !== undefined || state.printTitles !== undefined)
    .map((state) => ({
      sheetName: state.sheetName,
      printArea: state.printArea,
      printTitles: state.printTitles,
    }))
  if (printAreas.length > 0) {
    workbookXml = applyPrintAreas(workbookXml, printAreas)
  }

  // New pivots run last: their worksheet rels ride the package overlay, and
  // the <pivotCaches> entry lands on the final workbook XML string.
  if (pivotAdditions.length > 0) {
    const resolvedPivots = []
    for (const addition of pivotAdditions) {
      resolvedPivots.push({
        worksheetPath:
          additionPaths.get(addition.sheetName) ??
          (await resolveWorksheetPath(pkg, addition.sheetName)),
        sourceSheetName: addition.sourceSheetName,
        sourceArea: addition.sourceArea,
        location: addition.location,
        name: addition.name,
        fieldNames: addition.fieldNames,
        rowFieldIndices: addition.rowFieldIndices,
        columnFieldIndex: addition.columnFieldIndex,
        pageFieldIndices: addition.pageFieldIndices,
        rowItems: addition.rowItems,
        rowLevelItems: addition.rowLevelItems,
        rowLines: addition.rowLines,
        columnItems: addition.columnItems,
        columnFieldIndices: addition.columnFieldIndices,
        colLevelItems: addition.colLevelItems,
        colLines: addition.colLines,
        groupings: addition.groupings,
        filters: addition.filters,
        rowHiddenItems: addition.rowHiddenItems,
        colHiddenItems: addition.colHiddenItems,
        values: addition.values,
      })
    }
    workbookXml = await applyPivotAdditions(pkg, resolvedPivots, workbookXml, touchedEntries)
  }

  // Excel trusts cached formula values on open, so formulas that depend on an
  // edited cell would show stale results without a forced recalculation.
  workbookXml = ensureFullCalcOnLoad(workbookXml)
  if (workbookXml !== originalWorkbookXml) {
    pkg.write(workbookPath, workbookXml)
    touchedEntries.add(workbookPath)
  }

  const plan = pkg.toPlan(touchedEntries)
  return addedVisuals === undefined ? plan : { ...plan, addedVisuals }
}

/// Assigns non-colliding part paths, sheetId attributes, and relationship
/// ids to the sheets being added.
async function allocateAddedSheets(
  pkg: PackageEditor,
  names: readonly string[],
): Promise<SheetAllocation[]> {
  if (names.length === 0) return []
  const workbookXml = await pkg.readText('xl/workbook.xml')
  const relationshipsXml = await pkg.readText('xl/_rels/workbook.xml.rels')
  let nextPartNumber = 1
  for (const path of await pkg.paths()) {
    const match = /^xl\/worksheets\/sheet([0-9]+)\.xml$/.exec(path)
    if (match) nextPartNumber = Math.max(nextPartNumber, Number(match[1]) + 1)
  }
  const nextSheetId = maxSheetIdInWorkbook(workbookXml) + 1
  const nextRelationshipId = maxRelationshipId(relationshipsXml) + 1
  return names.map((name, index) => ({
    name,
    path: `xl/worksheets/sheet${nextPartNumber + index}.xml`,
    sheetId: nextSheetId + index,
    relationshipId: `rId${nextRelationshipId + index}`,
  }))
}

async function applySheetPlanToPackage(
  pkg: PackageEditor,
  plan: SheetEditPlan,
  additions: readonly SheetAllocation[],
  workbookXml: string,
  touchedEntries: Set<string>,
): Promise<string> {
  const elements = parseSheetElements(workbookXml)
  const originalNames = elements.map((element) => element.name)
  const originalSet = new Set(originalNames)
  const renameByOriginal = new Map(plan.renames.map((rename) => [rename.sheetName, rename.newName]))
  for (const rename of plan.renames) {
    validateSheetName(rename.newName)
    if (!originalSet.has(rename.sheetName)) {
      throw new SheetEditError(`Sheet "${rename.sheetName}" was not found in the workbook.`)
    }
  }
  for (const addition of additions) validateSheetName(addition.name)
  for (const removal of plan.removals) {
    if (!originalSet.has(removal)) {
      throw new SheetEditError(`Sheet "${removal}" was not found in the workbook.`)
    }
    if (renameByOriginal.has(removal)) {
      throw new SheetEditError(`Sheet "${removal}" cannot be renamed and removed in one save.`)
    }
  }
  const finalNames = [
    ...originalNames
      .filter((name) => !plan.removals.includes(name))
      .map((name) => renameByOriginal.get(name) ?? name),
    ...additions.map((addition) => addition.name),
  ]
  if (new Set(finalNames).size !== finalNames.length) {
    throw new SheetEditError('Two sheets would end up with the same name — aborted.')
  }

  const removalPaths = new Map<string, string>()
  for (const removal of plan.removals) {
    removalPaths.set(removal, await resolveWorksheetPath(pkg, removal))
  }
  const removedPathSet = new Set(removalPaths.values())
  const packagePaths = await pkg.paths()
  const survivingWorksheetPaths = packagePaths.filter(
    (path) => /^xl\/worksheets\/[^/]+\.xml$/.test(path) && !removedPathSet.has(path),
  )
  const chartPaths = packagePaths.filter(
    (path) => path.startsWith('xl/charts/') && path.endsWith('.xml'),
  )
  const pivotCacheDefinitionPaths = packagePaths.filter((path) =>
    /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(path),
  )

  // Satellite parts owned by the removed sheets — drawings with their images
  // and charts, legacy VML, comments, tables — die with the sheet. The
  // closure walks each owned part's own relationships; unsupported sheet
  // relationships (pivot tables, slicers) fail closed inside
  // classifyRemovedSheetRels.
  const removalRelsPaths = new Map<string, string>()
  const ownedPartsByRemoval = new Map<string, ReadonlySet<string>>()
  const removedOwnedParts = new Set<string>()
  for (const removal of plan.removals) {
    const removalPath = removalPaths.get(removal) ?? ''
    const relsPath = relsPathFor(removalPath)
    removalRelsPaths.set(removal, relsPath)
    const owned = new Set<string>()
    if (await pkg.has(relsPath)) {
      const targets = classifyRemovedSheetRels(await pkg.readText(relsPath), removal)
      const queue = targets.map((target) => resolveRelTarget(removalPath, target))
      while (queue.length > 0) {
        const part = queue.pop() as string
        if (owned.has(part) || !(await pkg.has(part))) continue
        owned.add(part)
        const childRelsPath = relsPathFor(part)
        if (!(await pkg.has(childRelsPath))) continue
        owned.add(childRelsPath)
        for (const entry of parseRelationships(await pkg.readText(childRelsPath))) {
          if (!entry.external) queue.push(resolveRelTarget(part, entry.target))
        }
      }
    }
    ownedPartsByRemoval.set(removal, owned)
    for (const part of owned) removedOwnedParts.add(part)
  }

  // A part in the closure may also be referenced from a part that survives —
  // an image placed on two sheets shares one xl/media entry. Walk every
  // surviving rels part and pull such targets (with their own subtrees) back
  // out of the removal set.
  const dyingRelsParts = new Set([
    ...removalRelsPaths.values(),
    ...[...removedOwnedParts].filter((part) => part.endsWith('.rels')),
  ])
  const keepQueue: string[] = []
  for (const relsPath of packagePaths) {
    if (!relsPath.endsWith('.rels') || dyingRelsParts.has(relsPath)) continue
    const owner = partPathForRels(relsPath)
    if (removedPathSet.has(owner)) continue
    for (const entry of parseRelationships(await pkg.readText(relsPath))) {
      if (entry.external) continue
      const target = resolveRelTarget(owner, entry.target)
      if (removedOwnedParts.has(target)) keepQueue.push(target)
    }
  }
  while (keepQueue.length > 0) {
    const part = keepQueue.pop() as string
    if (!removedOwnedParts.delete(part)) continue
    const childRelsPath = relsPathFor(part)
    if (removedOwnedParts.delete(childRelsPath)) {
      for (const entry of parseRelationships(await pkg.readText(childRelsPath))) {
        if (!entry.external) keepQueue.push(resolveRelTarget(part, entry.target))
      }
    }
  }

  // Removals fail closed while every reference to the sheet is still intact.
  const removedLocalIds = new Set(
    originalNames.flatMap((name, index) => (plan.removals.includes(name) ? [index] : [])),
  )
  for (const removal of plan.removals) {
    for (const path of survivingWorksheetPaths) {
      if (!(await pkg.canPatch(path))) {
        if (await pkg.containsText(path, removal)) {
          throw new SheetEditError(
            `Another sheet's formulas reference "${removal}" — deleting it is not allowed.`,
          )
        }
        continue
      }
      if (worksheetReferencesSheet(await pkg.readText(path), removal)) {
        throw new SheetEditError(
          `Another sheet's formulas reference "${removal}" — deleting it is not allowed.`,
        )
      }
    }
    for (const chartPath of chartPaths) {
      // Charts dying with a removed sheet chart that sheet's own data; only
      // surviving charts can hold a genuinely dangling reference.
      if (removedOwnedParts.has(chartPath)) continue
      if (chartReferencesSheet(await pkg.readText(chartPath), removal)) {
        throw new SheetEditError(
          `A chart reads its data from "${removal}" — deleting it is not allowed.`,
        )
      }
    }
    if (definedNamesReferenceSheet(workbookXml, removal, removedLocalIds)) {
      throw new SheetEditError(
        `A workbook defined name references "${removal}" — deleting it is not allowed.`,
      )
    }
    // A pivot hosted on a surviving sheet may read its source rows from the
    // removed sheet; only the pivotCacheDefinition records that link
    // (cacheSource/worksheetSource@sheet), so the hosting-sheet pivotTable
    // fail-close in classifyRemovedSheetRels cannot catch it.
    for (const cachePath of pivotCacheDefinitionPaths) {
      if (pivotCacheReadsFromSheet(await pkg.readText(cachePath), removal)) {
        throw new SheetEditError(
          `A pivot table reads its source data from "${removal}" — deleting it is not allowed.`,
        )
      }
    }
    // Structured references (DecoTable[Amount]) into a removed table are not
    // sheet-qualified, so the sheet-name checks above cannot catch them.
    // Excel treats table names as case-insensitive, so the scan must too;
    // entries too large to patch fall back to the sidecar's exact-case scan.
    for (const part of ownedPartsByRemoval.get(removal) ?? []) {
      if (!removedOwnedParts.has(part) || !/^xl\/tables\/[^/]+\.xml$/.test(part)) continue
      const name = tableDisplayName(await pkg.readText(part))
      if (name === undefined) continue
      const needle = `${name}[`
      const needleLower = needle.toLowerCase()
      for (const path of survivingWorksheetPaths) {
        const referenced = (await pkg.canPatch(path))
          ? (await pkg.readText(path)).toLowerCase().includes(needleLower)
          : await pkg.containsText(path, needle)
        if (referenced) {
          throw new SheetEditError(
            `Another sheet's formulas use table "${name}" on "${removal}" — deleting it is not allowed.`,
          )
        }
      }
      // Defined names scoped to a removed sheet die with it (matching the
      // sheet-name check above), so only surviving names can block.
      if (definedNamesUseToken(workbookXml, needle, removedLocalIds)) {
        throw new SheetEditError(
          `A workbook defined name uses table "${name}" on "${removal}" — deleting it is not allowed.`,
        )
      }
    }
  }

  for (const removal of plan.removals) {
    const relsPath = removalRelsPaths.get(removal) ?? ''
    if (await pkg.has(relsPath)) pkg.remove(relsPath)
    pkg.remove(removalPaths.get(removal) ?? '')
  }
  for (const part of removedOwnedParts) pkg.remove(part)

  // Renames rewrite every qualified reference in the surviving parts.
  for (const rename of plan.renames) {
    for (const path of survivingWorksheetPaths) {
      if (!(await pkg.canPatch(path))) {
        if (await pkg.containsText(path, rename.sheetName)) {
          throw new SheetEditError(
            `${path} references "${rename.sheetName}" but is too large to rewrite — ` +
              'renaming this sheet cannot be saved.',
          )
        }
        continue
      }
      const xml = await pkg.readText(path)
      const renamed = renameSheetReferencesInWorksheet(xml, rename.sheetName, rename.newName)
      if (renamed !== xml) {
        pkg.write(path, renamed)
        touchedEntries.add(path)
      }
    }
    for (const chartPath of chartPaths) {
      // Charts cascade-deleted with a removed sheet are already gone from the
      // package by this point — reading them would throw.
      if (removedOwnedParts.has(chartPath)) continue
      const xml = await pkg.readText(chartPath)
      const renamed = renameSheetReferencesInChart(xml, rename.sheetName, rename.newName)
      if (renamed !== xml) {
        pkg.write(chartPath, renamed)
        touchedEntries.add(chartPath)
      }
    }
    // Pivot caches sourced from the renamed sheet keep working only if their
    // worksheetSource@sheet follows the rename.
    for (const cachePath of pivotCacheDefinitionPaths) {
      const xml = await pkg.readText(cachePath)
      const renamed = renameSheetInPivotCacheSource(xml, rename.sheetName, rename.newName)
      if (renamed !== xml) {
        pkg.write(cachePath, renamed)
        touchedEntries.add(cachePath)
      }
    }
    workbookXml = renameSheetReferencesInDefinedNames(workbookXml, rename.sheetName, rename.newName)
  }

  const relationshipsPath = 'xl/_rels/workbook.xml.rels'
  const originalRelationships = await pkg.readText(relationshipsPath)
  let relationshipsXml = originalRelationships
  for (const removal of plan.removals) {
    const relationshipId = elements.find((element) => element.name === removal)?.relationshipId
    if (relationshipId) relationshipsXml = removeRelationshipById(relationshipsXml, relationshipId)
  }
  for (const addition of additions) {
    relationshipsXml = addWorksheetRelationship(
      relationshipsXml,
      addition.relationshipId,
      addition.path.replace(/^xl\//, ''),
    )
  }
  if (relationshipsXml !== originalRelationships) {
    pkg.write(relationshipsPath, relationshipsXml)
    touchedEntries.add(relationshipsPath)
  }

  const contentTypesPath = '[Content_Types].xml'
  const originalContentTypes = await pkg.readText(contentTypesPath)
  let contentTypesXml = originalContentTypes
  for (const removal of plan.removals) {
    contentTypesXml = removePartOverride(contentTypesXml, removalPaths.get(removal) ?? '')
  }
  for (const part of removedOwnedParts) {
    contentTypesXml = removePartOverride(contentTypesXml, part)
  }
  for (const addition of additions) {
    contentTypesXml = addWorksheetOverride(contentTypesXml, addition.path)
  }
  if (contentTypesXml !== originalContentTypes) {
    pkg.write(contentTypesPath, contentTypesXml)
    touchedEntries.add(contentTypesPath)
  }

  return applySheetPlanToWorkbookXml(workbookXml, plan, additions)
}

/// Fails closed when a mutation altered any package entry it did not intend
/// to touch, or dropped/created entries beyond the declared removals and
/// additions.
export function assertOnlyTouchedEntriesChanged(mutation: XlsxMutation): void {
  const touched = new Set(mutation.touchedEntries)
  const removed = new Set(mutation.removedEntries)
  const added = new Set(mutation.addedEntries)
  const before = new Set(mutation.beforeEntries.map((entry) => entry.path))
  const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
  if (mutation.beforeEntries.length !== mutation.afterEntries.length + removed.size - added.size) {
    throw new Error('Saving would change the workbook package structure — aborted.')
  }
  for (const entry of mutation.beforeEntries) {
    const afterHash = after.get(entry.path)
    if (afterHash === undefined) {
      if (removed.has(entry.path)) continue
      throw new Error(`Saving would drop ${entry.path} — aborted.`)
    }
    if (removed.has(entry.path)) {
      throw new Error(`Saving should have removed ${entry.path} but did not — aborted.`)
    }
    if (!touched.has(entry.path) && afterHash !== entry.sha256) {
      throw new Error(`Saving would unexpectedly modify ${entry.path} — aborted.`)
    }
  }
  for (const path of added) {
    if (before.has(path)) {
      throw new Error(`Saving should have created ${path} but it already existed — aborted.`)
    }
    if (!after.has(path)) {
      throw new Error(`Saving should have created ${path} but did not — aborted.`)
    }
  }
  for (const entry of mutation.afterEntries) {
    if (!before.has(entry.path) && !added.has(entry.path)) {
      throw new Error(`Saving would unexpectedly create ${entry.path} — aborted.`)
    }
  }
}

export function toA1Address(row: number, column: number): string {
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(column) || column < 0) {
    throw new Error(`Invalid cell coordinates: ${row},${column}`)
  }
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return `${letters}${row + 1}`
}

// NOTE (Increment 3F): The runtime filesystem functions (syncFileBestEffort,
// writeXlsxAtomically, mutateXlsxFile, sha256) were moved to
// packages/platform-electron/src/capabilities/xlsx-file-ops.ts because they
// use node:crypto, node:fs/promises, node:path — the xlsx-gateway package
// must be pure (ZERO node:* imports).

async function loadSafeZip(buffer: Buffer): Promise<JSZip> {
  const zip = await JSZip.loadAsync(buffer, { checkCRC32: true })
  const paths = Object.keys(zip.files)
  if (paths.length > MAX_ENTRY_COUNT) throw new Error('Workbook contains too many ZIP entries.')
  if (paths.some((path) => path.startsWith('/') || path.split('/').includes('..'))) {
    throw new Error('Workbook contains an unsafe ZIP path.')
  }
  return zip
}

async function readTextEntry(zip: JSZip, path: string): Promise<string> {
  const entry = zip.file(path)
  if (!entry) throw new Error(`Workbook is missing ${path}.`)
  return entry.async('text')
}

/// Drawing anchors and table ranges live in sibling parts wired through the
/// worksheet rels; they must shift with the same structural op batch or the
/// sheet's visuals and tables would drift.
async function shiftAnchoredSheetParts(
  pkg: PackageEditor,
  worksheetPath: string,
  worksheetXml: string,
  ops: readonly StructuralOp[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (!ops.some(isShiftingOp)) return
  const parts: Array<{ relId: string; kind: 'drawing' | 'table' }> = []
  const drawingRelId = /<drawing\b[^>]*\br:id="([^"]+)"/.exec(worksheetXml)?.[1]
  if (drawingRelId !== undefined) parts.push({ relId: drawingRelId, kind: 'drawing' })
  for (const match of worksheetXml.matchAll(/<tablePart\b[^>]*\br:id="([^"]+)"/g)) {
    if (match[1] !== undefined) parts.push({ relId: match[1], kind: 'table' })
  }
  if (parts.length === 0) return
  const relsPath = relsPathFor(worksheetPath)
  if (!(await pkg.has(relsPath))) {
    throw new StructuralShiftError(
      `${worksheetPath} has anchored parts but ${relsPath} is missing — ` +
        'rows/columns cannot shift here.',
    )
  }
  const relsXml = await pkg.readText(relsPath)
  for (const { relId, kind } of parts) {
    // Two-step lookup: attribute order varies by producer (openpyxl puts
    // Target before Id), so never assume Id precedes Target.
    const relationshipXml = new RegExp(
      `<Relationship\\b[^>]*\\bId="${escapeRegExp(relId)}"[^>]*/?>`,
    ).exec(relsXml)?.[0]
    const target =
      relationshipXml === undefined ? undefined : /\bTarget="([^"]+)"/.exec(relationshipXml)?.[1]
    if (target === undefined) {
      throw new StructuralShiftError(
        `${worksheetPath} references ${kind} ${relId} but its relationship is missing — ` +
          'rows/columns cannot shift here.',
      )
    }
    const partPath = resolveRelTarget(worksheetPath, target)
    if (!(await pkg.canPatch(partPath))) {
      throw new StructuralShiftError(
        `${partPath} is too large to rewrite — rows/columns cannot shift here.`,
      )
    }
    const xml = await pkg.readText(partPath)
    const shifted = kind === 'drawing' ? shiftDrawingAnchors(xml, ops) : shiftTablePart(xml, ops)
    if (shifted === xml) continue
    pkg.write(partPath, shifted)
    touchedEntries.add(partPath)
  }
}

/// Attribute order and entity encoding in <sheet> elements vary by producer,
/// so never pattern-match the serialized XML for a name — parse each element
/// and compare decoded names instead (issue #10: valid workbooks failed to
/// save because r:id preceded name, or the name used numeric char refs).
function findSheetElement(workbookXml: string, sheetName: string): SheetElement | undefined {
  return parseSheetElements(workbookXml).find((element) => element.name === sheetName)
}

/// Resolve a worksheet's comments part via its rels (COMMENTS_REL_TYPE),
/// mirroring the desktop's sidecar lookup. Null when the sheet has none.
async function resolveCommentsPath(
  reader: Pick<EntrySource, 'readText' | 'has'>,
  worksheetPath: string,
): Promise<string | null> {
  const relsPath = worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
  if (!(await reader.has(relsPath))) return null
  const relsXml = await reader.readText(relsPath)
  const relationship = new RegExp(`<Relationship[^>]*Type="${COMMENTS_REL_TYPE}"[^>]*/?>`).exec(
    relsXml,
  )?.[0]
  const target =
    relationship === undefined ? undefined : /\bTarget="([^"]+)"/.exec(relationship)?.[1]
  if (target === undefined) return null
  if (target.startsWith('/')) return target.slice(1)
  const base = worksheetPath.split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part === '..') base.pop()
    else if (part !== '.') base.push(part)
  }
  return base.join('/')
}

async function resolveWorksheetPath(
  reader: Pick<EntrySource, 'readText'>,
  sheetName: string,
): Promise<string> {
  const workbookXml = await reader.readText('xl/workbook.xml')
  const relationshipId = findSheetElement(workbookXml, sheetName)?.relationshipId
  if (relationshipId === undefined)
    throw new Error(`Sheet "${sheetName}" was not found in workbook.xml.`)

  const relationshipsXml = await reader.readText('xl/_rels/workbook.xml.rels')
  // Two-step lookup: attribute order varies by producer (openpyxl puts
  // Target before Id), so never assume Id precedes Target.
  const relationshipXml = new RegExp(
    `<Relationship\\b[^>]*\\bId="${escapeRegExp(relationshipId)}"[^>]*/?>`,
  ).exec(relationshipsXml)?.[0]
  const targetMatch =
    relationshipXml === undefined ? undefined : /\bTarget="([^"]+)"/.exec(relationshipXml)?.[1]
  if (!targetMatch) throw new Error(`Relationship ${relationshipId} was not found.`)
  const target = targetMatch.replace(/^\/?xl\//, '')
  return `xl/${target.replace(/^\.\//, '')}`
}

function replaceSheetName(workbookXml: string, before: string, after: string): string {
  const element = findSheetElement(workbookXml, before)
  if (!element) throw new Error(`Sheet "${before}" was not found.`)
  const renamedXml = element.xml.replace(
    /(\bname=")[^"]*(")/,
    (_match, prefix: string, suffix: string) => `${prefix}${escapeXmlAttribute(after)}${suffix}`,
  )
  return workbookXml.replace(element.xml, () => renamedXml)
}

function patchCell(worksheetXml: string, address: string, cell: CellState): string {
  const cellPattern = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*(?:/>|>[\\s\\S]*?</c>)`)
  const replacement = serializeCell(address, cell)
  if (cellPattern.test(worksheetXml)) {
    return worksheetXml.replace(cellPattern, replacement)
  }
  if (replacement === '') return worksheetXml

  const rowNumber = address.match(/[1-9][0-9]*$/)?.[0]
  if (!rowNumber) throw new Error(`Invalid cell address: ${address}`)
  const rowPattern = new RegExp(`(<row\\b[^>]*\\br="${rowNumber}"[^>]*>)([\\s\\S]*?)(</row>)`)
  if (rowPattern.test(worksheetXml)) {
    return worksheetXml.replace(rowPattern, `$1$2${replacement}$3`)
  }
  const sheetDataClose = '</sheetData>'
  if (!worksheetXml.includes(sheetDataClose)) throw new Error('Worksheet has no sheetData element.')
  return worksheetXml.replace(
    sheetDataClose,
    `<row r="${rowNumber}">${replacement}</row>${sheetDataClose}`,
  )
}

function readCellStyleIndex(worksheetXml: string, address: string): number | undefined {
  const match = new RegExp(`<c\\b([^>]*)\\br="${address}"([^>]*?)[/>]`).exec(worksheetXml)
  if (!match) return undefined
  const index = readXmlAttribute(`${match[1] ?? ''} ${match[2] ?? ''}`, 's')
  return index === undefined ? undefined : Number(index)
}

/// Sets only the style index of a cell, leaving its stored content (value,
/// type, formula) byte-for-byte untouched. Missing cells become empty styled
/// cells so formatting applies to blanks too.
function patchCellStyleOnly(
  worksheetXml: string,
  address: string,
  styleIndex: number | undefined,
): string {
  if (styleIndex === undefined) return worksheetXml
  const cellPattern = new RegExp(`(<c\\b[^>]*?\\br="${address}"[^>]*?)(\\s*/>|>)`)
  const existing = cellPattern.exec(worksheetXml)
  if (existing) {
    const opening = existing[1] ?? ''
    const patched = /\bs="[^"]*"/.test(opening)
      ? opening.replace(/\bs="[^"]*"/, () => `s="${styleIndex}"`)
      : opening.replace(`r="${address}"`, () => `r="${address}" s="${styleIndex}"`)
    return worksheetXml.replace(cellPattern, () => `${patched}${existing[2] ?? ''}`)
  }
  return insertMissingCell(worksheetXml, address, `<c r="${address}" s="${styleIndex}"/>`)
}

/// Rewrites one cell in worksheet XML. The original cell's style index is
/// kept (or replaced by styleOverride) so edits don't strip formatting, and
/// missing cells or rows are inserted in ascending order (Excel expects
/// sorted sheetData).
function patchCellKeepingStyle(
  worksheetXml: string,
  address: string,
  cell: CellState,
  styleOverride?: number,
  rich?: readonly WorkbookRichRun[],
): string {
  const cellPattern = new RegExp(`<c\\b([^>]*)\\br="${address}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`)
  const existing = cellPattern.exec(worksheetXml)
  const styleIndex =
    styleOverride !== undefined
      ? String(styleOverride)
      : existing
        ? readXmlAttribute(`${existing[1] ?? ''} ${existing[2] ?? ''}`, 's')
        : undefined
  const replacement = serializeStyledCell(address, cell, styleIndex, rich)
  // Function replacements throughout: user text can contain `$1`/`$&`, which
  // string replacements would expand as backreferences and corrupt the XML.
  if (existing) return worksheetXml.replace(cellPattern, () => replacement)
  if (replacement === '') return worksheetXml
  return insertMissingCell(worksheetXml, address, replacement)
}

/**
 * Refresh a formula cell's cached value: replace (or insert) <v> inside the existing
 * <c>, keeping <f> and every attribute. Cells that don't exist or aren't formulas are
 * left alone — the recalc overlay only ever names formula cells, and a cell the user
 * turned into a literal must keep the literal.
 */
function patchFormulaCachedValue(
  worksheetXml: string,
  address: string,
  value: string | number | boolean | null,
): string {
  // Paired form only: a self-closing <c/> has no formula to keep.
  const cellPattern = new RegExp(`<c\\b([^>]*)\\br="${address}"([^>]*)>([\\s\\S]*?)</c>`)
  const existing = cellPattern.exec(worksheetXml)
  if (!existing) return worksheetXml
  const body = existing[3] ?? ''
  if (!/<f[\s/>]/.test(body)) return worksheetXml
  const attrs = `${existing[1] ?? ''}${existing[2] ?? ''}`
  // Formula results carry t="str" for text, no t (numeric default) otherwise;
  // booleans use t="b" with 1/0. A null result drops the cached value entirely.
  const numeric = typeof value === 'number' && Number.isFinite(value)
  const stripped = attrs.replace(/\st="[^"]*"/g, '')
  let typeAttr = ''
  let valueXml = ''
  if (numeric) {
    valueXml = `<v>${value}</v>`
  } else if (typeof value === 'boolean') {
    typeAttr = ' t="b"'
    valueXml = `<v>${value ? 1 : 0}</v>`
  } else if (value !== null && value !== undefined && value !== '') {
    typeAttr = ' t="str"'
    valueXml = `<v>${escapeXmlText(String(value))}</v>`
  }
  // Keep <f> (and any other children apart from the cached value) verbatim.
  const kept = body.replace(/<v\b[^>]*\/>|<v\b[^>]*>[\s\S]*?<\/v>/g, '')
  const replacement = `<c r="${address}"${stripped}${typeAttr}>${kept}${valueXml}</c>`
  return worksheetXml.replace(cellPattern, () => replacement)
}

function insertMissingCell(worksheetXml: string, address: string, cellXml: string): string {
  const rowNumber = Number(/[1-9][0-9]*$/.exec(address)?.[0])
  if (!Number.isFinite(rowNumber)) throw new Error(`Invalid cell address: ${address}`)
  const targetColumn = parseA1Column(address)
  const rowPattern = new RegExp(`<row\\b([^>]*?\\br="${rowNumber}"[^>]*?)(/>|>([\\s\\S]*?)</row>)`)
  const rowMatch = rowPattern.exec(worksheetXml)
  if (rowMatch) {
    const attributes = rowMatch[1] ?? ''
    const body = rowMatch[2] === '/>' ? '' : (rowMatch[3] ?? '')
    return worksheetXml.replace(
      rowPattern,
      () => `<row${attributes}>${insertCellInColumnOrder(body, cellXml, targetColumn)}</row>`,
    )
  }
  return insertRowInOrder(worksheetXml, rowNumber, cellXml)
}

function insertCellInColumnOrder(rowBody: string, cellXml: string, targetColumn: number): string {
  const siblingPattern = /<c\b[^>]*?\br="([A-Z]{1,3})[1-9][0-9]*"/g
  let match: RegExpExecArray | null
  while ((match = siblingPattern.exec(rowBody)) !== null) {
    if (lettersToColumn(match[1] ?? '') > targetColumn) {
      return rowBody.slice(0, match.index) + cellXml + rowBody.slice(match.index)
    }
  }
  return rowBody + cellXml
}

function insertRowInOrder(worksheetXml: string, rowNumber: number, cellXml: string): string {
  const newRow = `<row r="${rowNumber}">${cellXml}</row>`
  const rowStartPattern = /<row\b[^>]*?\br="([1-9][0-9]*)"/g
  let match: RegExpExecArray | null
  while ((match = rowStartPattern.exec(worksheetXml)) !== null) {
    if (Number(match[1]) > rowNumber) {
      return worksheetXml.slice(0, match.index) + newRow + worksheetXml.slice(match.index)
    }
  }
  if (worksheetXml.includes('</sheetData>')) {
    return worksheetXml.replace('</sheetData>', () => `${newRow}</sheetData>`)
  }
  const emptySheetData = /<sheetData\s*\/>/
  if (emptySheetData.test(worksheetXml)) {
    return worksheetXml.replace(emptySheetData, () => `<sheetData>${newRow}</sheetData>`)
  }
  throw new Error('Worksheet has no sheetData element.')
}

/// The sidecar uses worksheet dimension metadata to choose the streamable
/// viewport. Some minimal workbooks start at A1:A1; after inserting cells the
/// dimension must grow too, or a successful save appears to lose every cell
/// outside A1 when the workbook is reopened.
function expandWorksheetDimensionToCells(worksheetXml: string): string {
  let maximumRow = 1
  let maximumColumn = 0
  const include = (reference: string): void => {
    const cleaned = reference.replace(/\$/g, '')
    const cell = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(cleaned)
    if (!cell?.[1] || !cell[2]) return
    const row = Number(cell[2])
    const column = lettersToColumn(cell[1])
    maximumRow = Math.max(maximumRow, row)
    maximumColumn = Math.max(maximumColumn, column)
  }

  const dimension = /<dimension\b[^>]*\bref="([^"]+)"[^>]*\/?>/.exec(worksheetXml)
  if (dimension?.[1]) {
    for (const reference of dimension[1].split(':')) include(reference)
  }
  for (const match of worksheetXml.matchAll(/<c\b[^>]*\br="([A-Z]{1,3}[1-9][0-9]*)"/g)) {
    if (match[1]) include(match[1])
  }

  const last = toA1Address(maximumRow - 1, maximumColumn)
  const reference = last === 'A1' ? 'A1' : `A1:${last}`
  if (dimension) {
    return worksheetXml.replace(/(<dimension\b[^>]*\bref=")[^"]+("[^>]*\/?>)/, `$1${reference}$2`)
  }
  return worksheetXml.replace(/(<worksheet\b[^>]*>)/, `$1<dimension ref="${reference}"/>`)
}

function serializeStyledCell(
  address: string,
  cell: CellState,
  styleIndex: string | undefined,
  rich?: readonly WorkbookRichRun[],
): string {
  const style = styleIndex === undefined ? '' : ` s="${styleIndex}"`
  if (cell.formula) {
    return `<c r="${address}"${style}><f>${escapeXmlText(cell.formula.replace(/^=/, ''))}</f></c>`
  }
  if (cell.value === null) {
    // A cleared cell keeps its formatting only if it keeps a style index.
    return styleIndex === undefined ? '' : `<c r="${address}"${style}/>`
  }
  if (typeof cell.value === 'string') {
    if (rich && rich.length > 0) {
      const runs = rich
        .map(
          (run) =>
            `<r>${serializeRunProperties(run)}<t xml:space="preserve">${escapeXmlText(run.text)}</t></r>`,
        )
        .join('')
      return `<c r="${address}"${style} t="inlineStr"><is>${runs}</is></c>`
    }
    return `<c r="${address}"${style} t="inlineStr"><is><t xml:space="preserve">${escapeXmlText(cell.value)}</t></is></c>`
  }
  if (typeof cell.value === 'boolean') {
    return `<c r="${address}"${style} t="b"><v>${cell.value ? 1 : 0}</v></c>`
  }
  return `<c r="${address}"${style}><v>${cell.value}</v></c>`
}

/// `<rPr>` for one inline-string run; empty string when the run is plain.
function serializeRunProperties(run: WorkbookRichRun): string {
  const parts: string[] = []
  if (run.bold) parts.push('<b/>')
  if (run.italic) parts.push('<i/>')
  if (run.strikethrough) parts.push('<strike/>')
  if (run.underline) parts.push('<u/>')
  if (run.size !== undefined) parts.push(`<sz val="${run.size}"/>`)
  if (run.color !== undefined) {
    const hex = run.color.replace(/^#/, '').toUpperCase()
    if (/^[0-9A-F]{6}$/.test(hex)) parts.push(`<color rgb="FF${hex}"/>`)
  }
  if (run.family !== undefined) {
    parts.push(`<rFont val="${escapeXmlAttribute(run.family)}"/>`)
  }
  return parts.length === 0 ? '' : `<rPr>${parts.join('')}</rPr>`
}

function ensureFullCalcOnLoad(workbookXml: string): string {
  if (/<calcPr\b[^>]*\bfullCalcOnLoad="1"/.test(workbookXml)) return workbookXml
  if (/<calcPr\b/.test(workbookXml)) {
    return workbookXml.replace(
      /<calcPr\b([^>]*?)(\/?>)/,
      (_full, attributes: string, close: string) => {
        const cleaned = attributes.replace(/\s*fullCalcOnLoad="[^"]*"/, '')
        return `<calcPr${cleaned} fullCalcOnLoad="1"${close}`
      },
    )
  }
  // Schema order places calcPr after definedNames (or after sheets).
  const anchor = workbookXml.includes('</definedNames>') ? '</definedNames>' : '</sheets>'
  if (!workbookXml.includes(anchor)) return workbookXml
  return workbookXml.replace(anchor, `${anchor}<calcPr fullCalcOnLoad="1"/>`)
}

function parseA1Column(address: string): number {
  const letters = /^[A-Z]{1,3}/.exec(address)?.[0]
  if (!letters) throw new Error(`Invalid cell address: ${address}`)
  return lettersToColumn(letters)
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const character of letters) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}

function serializeCell(address: string, cell: CellState): string {
  if (cell.formula) {
    return `<c r="${address}"><f>${escapeXmlText(cell.formula.slice(1))}</f></c>`
  }
  if (cell.value === null) return ''
  if (typeof cell.value === 'string') {
    return `<c r="${address}" t="inlineStr"><is><t xml:space="preserve">${escapeXmlText(cell.value)}</t></is></c>`
  }
  if (typeof cell.value === 'boolean') {
    return `<c r="${address}" t="b"><v>${cell.value ? 1 : 0}</v></c>`
  }
  return `<c r="${address}"><v>${cell.value}</v></c>`
}

function parseCell(worksheetXml: string, address: string): CellState {
  const cellPattern = new RegExp(`<c\\b([^>]*)\\br="${address}"([^>]*)(?:/>|>([\\s\\S]*?)</c>)`)
  const match = cellPattern.exec(worksheetXml)
  if (!match) return { value: null }
  const attributes = `${match[1] ?? ''}${match[2] ?? ''}`
  const body = match[3] ?? ''
  const formula = /<f(?:\s[^>]*[^/>])?>([\s\S]*?)<\/f>/.exec(body)?.[1]
  if (formula !== undefined) return { value: null, formula: `=${decodeXmlText(formula)}` }
  const type = /\bt="([^"]+)"/.exec(attributes)?.[1]
  if (type === 's') throw new Error(`Shared-string cell ${address} is not writable in this PoC.`)
  if (type === 'inlineStr') {
    const text = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/.exec(body)?.[1] ?? ''
    return { value: decodeXmlText(text) }
  }
  const rawValue = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(body)?.[1]
  if (rawValue === undefined) return { value: null }
  if (type === 'b') return { value: rawValue === '1' }
  const numericValue = Number(rawValue)
  if (!Number.isFinite(numericValue))
    throw new Error(`Cell ${address} has an unsupported numeric value.`)
  return { value: numericValue }
}

function parseWorksheetCells(
  worksheetXml: string,
  sharedStrings: readonly string[],
): Readonly<Record<string, CellState>> {
  const cells: Record<string, CellState> = {}
  const cellPattern = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
  let match: RegExpExecArray | null
  while ((match = cellPattern.exec(worksheetXml)) !== null) {
    const attributes = match[1] ?? ''
    const address = readXmlAttribute(attributes, 'r')
    if (!address || !/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(address)) continue
    const body = match[2] ?? ''
    const formula = /<f(?:\s[^>]*[^/>])?>([\s\S]*?)<\/f>/.exec(body)?.[1]
    if (formula !== undefined) {
      cells[address] = { value: null, formula: `=${decodeXmlText(formula)}` }
      continue
    }
    const type = readXmlAttribute(attributes, 't')
    if (type === 'inlineStr') {
      const text = [...body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
        .map((textMatch) => decodeXmlText(textMatch[1] ?? ''))
        .join('')
      cells[address] = { value: text }
      continue
    }
    const rawValue = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(body)?.[1]
    if (rawValue === undefined) {
      cells[address] = { value: null }
    } else if (type === 's') {
      const index = Number(rawValue)
      cells[address] = { value: sharedStrings[index] ?? '' }
    } else if (type === 'b') {
      cells[address] = { value: rawValue === '1' }
    } else if (type === 'str') {
      cells[address] = { value: decodeXmlText(rawValue) }
    } else {
      const numericValue = Number(rawValue)
      cells[address] = {
        value: Number.isFinite(numericValue) ? numericValue : decodeXmlText(rawValue),
      }
    }
  }
  return cells
}

async function readSharedStrings(source: EntrySource): Promise<readonly string[]> {
  if (!(await source.has('xl/sharedStrings.xml'))) return []
  const xml = await source.readText('xl/sharedStrings.xml')
  return [...xml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)].map((itemMatch) =>
    [...(itemMatch[1] ?? '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
      .map((textMatch) => decodeXmlText(textMatch[1] ?? ''))
      .join(''),
  )
}

function cellsEqual(left: CellState, right: CellState): boolean {
  return left.value === right.value && left.formula === right.formula
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

const XML_NAMED_ENTITIES: Record<string, string> = {
  quot: '"',
  apos: "'",
  lt: '<',
  gt: '>',
  amp: '&',
}

/// Single-pass decode of the XML named entities plus numeric character
/// references (&#dd; / &#xhh;), which some producers use for non-ASCII text.
function decodeXmlText(input: string): string {
  return input.replace(
    /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|(quot|apos|lt|gt|amp));/g,
    (match, hex: string | undefined, dec: string | undefined, named: string | undefined) => {
      if (named !== undefined) return XML_NAMED_ENTITIES[named] ?? match
      const code = hex !== undefined ? Number.parseInt(hex, 16) : Number(dec)
      return code <= 0x10ffff ? String.fromCodePoint(code) : match
    },
  )
}

function escapeXmlAttribute(input: string): string {
  return escapeXmlText(input).replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function readXmlAttribute(attributes: string, name: string): string | undefined {
  return new RegExp(`(?:^|\\s)${escapeRegExp(name)}="([^"]*)"`).exec(attributes)?.[1]
}
