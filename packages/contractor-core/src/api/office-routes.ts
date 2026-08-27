/**
 * Office API routes — pure, server-side handlers for Word/Excel file I/O.
 *
 * Lives in @contractor/core because the routes depend on the pure office
 * engines (`@genoffice/xlsx-gateway`, `@genoffice/docx-engine`). The host
 * (web-host / vercel-handler) delegates `/api/office/*` requests here.
 *
 * The routes:
 *   POST /api/office/workbooks/open  — read XLSX → WorkbookSnapshot
 *   POST /api/office/workbooks/save  — apply CellEdits to source XLSX bytes
 *   POST /api/office/documents/open  — read DOCX → simplified SerializedBlock[]
 *   POST /api/office/documents/save  — patch DOCX with the editor's blocks
 *
 * File bytes are base64-encoded inside JSON envelopes (per Phase-2 spec —
 * acceptable for fixture-sized files; the 10MB cap is enforced post-decode).
 * The actual byte transport is abstracted behind OfficeBinaryCodec so the
 * route handlers never call Buffer.from(..., 'base64') directly — the codec
 * is the single replaceable seam between the wire format and the engine.
 *
 * Purity: zero Electron imports, zero node:* imports except for the Buffer
 * global (used by @genoffice/xlsx-gateway and the default Base64Codec). The
 * host owns the HTTP transport; this module owns request validation +
 * engine delegation.
 */

import {
  applyCellEditsToXlsx,
  readBasicWorkbook,
  type AddedVisualLocator,
  type CellEdit,
  type ChartAdd,
  type ChartAddSeries,
  type ChartSeriesEdit,
  type ChartSeriesSetEntry,
  type EditableBorderStyle,
  type SheetDvState,
  type SheetFilterState,
  type SheetNoteState,
  type SheetPageSetupState,
  type SheetProtectionState,
  type SheetStructuralOps,
  type SheetTableAddition,
  type SheetVisualAddition,
  type StructuralOp,
  type WorkbookChartEdit,
  type WorkbookSnapshot,
  type WorkbookStyleEdit,
  type WorkbookVisualEdit,
} from '@genoffice/xlsx-gateway'
import {
  applyImageWrap,
  generateTableModelXml,
  parseDocx,
  patchImageParagraphXml,
  saveDocx,
  type Block,
  type CellBorder,
  type CellBorders,
  type ImagePatch,
  type ImageWrap,
  type NewImage,
  type ParsedDocFull,
  type Run,
  type SaveBlock,
  type TableModel,
} from '@genoffice/docx-engine'

// ── Wire types (JSON-serializable; the browser sends/receives these shapes) ──

export interface OpenWorkbookRequest {
  readonly fileName: string
  readonly fileBytes: string // base64-encoded XLSX
}
export interface OpenWorkbookResponse {
  readonly snapshot: WorkbookSnapshot
  readonly sheetNamesById: Readonly<Record<string, string>>
}

/**
 * Save-mutation plan for a workbook. Modeled after the canonical
 * `SavePlan` from `@genoffice/runtime-contracts` but trimmed to the
 * mutation families the browser actually emits today (cell edits). The
 * `edits` field carries the existing `CellEdit[]` shape; future mutation
 * families (structural ops, charts, hyperlinks, …) get added here as
 * optional fields, preserving backward compatibility with older clients
 * that send only the fields they know about.
 */
export interface BrowserWorkbookSavePlan {
  readonly edits: readonly CellEdit[]
  /**
   * Row/column structural operations (insert/delete rows/columns) per sheet.
   * The engine replays these BEFORE the cell edits — journaled cell edits
   * are in the POST-operation coordinate space, which is exactly what the
   * browser's mutation-captured dirty map produces.
   */
  readonly structuralOps?: readonly SheetStructuralOps[]
  /**
   * Per-sheet page-setup states (freeze panes, …). The engine applies these
   * AFTER structural ops and cell edits. Only `frozenRows`/`frozenColumns`
   * are emitted by the web shell today; the underlying `SheetPageSetupState`
   * carries the full page-setup model so future View commands can land here
   * without a wire-breaking change.
   */
  readonly pageSetupStates?: readonly SheetPageSetupState[]
  /**
   * Per-sheet AutoFilter states (Data → Filter). The engine applies these
   * AFTER structural ops, cell edits, and page setup — filter coordinates
   * and row sets match the sheet's final content. Each entry is the
   * canonical `SheetFilterState`; `filter: null` clears the filter.
   */
  readonly filterStates?: readonly SheetFilterState[]
  /**
   * Per-sheet data-validation states (Data → Data Validation). The engine
   * applies these AFTER structural ops, cell edits, and filters — each
   * entry is the canonical `SheetDvState` (the full declarative rule set
   * of a DV-dirty sheet). An empty rules list clears the sheet's
   * `<dataValidations>`.
   */
  readonly dvStates?: readonly SheetDvState[]
  /**
   * Per-sheet legacy-note states (Review → New Comment). Each entry REPLACES
   * the sheet's whole comment set with the canonical `SheetNote[]` — an
   * empty notes array removes the sheet's comment part entirely.
   */
  readonly noteStates?: readonly SheetNoteState[]
  /**
   * Per-sheet protection states (Review → Protect Sheet, EXCEL-020). Each
   * entry carries the desired protected flag for one sheet — the engine
   * adds/removes the worksheet's `<sheetProtection>` element (no password
   * support: unprotecting a password-protected sheet fails closed in the
   * engine, surfacing as a 4xx malformed error).
   */
  readonly sheetProtections?: readonly SheetProtectionState[]
  /**
   * Desired workbook structure protection (Review → Protect Workbook,
   * EXCEL-020). null/absent = untouched — the engine leaves workbook.xml's
   * `<workbookProtection>` alone.
   */
  readonly workbookProtectionState?: { readonly lockStructure: boolean } | null
  /**
   * Session-created tables (Insert → Table, EXCEL-021). Each entry carries
   * the journaled creation for one table — the engine writes a brand-new
   * xl/tables/tableN.xml part, the worksheet's `<tableParts>` element, the
   * relationship, and the [Content_Types] override, failing closed on name
   * collisions, overlaps, and bad column names. Mirrors the desktop's
   * tableAdditions journal semantics (deleting a session table drops its
   * entry, so it is never persisted — convert-to-range).
   */
  readonly tableAdditions?: readonly SheetTableAddition[]
  /**
   * Surgical semantic edits to file-native charts (EXCEL-023 — Chart
   * Design). Each entry targets one chart part by its canonical
   * xl/charts/*.xml path (the open snapshot's charts carry it) and carries
   * at least one bounded property change: title, the six convertible
   * chart types, series colors, legend, data labels (+ position/format),
   * axis titles, point colors, grouping, gridlines, value-axis bounds,
   * gap width, hole size, explosion, or a full series replacement.
   * Mirrors the desktop preload's workbookChartEditSchema bounds.
   */
  readonly chartEdits?: readonly WorkbookChartEdit[]
  /**
   * Session-created visuals (Insert → Picture / Insert → Chart,
   * EXCEL-022 / EXCEL-023). Each entry carries a typed visual — an image
   * addition or a chart addition — persisted by the engine as new OOXML
   * parts. Mirrors the desktop visualAdditions journal semantics
   * (deleting a session visual drops its entry, so it is never
   * persisted).
   */
  readonly visualAdditions?: readonly SheetVisualAddition[]
  /**
   * Surgical edits to file-native visuals (move / resize / delete,
   * EXCEL-022). Each entry targets the canonical (drawingPath,
   * drawingIndex) locator the open snapshot's images carry — the engine
   * rewrites the anchor's from/to markers or splices the anchor out,
   * cascading the image relationship + media part on delete (media is
   * removed only when no remaining relationship references it).
   */
  readonly visualEdits?: readonly WorkbookVisualEdit[]
  // Extensibility seam — future mutation families land here as optional
  // readonly fields (chartEdits?, hyperlinkEdits?, …).
  // The route handler ignores unknown keys, so adding a field is a
  // forward-compatible wire change.
  readonly [key: string]: unknown
}

/**
 * Canonical workbook save request — typed envelope carrying a savePlan.
 *
 * The browser sends this shape. For backward compatibility with the
 * pre-increment client (and with the round-trip tests, which predate the
 * savePlan field), the parser also accepts the legacy
 * `{ fileName, fileBytes, cellEdits }` shape and normalizes it into a
 * savePlan with `{ edits: cellEdits }`.
 */
export interface BrowserWorkbookSaveRequest {
  readonly fileName: string
  readonly fileBytes: string // base64-encoded source XLSX
  readonly savePlan: BrowserWorkbookSavePlan
}

/**
 * Legacy workbook save request — accepted for backward compatibility.
 * Normalized into a BrowserWorkbookSaveRequest by the parser.
 */
export interface LegacyWorkbookSaveRequest {
  readonly fileName: string
  readonly fileBytes: string
  readonly cellEdits: readonly CellEdit[]
}

export type SaveWorkbookRequest = BrowserWorkbookSaveRequest | LegacyWorkbookSaveRequest

export interface SaveWorkbookResponse {
  readonly fileBytes: string // base64-encoded mutated XLSX
  /**
   * EXCEL-022: locators of visuals persisted by this save's
   * visualAdditions (drawingPath + drawingIndex per addition, in order).
   * The browser merges them into its live image state so a later
   * move/resize/delete targets the exact appended anchor.
   */
  readonly addedVisuals?: readonly AddedVisualLocator[]
}

/**
 * Serialized run — browser-safe representation of a docx-engine Run.
 *
 * Carries inline formatting marks (bold, italic, underline, strike, links)
 * so the browser editor can faithfully render and edit rich text.
 * Derived from docx-engine's Run type; only the browser-relevant subset
 * is serialized (display-only fields like rawRPr, charSpacingTwips, etc.
 * are preserved by the server's original-block passthrough).
 */
export interface SerializedRun {
  readonly text: string
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  readonly link?: { readonly href: string; readonly tooltip?: string }
}

// ── Serialized table (wire mirror of the canonical docx-engine TableModel) ──

/**
 * One cell border edge (wire mirror of docx-engine's CellBorder).
 * style is a w:val keyword (single/dashed/double/none/…).
 */
export interface SerializedCellBorder {
  readonly style: string
  readonly szEighths?: number
  readonly color?: string
}

export interface SerializedCellBorders {
  readonly top?: SerializedCellBorder
  readonly left?: SerializedCellBorder
  readonly bottom?: SerializedCellBorder
  readonly right?: SerializedCellBorder
}

export interface SerializedTableBorders extends SerializedCellBorders {
  readonly insideH?: SerializedCellBorder
  readonly insideV?: SerializedCellBorder
}

export interface SerializedCellMargins {
  readonly top?: number
  readonly left?: number
  readonly bottom?: number
  readonly right?: number
}

/** One paragraph inside a table cell (wire mirror of TableParagraph). */
export interface SerializedTableParagraph {
  readonly runs: readonly SerializedRun[]
  readonly align?: 'left' | 'center' | 'right' | 'justify' | 'distribute'
  readonly styleId?: string
}

/**
 * One table cell (wire mirror of docx-engine's TableCell, editable subset).
 *
 * `rawTcPr` carries the original <w:tcPr> bytes so unmodeled properties
 * (per-cell tcMar, textDirection, noWrap…) survive regeneration of an edited
 * table — the generator surgically patches the modeled children into it.
 */
export interface SerializedTableCell {
  readonly paras: readonly string[]
  readonly richParas?: readonly SerializedTableParagraph[]
  readonly colSpan?: number
  readonly vMerge?: 'restart' | 'continue'
  readonly fill?: string
  readonly color?: string
  readonly bold?: boolean
  readonly align?: 'left' | 'center' | 'right' | 'justify' | 'distribute'
  readonly vAlign?: 'top' | 'center' | 'bottom'
  readonly borders?: SerializedCellBorders
  readonly rawTcPr?: string
}

/**
 * Editable table interchange (wire mirror of docx-engine's TableModel).
 *
 * NOT a lossy text-only DTO: structure (colSpan/vMerge), cell formatting
 * (fill/vAlign/borders), per-row raw <w:trPr> bytes (row heights, tblHeader,
 * cantSplit, row revisions) and per-cell raw <w:tcPr> bytes all round-trip,
 * and an edited table regenerates through the canonical engine generator
 * (generateTableModelXml) — never a web-side XML serializer.
 *
 * `headerRows` is the wire representation of trPr <w:tblHeader/> (the
 * canonical model keeps it inside rawTrPrs); the server patches it into the
 * row properties on save.
 */
export interface SerializedTable {
  readonly rows: readonly (readonly SerializedTableCell[])[]
  readonly colWidthsPct?: readonly number[]
  readonly colWidthsTwips?: readonly number[]
  readonly widthPct?: number
  readonly autoLayout?: boolean
  readonly cellMarTwips?: SerializedCellMargins
  readonly borders?: SerializedTableBorders
  readonly align?: 'left' | 'center' | 'right'
  readonly indentTwips?: number
  readonly rowHeightsTwips?: readonly (number | null)[]
  readonly rowHeightRules?: readonly ('atLeast' | 'exact' | null)[]
  readonly rawTrPrs?: readonly (string | null)[]
  readonly tblStyleId?: string
  readonly bidiVisual?: boolean
  /** per-row trPr <w:tblHeader/> flag (row is a header row) */
  readonly headerRows?: readonly boolean[]
}

// ── Serialized image (wire mirror of the canonical docx-engine image model) ──

/** Wrap modes of a floating image (wire mirror of the canonical ImageWrap). */
export type SerializedImageWrap =
  | 'inline'
  | 'square-left'
  | 'square-right'
  | 'tight-left'
  | 'tight-right'
  | 'through-left'
  | 'through-right'
  | 'topBottom'
  | 'behind'
  | 'front'

/** Source crop / fill placement rect as per-side fractions of the picture. */
export interface SerializedImageRect {
  readonly l: number
  readonly t: number
  readonly r: number
  readonly b: number
}

/**
 * Editable image payload (type === 'image').
 *
 * Wire mirror of the canonical Block image fields. The browser renders pixels
 * from `imageDataUrl` and edits the typed properties; the server diffs the
 * edited state against the parsed original block and emits ONLY the canonical
 * engine patches (patchImageParagraphXml / applyImageWrap) — the browser
 * never constructs image XML. `posHRel`/`posVRel` and `fillRect` are echo/
 * display fields for floating-position fidelity.
 */
export interface SerializedImage {
  readonly imageDataUrl: string | null
  readonly widthPx?: number
  readonly heightPx?: number
  readonly crop?: SerializedImageRect
  readonly fillRect?: SerializedImageRect
  readonly align?: 'left' | 'center' | 'right'
  readonly wrap?: SerializedImageWrap
  readonly offsetXEmu?: number
  readonly offsetYEmu?: number
  readonly posH?: 'left' | 'center' | 'right'
  readonly posV?: 'top' | 'center' | 'bottom'
  readonly posHRel?: 'margin' | 'page' | 'column' | 'paragraph' | 'character'
  readonly posVRel?: 'margin' | 'page' | 'paragraph' | 'line'
  readonly rotDeg?: number
  readonly flipH?: boolean
  readonly flipV?: boolean
  /**
   * Accessibility alt text (wp:docPr descr). Tri-state:
   *  - undefined: keep the existing descr (field absent — unchanged echo)
   *  - null:       clear (remove the descr attribute)
   *  - non-empty:  set the descr (bounded + control-char stripped)
   */
  readonly alt?: string | null
}

/**
 * A NEW image to embed at save time (docxIndex === null image blocks).
 * Wire mirror of the canonical NewImage: bytes become a word/media part +
 * relationship + drawing through the engine's embed path.
 */
export interface SerializedNewImage {
  readonly base64: string
  readonly mime: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly widthPx: number
  readonly heightPx: number
  readonly align?: 'left' | 'center' | 'right'
  readonly wrap?: SerializedImageWrap
  readonly rotDeg?: number
  readonly flipH?: boolean
  readonly flipV?: boolean
}

/**
 * Simplified Tiptap-compatible block representation.
 *
 * The browser renders blocks in Tiptap; the docx-engine's rich Block type
 * carries display-only fields (charts, ink, image data URLs…) that we do not
 * round-trip through JSON. SerializedBlock keeps just enough to:
 *   - render the block (type + runs + level)
 *   - save the block back: `docxIndex` lets the server re-emit the original
 *     XML byte-identically when the block is unchanged; `edited` lets the
 *     server regenerate the paragraph from `runs` when the user typed into it.
 *
 * Stable identity rule: `docxIndex` is the index of the original
 * <w:body> child the block came from. New blocks inserted by the editor
 * carry `docxIndex: null`. Deleted blocks are simply absent from the save
 * request. The browser must NOT renumber blocks by their visible position —
 * that would corrupt patch anchors when an edit changes the block order.
 *
 * Run-level formatting: `runs` carries the inline marks (bold, italic, etc.)
 * so the browser can render faithful rich text. When `edited` is false and
 * `docxIndex` is not null, the server copies original bytes; runs are only
 * used when `edited` is true.
 */
export interface SerializedBlock {
  readonly docxIndex: number | null
  readonly type: 'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'passthrough' | 'hidden'
  readonly text: string
  readonly runs?: readonly SerializedRun[]
  /**
   * Editable table payload (type === 'table'). Present when the canonical
   * TableModel was extracted and the table is browser-editable; absent for
   * tables the editor cannot safely regenerate (nested tables, anchored
   * shapes in cells) — those stay byte-preserved read-only blocks.
   */
  readonly table?: SerializedTable
  /**
   * Editable image payload (type === 'image'). Present when the canonical
   * image model was extracted (media readable); absent for broken images —
   * those stay byte-preserved read-only blocks.
   */
  readonly image?: SerializedImage
  /**
   * New-image embedding spec (type === 'image', docxIndex === null only).
   * The engine creates the media part + relationship + drawing.
   */
  readonly newImage?: SerializedNewImage
  readonly level?: number
  readonly listKind?: 'bullet' | 'ordered'
  /** true when the browser editor modified the block; the server regenerates it. */
  readonly edited?: boolean
  /** true for body-trailing elements (w:sectPr) never shown in the editor. */
  readonly hidden?: boolean
}

export interface OpenDocumentRequest {
  readonly fileName: string
  readonly fileBytes: string // base64-encoded DOCX
}
export interface OpenDocumentResponse {
  readonly blocks: readonly SerializedBlock[]
}
export interface SaveDocumentRequest {
  readonly fileName: string
  readonly fileBytes: string // base64-encoded source DOCX
  readonly blocks: readonly SerializedBlock[]
}
export interface SaveDocumentResponse {
  readonly fileBytes: string // base64-encoded patched DOCX
}

// ── Transport shape ─────────────────────────────────────────────────────────

export interface OfficeApiRequest {
  readonly method: string
  /** Path relative to /api/office, e.g. "/workbooks/open". */
  readonly path: string
  readonly body: unknown
}

export interface OfficeApiResponse {
  readonly status: number
  readonly body: unknown
}

// ── Byte transport codec ────────────────────────────────────────────────────

/**
 * Replaceable byte↔wire encoder. The default implementation is base64
 * (matching the JSON-envelope contract), but the route handlers depend
 * only on the interface so a future binary transport (e.g. multipart,
 * MessagePack) can be swapped in without touching the handlers.
 *
 * PURITY: the interface is transport-agnostic. The default Base64Codec
 * implementation uses the Node Buffer global; the browser-side mirror of
 * this interface (in office-client.ts) uses atob/btoa instead. Both
 * produce the same wire bytes.
 */
export interface OfficeBinaryCodec {
  encode(bytes: Uint8Array): string
  decode(encoded: string): Uint8Array
}

/**
 * Default base64 codec backed by the Node Buffer global. This is the only
 * place in office-routes.ts that touches Buffer — every handler goes
 * through `codec.encode` / `codec.decode` so the transport stays replaceable.
 */
export class Base64Codec implements OfficeBinaryCodec {
  encode(bytes: Uint8Array): string {
    const buf = bytes instanceof Buffer ? bytes : Buffer.from(bytes)
    return buf.toString('base64')
  }
  decode(encoded: string): Uint8Array {
    return Buffer.from(encoded, 'base64')
  }
}

const DEFAULT_CODEC: OfficeBinaryCodec = new Base64Codec()

// ── Security ────────────────────────────────────────────────────────────────

const MAX_FILE_BYTES = 10 * 1024 * 1024
const ALLOWED_EXTENSIONS = new Set(['.xlsx', '.csv', '.xls', '.docx'])
const SAFE_NAME_RE = /^[A-Za-z0-9.\-_]+$/

export class OfficeValidationError extends Error {
  constructor(
    public readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'OfficeValidationError'
  }
}

// ── Runtime type guards (replace `as Record<string, unknown>` / `as readonly T[]`) ──

/**
 * Narrow `unknown` to `Record<string, unknown>` with a runtime check.
 * Rejects arrays (arrays are objects in JS, but they are not record-shaped
 * request bodies) and null/undefined.
 */
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * Validate that `value` is a string and return it; throw OfficeValidationError
 * with `field` context otherwise. Empty strings are rejected when
 * `requireNonEmpty` is true (the default — every string field in our wire
 * shape is required to carry a value).
 */
function expectString(value: unknown, field: string, requireNonEmpty = true): string {
  if (typeof value !== 'string') {
    throw new OfficeValidationError('validation', `${field} must be a string`)
  }
  if (requireNonEmpty && value.length === 0) {
    throw new OfficeValidationError('validation', `${field} is required`)
  }
  return value
}

function expectNumber(value: unknown, field: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new OfficeValidationError('validation', `${field} must be a finite number`)
  }
  return value
}

function expectBoolean(value: unknown, field: string): boolean {
  if (typeof value !== 'boolean') {
    throw new OfficeValidationError('validation', `${field} must be a boolean`)
  }
  return value
}

/**
 * Validate that `value` is an array and run `predicate` on each entry.
 * Returns a readonly array of the narrowed type. Replaces the unsafe
 * `asArray<T>` cast that previously trusted the caller's claim about T.
 */
function expectArray<T>(
  value: unknown,
  field: string,
  predicate: (entry: unknown, index: number) => T,
): readonly T[] {
  if (!Array.isArray(value)) {
    throw new OfficeValidationError('validation', `${field} must be an array`)
  }
  return value.map((entry, index) => predicate(entry, index))
}

/**
 * Validate a `CellState` (the cell payload inside a CellEdit). The wire
 * shape mirrors `@genoffice/xlsx-gateway`'s `CellState` — `value` is a
 * scalar (string | number | boolean | null), `formula` is an optional
 * string.
 */
function expectCellState(
  value: unknown,
  field: string,
): {
  readonly value: string | number | boolean | null
  readonly formula?: string
} {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const v = value.value
  // `value` may be string | number | boolean | null. Anything else (object,
  // array, undefined) is rejected.
  if (v !== null && typeof v !== 'string' && typeof v !== 'number' && typeof v !== 'boolean') {
    throw new OfficeValidationError(
      'validation',
      `${field}.value must be a string, number, boolean, or null`,
    )
  }
  const formula = value.formula
  if (formula !== undefined && typeof formula !== 'string') {
    throw new OfficeValidationError('validation', `${field}.formula must be a string when present`)
  }
  return {
    value: v as string | number | boolean | null,
    ...(formula !== undefined ? { formula } : {}),
  }
}

/**
 * Validate a `CellEdit`. The wire shape mirrors
 * `@genoffice/xlsx-gateway`'s `CellEdit` — we validate the required fields
 * (sheetName, row, column, writeValue, cell) and ignore optional style
 * fields the browser may attach (style, rich, styleReset). Those optional
 * fields are passed through verbatim to the engine, which re-validates
 * them against its own schema.
 */
// ── WorkbookStyleEdit validation (Excel cell formatting) ────────────────────

/** WorkbookStyleEdit HexColor on the wire: '#'-prefixed 6-digit hex, or null to clear. */
function expectStyleHexColor(value: unknown, field: string): string | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const s = expectString(value, field)
  if (!/^#[0-9A-Fa-f]{6}$/.test(s)) {
    throw new OfficeValidationError('validation', `${field} must be a '#RRGGBB' hex color or null`)
  }
  return s.toUpperCase()
}

const STYLE_ALIGNS_H = ['left', 'center', 'right', 'justify', 'distributed'] as const
const STYLE_ALIGNS_V = ['top', 'center', 'bottom'] as const
const STYLE_BORDER_STYLES = [
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
] as const

function expectStyleBorder(
  value: unknown,
  field: string,
): { style: EditableBorderStyle; color?: string } {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const style = expectString(value.style, `${field}.style`)
  if (!(STYLE_BORDER_STYLES as readonly string[]).includes(style)) {
    throw new OfficeValidationError('validation', `${field}.style must be a border style keyword`)
  }
  const color = expectStyleHexColor(value.color, `${field}.color`)
  return {
    style: style as EditableBorderStyle,
    ...(color !== undefined && color !== null ? { color } : {}),
  }
}

/**
 * Validate a `WorkbookStyleEdit` delta (Excel cell formatting). Malformed
 * deltas throw OfficeValidationError → the existing 400 validation error
 * shape. Only typed fields pass — the browser can never inject raw XML.
 */
function expectStyleEdit(value: unknown, field: string): WorkbookStyleEdit {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const bold = expectOptionalBoolean(value.bold, `${field}.bold`)
  const italic = expectOptionalBoolean(value.italic, `${field}.italic`)
  const underline = expectOptionalBoolean(value.underline, `${field}.underline`)
  const underlineStyleRaw = expectOptionalString(
    value.underlineStyle,
    `${field}.underlineStyle`,
    10,
  )
  if (
    underlineStyleRaw !== undefined &&
    underlineStyleRaw !== 'single' &&
    underlineStyleRaw !== 'double'
  ) {
    throw new OfficeValidationError(
      'validation',
      `${field}.underlineStyle must be 'single' or 'double'`,
    )
  }
  const strikethrough = expectOptionalBoolean(value.strikethrough, `${field}.strikethrough`)
  const fontFamily = expectOptionalString(value.fontFamily, `${field}.fontFamily`, 128)
  const fontSize = expectOptionalNumber(value.fontSize, `${field}.fontSize`)
  if (fontSize !== undefined && (fontSize < 1 || fontSize > 409)) {
    throw new OfficeValidationError('validation', `${field}.fontSize must be within 1..409`)
  }
  const fontColor = expectStyleHexColor(value.fontColor, `${field}.fontColor`)
  const fillColor = expectStyleHexColor(value.fillColor, `${field}.fillColor`)
  const alignH = expectOptionalString(value.horizontalAlignment, `${field}.horizontalAlignment`, 20)
  if (alignH !== undefined && !(STYLE_ALIGNS_H as readonly string[]).includes(alignH)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.horizontalAlignment must be one of: ${STYLE_ALIGNS_H.join(', ')}`,
    )
  }
  const alignV = expectOptionalString(value.verticalAlignment, `${field}.verticalAlignment`, 20)
  if (alignV !== undefined && !(STYLE_ALIGNS_V as readonly string[]).includes(alignV)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.verticalAlignment must be one of: ${STYLE_ALIGNS_V.join(', ')}`,
    )
  }
  const wrapText = expectOptionalBoolean(value.wrapText, `${field}.wrapText`)
  const textRotation = expectOptionalNumber(value.textRotation, `${field}.textRotation`)
  if (textRotation !== undefined && (textRotation < 0 || textRotation > 255)) {
    throw new OfficeValidationError('validation', `${field}.textRotation must be within 0..255`)
  }
  const indent = expectOptionalNumber(value.indent, `${field}.indent`)
  if (indent !== undefined && (indent < 0 || indent > 15)) {
    throw new OfficeValidationError('validation', `${field}.indent must be within 0..15`)
  }
  const numberFormat = expectOptionalString(value.numberFormat, `${field}.numberFormat`, 255)
  const protectionLocked = expectOptionalBoolean(
    value.protectionLocked,
    `${field}.protectionLocked`,
  )
  const protectionHidden = expectOptionalBoolean(
    value.protectionHidden,
    `${field}.protectionHidden`,
  )
  const borderTop =
    value.borderTop !== undefined && value.borderTop !== null
      ? expectStyleBorder(value.borderTop, `${field}.borderTop`)
      : undefined
  const borderBottom =
    value.borderBottom !== undefined && value.borderBottom !== null
      ? expectStyleBorder(value.borderBottom, `${field}.borderBottom`)
      : undefined
  const borderLeft =
    value.borderLeft !== undefined && value.borderLeft !== null
      ? expectStyleBorder(value.borderLeft, `${field}.borderLeft`)
      : undefined
  const borderRight =
    value.borderRight !== undefined && value.borderRight !== null
      ? expectStyleBorder(value.borderRight, `${field}.borderRight`)
      : undefined
  return {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(underlineStyleRaw !== undefined
      ? { underlineStyle: underlineStyleRaw as 'single' | 'double' }
      : {}),
    ...(strikethrough !== undefined ? { strikethrough } : {}),
    ...(fontFamily !== undefined ? { fontFamily } : {}),
    ...(fontSize !== undefined ? { fontSize } : {}),
    ...(fontColor !== undefined ? { fontColor } : {}),
    ...(fillColor !== undefined ? { fillColor } : {}),
    ...(alignH !== undefined
      ? { horizontalAlignment: alignH as WorkbookStyleEdit['horizontalAlignment'] }
      : {}),
    ...(alignV !== undefined
      ? { verticalAlignment: alignV as WorkbookStyleEdit['verticalAlignment'] }
      : {}),
    ...(wrapText !== undefined ? { wrapText } : {}),
    ...(textRotation !== undefined ? { textRotation } : {}),
    ...(indent !== undefined ? { indent } : {}),
    ...(protectionLocked !== undefined ? { protectionLocked } : {}),
    ...(protectionHidden !== undefined ? { protectionHidden } : {}),
    ...(numberFormat !== undefined ? { numberFormat } : {}),
    ...(borderTop !== undefined ? { borderTop } : {}),
    ...(borderBottom !== undefined ? { borderBottom } : {}),
    ...(borderLeft !== undefined ? { borderLeft } : {}),
    ...(borderRight !== undefined ? { borderRight } : {}),
  }
}

function expectCellEdit(value: unknown, index: number): CellEdit {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `cellEdits[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `cellEdits[${index}].sheetName`)
  const row = expectNumber(value.row, `cellEdits[${index}].row`)
  const column = expectNumber(value.column, `cellEdits[${index}].column`)
  const writeValue = expectBoolean(value.writeValue, `cellEdits[${index}].writeValue`)
  const cell = expectCellState(value.cell, `cellEdits[${index}].cell`)
  // Style delta (Excel cell formatting): strictly validated typed fields —
  // the browser can never inject raw XML through the style payload.
  let style: WorkbookStyleEdit | undefined
  if (value.style !== undefined && value.style !== null) {
    style = expectStyleEdit(value.style, `cellEdits[${index}].style`)
  }
  const edit: CellEdit = {
    sheetName,
    row,
    column,
    writeValue,
    cell,
    ...(style !== undefined ? { style } : {}),
    ...(Array.isArray(value.rich) ? { rich: value.rich as CellEdit['rich'] } : {}),
    ...(typeof value.styleReset === 'boolean' ? { styleReset: value.styleReset } : {}),
  }
  return edit
}

/**
 * Validate a `SerializedRun` from the wire. The browser sends these inside
 * blocks for the DOCX save route — they carry the inline formatting marks
 * (bold, italic, underline, strike, link) the editor round-trips.
 */
function expectSerializedRun(value: unknown, index: number): SerializedRun {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `runs[${index}] must be an object`)
  }
  const text = expectString(value.text, `runs[${index}].text`, false)
  const bold = value.bold
  if (bold !== undefined && typeof bold !== 'boolean') {
    throw new OfficeValidationError('validation', `runs[${index}].bold must be a boolean`)
  }
  const italic = value.italic
  if (italic !== undefined && typeof italic !== 'boolean') {
    throw new OfficeValidationError('validation', `runs[${index}].italic must be a boolean`)
  }
  const underline = value.underline
  if (underline !== undefined && typeof underline !== 'boolean') {
    throw new OfficeValidationError('validation', `runs[${index}].underline must be a boolean`)
  }
  const strike = value.strike
  if (strike !== undefined && typeof strike !== 'boolean') {
    throw new OfficeValidationError('validation', `runs[${index}].strike must be a boolean`)
  }
  let link: { href: string; tooltip?: string } | undefined
  if (value.link !== undefined) {
    if (!isRecord(value.link)) {
      throw new OfficeValidationError('validation', `runs[${index}].link must be an object`)
    }
    const href = expectString(value.link.href, `runs[${index}].link.href`)
    const tooltip =
      value.link.tooltip !== undefined
        ? expectString(value.link.tooltip, `runs[${index}].link.tooltip`, false)
        : undefined
    link = { href, ...(tooltip !== undefined ? { tooltip } : {}) }
  }
  return {
    text,
    ...(bold === true ? { bold: true } : {}),
    ...(italic === true ? { italic: true } : {}),
    ...(underline === true ? { underline: true } : {}),
    ...(strike === true ? { strike: true } : {}),
    ...(link !== undefined ? { link } : {}),
  }
}

// ── Table payload validation (mirrors SerializedRun/SerializedBlock rigor) ──

/** Size guards: a malformed table payload must never become a memory bomb. */
const MAX_TABLE_ROWS = 1000
const MAX_TABLE_COLS = 63

function expectOptionalNumber(value: unknown, field: string): number | undefined {
  if (value === undefined || value === null) return undefined
  return expectNumber(value, field)
}

function expectOptionalBoolean(value: unknown, field: string): boolean | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'boolean') {
    throw new OfficeValidationError('validation', `${field} must be a boolean`)
  }
  return value
}

function expectOptionalString(value: unknown, field: string, maxLength = 200): string | undefined {
  if (value === undefined || value === null) return undefined
  const s = expectString(value, field, false)
  if (s.length > maxLength) {
    throw new OfficeValidationError('validation', `${field} exceeds ${maxLength} characters`)
  }
  return s
}

function expectHexColor(value: unknown, field: string): string | undefined {
  if (value === undefined || value === null) return undefined
  const s = expectString(value, field)
  if (!/^[0-9A-Za-z]{0,8}$/.test(s)) {
    throw new OfficeValidationError('validation', `${field} must be a hex color without '#'`)
  }
  return s
}

function expectCellBorder(value: unknown, field: string): SerializedCellBorder {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const style = expectString(value.style, `${field}.style`)
  if (style.length > 40) {
    throw new OfficeValidationError('validation', `${field}.style is too long`)
  }
  const szEighths = expectOptionalNumber(value.szEighths, `${field}.szEighths`)
  if (szEighths !== undefined && (szEighths < 0 || szEighths > 96)) {
    throw new OfficeValidationError('validation', `${field}.szEighths must be within 0..96`)
  }
  const color = expectHexColor(value.color, `${field}.color`)
  return {
    style,
    ...(szEighths !== undefined ? { szEighths } : {}),
    ...(color !== undefined ? { color } : {}),
  }
}

function expectCellBorders(value: unknown, field: string): SerializedCellBorders {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const out: { -readonly [K in keyof SerializedCellBorders]: SerializedCellBorders[K] } = {}
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    if (value[edge] !== undefined && value[edge] !== null) {
      out[edge] = expectCellBorder(value[edge], `${field}.${edge}`)
    }
  }
  return out
}

function expectTableBorders(value: unknown, field: string): SerializedTableBorders {
  const base = expectCellBorders(value, field)
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const out: { -readonly [K in keyof SerializedTableBorders]: SerializedTableBorders[K] } = {
    ...base,
  }
  for (const edge of ['insideH', 'insideV'] as const) {
    if (value[edge] !== undefined && value[edge] !== null) {
      out[edge] = expectCellBorder(value[edge], `${field}.${edge}`)
    }
  }
  return out
}

function expectCellMargins(value: unknown, field: string): SerializedCellMargins {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const out: { -readonly [K in keyof SerializedCellMargins]: SerializedCellMargins[K] } = {}
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const v = expectOptionalNumber(value[side], `${field}.${side}`)
    if (v !== undefined) {
      if (v < 0 || v > 31680) {
        throw new OfficeValidationError(
          'validation',
          `${field}.${side} must be within 0..31680 twips`,
        )
      }
      out[side] = v
    }
  }
  return out
}

function expectTableParagraph(value: unknown, field: string): SerializedTableParagraph {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const runs = expectArray(value.runs, `${field}.runs`, (r, i) => expectSerializedRun(r, i))
  if (runs.length > 200) {
    throw new OfficeValidationError('validation', `${field}.runs exceeds 200 entries`)
  }
  const alignRaw = expectOptionalString(value.align, `${field}.align`, 20)
  if (
    alignRaw !== undefined &&
    !['left', 'center', 'right', 'justify', 'distribute'].includes(alignRaw)
  ) {
    throw new OfficeValidationError('validation', `${field}.align must be a paragraph alignment`)
  }
  const styleId = expectOptionalString(value.styleId, `${field}.styleId`, 64)
  return {
    runs,
    ...(alignRaw !== undefined ? { align: alignRaw as SerializedTableParagraph['align'] } : {}),
    ...(styleId !== undefined ? { styleId } : {}),
  }
}

function expectTableCell(value: unknown, field: string): SerializedTableCell {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const paras = expectArray(value.paras, `${field}.paras`, (p) => {
    if (typeof p !== 'string') {
      throw new OfficeValidationError('validation', `${field}.paras entries must be strings`)
    }
    return p
  })
  if (paras.length > 200) {
    throw new OfficeValidationError('validation', `${field}.paras exceeds 200 entries`)
  }
  const richParas =
    value.richParas !== undefined && value.richParas !== null
      ? expectArray(value.richParas, `${field}.richParas`, (p, i) =>
          expectTableParagraph(p, `${field}.richParas[${i}]`),
        )
      : undefined
  const colSpan = expectOptionalNumber(value.colSpan, `${field}.colSpan`)
  if (
    colSpan !== undefined &&
    (!Number.isInteger(colSpan) || colSpan < 1 || colSpan > MAX_TABLE_COLS)
  ) {
    throw new OfficeValidationError(
      'validation',
      `${field}.colSpan must be an integer 1..${MAX_TABLE_COLS}`,
    )
  }
  const vMergeRaw = expectOptionalString(value.vMerge, `${field}.vMerge`, 10)
  if (vMergeRaw !== undefined && vMergeRaw !== 'restart' && vMergeRaw !== 'continue') {
    throw new OfficeValidationError('validation', `${field}.vMerge must be 'restart' or 'continue'`)
  }
  const fill = expectHexColor(value.fill, `${field}.fill`)
  const color = expectHexColor(value.color, `${field}.color`)
  const bold = expectOptionalBoolean(value.bold, `${field}.bold`)
  const alignRaw = expectOptionalString(value.align, `${field}.align`, 20)
  if (
    alignRaw !== undefined &&
    !['left', 'center', 'right', 'justify', 'distribute'].includes(alignRaw)
  ) {
    throw new OfficeValidationError('validation', `${field}.align must be a paragraph alignment`)
  }
  const vAlignRaw = expectOptionalString(value.vAlign, `${field}.vAlign`, 10)
  if (vAlignRaw !== undefined && !['top', 'center', 'bottom'].includes(vAlignRaw)) {
    throw new OfficeValidationError('validation', `${field}.vAlign must be top/center/bottom`)
  }
  const borders =
    value.borders !== undefined && value.borders !== null
      ? expectCellBorders(value.borders, `${field}.borders`)
      : undefined
  // rawTcPr: original <w:tcPr> bytes echoed back for byte preservation.
  // Must look like a tcPr fragment and be bounded.
  const rawTcPrRaw = value.rawTcPr
  let rawTcPr: string | undefined
  if (rawTcPrRaw !== undefined && rawTcPrRaw !== null) {
    rawTcPr = expectString(rawTcPrRaw, `${field}.rawTcPr`, false)
    if (rawTcPr.length > 4096 || !/^<w:tcPr[\s>]/.test(rawTcPr) || !rawTcPr.endsWith('</w:tcPr>')) {
      throw new OfficeValidationError(
        'validation',
        `${field}.rawTcPr must be a <w:tcPr>…</w:tcPr> fragment (max 4096 chars)`,
      )
    }
    // Every element must be w:-namespaced; no comments or processing
    // instructions. The fragment is spliced verbatim into generated OOXML,
    // so foreign content must never ride through it.
    if (/<[!?]/.test(rawTcPr)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.rawTcPr may not contain comments or processing instructions`,
      )
    }
    const tagNames = rawTcPr.match(/<\/?([A-Za-z:][^\s>/]*)/g) ?? []
    for (const tag of tagNames) {
      const name = tag.replace(/^<\/?/, '')
      if (!name.startsWith('w:')) {
        throw new OfficeValidationError(
          'validation',
          `${field}.rawTcPr may only contain w:-namespaced elements (found <${name}>)`,
        )
      }
    }
  }
  return {
    paras,
    ...(richParas !== undefined ? { richParas } : {}),
    ...(colSpan !== undefined ? { colSpan } : {}),
    ...(vMergeRaw !== undefined ? { vMerge: vMergeRaw as 'restart' | 'continue' } : {}),
    ...(fill !== undefined ? { fill } : {}),
    ...(color !== undefined ? { color } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(alignRaw !== undefined ? { align: alignRaw as SerializedTableCell['align'] } : {}),
    ...(vAlignRaw !== undefined ? { vAlign: vAlignRaw as 'top' | 'center' | 'bottom' } : {}),
    ...(borders !== undefined ? { borders } : {}),
    ...(rawTcPr !== undefined ? { rawTcPr } : {}),
  }
}

/**
 * Validate a `SerializedTable` from the wire. Malformed tables throw
 * OfficeValidationError → the existing 400 validation error shape.
 */
function expectSerializedTable(value: unknown, field: string): SerializedTable {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const rows = expectArray(value.rows, `${field}.rows`, (row, ri) => {
    if (!Array.isArray(row)) {
      throw new OfficeValidationError('validation', `${field}.rows[${ri}] must be an array`)
    }
    if (row.length === 0 || row.length > MAX_TABLE_COLS) {
      throw new OfficeValidationError(
        'validation',
        `${field}.rows[${ri}] must have 1..${MAX_TABLE_COLS} cells`,
      )
    }
    return row.map((cell, ci) => expectTableCell(cell, `${field}.rows[${ri}][${ci}]`))
  })
  if (rows.length === 0 || rows.length > MAX_TABLE_ROWS) {
    throw new OfficeValidationError(
      'validation',
      `${field}.rows must have 1..${MAX_TABLE_ROWS} rows`,
    )
  }
  // Total grid width per row (colSpan sum) must stay bounded and consistent-ish;
  // vMerge continuations may legitimately shorten a row, so we only cap the max.
  for (let ri = 0; ri < rows.length; ri++) {
    const width = rows[ri].reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0)
    if (width > MAX_TABLE_COLS) {
      throw new OfficeValidationError(
        'validation',
        `${field}.rows[${ri}] spans ${width} grid columns (max ${MAX_TABLE_COLS})`,
      )
    }
  }
  const numArray = (v: unknown, f: string): readonly number[] | undefined =>
    v === undefined || v === null
      ? undefined
      : expectArray(v, f, (entry) => {
          if (typeof entry !== 'number' || !Number.isFinite(entry)) {
            throw new OfficeValidationError('validation', `${f} entries must be numbers`)
          }
          return entry
        })
  const colWidthsPct = numArray(value.colWidthsPct, `${field}.colWidthsPct`)
  if (colWidthsPct !== undefined && colWidthsPct.length > MAX_TABLE_COLS) {
    throw new OfficeValidationError(
      'validation',
      `${field}.colWidthsPct exceeds ${MAX_TABLE_COLS} entries`,
    )
  }
  const colWidthsTwips = numArray(value.colWidthsTwips, `${field}.colWidthsTwips`)
  if (colWidthsTwips !== undefined && colWidthsTwips.length > MAX_TABLE_COLS) {
    throw new OfficeValidationError(
      'validation',
      `${field}.colWidthsTwips exceeds ${MAX_TABLE_COLS} entries`,
    )
  }
  const widthPct = expectOptionalNumber(value.widthPct, `${field}.widthPct`)
  if (widthPct !== undefined && (widthPct <= 0 || widthPct > 100)) {
    throw new OfficeValidationError('validation', `${field}.widthPct must be within (0..100]`)
  }
  const autoLayout = expectOptionalBoolean(value.autoLayout, `${field}.autoLayout`)
  const cellMarTwips =
    value.cellMarTwips !== undefined && value.cellMarTwips !== null
      ? expectCellMargins(value.cellMarTwips, `${field}.cellMarTwips`)
      : undefined
  const borders =
    value.borders !== undefined && value.borders !== null
      ? expectTableBorders(value.borders, `${field}.borders`)
      : undefined
  const alignRaw = expectOptionalString(value.align, `${field}.align`, 10)
  if (alignRaw !== undefined && !['left', 'center', 'right'].includes(alignRaw)) {
    throw new OfficeValidationError('validation', `${field}.align must be left/center/right`)
  }
  const indentTwips = expectOptionalNumber(value.indentTwips, `${field}.indentTwips`)
  if (indentTwips !== undefined && Math.abs(indentTwips) > 31680) {
    throw new OfficeValidationError('validation', `${field}.indentTwips is out of range`)
  }
  const rowHeightsTwips =
    value.rowHeightsTwips !== undefined && value.rowHeightsTwips !== null
      ? expectArray(value.rowHeightsTwips, `${field}.rowHeightsTwips`, (entry) => {
          if (entry === null) return null
          if (typeof entry !== 'number' || !Number.isFinite(entry)) {
            throw new OfficeValidationError(
              'validation',
              `${field}.rowHeightsTwips entries must be numbers or null`,
            )
          }
          if (entry < 0 || entry > 31680) {
            throw new OfficeValidationError(
              'validation',
              `${field}.rowHeightsTwips entries must be within 0..31680`,
            )
          }
          return entry
        })
      : undefined
  const rowHeightRules =
    value.rowHeightRules !== undefined && value.rowHeightRules !== null
      ? expectArray(
          value.rowHeightRules,
          `${field}.rowHeightRules`,
          (entry): 'atLeast' | 'exact' | null => {
            if (entry === null) return null
            if (entry !== 'atLeast' && entry !== 'exact') {
              throw new OfficeValidationError(
                'validation',
                `${field}.rowHeightRules entries must be 'atLeast'|'exact'|null`,
              )
            }
            return entry
          },
        )
      : undefined
  const rawTrPrs =
    value.rawTrPrs !== undefined && value.rawTrPrs !== null
      ? expectArray(value.rawTrPrs, `${field}.rawTrPrs`, (entry) => {
          if (entry === null) return null
          if (typeof entry !== 'string') {
            throw new OfficeValidationError(
              'validation',
              `${field}.rawTrPrs entries must be strings or null`,
            )
          }
          if (entry.length > 4096 || !/^<w:trPr[\s>]/.test(entry) || !entry.endsWith('</w:trPr>')) {
            throw new OfficeValidationError(
              'validation',
              `${field}.rawTrPrs entries must be <w:trPr>…</w:trPr> fragments (max 4096 chars)`,
            )
          }
          if (/<[!?]/.test(entry)) {
            throw new OfficeValidationError(
              'validation',
              `${field}.rawTrPrs entries may not contain comments or processing instructions`,
            )
          }
          const trTagNames = entry.match(/<\/?([A-Za-z:][^\s>/]*)/g) ?? []
          for (const tag of trTagNames) {
            const name = tag.replace(/^<\/?/, '')
            if (!name.startsWith('w:')) {
              throw new OfficeValidationError(
                'validation',
                `${field}.rawTrPrs entries may only contain w:-namespaced elements (found <${name}>)`,
              )
            }
          }
          return entry
        })
      : undefined
  const tblStyleId = expectOptionalString(value.tblStyleId, `${field}.tblStyleId`, 64)
  const bidiVisual = expectOptionalBoolean(value.bidiVisual, `${field}.bidiVisual`)
  const headerRows =
    value.headerRows !== undefined && value.headerRows !== null
      ? expectArray(value.headerRows, `${field}.headerRows`, (entry) => {
          if (typeof entry !== 'boolean') {
            throw new OfficeValidationError(
              'validation',
              `${field}.headerRows entries must be booleans`,
            )
          }
          return entry
        })
      : undefined
  return {
    rows,
    ...(colWidthsPct !== undefined ? { colWidthsPct: [...colWidthsPct] } : {}),
    ...(colWidthsTwips !== undefined ? { colWidthsTwips: [...colWidthsTwips] } : {}),
    ...(widthPct !== undefined ? { widthPct } : {}),
    ...(autoLayout !== undefined ? { autoLayout } : {}),
    ...(cellMarTwips !== undefined ? { cellMarTwips } : {}),
    ...(borders !== undefined ? { borders } : {}),
    ...(alignRaw !== undefined ? { align: alignRaw as 'left' | 'center' | 'right' } : {}),
    ...(indentTwips !== undefined ? { indentTwips } : {}),
    ...(rowHeightsTwips !== undefined ? { rowHeightsTwips } : {}),
    ...(rowHeightRules !== undefined ? { rowHeightRules } : {}),
    ...(rawTrPrs !== undefined ? { rawTrPrs } : {}),
    ...(tblStyleId !== undefined ? { tblStyleId } : {}),
    ...(bidiVisual !== undefined ? { bidiVisual } : {}),
    ...(headerRows !== undefined ? { headerRows } : {}),
  }
}

// ── Image payload validation (mirrors the table/run validation rigor) ──────

/** Wrap modes the wire accepts (canonical ImageWrap + the inline marker). */
const IMAGE_WRAPS: readonly SerializedImageWrap[] = [
  'inline',
  'square-left',
  'square-right',
  'tight-left',
  'tight-right',
  'through-left',
  'through-right',
  'topBottom',
  'behind',
  'front',
]
const IMAGE_ALIGNS = ['left', 'center', 'right'] as const
const IMAGE_POS_H = ['left', 'center', 'right'] as const
const IMAGE_POS_V = ['top', 'center', 'bottom'] as const
const IMAGE_POS_H_RELS = ['margin', 'page', 'column', 'paragraph', 'character'] as const
const IMAGE_POS_V_RELS = ['margin', 'page', 'paragraph', 'line'] as const
/** display size bounds in CSS px (10k px ≈ 104 in — beyond any Word page) */
const MAX_IMAGE_DIM_PX = 10_000
/** posOffset bounds in EMU (±50M EMU ≈ ±54 in) */
const MAX_IMAGE_OFFSET_EMU = 50_000_000
/** base64 payload cap for a single image (~8 MB binary) */
const MAX_IMAGE_BASE64_CHARS = 11_000_000
const IMAGE_DATA_URL_RE = /^data:image\/(?:png|jpeg|gif);base64,[A-Za-z0-9+/=]*$/
const IMAGE_BASE64_RE = /^[A-Za-z0-9+/=]+$/
const IMAGE_MIMES = ['image/png', 'image/jpeg', 'image/gif'] as const
/**
 * Accessibility alt text bounds (wp:docPr descr). Word's UI caps the descr
 * field at ~125 chars but the format allows more; 500 is a generous ceiling
 * that still blocks oversized payloads. Control characters (except \t \n \r)
 * are stripped at the wire boundary so they cannot break the XML attribute
 * the canonical generator emits.
 */
const MAX_ALT_CHARS = 500

/**
 * Sanitize a wire alt-text value.
 *  - undefined → undefined (field absent: keep the existing descr)
 *  - null      → null       (explicit clear: remove the descr attribute)
 *  - ""        → null       (empty string is treated as clear)
 *  - non-empty → stripped string (set the descr)
 * Control characters (except \t \n \r — invalid in XML attrs) are removed
 * and length is bounded to MAX_ALT_CHARS.
 */
function sanitizeAltText(value: unknown, field: string): string | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const s = expectString(value, field, false)
  // strip control chars except \t (0x09), \n (0x0A), \r (0x0D) — invalid in XML attrs
  // (intentional: XML 1.0 forbids these control chars in attribute values)
  // eslint-disable-next-line no-control-regex
  const stripped = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
  if (stripped.length > MAX_ALT_CHARS) {
    throw new OfficeValidationError('validation', `${field} exceeds ${MAX_ALT_CHARS} characters`)
  }
  return stripped.length > 0 ? stripped : null
}

function expectImageRect(value: unknown, field: string): SerializedImageRect {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object {l,t,r,b}`)
  }
  const rect: { l?: number; t?: number; r?: number; b?: number } = {}
  for (const side of ['l', 't', 'r', 'b'] as const) {
    const v = value[side]
    if (v === undefined || v === null) {
      rect[side] = 0
      continue
    }
    if (typeof v !== 'number' || !Number.isFinite(v)) {
      throw new OfficeValidationError('validation', `${field}.${side} must be a number`)
    }
    if (v < 0 || v > 1) {
      throw new OfficeValidationError(
        'validation',
        `${field}.${side} must be a fraction within 0..1`,
      )
    }
    rect[side] = v
  }
  return { l: rect.l ?? 0, t: rect.t ?? 0, r: rect.r ?? 0, b: rect.b ?? 0 }
}

function expectImageDim(value: unknown, field: string): number | undefined {
  const v = expectOptionalNumber(value, field)
  if (v !== undefined && (!Number.isInteger(v) || v < 1 || v > MAX_IMAGE_DIM_PX)) {
    throw new OfficeValidationError(
      'validation',
      `${field} must be an integer 1..${MAX_IMAGE_DIM_PX}`,
    )
  }
  return v
}

function expectImageOffset(value: unknown, field: string): number | undefined {
  const v = expectOptionalNumber(value, field)
  if (v !== undefined && (!Number.isInteger(v) || Math.abs(v) > MAX_IMAGE_OFFSET_EMU)) {
    throw new OfficeValidationError(
      'validation',
      `${field} must be an integer within ±${MAX_IMAGE_OFFSET_EMU} EMU`,
    )
  }
  return v
}

function expectEnumString<T extends string>(
  value: unknown,
  allowed: readonly T[],
  field: string,
): T | undefined {
  const s = expectOptionalString(value, field, 20)
  if (s === undefined) return undefined
  if (!allowed.includes(s as T)) {
    throw new OfficeValidationError('validation', `${field} must be one of: ${allowed.join(', ')}`)
  }
  return s as T
}

/**
 * Validate a `SerializedImage` from the wire. Malformed images throw
 * OfficeValidationError → the existing 400 validation error shape.
 */
function expectSerializedImage(value: unknown, field: string): SerializedImage {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  let imageDataUrl: string | null = null
  if (value.imageDataUrl !== undefined && value.imageDataUrl !== null) {
    const s = expectString(value.imageDataUrl, `${field}.imageDataUrl`)
    if (!IMAGE_DATA_URL_RE.test(s) || s.length > MAX_IMAGE_BASE64_CHARS) {
      throw new OfficeValidationError(
        'validation',
        `${field}.imageDataUrl must be a data:image/(png|jpeg|gif);base64 URL (max ${MAX_IMAGE_BASE64_CHARS} chars)`,
      )
    }
    imageDataUrl = s
  }
  const widthPx = expectImageDim(value.widthPx, `${field}.widthPx`)
  const heightPx = expectImageDim(value.heightPx, `${field}.heightPx`)
  if ((widthPx === undefined) !== (heightPx === undefined)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.widthPx and ${field}.heightPx must be present together`,
    )
  }
  const crop =
    value.crop !== undefined && value.crop !== null
      ? expectImageRect(value.crop, `${field}.crop`)
      : undefined
  const fillRect =
    value.fillRect !== undefined && value.fillRect !== null
      ? expectImageRect(value.fillRect, `${field}.fillRect`)
      : undefined
  const align = expectEnumString(value.align, IMAGE_ALIGNS, `${field}.align`)
  const wrap = expectEnumString(value.wrap, IMAGE_WRAPS, `${field}.wrap`)
  const offsetXEmu = expectImageOffset(value.offsetXEmu, `${field}.offsetXEmu`)
  const offsetYEmu = expectImageOffset(value.offsetYEmu, `${field}.offsetYEmu`)
  const posH = expectEnumString(value.posH, IMAGE_POS_H, `${field}.posH`)
  const posV = expectEnumString(value.posV, IMAGE_POS_V, `${field}.posV`)
  const posHRel = expectEnumString(value.posHRel, IMAGE_POS_H_RELS, `${field}.posHRel`)
  const posVRel = expectEnumString(value.posVRel, IMAGE_POS_V_RELS, `${field}.posVRel`)
  const rotRaw = expectOptionalNumber(value.rotDeg, `${field}.rotDeg`)
  if (rotRaw !== undefined && (!Number.isInteger(rotRaw) || rotRaw < 0 || rotRaw > 359)) {
    throw new OfficeValidationError('validation', `${field}.rotDeg must be an integer 0..359`)
  }
  const flipH = expectOptionalBoolean(value.flipH, `${field}.flipH`)
  const flipV = expectOptionalBoolean(value.flipV, `${field}.flipV`)
  const altSanitized = sanitizeAltText(value.alt, `${field}.alt`)
  return {
    imageDataUrl,
    ...(widthPx !== undefined ? { widthPx } : {}),
    ...(heightPx !== undefined ? { heightPx } : {}),
    ...(crop !== undefined ? { crop } : {}),
    ...(fillRect !== undefined ? { fillRect } : {}),
    ...(align !== undefined ? { align } : {}),
    ...(wrap !== undefined ? { wrap } : {}),
    ...(offsetXEmu !== undefined ? { offsetXEmu } : {}),
    ...(offsetYEmu !== undefined ? { offsetYEmu } : {}),
    ...(posH !== undefined ? { posH } : {}),
    ...(posV !== undefined ? { posV } : {}),
    ...(posHRel !== undefined ? { posHRel } : {}),
    ...(posVRel !== undefined ? { posVRel } : {}),
    ...(rotRaw !== undefined ? { rotDeg: rotRaw } : {}),
    ...(flipH !== undefined ? { flipH } : {}),
    ...(flipV !== undefined ? { flipV } : {}),
    ...(altSanitized !== undefined ? { alt: altSanitized } : {}),
  }
}

/**
 * Validate a `SerializedNewImage` (image insertion) from the wire.
 */
function expectSerializedNewImage(value: unknown, field: string): SerializedNewImage {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const base64 = expectString(value.base64, `${field}.base64`)
  if (
    base64.length < 32 ||
    base64.length > MAX_IMAGE_BASE64_CHARS ||
    !IMAGE_BASE64_RE.test(base64)
  ) {
    throw new OfficeValidationError(
      'validation',
      `${field}.base64 must be standard base64 (32..${MAX_IMAGE_BASE64_CHARS} chars)`,
    )
  }
  const mimeRaw = expectString(value.mime, `${field}.mime`)
  if (!IMAGE_MIMES.includes(mimeRaw as (typeof IMAGE_MIMES)[number])) {
    throw new OfficeValidationError(
      'validation',
      `${field}.mime must be one of: ${IMAGE_MIMES.join(', ')}`,
    )
  }
  const widthPx = expectImageDim(value.widthPx, `${field}.widthPx`)
  const heightPx = expectImageDim(value.heightPx, `${field}.heightPx`)
  if (widthPx === undefined || heightPx === undefined) {
    throw new OfficeValidationError(
      'validation',
      `${field}.widthPx and ${field}.heightPx are required for a new image`,
    )
  }
  const align = expectEnumString(value.align, IMAGE_ALIGNS, `${field}.align`)
  const wrap = expectEnumString(value.wrap, IMAGE_WRAPS, `${field}.wrap`)
  const rotRaw = expectOptionalNumber(value.rotDeg, `${field}.rotDeg`)
  if (rotRaw !== undefined && (!Number.isInteger(rotRaw) || rotRaw < 0 || rotRaw > 359)) {
    throw new OfficeValidationError('validation', `${field}.rotDeg must be an integer 0..359`)
  }
  const flipH = expectOptionalBoolean(value.flipH, `${field}.flipH`)
  const flipV = expectOptionalBoolean(value.flipV, `${field}.flipV`)
  return {
    base64,
    mime: mimeRaw as SerializedNewImage['mime'],
    widthPx,
    heightPx,
    ...(align !== undefined ? { align } : {}),
    ...(wrap !== undefined ? { wrap } : {}),
    ...(rotRaw !== undefined ? { rotDeg: rotRaw } : {}),
    ...(flipH !== undefined ? { flipH } : {}),
    ...(flipV !== undefined ? { flipV } : {}),
  }
}

/**
 * Validate a `SerializedBlock` from the wire. The browser sends these for
 * the DOCX save route.
 */
function expectSerializedBlock(value: unknown, index: number): SerializedBlock {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `blocks[${index}] must be an object`)
  }
  const docxIndex =
    value.docxIndex === null || value.docxIndex === undefined
      ? null
      : expectNumber(value.docxIndex, `blocks[${index}].docxIndex`)
  const type = expectString(value.type, `blocks[${index}].type`)
  const allowedTypes: readonly SerializedBlock['type'][] = [
    'paragraph',
    'heading',
    'listItem',
    'table',
    'image',
    'passthrough',
    'hidden',
  ]
  if (!allowedTypes.includes(type as SerializedBlock['type'])) {
    throw new OfficeValidationError(
      'validation',
      `blocks[${index}].type must be one of: ${allowedTypes.join(', ')}`,
    )
  }
  const text = expectString(value.text, `blocks[${index}].text`, false)
  // Runs carry the inline formatting marks. Without this field the server
  // would regenerate edited blocks as plain text, silently stripping every
  // bold/italic/underline/strike/link mark the browser editor applied.
  const runs =
    value.runs !== undefined
      ? expectArray(value.runs, `blocks[${index}].runs`, expectSerializedRun)
      : undefined
  // Editable table payload. Only valid on table blocks; a non-table block
  // carrying a table payload (or vice versa) is malformed.
  let table: SerializedTable | undefined
  if (value.table !== undefined && value.table !== null) {
    if (type !== 'table') {
      throw new OfficeValidationError(
        'validation',
        `blocks[${index}].table is only allowed on type 'table' blocks`,
      )
    }
    table = expectSerializedTable(value.table, `blocks[${index}].table`)
  }
  if (type === 'table' && value.edited === true && table === undefined) {
    throw new OfficeValidationError(
      'validation',
      `blocks[${index}] is an edited table but carries no table payload`,
    )
  }
  // Editable image payload. Only valid on image blocks; the payload is the
  // edited state the server diffs against the parsed original.
  let image: SerializedImage | undefined
  if (value.image !== undefined && value.image !== null) {
    if (type !== 'image') {
      throw new OfficeValidationError(
        'validation',
        `blocks[${index}].image is only allowed on type 'image' blocks`,
      )
    }
    image = expectSerializedImage(value.image, `blocks[${index}].image`)
  }
  // New-image embedding spec. Only valid for blocks the editor created
  // (docxIndex === null) — an original block never re-embeds as new media.
  let newImage: SerializedNewImage | undefined
  if (value.newImage !== undefined && value.newImage !== null) {
    if (type !== 'image' || docxIndex !== null) {
      throw new OfficeValidationError(
        'validation',
        `blocks[${index}].newImage is only allowed on new image blocks (docxIndex null)`,
      )
    }
    newImage = expectSerializedNewImage(value.newImage, `blocks[${index}].newImage`)
  }
  if (type === 'image' && docxIndex === null && newImage === undefined) {
    throw new OfficeValidationError(
      'validation',
      `blocks[${index}] is a new image block but carries no newImage payload`,
    )
  }
  if (type === 'image' && value.edited === true && docxIndex !== null && image === undefined) {
    throw new OfficeValidationError(
      'validation',
      `blocks[${index}] is an edited image but carries no image payload`,
    )
  }
  const block: SerializedBlock = {
    docxIndex,
    type: type as SerializedBlock['type'],
    text,
    ...(runs !== undefined ? { runs } : {}),
    ...(table !== undefined ? { table } : {}),
    ...(image !== undefined ? { image } : {}),
    ...(newImage !== undefined ? { newImage } : {}),
    ...(value.level !== undefined
      ? { level: expectNumber(value.level, `blocks[${index}].level`) }
      : {}),
    ...(value.listKind !== undefined
      ? (() => {
          const k = value.listKind
          if (k !== 'bullet' && k !== 'ordered') {
            throw new OfficeValidationError(
              'validation',
              `blocks[${index}].listKind must be 'bullet' or 'ordered'`,
            )
          }
          return { listKind: k }
        })()
      : {}),
    ...(typeof value.edited === 'boolean' ? { edited: value.edited } : {}),
    ...(typeof value.hidden === 'boolean' ? { hidden: value.hidden } : {}),
  }
  return block
}

// ── Parsers: body → typed request (replace ad-hoc asObject/asArray) ──────────

function validateFileName(name: unknown): string {
  expectString(name, 'fileName')
  if (typeof name !== 'string')
    throw new OfficeValidationError('validation', 'fileName is required') // narrowed above
  if (name.length > 255) {
    throw new OfficeValidationError('validation', 'fileName exceeds 255 characters')
  }
  // Reject any path separator or traversal attempt — even percent-encoded.
  if (name.includes('/') || name.includes('\\') || name.includes('..') || name.includes('\0')) {
    throw new OfficeValidationError(
      'validation',
      'fileName may not contain path separators or traversal sequences',
    )
  }
  if (!SAFE_NAME_RE.test(name)) {
    throw new OfficeValidationError(
      'validation',
      'fileName contains invalid characters (allowed: A-Z a-z 0-9 . - _)',
    )
  }
  const lower = name.toLowerCase()
  const ext = lower.match(/\.[^.]+$/)?.[0] ?? ''
  if (!ALLOWED_EXTENSIONS.has(ext)) {
    throw new OfficeValidationError(
      'validation',
      `fileName must end with one of: ${[...ALLOWED_EXTENSIONS].join(', ')}`,
    )
  }
  return name
}

function decodeFileBytes(b64: unknown, codec: OfficeBinaryCodec): Uint8Array {
  expectString(b64, 'fileBytes')
  if (typeof b64 !== 'string')
    throw new OfficeValidationError('validation', 'fileBytes (base64) is required') // narrowed above
  let bytes: Uint8Array
  try {
    bytes = codec.decode(b64)
  } catch {
    throw new OfficeValidationError('validation', 'fileBytes is not valid base64')
  }
  if (bytes.length === 0) {
    throw new OfficeValidationError('validation', 'fileBytes decoded to an empty buffer')
  }
  if (bytes.length > MAX_FILE_BYTES) {
    throw new OfficeValidationError(
      'validation',
      `fileBytes exceeds the 10MB limit (decoded to ${bytes.length} bytes)`,
    )
  }
  return bytes
}

function encodeFileBytes(bytes: Uint8Array, codec: OfficeBinaryCodec): string {
  return codec.encode(bytes)
}

/**
 * Parsed (decoded) form of an open-workbook request. The wire type
 * (`OpenWorkbookRequest`) carries the base64-encoded fileBytes string; the
 * parsed form carries the decoded Uint8Array the engine consumes.
 */
interface ParsedOpenWorkbookRequest {
  readonly fileName: string
  readonly fileBytes: Uint8Array
}

/**
 * Parsed (decoded) form of a workbook save request. Same shape as
 * `BrowserWorkbookSaveRequest` except `fileBytes` is decoded to bytes and
 * `savePlan.edits` is type-checked at runtime (the wire type only guarantees
 * JSON shape, not CellEdit conformance).
 */
interface ParsedSaveWorkbookRequest {
  readonly fileName: string
  readonly fileBytes: Uint8Array
  readonly savePlan: BrowserWorkbookSavePlan
}

interface ParsedOpenDocumentRequest {
  readonly fileName: string
  readonly fileBytes: Uint8Array
}

interface ParsedSaveDocumentRequest {
  readonly fileName: string
  readonly fileBytes: Uint8Array
  readonly blocks: readonly SerializedBlock[]
}

/**
 * Parse and validate an `OpenWorkbookRequest` from the raw request body.
 * Throws OfficeValidationError on any shape violation.
 */
function parseOpenWorkbookRequest(
  body: unknown,
  codec: OfficeBinaryCodec,
): ParsedOpenWorkbookRequest {
  if (!isRecord(body)) {
    throw new OfficeValidationError('validation', 'Request body must be a JSON object')
  }
  const fileName = validateFileName(body.fileName)
  const fileBytes = decodeFileBytes(body.fileBytes, codec)
  return { fileName, fileBytes }
}

/**
 * Parse and validate a workbook save request. Accepts BOTH the canonical
 * `BrowserWorkbookSaveRequest` shape (`{ fileName, fileBytes, savePlan: { edits } }`)
 * and the legacy `{ fileName, fileBytes, cellEdits }` shape, normalizing
 * them into the canonical shape. This is the single backward-compat seam.
 */
// ── StructuralOp validation (row/column insert/delete) ──────────────────────

const ROWCOL_STRUCTURAL_KINDS = [
  'insert-rows',
  'remove-rows',
  'insert-cols',
  'remove-cols',
] as const
const MERGE_STRUCTURAL_KINDS = ['merge-cells', 'unmerge-cells'] as const
const SORT_STRUCTURAL_KIND = 'reorder-rows' as const
const STRUCTURAL_KINDS = [
  ...ROWCOL_STRUCTURAL_KINDS,
  ...MERGE_STRUCTURAL_KINDS,
  SORT_STRUCTURAL_KIND,
]
const MAX_STRUCTURAL_COUNT = 10_000
/// Upper bound on the sort-range row count — keeps the wire payload small
/// and matches Excel's hard 1,048,576-row ceiling without inviting a
/// multi-megabyte permutation map to block the event loop.
const MAX_REORDER_ROWS = 1_048_576

/**
 * Validate a StructuralOp from the wire. The browser emits
 * insert/remove row/column ops (with index/count), merge/unmerge ops
 * (with range), and sort/reorder-rows ops (with range + order permutation
 * map). Malformed ops throw OfficeValidationError → the existing 400
 * validation error shape.
 */
function expectStructuralOp(value: unknown, field: string): StructuralOp {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const kind = expectString(value.kind, `${field}.kind`)
  if (!(STRUCTURAL_KINDS as readonly string[]).includes(kind)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.kind must be one of: ${STRUCTURAL_KINDS.join(', ')}`,
    )
  }
  // merge-cells / unmerge-cells carry a range (not index/count). The
  // canonical StructuralOp uses the same range shape as the wire.
  if (kind === 'merge-cells' || kind === 'unmerge-cells') {
    const range = expectStructuralRange(value.range, `${field}.range`)
    return { kind, range } as unknown as StructuralOp
  }
  // reorder-rows (sort) carries a range + an order permutation map. The
  // map is in UNIVER'S NATIVE DEST→SRC shape: NEW[destRow] =
  // OLD[order[destRow]] (the ReorderRangeMutation reads getCellRaw(order
  // [row]) and writes it to row). Passed through verbatim — the gateway
  // inverts it internally. The gateway permutes <row> blocks atomically;
  // the entire cell record (styles, numfmt, formulas, hyperlinks) travels
  // with the row.
  if (kind === SORT_STRUCTURAL_KIND) {
    const range = expectStructuralRange(value.range, `${field}.range`)
    const order = expectReorderOrder(value.order, `${field}.order`, range)
    return { kind: SORT_STRUCTURAL_KIND, range, order } as unknown as StructuralOp
  }
  const index = expectNumber(value.index, `${field}.index`)
  if (!Number.isInteger(index) || index < 0 || index > 1_048_576) {
    throw new OfficeValidationError(
      'validation',
      `${field}.index must be a non-negative integer row/column index`,
    )
  }
  const count = expectNumber(value.count, `${field}.count`)
  if (!Number.isInteger(count) || count < 1 || count > MAX_STRUCTURAL_COUNT) {
    throw new OfficeValidationError(
      'validation',
      `${field}.count must be an integer 1..${MAX_STRUCTURAL_COUNT}`,
    )
  }
  return { kind: kind as (typeof ROWCOL_STRUCTURAL_KINDS)[number], index, count }
}

/**
 * Validate a reorder-rows `order` permutation map. The map is in Univer's
 * native DEST→SRC shape: keys are 0-based DESTINATION row indices inside
 * the sort range; values are 0-based SOURCE row indices inside the same
 * range (NEW[dest] = OLD[order[dest]]). The map must be a bijection over a
 * subset of the range's rows (every key and value sits inside
 * [startRow, endRow], no duplicates among values). Rows in the range
 * absent from the map stay put (Univer skips filtered/hidden/merge-child
 * rows — those are omitted from the order map and the gateway leaves them
 * alone).
 */
function expectReorderOrder(
  value: unknown,
  field: string,
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): Record<number, number> {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const { startRow, endRow } = range
  if (endRow < startRow) {
    throw new OfficeValidationError('validation', `${field}: range endRow < startRow`)
  }
  const rangeRowCount = endRow - startRow + 1
  if (rangeRowCount > MAX_REORDER_ROWS) {
    throw new OfficeValidationError(
      'validation',
      `${field}: sort range exceeds ${MAX_REORDER_ROWS} rows`,
    )
  }
  const out: Record<number, number> = {}
  const seenDest = new Set<number>()
  for (const key of Object.keys(value)) {
    const src = Number(key)
    if (!Number.isInteger(src) || src < 0) {
      throw new OfficeValidationError(
        'validation',
        `${field}: source row "${key}" must be a non-negative integer`,
      )
    }
    if (src < startRow || src > endRow) {
      throw new OfficeValidationError(
        'validation',
        `${field}: source row ${src} is outside the sort range [${startRow}, ${endRow}]`,
      )
    }
    const dest = expectNumber((value as Record<string, unknown>)[key], `${field}[${src}]`)
    if (!Number.isInteger(dest) || dest < 0) {
      throw new OfficeValidationError(
        'validation',
        `${field}[${src}] must be a non-negative integer`,
      )
    }
    if (dest < startRow || dest > endRow) {
      throw new OfficeValidationError(
        'validation',
        `${field}[${src}] = ${dest} is outside the sort range [${startRow}, ${endRow}]`,
      )
    }
    if (seenDest.has(dest)) {
      throw new OfficeValidationError(
        'validation',
        `${field}: destination row ${dest} is targeted by more than one source row (not a bijection)`,
      )
    }
    seenDest.add(dest)
    out[src] = dest
  }
  return out
}

/** Validate a range (startRow/endRow/startColumn/endColumn, 0-based integers). */
function expectStructuralRange(
  value: unknown,
  field: string,
): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const startRow = expectNumber(value.startRow, `${field}.startRow`)
  const endRow = expectNumber(value.endRow, `${field}.endRow`)
  const startColumn = expectNumber(value.startColumn, `${field}.startColumn`)
  const endColumn = expectNumber(value.endColumn, `${field}.endColumn`)
  for (const [n, label] of [
    [startRow, `${field}.startRow`],
    [endRow, `${field}.endRow`],
    [startColumn, `${field}.startColumn`],
    [endColumn, `${field}.endColumn`],
  ] as const) {
    if (!Number.isInteger(n) || n < 0) {
      throw new OfficeValidationError('validation', `${label} must be a non-negative integer`)
    }
  }
  if (endRow < startRow || endColumn < startColumn) {
    throw new OfficeValidationError('validation', `${field} end must be >= start`)
  }
  return { startRow, endRow, startColumn, endColumn }
}

function expectSheetStructuralOps(value: unknown, index: number): SheetStructuralOps {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `structuralOps[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `structuralOps[${index}].sheetName`)
  const ops = expectArray(value.ops, `structuralOps[${index}].ops`, (op, i) =>
    expectStructuralOp(op, `structuralOps[${index}].ops[${i}]`),
  )
  if (ops.length > 100) {
    throw new OfficeValidationError('validation', `structuralOps[${index}].ops exceeds 100 entries`)
  }
  return { sheetName, ops }
}

/**
 * Validate one per-sheet page-setup state from the wire. Only the
 * `frozenRows` / `frozenColumns` fields are wired by the web shell today;
 * the remaining optional SheetPageSetupState fields are accepted (and
 * type-validated) so future View commands can land here without a
 * wire-breaking change. Frozen-row/column counts are bounded by the
 * OOXML maximum row/column counts.
 */
function expectSheetPageSetupState(value: unknown, index: number): SheetPageSetupState {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `pageSetupStates[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `pageSetupStates[${index}].sheetName`)
  const out: {
    sheetName: string
    frozenRows?: number
    frozenColumns?: number
    [k: string]: unknown
  } = { sheetName }
  if (value.frozenRows !== undefined) {
    const n = expectNumber(value.frozenRows, `pageSetupStates[${index}].frozenRows`)
    if (!Number.isInteger(n) || n < 0 || n > 1_048_576) {
      throw new OfficeValidationError(
        'validation',
        `pageSetupStates[${index}].frozenRows must be a non-negative integer`,
      )
    }
    out.frozenRows = n
  }
  if (value.frozenColumns !== undefined) {
    const n = expectNumber(value.frozenColumns, `pageSetupStates[${index}].frozenColumns`)
    if (!Number.isInteger(n) || n < 0 || n > 16_384) {
      throw new OfficeValidationError(
        'validation',
        `pageSetupStates[${index}].frozenColumns must be a non-negative integer`,
      )
    }
    out.frozenColumns = n
  }
  // Forward-compatibility seam: ignore unknown keys (the canonical
  // SheetPageSetupState carries many optional fields — orientation,
  // paperSize, margins, … — that the web shell does not yet emit; future
  // increments can add validated readers here without bumping the wire).
  return out as SheetPageSetupState
}

// ── SheetFilterState validation (Data → Filter, Phase 4 Increment 4) ────────

/// Custom filter operators the canonical gateway can serialize — the exact
/// set Univer's CustomFilterOperator enum emits (and xlsx-filter.ts accepts).
const FILTER_CUSTOM_OPERATORS = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])

/** Upper bound on filter rows/cols — Excel's sheet dimensions. */
const MAX_FILTER_ROW_OR_COLUMN = 1_048_576
/** Excel's column count ceiling. */
const MAX_FILTER_COLUMN_INDEX = 16_384
/** Guard against absurd payloads (one filter per sheet, plus headroom). */
const MAX_FILTER_STATES = 100
/** Guard against absurd per-column value lists. */
const MAX_FILTER_VALUES = 10_000
/** OOXML allows at most two customFilters per customFilters element. */
const MAX_CUSTOM_FILTERS = 2

/**
 * Validate one per-sheet AutoFilter state from the wire. The browser only
 * emits states snapshotted from Univer's live filter model through the
 * canonical `SheetFilterState` shape — anything else (unknown kinds,
 * arbitrary filter objects, out-of-range coordinates, unsupported
 * operators) is rejected with a 400 rather than reaching the engine.
 * `filter: null` is the explicit cleared-filter state.
 */
function expectSheetFilterState(value: unknown, index: number): SheetFilterState {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `filterStates[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `filterStates[${index}].sheetName`)
  const out: {
    sheetName: string
    filter: {
      range: SheetFilterArea
      columns: readonly SheetFilterColumn[]
    } | null
    hiddenRows: readonly number[]
    visibilityRange: SheetFilterArea
  } = {
    sheetName,
    filter: null,
    hiddenRows: [],
    visibilityRange: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
  }
  if (value.filter !== null && value.filter !== undefined) {
    if (!isRecord(value.filter)) {
      throw new OfficeValidationError(
        'validation',
        `filterStates[${index}].filter must be an object or null`,
      )
    }
    const range = expectFilterArea(value.filter.range, `filterStates[${index}].filter.range`)
    const columns = expectArray(
      value.filter.columns,
      `filterStates[${index}].filter.columns`,
      (column, i) => expectFilterColumn(column, `filterStates[${index}].filter.columns[${i}]`),
    )
    if (columns.length > MAX_FILTER_COLUMN_INDEX) {
      throw new OfficeValidationError(
        'validation',
        `filterStates[${index}].filter.columns exceeds ${MAX_FILTER_COLUMN_INDEX} entries`,
      )
    }
    for (const column of columns) {
      if (column.colId < 0 || range.startColumn + column.colId > range.endColumn) {
        throw new OfficeValidationError(
          'validation',
          `filterStates[${index}]: column colId ${column.colId} is outside the filter range`,
        )
      }
    }
    out.filter = { range, columns }
  } else if (value.filter === undefined) {
    throw new OfficeValidationError(
      'validation',
      `filterStates[${index}].filter is required (object or null)`,
    )
  }
  const hiddenRows = expectArray(
    value.hiddenRows,
    `filterStates[${index}].hiddenRows`,
    (row, i) => {
      const n = expectNumber(row, `filterStates[${index}].hiddenRows[${i}]`)
      if (!Number.isInteger(n) || n < 0 || n > MAX_FILTER_ROW_OR_COLUMN) {
        throw new OfficeValidationError(
          'validation',
          `filterStates[${index}].hiddenRows[${i}] must be a 0-based row index`,
        )
      }
      return n
    },
  )
  out.hiddenRows = hiddenRows
  out.visibilityRange = expectFilterArea(
    value.visibilityRange,
    `filterStates[${index}].visibilityRange`,
  )
  return out as unknown as SheetFilterState
}

type SheetFilterArea = {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}
type SheetFilterColumn = {
  colId: number
  values?: readonly string[]
  blank?: boolean
  customs?: {
    and?: boolean
    filters: readonly { val: string | number; operator?: string }[]
  }
}

function expectFilterArea(value: unknown, field: string): SheetFilterArea {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const startRow = expectNumber(value.startRow, `${field}.startRow`)
  const endRow = expectNumber(value.endRow, `${field}.endRow`)
  const startColumn = expectNumber(value.startColumn, `${field}.startColumn`)
  const endColumn = expectNumber(value.endColumn, `${field}.endColumn`)
  for (const [n, label] of [
    [startRow, `${field}.startRow`],
    [endRow, `${field}.endRow`],
    [startColumn, `${field}.startColumn`],
    [endColumn, `${field}.endColumn`],
  ] as const) {
    if (!Number.isInteger(n) || n < 0 || n > MAX_FILTER_ROW_OR_COLUMN) {
      throw new OfficeValidationError('validation', `${label} must be a non-negative integer`)
    }
  }
  if (endRow < startRow || endColumn < startColumn) {
    throw new OfficeValidationError('validation', `${field} end must be >= start`)
  }
  if (endColumn > MAX_FILTER_COLUMN_INDEX) {
    throw new OfficeValidationError(
      'validation',
      `${field}.endColumn exceeds Excel's column ceiling`,
    )
  }
  return { startRow, endRow, startColumn, endColumn }
}

function expectFilterColumn(value: unknown, field: string): SheetFilterColumn {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const colId = expectNumber(value.colId, `${field}.colId`)
  if (!Number.isInteger(colId) || colId < 0 || colId > MAX_FILTER_COLUMN_INDEX) {
    throw new OfficeValidationError('validation', `${field}.colId must be a non-negative integer`)
  }
  const out: SheetFilterColumn = { colId }
  if (value.values !== undefined && value.values !== null) {
    const values = expectArray(value.values, `${field}.values`, (v, i) => {
      if (typeof v !== 'string') {
        throw new OfficeValidationError('validation', `${field}.values[${i}] must be a string`)
      }
      return v
    })
    if (values.length > MAX_FILTER_VALUES) {
      throw new OfficeValidationError(
        'validation',
        `${field}.values exceeds ${MAX_FILTER_VALUES} entries`,
      )
    }
    out.values = values
  }
  if (value.blank !== undefined) {
    if (typeof value.blank !== 'boolean') {
      throw new OfficeValidationError('validation', `${field}.blank must be a boolean`)
    }
    out.blank = value.blank
  }
  if (value.customs !== undefined && value.customs !== null) {
    if (!isRecord(value.customs)) {
      throw new OfficeValidationError('validation', `${field}.customs must be an object`)
    }
    if (value.customs.and !== undefined && typeof value.customs.and !== 'boolean') {
      throw new OfficeValidationError('validation', `${field}.customs.and must be a boolean`)
    }
    const filters = expectArray(value.customs.filters, `${field}.customs.filters`, (custom, i) => {
      if (!isRecord(custom)) {
        throw new OfficeValidationError(
          'validation',
          `${field}.customs.filters[${i}] must be an object`,
        )
      }
      const val = custom.val
      if (typeof val !== 'string' && typeof val !== 'number') {
        throw new OfficeValidationError(
          'validation',
          `${field}.customs.filters[${i}].val must be a string or number`,
        )
      }
      if (custom.operator !== undefined && custom.operator !== null) {
        if (typeof custom.operator !== 'string' || !FILTER_CUSTOM_OPERATORS.has(custom.operator)) {
          throw new OfficeValidationError(
            'validation',
            `${field}.customs.filters[${i}].operator "${String(custom.operator)}" is not a supported filter condition`,
          )
        }
        return { val, operator: custom.operator as string }
      }
      return { val }
    })
    if (filters.length === 0 || filters.length > MAX_CUSTOM_FILTERS) {
      throw new OfficeValidationError(
        'validation',
        `${field}.customs.filters must carry 1..${MAX_CUSTOM_FILTERS} conditions`,
      )
    }
    const and = value.customs.and === true
    out.customs = and ? { and: true, filters } : { filters }
  }
  if (out.values === undefined && !out.blank && out.customs === undefined) {
    throw new OfficeValidationError(
      'validation',
      `${field} carries no criteria (need values, blank, or customs)`,
    )
  }
  return out
}

// ── SheetDvState validation (Data → Data Validation, Phase 4 Increment 5) ────

/// Canonical DV types the gateway can serialize — the Univer DataValidationType
/// subset that maps to OOXML (plus 'any'/'none', the messages-only form).
const DV_TYPES = new Set([
  'whole',
  'decimal',
  'list',
  'date',
  'time',
  'textLength',
  'custom',
  'any',
  'none',
])
/// Canonical DV operators — identical to Univer's DataValidationOperator enum.
const DV_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
/// Univer DataValidationErrorStyle numbers the gateway understands.
const DV_ERROR_STYLES = new Set([0, 1, 2])

/** Caps mirroring the desktop's workbookDvStateSchema. */
const MAX_DV_STATES = 1_000
const MAX_DV_RULES_PER_SHEET = 500
const MAX_DV_RANGES_PER_RULE = 100
/** OOXML message fields are Excel-bounded to 255 characters. */
const MAX_DV_MESSAGE_LENGTH = 255
/** Excel's row/column ceilings. */
const MAX_DV_ROW = 1_048_575
const MAX_DV_COLUMN = 16_383

/**
 * Validate one per-sheet data-validation state from the wire. The browser
 * only emits full declarative snapshots of Univer's live validation model
 * (the same shape the desktop ships) — anything else (unknown types,
 * unknown operators, malformed ranges, oversized messages, excessive rule
 * counts, non-object rules) is rejected with a 400 rather than reaching the
 * engine. An empty rules array is VALID: it means "all validation on this
 * sheet was cleared".
 */
function expectSheetDvState(value: unknown, index: number): SheetDvState {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `dvStates[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `dvStates[${index}].sheetName`)
  const rules = expectArray(value.rules, `dvStates[${index}].rules`, (rule, i) =>
    expectDvWireRule(rule, `dvStates[${index}].rules[${i}]`),
  )
  if (rules.length > MAX_DV_RULES_PER_SHEET) {
    throw new OfficeValidationError(
      'validation',
      `dvStates[${index}].rules exceeds ${MAX_DV_RULES_PER_SHEET} entries`,
    )
  }
  return { sheetName, rules }
}

function expectDvWireRule(
  value: unknown,
  field: string,
): {
  ranges: ReadonlyArray<{
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
  }>
  rule: Record<string, unknown>
} {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const ranges = expectArray(value.ranges, `${field}.ranges`, (area, i) => {
    if (!isRecord(area)) {
      throw new OfficeValidationError('validation', `${field}.ranges[${i}] must be an object`)
    }
    return expectDvArea(area, `${field}.ranges[${i}]`)
  })
  if (ranges.length === 0 || ranges.length > MAX_DV_RANGES_PER_RULE) {
    throw new OfficeValidationError(
      'validation',
      `${field}.ranges must carry 1..${MAX_DV_RANGES_PER_RULE} areas`,
    )
  }
  if (!isRecord(value.rule)) {
    throw new OfficeValidationError('validation', `${field}.rule must be an object`)
  }
  const rule = expectDvRule(value.rule, `${field}.rule`)
  return { ranges, rule }
}

function expectDvArea(
  value: Record<string, unknown>,
  field: string,
): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
  const startRow = expectNumber(value.startRow, `${field}.startRow`)
  const endRow = expectNumber(value.endRow, `${field}.endRow`)
  const startColumn = expectNumber(value.startColumn, `${field}.startColumn`)
  const endColumn = expectNumber(value.endColumn, `${field}.endColumn`)
  for (const [n, label] of [
    [startRow, `${field}.startRow`],
    [endRow, `${field}.endRow`],
    [startColumn, `${field}.startColumn`],
    [endColumn, `${field}.endColumn`],
  ] as const) {
    if (!Number.isInteger(n) || n < 0) {
      throw new OfficeValidationError('validation', `${label} must be a non-negative integer`)
    }
  }
  if (endRow < startRow || endColumn < startColumn) {
    throw new OfficeValidationError('validation', `${field} end must be >= start`)
  }
  if (endRow > MAX_DV_ROW || endColumn > MAX_DV_COLUMN) {
    throw new OfficeValidationError('validation', `${field} exceeds Excel's sheet dimensions`)
  }
  return { startRow, endRow, startColumn, endColumn }
}

function expectDvRule(value: Record<string, unknown>, field: string): Record<string, unknown> {
  const type = value.type
  if (typeof type !== 'string' || !DV_TYPES.has(type)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.type "${String(type)}" is not a supported data-validation type`,
    )
  }
  const out: Record<string, unknown> = { type }
  if (value.operator !== undefined && value.operator !== null) {
    if (typeof value.operator !== 'string' || !DV_OPERATORS.has(value.operator)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.operator "${String(value.operator)}" is not a supported data-validation operator`,
      )
    }
    out.operator = value.operator
  }
  for (const key of ['formula1', 'formula2'] as const) {
    const v = value[key]
    if (v === undefined || v === null) continue
    if (typeof v !== 'string' && typeof v !== 'number') {
      throw new OfficeValidationError('validation', `${field}.${key} must be a string or number`)
    }
    const text = String(v)
    if (text.length > 1_000) {
      throw new OfficeValidationError('validation', `${field}.${key} exceeds 1000 characters`)
    }
    out[key] = text
  }
  if (value.allowBlank !== undefined) {
    if (typeof value.allowBlank !== 'boolean') {
      throw new OfficeValidationError('validation', `${field}.allowBlank must be a boolean`)
    }
    out.allowBlank = value.allowBlank
  }
  for (const key of ['showDropDown', 'showInputMessage', 'showErrorMessage'] as const) {
    if (value[key] === undefined) continue
    if (typeof value[key] !== 'boolean') {
      throw new OfficeValidationError('validation', `${field}.${key} must be a boolean`)
    }
    out[key] = value[key]
  }
  if (value.errorStyle !== undefined && value.errorStyle !== null) {
    const style = value.errorStyle
    if (typeof style !== 'number' || !Number.isInteger(style) || !DV_ERROR_STYLES.has(style)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.errorStyle "${String(style)}" is not a supported error style`,
      )
    }
    out.errorStyle = style
  }
  for (const key of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
    const v = value[key]
    if (v === undefined || v === null) continue
    if (typeof v !== 'string') {
      throw new OfficeValidationError('validation', `${field}.${key} must be a string`)
    }
    if (v.length > MAX_DV_MESSAGE_LENGTH) {
      throw new OfficeValidationError(
        'validation',
        `${field}.${key} exceeds ${MAX_DV_MESSAGE_LENGTH} characters`,
      )
    }
    out[key] = v
  }
  // Reject unknown extra keys: the canonical model carries only the fields
  // above plus uid/renderMode (browser-side chrome the serializer ignores).
  for (const key of Object.keys(value)) {
    if (
      ![
        'type',
        'operator',
        'formula1',
        'formula2',
        'allowBlank',
        'showDropDown',
        'showInputMessage',
        'showErrorMessage',
        'errorStyle',
        'errorTitle',
        'error',
        'promptTitle',
        'prompt',
        'uid',
        'renderMode',
      ].includes(key)
    ) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return out
}

// ── Protection validation (Review → Protection, EXCEL-020) ───────────────────────────────────────────────────────

/** Caps mirroring the desktop's workbookSheetProtectionSchema. */
const MAX_SHEET_PROTECTIONS = 1_000
const MAX_TABLE_ADDITIONS = 50
const TABLE_STYLE_PATTERN = /^TableStyle(?:Light|Medium|Dark)[1-9][0-9]?$/

// ── Visual validation (Insert → Picture / image edit, EXCEL-022) ─────────────

/** Caps for the visual families (image-only additions + surgical edits). */
const MAX_VISUAL_ADDITIONS = 50
const MAX_VISUAL_EDITS = 200
const MAX_VISUAL_IMAGE_BASE64_CHARS = 11_000_000
const MAX_ANCHOR_ROW = 1_048_575
const MAX_ANCHOR_COLUMN = 16_383
const MAX_ANCHOR_OFFSET_EMU = 50_000_000
const DRAWING_PATH_PATTERN = /^xl\/drawings\/[A-Za-z0-9._/-]+\.xml$/
const VISUAL_MEDIA_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif'])

/**
 * Validate one drawing anchor from the wire — eight non-negative bounded
 * integers (rows 0..1048575, columns 0..16383, EMU offsets 0..50M).
 */
function expectDrawingAnchor(
  value: unknown,
  field: string,
): {
  readonly fromRow: number
  readonly fromColumn: number
  readonly fromRowOffset: number
  readonly fromColumnOffset: number
  readonly toRow: number
  readonly toColumn: number
  readonly toRowOffset: number
  readonly toColumnOffset: number
} {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const bounds: Array<[string, number]> = [
    ['fromRow', MAX_ANCHOR_ROW],
    ['fromColumn', MAX_ANCHOR_COLUMN],
    ['fromRowOffset', MAX_ANCHOR_OFFSET_EMU],
    ['fromColumnOffset', MAX_ANCHOR_OFFSET_EMU],
    ['toRow', MAX_ANCHOR_ROW],
    ['toColumn', MAX_ANCHOR_COLUMN],
    ['toRowOffset', MAX_ANCHOR_OFFSET_EMU],
    ['toColumnOffset', MAX_ANCHOR_OFFSET_EMU],
  ]
  const out: Record<string, number> = {}
  for (const [key, max] of bounds) {
    const raw = value[key]
    if (typeof raw !== 'number' || !Number.isInteger(raw) || raw < 0 || raw > max) {
      throw new OfficeValidationError(
        'validation',
        `${field}.${key} must be an integer between 0 and ${max}`,
      )
    }
    out[key] = raw
  }
  for (const key of Object.keys(value)) {
    if (!bounds.some(([name]) => name === key)) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return out as {
    readonly fromRow: number
    readonly fromColumn: number
    readonly fromRowOffset: number
    readonly fromColumnOffset: number
    readonly toRow: number
    readonly toColumn: number
    readonly toRowOffset: number
    readonly toColumnOffset: number
  }
}

/**
 * Validate one visual addition from the wire (EXCEL-022 / EXCEL-023).
 * IMAGE or CHART: the payload carries a sheetName, a bounded drawing
 * anchor, and exactly one typed visual — an image (supported media type
 * + base64 bytes) or a chart (canonical ChartAdd with a supported chart
 * type, 1-24 bounded series, and bounded style options). Shape additions
 * remain rejected — the web editor has no shape UI. Unknown fields are
 * rejected; nothing unvalidated reaches the generic engine family.
 */
function expectSheetVisualAddition(value: unknown, index: number): SheetVisualAddition {
  const field = `visualAdditions[${index}]`
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  if (value.shape !== undefined) {
    throw new OfficeValidationError(
      'validation',
      `${field} carries an unsupported visual kind — image and chart additions are supported`,
    )
  }
  const hasImage = value.image !== undefined
  const hasChart = value.chart !== undefined
  if (hasImage && hasChart) {
    throw new OfficeValidationError(
      'validation',
      `${field} carries both image and chart — exactly one is required`,
    )
  }
  if (!hasImage && !hasChart) {
    throw new OfficeValidationError(
      'validation',
      `${field} needs an image or a chart payload`,
    )
  }
  const sheetName = expectString(value.sheetName, `${field}.sheetName`)
  const anchor = expectDrawingAnchor(value.anchor, `${field}.anchor`)
  let chart: ChartAdd | undefined
  if (hasChart) {
    chart = expectChartAdd(value.chart, `${field}.chart`)
  }
  let image: SheetVisualAddition['image'] | undefined
  if (hasImage) {
    if (!isRecord(value.image)) {
      throw new OfficeValidationError('validation', `${field}.image must be an object`)
    }
    const mediaTypeRaw = expectString(value.image.mediaType, `${field}.image.mediaType`)
    if (!VISUAL_MEDIA_TYPES.has(mediaTypeRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.image.mediaType must be one of image/png, image/jpeg, image/gif`,
      )
    }
    const mediaType = mediaTypeRaw as 'image/png' | 'image/jpeg' | 'image/gif'
    const base64 = expectString(value.image.base64, `${field}.image.base64`)
    if (base64.length > MAX_VISUAL_IMAGE_BASE64_CHARS) {
      throw new OfficeValidationError(
        'validation',
        `${field}.image.base64 exceeds ${MAX_VISUAL_IMAGE_BASE64_CHARS} characters`,
      )
    }
    for (const key of Object.keys(value.image)) {
      if (key !== 'mediaType' && key !== 'base64') {
        throw new OfficeValidationError(
          'validation',
          `${field}.image carries an unknown field "${key}"`,
        )
      }
    }
    image = { mediaType, base64 }
  }
  for (const key of Object.keys(value)) {
    if (!['sheetName', 'anchor', 'image', 'chart'].includes(key)) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return {
    sheetName,
    anchor,
    ...(image !== undefined ? { image } : {}),
    ...(chart !== undefined ? { chart } : {}),
  }
}

// ── Chart payload validation (EXCEL-023) ────────────────────────────

const CHART_PATH_PATTERN = /^xl\/charts\/[A-Za-z0-9._-]+\.xml$/
const CHART_HEX_COLOR_PATTERN = /^#[0-9A-Fa-f]{6}$/
/// Series/point index keys: 0-999 (desktop preload parity).
const CHART_INDEX_KEY_PATTERN = /^[0-9]{1,3}$/
const MAX_CHART_EDITS = 200
const MAX_CHART_SERIES = 24
const MAX_CHART_POINTS = 1_000

const CHART_ADD_TYPES = new Set([
  'column',
  'bar',
  'line',
  'area',
  'pie',
  'scatter',
  'radar',
  'doughnut',
  'combo',
])
const CHART_EDIT_TYPES = new Set(['column', 'bar', 'line', 'area', 'pie', 'doughnut'])
const CHART_LEGEND_POSITIONS = new Set(['none', 'right', 'bottom', 'top', 'left'])
const CHART_DATA_LABEL_MODES = new Set(['none', 'value', 'percent', 'category-percent'])
const CHART_LABEL_POSITIONS = new Set(['center', 'inside-end', 'outside-end'])
const CHART_GROUPINGS = new Set(['clustered', 'stacked', 'percentStacked', 'standard'])

function expectChartHexColor(value: unknown, field: string): string {
  const s = expectString(value, field)
  if (!CHART_HEX_COLOR_PATTERN.test(s)) {
    throw new OfficeValidationError('validation', `${field} must be a #RRGGBB hex color`)
  }
  return s
}

function expectChartIndexKey(value: unknown, field: string): string {
  const s = expectString(value, field)
  if (!CHART_INDEX_KEY_PATTERN.test(s)) {
    throw new OfficeValidationError('validation', `${field} must be a series/point index (0-999)`)
  }
  return s
}

function expectFiniteNumberArray(value: unknown, field: string): readonly number[] {
  const values = expectArray(value, field, (entry, position) => {
    if (typeof entry !== 'number' || !Number.isFinite(entry)) {
      throw new OfficeValidationError('validation', `${field}[${position}] must be a finite number`)
    }
    return entry
  })
  if (values.length > MAX_CHART_POINTS) {
    throw new OfficeValidationError('validation', `${field} exceeds ${MAX_CHART_POINTS} entries`)
  }
  return values
}

function expectBoundedStringArray(value: unknown, field: string): readonly string[] {
  const values = expectArray(value, field, (entry, position) => {
    const text = expectString(entry, `${field}[${position}]`)
    if (text.length > 255) {
      throw new OfficeValidationError('validation', `${field}[${position}] exceeds 255 characters`)
    }
    return text
  })
  if (values.length > MAX_CHART_POINTS) {
    throw new OfficeValidationError('validation', `${field} exceeds ${MAX_CHART_POINTS} entries`)
  }
  return values
}

function expectChartRefString(value: unknown, field: string): string | undefined {
  const s = expectString(value, field)
  if (s.length > 512) {
    throw new OfficeValidationError('validation', `${field} exceeds 512 characters`)
  }
  return s
}

/// One chart addition series (mirrors the desktop preload's
/// workbookVisualAddSchema.chart.series entry).
function expectChartAddSeries(value: unknown, field: string): ChartAddSeries {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const name = expectString(value.name, `${field}.name`)
  if (name.length > 255) {
    throw new OfficeValidationError('validation', `${field}.name exceeds 255 characters`)
  }
  const values = expectFiniteNumberArray(value.values, `${field}.values`)
  if (values.length === 0) {
    throw new OfficeValidationError('validation', `${field}.values must not be empty`)
  }
  const categories = expectBoundedStringArray(value.categories, `${field}.categories`)
  const valuesRef =
    value.valuesRef !== undefined && value.valuesRef !== null
      ? expectChartRefString(value.valuesRef, `${field}.valuesRef`)
      : undefined
  const categoriesRef =
    value.categoriesRef !== undefined && value.categoriesRef !== null
      ? expectChartRefString(value.categoriesRef, `${field}.categoriesRef`)
      : undefined
  const color =
    value.color !== undefined && value.color !== null
      ? expectChartHexColor(value.color, `${field}.color`)
      : undefined
  let pointColors: ChartAddSeries['pointColors'] | undefined
  if (value.pointColors !== undefined && value.pointColors !== null) {
    if (!isRecord(value.pointColors)) {
      throw new OfficeValidationError('validation', `${field}.pointColors must be an object`)
    }
    const out: Record<string, string> = {}
    for (const [key, entry] of Object.entries(value.pointColors)) {
      expectChartIndexKey(key, `${field}.pointColors key "${key}"`)
      out[key] = expectChartHexColor(entry, `${field}.pointColors["${key}"]`)
    }
    pointColors = out
  }
  let explosionPct: number | undefined
  if (value.explosionPct !== undefined && value.explosionPct !== null) {
    explosionPct = expectBoundedInteger(value.explosionPct, 0, 400, `${field}.explosionPct`)
  }
  let pointExplosions: ChartAddSeries['pointExplosions'] | undefined
  if (value.pointExplosions !== undefined && value.pointExplosions !== null) {
    if (!isRecord(value.pointExplosions)) {
      throw new OfficeValidationError('validation', `${field}.pointExplosions must be an object`)
    }
    const out: Record<string, number> = {}
    for (const [key, entry] of Object.entries(value.pointExplosions)) {
      expectChartIndexKey(key, `${field}.pointExplosions key "${key}"`)
      out[key] = expectBoundedInteger(
        entry,
        0,
        400,
        `${field}.pointExplosions["${key}"]`,
      )
    }
    pointExplosions = out
  }
  for (const key of Object.keys(value)) {
    if (
      ![
        'name',
        'values',
        'categories',
        'valuesRef',
        'categoriesRef',
        'color',
        'pointColors',
        'explosionPct',
        'pointExplosions',
      ].includes(key)
    ) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return {
    name,
    values,
    categories,
    ...(valuesRef !== undefined ? { valuesRef } : {}),
    ...(categoriesRef !== undefined ? { categoriesRef } : {}),
    ...(color !== undefined ? { color } : {}),
    ...(pointColors !== undefined ? { pointColors } : {}),
    ...(explosionPct !== undefined ? { explosionPct } : {}),
    ...(pointExplosions !== undefined ? { pointExplosions } : {}),
  }
}

function expectBoundedInteger(
  value: unknown,
  min: number,
  max: number,
  field: string,
): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < min || value > max) {
    throw new OfficeValidationError(
      'validation',
      `${field} must be an integer between ${min} and ${max}`,
    )
  }
  return value
}

/// One chart addition (mirrors the desktop preload's
/// workbookVisualAddSchema.chart — the canonical engine ChartAdd).
function expectChartAdd(value: unknown, field: string): ChartAdd {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const chartTypeRaw = expectString(value.chartType, `${field}.chartType`)
  if (!CHART_ADD_TYPES.has(chartTypeRaw)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.chartType must be one of column, bar, line, area, pie, scatter, radar, doughnut, combo`,
    )
  }
  const chartType = chartTypeRaw as ChartAdd['chartType']
  const title = expectString(value.title, `${field}.title`)
  if (title.length > 255) {
    throw new OfficeValidationError('validation', `${field}.title exceeds 255 characters`)
  }
  const series = expectArray(value.series, `${field}.series`, (entry, position) =>
    expectChartAddSeries(entry, `${field}.series[${position}]`),
  )
  if (series.length === 0) {
    throw new OfficeValidationError('validation', `${field}.series must not be empty`)
  }
  if (series.length > MAX_CHART_SERIES) {
    throw new OfficeValidationError(
      'validation',
      `${field}.series exceeds ${MAX_CHART_SERIES} entries`,
    )
  }
  let legend: ChartAdd['legend']
  if (value.legend !== undefined && value.legend !== null) {
    const legendRaw = expectString(value.legend, `${field}.legend`)
    if (!CHART_LEGEND_POSITIONS.has(legendRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.legend must be one of none, right, bottom, top, left`,
      )
    }
    legend = legendRaw as ChartAdd['legend']
  }
  let dataLabels: ChartAdd['dataLabels']
  if (value.dataLabels !== undefined && value.dataLabels !== null) {
    const labelsRaw = expectString(value.dataLabels, `${field}.dataLabels`)
    if (!CHART_DATA_LABEL_MODES.has(labelsRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.dataLabels must be one of none, value, percent, category-percent`,
      )
    }
    dataLabels = labelsRaw as ChartAdd['dataLabels']
  }
  let dataLabelPosition: ChartAdd['dataLabelPosition']
  if (value.dataLabelPosition !== undefined && value.dataLabelPosition !== null) {
    const positionRaw = expectString(value.dataLabelPosition, `${field}.dataLabelPosition`)
    if (!CHART_LABEL_POSITIONS.has(positionRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.dataLabelPosition must be one of center, inside-end, outside-end`,
      )
    }
    dataLabelPosition = positionRaw as ChartAdd['dataLabelPosition']
  }
  let dataLabelFormat: string | undefined
  if (value.dataLabelFormat !== undefined && value.dataLabelFormat !== null) {
    dataLabelFormat = expectString(value.dataLabelFormat, `${field}.dataLabelFormat`)
    if (dataLabelFormat.length > 64) {
      throw new OfficeValidationError(
        'validation',
        `${field}.dataLabelFormat exceeds 64 characters`,
      )
    }
  }
  let axisTitles: ChartAdd['axisTitles']
  if (value.axisTitles !== undefined && value.axisTitles !== null) {
    if (!isRecord(value.axisTitles)) {
      throw new OfficeValidationError('validation', `${field}.axisTitles must be an object`)
    }
    const category =
      value.axisTitles.category !== undefined && value.axisTitles.category !== null
        ? expectString(value.axisTitles.category, `${field}.axisTitles.category`)
        : undefined
    const valueTitle =
      value.axisTitles.value !== undefined && value.axisTitles.value !== null
        ? expectString(value.axisTitles.value, `${field}.axisTitles.value`)
        : undefined
    for (const side of [category, valueTitle]) {
      if (side !== undefined && side.length > 255) {
        throw new OfficeValidationError(
          'validation',
          `${field}.axisTitles entries exceed 255 characters`,
        )
      }
    }
    for (const key of Object.keys(value.axisTitles)) {
      if (key !== 'category' && key !== 'value') {
        throw new OfficeValidationError(
          'validation',
          `${field}.axisTitles carries an unknown field "${key}"`,
        )
      }
    }
    axisTitles = {
      ...(category !== undefined ? { category } : {}),
      ...(valueTitle !== undefined ? { value: valueTitle } : {}),
    }
  }
  let grouping: ChartAdd['grouping']
  if (value.grouping !== undefined && value.grouping !== null) {
    const groupingRaw = expectString(value.grouping, `${field}.grouping`)
    if (!CHART_GROUPINGS.has(groupingRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.grouping must be one of clustered, stacked, percentStacked, standard`,
      )
    }
    grouping = groupingRaw as ChartAdd['grouping']
  }
  let gridlines: boolean | undefined
  if (value.gridlines !== undefined && value.gridlines !== null) {
    gridlines = expectBoolean(value.gridlines, `${field}.gridlines`)
  }
  let valueAxis: ChartAdd['valueAxis']
  if (value.valueAxis !== undefined && value.valueAxis !== null) {
    if (!isRecord(value.valueAxis)) {
      throw new OfficeValidationError('validation', `${field}.valueAxis must be an object`)
    }
    const min =
      value.valueAxis.min !== undefined && value.valueAxis.min !== null
        ? expectFiniteNumber(value.valueAxis.min, `${field}.valueAxis.min`)
        : undefined
    const max =
      value.valueAxis.max !== undefined && value.valueAxis.max !== null
        ? expectFiniteNumber(value.valueAxis.max, `${field}.valueAxis.max`)
        : undefined
    for (const key of Object.keys(value.valueAxis)) {
      if (key !== 'min' && key !== 'max') {
        throw new OfficeValidationError(
          'validation',
          `${field}.valueAxis carries an unknown field "${key}"`,
        )
      }
    }
    valueAxis = {
      ...(min !== undefined ? { min } : {}),
      ...(max !== undefined ? { max } : {}),
    }
  }
  let gapWidthPct: number | undefined
  if (value.gapWidthPct !== undefined && value.gapWidthPct !== null) {
    gapWidthPct = expectBoundedInteger(value.gapWidthPct, 0, 500, `${field}.gapWidthPct`)
  }
  let holeSizePct: number | undefined
  if (value.holeSizePct !== undefined && value.holeSizePct !== null) {
    holeSizePct = expectBoundedInteger(value.holeSizePct, 10, 90, `${field}.holeSizePct`)
  }
  for (const key of Object.keys(value)) {
    if (
      ![
        'chartType',
        'title',
        'series',
        'legend',
        'dataLabels',
        'dataLabelPosition',
        'dataLabelFormat',
        'axisTitles',
        'grouping',
        'gridlines',
        'valueAxis',
        'gapWidthPct',
        'holeSizePct',
      ].includes(key)
    ) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return {
    chartType,
    title,
    series,
    ...(legend !== undefined ? { legend } : {}),
    ...(dataLabels !== undefined ? { dataLabels } : {}),
    ...(dataLabelPosition !== undefined ? { dataLabelPosition } : {}),
    ...(dataLabelFormat !== undefined ? { dataLabelFormat } : {}),
    ...(axisTitles !== undefined ? { axisTitles } : {}),
    ...(grouping !== undefined ? { grouping } : {}),
    ...(gridlines !== undefined ? { gridlines } : {}),
    ...(valueAxis !== undefined ? { valueAxis } : {}),
    ...(gapWidthPct !== undefined ? { gapWidthPct } : {}),
    ...(holeSizePct !== undefined ? { holeSizePct } : {}),
  }
}

function expectFiniteNumber(value: unknown, field: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new OfficeValidationError('validation', `${field} must be a finite number`)
  }
  return value
}

/// One full-replacement series entry for a chart edit (desktop preload
/// workbookChartEditSchema.seriesSet entry).
function expectChartSeriesSetEntry(value: unknown, field: string): ChartSeriesSetEntry {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const name = expectString(value.name, `${field}.name`)
  if (name.length > 255) {
    throw new OfficeValidationError('validation', `${field}.name exceeds 255 characters`)
  }
  const values = expectFiniteNumberArray(value.values, `${field}.values`)
  const categories =
    value.categories !== undefined && value.categories !== null
      ? expectBoundedStringArray(value.categories, `${field}.categories`)
      : undefined
  const valuesRef =
    value.valuesRef !== undefined && value.valuesRef !== null
      ? expectChartRefString(value.valuesRef, `${field}.valuesRef`)
      : undefined
  const categoriesRef =
    value.categoriesRef !== undefined && value.categoriesRef !== null
      ? expectChartRefString(value.categoriesRef, `${field}.categoriesRef`)
      : undefined
  const color =
    value.color !== undefined && value.color !== null
      ? expectChartHexColor(value.color, `${field}.color`)
      : undefined
  for (const key of Object.keys(value)) {
    if (!['name', 'values', 'valuesRef', 'categories', 'categoriesRef', 'color'].includes(key)) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return {
    name,
    values,
    ...(valuesRef !== undefined ? { valuesRef } : {}),
    ...(categories !== undefined ? { categories } : {}),
    ...(categoriesRef !== undefined ? { categoriesRef } : {}),
    ...(color !== undefined ? { color } : {}),
  }
}

/// One index-keyed series edit (desktop workbookChartEditSchema.series
/// entry — needs a name or data).
function expectChartSeriesEdit(value: unknown, field: string): ChartSeriesEdit {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const index = expectBoundedInteger(value.index, 0, 255, `${field}.index`)
  const name =
    value.name !== undefined && value.name !== null
      ? expectString(value.name, `${field}.name`)
      : undefined
  if (name !== undefined && name.length > 255) {
    throw new OfficeValidationError('validation', `${field}.name exceeds 255 characters`)
  }
  const valuesRef =
    value.valuesRef !== undefined && value.valuesRef !== null
      ? expectChartRefString(value.valuesRef, `${field}.valuesRef`)
      : undefined
  const values =
    value.values !== undefined && value.values !== null
      ? expectFiniteNumberArray(value.values, `${field}.values`)
      : undefined
  const categoriesRef =
    value.categoriesRef !== undefined && value.categoriesRef !== null
      ? expectChartRefString(value.categoriesRef, `${field}.categoriesRef`)
      : undefined
  const categories =
    value.categories !== undefined && value.categories !== null
      ? expectBoundedStringArray(value.categories, `${field}.categories`)
      : undefined
  if (
    name === undefined &&
    values === undefined &&
    categories === undefined &&
    valuesRef === undefined &&
    categoriesRef === undefined
  ) {
    throw new OfficeValidationError('validation', `${field} needs a name or data`)
  }
  for (const key of Object.keys(value)) {
    if (
      !['index', 'name', 'valuesRef', 'values', 'categoriesRef', 'categories'].includes(key)
    ) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return {
    index,
    ...(name !== undefined ? { name } : {}),
    ...(valuesRef !== undefined ? { valuesRef } : {}),
    ...(values !== undefined ? { values } : {}),
    ...(categoriesRef !== undefined ? { categoriesRef } : {}),
    ...(categories !== undefined ? { categories } : {}),
  }
}

/**
 * Validate one chart edit from the wire (EXCEL-023 — Chart Design edits
 * on file-native charts). The edit targets a chart part by its canonical
 * xl/charts/*.xml path and carries at least one supported property
 * change, all bounded (desktop preload workbookChartEditSchema parity).
 * Unknown fields are rejected; nothing unvalidated reaches the engine.
 */
function expectWorkbookChartEdit(value: unknown, index: number): WorkbookChartEdit {
  const field = `chartEdits[${index}]`
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const chartPath = expectString(value.chartPath, `${field}.chartPath`)
  if (!CHART_PATH_PATTERN.test(chartPath)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.chartPath must be an xl/charts/*.xml package path`,
    )
  }
  let title: string | undefined
  if (value.title !== undefined && value.title !== null) {
    title = expectString(value.title, `${field}.title`)
    if (title.length > 255) {
      throw new OfficeValidationError('validation', `${field}.title exceeds 255 characters`)
    }
  }
  let chartType: WorkbookChartEdit['chartType']
  if (value.chartType !== undefined && value.chartType !== null) {
    const typeRaw = expectString(value.chartType, `${field}.chartType`)
    if (!CHART_EDIT_TYPES.has(typeRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.chartType must be one of column, bar, line, area, pie, doughnut`,
      )
    }
    chartType = typeRaw as WorkbookChartEdit['chartType']
  }
  let seriesColors: Record<string, string> | undefined
  if (value.seriesColors !== undefined && value.seriesColors !== null) {
    if (!isRecord(value.seriesColors)) {
      throw new OfficeValidationError('validation', `${field}.seriesColors must be an object`)
    }
    const out: Record<string, string> = {}
    for (const [key, entry] of Object.entries(value.seriesColors)) {
      expectChartIndexKey(key, `${field}.seriesColors key "${key}"`)
      out[key] = expectChartHexColor(entry, `${field}.seriesColors["${key}"]`)
    }
    seriesColors = out
  }
  let legend: WorkbookChartEdit['legend']
  if (value.legend !== undefined && value.legend !== null) {
    const legendRaw = expectString(value.legend, `${field}.legend`)
    if (!CHART_LEGEND_POSITIONS.has(legendRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.legend must be one of none, right, bottom, top, left`,
      )
    }
    legend = legendRaw as WorkbookChartEdit['legend']
  }
  let dataLabels: WorkbookChartEdit['dataLabels']
  if (value.dataLabels !== undefined && value.dataLabels !== null) {
    const labelsRaw = expectString(value.dataLabels, `${field}.dataLabels`)
    if (!CHART_DATA_LABEL_MODES.has(labelsRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.dataLabels must be one of none, value, percent, category-percent`,
      )
    }
    dataLabels = labelsRaw as WorkbookChartEdit['dataLabels']
  }
  let dataLabelPosition: WorkbookChartEdit['dataLabelPosition']
  if (value.dataLabelPosition !== undefined && value.dataLabelPosition !== null) {
    const positionRaw = expectString(value.dataLabelPosition, `${field}.dataLabelPosition`)
    if (!CHART_LABEL_POSITIONS.has(positionRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.dataLabelPosition must be one of center, inside-end, outside-end`,
      )
    }
    dataLabelPosition = positionRaw as WorkbookChartEdit['dataLabelPosition']
  }
  let dataLabelFormat: string | undefined
  if (value.dataLabelFormat !== undefined && value.dataLabelFormat !== null) {
    dataLabelFormat = expectString(value.dataLabelFormat, `${field}.dataLabelFormat`)
    if (dataLabelFormat.length > 64) {
      throw new OfficeValidationError(
        'validation',
        `${field}.dataLabelFormat exceeds 64 characters`,
      )
    }
  }
  let axisTitles: WorkbookChartEdit['axisTitles']
  if (value.axisTitles !== undefined && value.axisTitles !== null) {
    if (!isRecord(value.axisTitles)) {
      throw new OfficeValidationError('validation', `${field}.axisTitles must be an object`)
    }
    const category =
      value.axisTitles.category === undefined || value.axisTitles.category === null
        ? null
        : expectString(value.axisTitles.category, `${field}.axisTitles.category`)
    const valueTitle =
      value.axisTitles.value === undefined || value.axisTitles.value === null
        ? null
        : expectString(value.axisTitles.value, `${field}.axisTitles.value`)
    for (const side of [category, valueTitle]) {
      if (side !== null && side.length > 255) {
        throw new OfficeValidationError(
          'validation',
          `${field}.axisTitles entries exceed 255 characters`,
        )
      }
    }
    for (const key of Object.keys(value.axisTitles)) {
      if (key !== 'category' && key !== 'value') {
        throw new OfficeValidationError(
          'validation',
          `${field}.axisTitles carries an unknown field "${key}"`,
        )
      }
    }
    axisTitles = {
      ...(category !== null ? { category } : {}),
      ...(valueTitle !== null ? { value: valueTitle } : {}),
    }
  }
  let pointColors: Record<string, Record<string, string>> | undefined
  if (value.pointColors !== undefined && value.pointColors !== null) {
    if (!isRecord(value.pointColors)) {
      throw new OfficeValidationError('validation', `${field}.pointColors must be an object`)
    }
    const out: Record<string, Record<string, string>> = {}
    for (const [seriesKey, seriesEntry] of Object.entries(value.pointColors)) {
      expectChartIndexKey(seriesKey, `${field}.pointColors key "${seriesKey}"`)
      if (!isRecord(seriesEntry)) {
        throw new OfficeValidationError(
          'validation',
          `${field}.pointColors["${seriesKey}"] must be an object`,
        )
      }
      const points: Record<string, string> = {}
      for (const [pointKey, pointEntry] of Object.entries(seriesEntry)) {
        expectChartIndexKey(pointKey, `${field}.pointColors["${seriesKey}"] key "${pointKey}"`)
        points[pointKey] = expectChartHexColor(
          pointEntry,
          `${field}.pointColors["${seriesKey}"]["${pointKey}"]`,
        )
      }
      out[seriesKey] = points
    }
    pointColors = out
  }
  let grouping: WorkbookChartEdit['grouping']
  if (value.grouping !== undefined && value.grouping !== null) {
    const groupingRaw = expectString(value.grouping, `${field}.grouping`)
    if (!CHART_GROUPINGS.has(groupingRaw)) {
      throw new OfficeValidationError(
        'validation',
        `${field}.grouping must be one of clustered, stacked, percentStacked, standard`,
      )
    }
    grouping = groupingRaw as WorkbookChartEdit['grouping']
  }
  let gridlines: boolean | undefined
  if (value.gridlines !== undefined && value.gridlines !== null) {
    gridlines = expectBoolean(value.gridlines, `${field}.gridlines`)
  }
  let valueAxis: WorkbookChartEdit['valueAxis']
  if (value.valueAxis !== undefined && value.valueAxis !== null) {
    if (!isRecord(value.valueAxis)) {
      throw new OfficeValidationError('validation', `${field}.valueAxis must be an object`)
    }
    const min =
      value.valueAxis.min === undefined || value.valueAxis.min === null
        ? null
        : expectFiniteNumber(value.valueAxis.min, `${field}.valueAxis.min`)
    const max =
      value.valueAxis.max === undefined || value.valueAxis.max === null
        ? null
        : expectFiniteNumber(value.valueAxis.max, `${field}.valueAxis.max`)
    for (const key of Object.keys(value.valueAxis)) {
      if (key !== 'min' && key !== 'max') {
        throw new OfficeValidationError(
          'validation',
          `${field}.valueAxis carries an unknown field "${key}"`,
        )
      }
    }
    valueAxis = {
      ...(min !== null ? { min } : {}),
      ...(max !== null ? { max } : {}),
    }
    if (min === null && max === null) {
      throw new OfficeValidationError('validation', `${field}.valueAxis needs min or max`)
    }
  }
  let gapWidthPct: number | undefined
  if (value.gapWidthPct !== undefined && value.gapWidthPct !== null) {
    gapWidthPct = expectBoundedInteger(value.gapWidthPct, 0, 500, `${field}.gapWidthPct`)
  }
  let holeSizePct: number | undefined
  if (value.holeSizePct !== undefined && value.holeSizePct !== null) {
    holeSizePct = expectBoundedInteger(value.holeSizePct, 10, 90, `${field}.holeSizePct`)
  }
  let explosionPct: number | undefined
  if (value.explosionPct !== undefined && value.explosionPct !== null) {
    explosionPct = expectBoundedInteger(value.explosionPct, 0, 400, `${field}.explosionPct`)
  }
  let pointExplosions: Record<string, number> | undefined
  if (value.pointExplosions !== undefined && value.pointExplosions !== null) {
    if (!isRecord(value.pointExplosions)) {
      throw new OfficeValidationError('validation', `${field}.pointExplosions must be an object`)
    }
    const out: Record<string, number> = {}
    for (const [key, entry] of Object.entries(value.pointExplosions)) {
      expectChartIndexKey(key, `${field}.pointExplosions key "${key}"`)
      out[key] = expectBoundedInteger(entry, 0, 400, `${field}.pointExplosions["${key}"]`)
    }
    pointExplosions = out
  }
  let seriesSet: readonly ChartSeriesSetEntry[] | undefined
  if (value.seriesSet !== undefined && value.seriesSet !== null) {
    seriesSet = expectArray(value.seriesSet, `${field}.seriesSet`, (entry, position) =>
      expectChartSeriesSetEntry(entry, `${field}.seriesSet[${position}]`),
    )
    if (seriesSet.length === 0) {
      throw new OfficeValidationError('validation', `${field}.seriesSet must not be empty`)
    }
    if (seriesSet.length > MAX_CHART_SERIES) {
      throw new OfficeValidationError(
        'validation',
        `${field}.seriesSet exceeds ${MAX_CHART_SERIES} entries`,
      )
    }
  }
  let series: readonly ChartSeriesEdit[] | undefined
  if (value.series !== undefined && value.series !== null) {
    series = expectArray(value.series, `${field}.series`, (entry, position) =>
      expectChartSeriesEdit(entry, `${field}.series[${position}]`),
    )
    if (series.length > MAX_CHART_SERIES) {
      throw new OfficeValidationError(
        'validation',
        `${field}.series exceeds ${MAX_CHART_SERIES} entries`,
      )
    }
  }
  for (const key of Object.keys(value)) {
    if (
      ![
        'chartPath',
        'title',
        'chartType',
        'seriesColors',
        'legend',
        'dataLabels',
        'dataLabelPosition',
        'dataLabelFormat',
        'axisTitles',
        'pointColors',
        'grouping',
        'gridlines',
        'valueAxis',
        'gapWidthPct',
        'holeSizePct',
        'explosionPct',
        'pointExplosions',
        'seriesSet',
        'series',
      ].includes(key)
    ) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  const hasAnyChange =
    title !== undefined ||
    chartType !== undefined ||
    (seriesColors !== undefined && Object.keys(seriesColors).length > 0) ||
    (pointColors !== undefined && Object.keys(pointColors).length > 0) ||
    legend !== undefined ||
    axisTitles !== undefined ||
    dataLabels !== undefined ||
    dataLabelPosition !== undefined ||
    dataLabelFormat !== undefined ||
    grouping !== undefined ||
    gridlines !== undefined ||
    valueAxis !== undefined ||
    gapWidthPct !== undefined ||
    holeSizePct !== undefined ||
    explosionPct !== undefined ||
    (pointExplosions !== undefined && Object.keys(pointExplosions).length > 0) ||
    (seriesSet !== undefined && seriesSet.length > 0) ||
    (series !== undefined && series.length > 0)
  if (!hasAnyChange) {
    throw new OfficeValidationError('validation', `${field} needs at least one property change`)
  }
  return {
    chartPath,
    ...(title !== undefined ? { title } : {}),
    ...(chartType !== undefined ? { chartType } : {}),
    ...(seriesColors !== undefined ? { seriesColors } : {}),
    ...(legend !== undefined ? { legend } : {}),
    ...(dataLabels !== undefined ? { dataLabels } : {}),
    ...(dataLabelPosition !== undefined ? { dataLabelPosition } : {}),
    ...(dataLabelFormat !== undefined ? { dataLabelFormat } : {}),
    ...(axisTitles !== undefined ? { axisTitles } : {}),
    ...(pointColors !== undefined ? { pointColors } : {}),
    ...(grouping !== undefined ? { grouping } : {}),
    ...(gridlines !== undefined ? { gridlines } : {}),
    ...(valueAxis !== undefined ? { valueAxis } : {}),
    ...(gapWidthPct !== undefined ? { gapWidthPct } : {}),
    ...(holeSizePct !== undefined ? { holeSizePct } : {}),
    ...(explosionPct !== undefined ? { explosionPct } : {}),
    ...(pointExplosions !== undefined ? { pointExplosions } : {}),
    ...(seriesSet !== undefined ? { seriesSet } : {}),
    ...(series !== undefined ? { series } : {}),
  }
}

/**
 * Validate one visual edit from the wire (EXCEL-022). The edit targets the
 * canonical (drawingPath, drawingIndex) locator and carries EITHER a
 * removal flag OR a new anchor — never both, never neither. Unknown
 * fields are rejected; unvalidated objects never reach the engine.
 */
function expectWorkbookVisualEdit(value: unknown, index: number): WorkbookVisualEdit {
  const field = `visualEdits[${index}]`
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const drawingPath = expectString(value.drawingPath, `${field}.drawingPath`)
  if (!DRAWING_PATH_PATTERN.test(drawingPath)) {
    throw new OfficeValidationError(
      'validation',
      `${field}.drawingPath must be an xl/drawings/*.xml package path`,
    )
  }
  const drawingIndexRaw = value.drawingIndex
  if (
    typeof drawingIndexRaw !== 'number' ||
    !Number.isInteger(drawingIndexRaw) ||
    drawingIndexRaw < 0 ||
    drawingIndexRaw > 10_000
  ) {
    throw new OfficeValidationError(
      'validation',
      `${field}.drawingIndex must be an integer between 0 and 10000`,
    )
  }
  const hasRemove = value.remove !== undefined
  const hasAnchor = value.anchor !== undefined
  if (hasRemove && hasAnchor) {
    throw new OfficeValidationError(
      'validation',
      `${field} carries both remove and anchor — exactly one is required`,
    )
  }
  if (!hasRemove && !hasAnchor) {
    throw new OfficeValidationError('validation', `${field} needs either remove or anchor`)
  }
  if (hasRemove && value.remove !== true) {
    throw new OfficeValidationError('validation', `${field}.remove must be true`)
  }
  const anchor = hasAnchor ? expectDrawingAnchor(value.anchor, `${field}.anchor`) : undefined
  for (const key of Object.keys(value)) {
    if (!['drawingPath', 'drawingIndex', 'remove', 'anchor'].includes(key)) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return {
    drawingPath,
    drawingIndex: drawingIndexRaw,
    ...(hasRemove ? { remove: true } : {}),
    ...(anchor !== undefined ? { anchor } : {}),
  }
}

/**
 * Validate one table addition from the wire (EXCEL-021). Mirrors the
 * desktop preload's workbookTableAddSchema exactly — sheetName, a 0-based
 * ordered integer area, a 1-255 name, 1-1000 column names of at most 255
 * chars, an optional built-in style name, and the bandedRows flag. Unknown
 * fields are rejected with a 400 rather than reaching the engine.
 */
function expectSheetTableAddition(value: unknown, index: number): SheetTableAddition {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `tableAdditions[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `tableAdditions[${index}].sheetName`)
  const name = expectString(value.name, `tableAdditions[${index}].name`)
  if (name.length > 255) {
    throw new OfficeValidationError(
      'validation',
      `tableAdditions[${index}].name exceeds 255 characters`,
    )
  }
  const columnNames = expectArray(
    value.columnNames,
    `tableAdditions[${index}].columnNames`,
    (column, columnNumber) => {
      const text = expectString(column, `tableAdditions[${index}].columnNames[${columnNumber}]`)
      if (text.length > 255) {
        throw new OfficeValidationError(
          'validation',
          `tableAdditions[${index}].columnNames[${columnNumber}] exceeds 255 characters`,
        )
      }
      return text
    },
  )
  if (columnNames.length === 0) {
    throw new OfficeValidationError(
      'validation',
      `tableAdditions[${index}].columnNames must not be empty`,
    )
  }
  if (columnNames.length > 1_000) {
    throw new OfficeValidationError(
      'validation',
      `tableAdditions[${index}].columnNames exceeds 1000 entries`,
    )
  }
  const bandedRows = expectBoolean(value.bandedRows, `tableAdditions[${index}].bandedRows`)
  let style: string | undefined
  if (value.style !== undefined) {
    style = expectString(value.style, `tableAdditions[${index}].style`)
    if (!TABLE_STYLE_PATTERN.test(style)) {
      throw new OfficeValidationError(
        'validation',
        `tableAdditions[${index}].style must be a built-in TableStyle name`,
      )
    }
  }
  const area = expectTableArea(value.area, `tableAdditions[${index}].area`)
  for (const key of Object.keys(value)) {
    if (!['sheetName', 'area', 'name', 'columnNames', 'style', 'bandedRows'].includes(key)) {
      throw new OfficeValidationError(
        'validation',
        `tableAdditions[${index}] carries an unknown field "${key}"`,
      )
    }
  }
  return {
    sheetName,
    area,
    name,
    columnNames,
    ...(style !== undefined ? { style } : {}),
    bandedRows,
  }
}

/**
 * 0-based ordered integer area — the desktop preload's parseCellArea.
 */
function expectTableArea(
  value: unknown,
  field: string,
): {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
} {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const startRow = expectNumber(value.startRow, `${field}.startRow`)
  const startColumn = expectNumber(value.startColumn, `${field}.startColumn`)
  const endRow = expectNumber(value.endRow, `${field}.endRow`)
  const endColumn = expectNumber(value.endColumn, `${field}.endColumn`)
  for (const [key, number] of [
    ['startRow', startRow],
    ['startColumn', startColumn],
    ['endRow', endRow],
    ['endColumn', endColumn],
  ] as const) {
    if (!Number.isInteger(number) || number < 0) {
      throw new OfficeValidationError(
        'validation',
        `${field}.${key} must be a non-negative integer`,
      )
    }
  }
  if (endRow < startRow || endColumn < startColumn) {
    throw new OfficeValidationError('validation', `${field} is not an ordered area`)
  }
  return { startRow, startColumn, endRow, endColumn }
}

/**
 * Validate one per-sheet protection state from the wire (EXCEL-020). The
 * browser only emits the journal's toggle decisions — exactly the shape
 * the desktop ships (`{ sheetId, protected }` resolved to sheet names).
 * Unknown fields, non-boolean flags, and excessive counts are rejected
 * with a 400 rather than reaching the engine.
 */
function expectSheetProtectionState(value: unknown, index: number): SheetProtectionState {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `sheetProtections[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `sheetProtections[${index}].sheetName`)
  const protect = expectBoolean(value.protected, `sheetProtections[${index}].protected`)
  // Reject unknown extra keys — the canonical state carries exactly these
  // two fields; password-bearing payloads must not slip through the wire
  // (the engine takes no passwords by design).
  for (const key of Object.keys(value)) {
    if (!['sheetName', 'protected'].includes(key)) {
      throw new OfficeValidationError(
        'validation',
        `sheetProtections[${index}] carries an unknown field "${key}"`,
      )
    }
  }
  return { sheetName, protected: protect }
}

/**
 * Validate the workbook structure-protection state from the wire
 * (EXCEL-020). null = untouched. Mirrors the desktop's strict Zod
 * `{ lockStructure: boolean }` — anything else is a 400.
 */
function expectWorkbookProtectionState(value: unknown): { readonly lockStructure: boolean } {
  if (!isRecord(value)) {
    throw new OfficeValidationError(
      'validation',
      'savePlan.workbookProtectionState must be an object',
    )
  }
  const lockStructure = expectBoolean(
    value.lockStructure,
    'savePlan.workbookProtectionState.lockStructure',
  )
  for (const key of Object.keys(value)) {
    if (key !== 'lockStructure') {
      throw new OfficeValidationError(
        'validation',
        `savePlan.workbookProtectionState carries an unknown field "${key}"`,
      )
    }
  }
  return { lockStructure }
}

// ── SheetNoteState validation (Review → Notes/Comments, Phase 4 Inc. 6) ─────

/** Caps mirroring the desktop's workbookNoteStateSchema. */
const MAX_NOTE_STATES = 1_000
const MAX_NOTES_PER_SHEET = 1_000
const MAX_NOTE_AUTHOR_LENGTH = 255
const MAX_NOTE_TEXT_LENGTH = 32_767
const MAX_NOTE_ROW = 1_048_575
const MAX_NOTE_COLUMN = 16_383

/**
 * Validate one per-sheet note state from the wire. The browser only emits
 * full declarative snapshots of Univer's live note model (the same shape
 * the desktop ships) — anything else (invalid coordinates, missing text,
 * oversized strings, unknown fields, excessive counts) is rejected with a
 * 400 rather than reaching the engine. An empty notes array is VALID: it
 * means "all notes on the sheet were cleared".
 */
function expectSheetNoteState(value: unknown, index: number): SheetNoteState {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `noteStates[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `noteStates[${index}].sheetName`)
  const notes = expectArray(value.notes, `noteStates[${index}].notes`, (note, i) =>
    expectSheetNote(note, `noteStates[${index}].notes[${i}]`),
  )
  if (notes.length > MAX_NOTES_PER_SHEET) {
    throw new OfficeValidationError(
      'validation',
      `noteStates[${index}].notes exceeds ${MAX_NOTES_PER_SHEET} entries`,
    )
  }
  return { sheetName, notes }
}

function expectSheetNote(
  value: unknown,
  field: string,
): { row: number; column: number; author: string; text: string } {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const row = expectNumber(value.row, `${field}.row`)
  if (!Number.isInteger(row) || row < 0 || row > MAX_NOTE_ROW) {
    throw new OfficeValidationError('validation', `${field}.row must be a 0-based row index`)
  }
  const column = expectNumber(value.column, `${field}.column`)
  if (!Number.isInteger(column) || column < 0 || column > MAX_NOTE_COLUMN) {
    throw new OfficeValidationError('validation', `${field}.column must be a 0-based column index`)
  }
  const author = value.author
  if (typeof author !== 'string' || author.length > MAX_NOTE_AUTHOR_LENGTH) {
    throw new OfficeValidationError(
      'validation',
      `${field}.author must be a string of at most ${MAX_NOTE_AUTHOR_LENGTH} characters`,
    )
  }
  const text = value.text
  if (typeof text !== 'string' || text.length === 0 || text.length > MAX_NOTE_TEXT_LENGTH) {
    throw new OfficeValidationError(
      'validation',
      `${field}.text must be a non-empty string of at most ${MAX_NOTE_TEXT_LENGTH} characters`,
    )
  }
  // Reject unknown extra keys — the canonical note carries exactly these
  // four fields; extension markers (threaded-comment ids, reply refs, …)
  // must not slip through.
  for (const key of Object.keys(value)) {
    if (!['row', 'column', 'author', 'text'].includes(key)) {
      throw new OfficeValidationError('validation', `${field} carries an unknown field "${key}"`)
    }
  }
  return { row, column, author, text }
}

function parseSaveWorkbookRequest(
  body: unknown,
  codec: OfficeBinaryCodec,
): ParsedSaveWorkbookRequest {
  if (!isRecord(body)) {
    throw new OfficeValidationError('validation', 'Request body must be a JSON object')
  }
  const fileName = validateFileName(body.fileName)
  const fileBytes = decodeFileBytes(body.fileBytes, codec)

  // Canonical path: savePlan.edits (+ optional structuralOps + pageSetupStates).
  if (body.savePlan !== undefined) {
    if (!isRecord(body.savePlan)) {
      throw new OfficeValidationError('validation', 'savePlan must be an object')
    }
    const edits = expectArray(body.savePlan.edits, 'savePlan.edits', expectCellEdit)
    const structuralOps =
      body.savePlan.structuralOps !== undefined && body.savePlan.structuralOps !== null
        ? expectArray(
            body.savePlan.structuralOps,
            'savePlan.structuralOps',
            expectSheetStructuralOps,
          )
        : undefined
    const pageSetupStates =
      body.savePlan.pageSetupStates !== undefined && body.savePlan.pageSetupStates !== null
        ? expectArray(
            body.savePlan.pageSetupStates,
            'savePlan.pageSetupStates',
            expectSheetPageSetupState,
          )
        : undefined
    const filterStates =
      body.savePlan.filterStates !== undefined && body.savePlan.filterStates !== null
        ? expectArray(body.savePlan.filterStates, 'savePlan.filterStates', expectSheetFilterState)
        : undefined
    if (filterStates !== undefined && filterStates.length > MAX_FILTER_STATES) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.filterStates exceeds ${MAX_FILTER_STATES} entries`,
      )
    }
    const dvStates =
      body.savePlan.dvStates !== undefined && body.savePlan.dvStates !== null
        ? expectArray(body.savePlan.dvStates, 'savePlan.dvStates', expectSheetDvState)
        : undefined
    if (dvStates !== undefined && dvStates.length > MAX_DV_STATES) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.dvStates exceeds ${MAX_DV_STATES} entries`,
      )
    }
    const noteStates =
      body.savePlan.noteStates !== undefined && body.savePlan.noteStates !== null
        ? expectArray(body.savePlan.noteStates, 'savePlan.noteStates', expectSheetNoteState)
        : undefined
    if (noteStates !== undefined && noteStates.length > MAX_NOTE_STATES) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.noteStates exceeds ${MAX_NOTE_STATES} entries`,
      )
    }
    // EXCEL-020: per-sheet protection toggles + workbook structure lock.
    // Both are strictly validated — nothing unvalidated reaches the engine.
    const sheetProtections =
      body.savePlan.sheetProtections !== undefined && body.savePlan.sheetProtections !== null
        ? expectArray(
            body.savePlan.sheetProtections,
            'savePlan.sheetProtections',
            expectSheetProtectionState,
          )
        : undefined
    if (sheetProtections !== undefined && sheetProtections.length > MAX_SHEET_PROTECTIONS) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.sheetProtections exceeds ${MAX_SHEET_PROTECTIONS} entries`,
      )
    }
    const workbookProtectionState =
      body.savePlan.workbookProtectionState !== undefined &&
      body.savePlan.workbookProtectionState !== null
        ? expectWorkbookProtectionState(body.savePlan.workbookProtectionState)
        : null
    // EXCEL-021: session-created tables (Insert → Table). Strictly
    // validated — nothing unvalidated reaches the engine.
    const tableAdditions =
      body.savePlan.tableAdditions !== undefined && body.savePlan.tableAdditions !== null
        ? expectArray(
            body.savePlan.tableAdditions,
            'savePlan.tableAdditions',
            expectSheetTableAddition,
          )
        : undefined
    if (tableAdditions !== undefined && tableAdditions.length > MAX_TABLE_ADDITIONS) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.tableAdditions exceeds ${MAX_TABLE_ADDITIONS} entries`,
      )
    }
    // EXCEL-022: session-created images (Insert → Picture) + surgical
    // edits to file-native images (move / resize / delete). Strictly
    // validated — nothing unvalidated reaches the engine.
    const visualAdditions =
      body.savePlan.visualAdditions !== undefined && body.savePlan.visualAdditions !== null
        ? expectArray(
            body.savePlan.visualAdditions,
            'savePlan.visualAdditions',
            expectSheetVisualAddition,
          )
        : undefined
    if (visualAdditions !== undefined && visualAdditions.length > MAX_VISUAL_ADDITIONS) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.visualAdditions exceeds ${MAX_VISUAL_ADDITIONS} entries`,
      )
    }
    const visualEdits =
      body.savePlan.visualEdits !== undefined && body.savePlan.visualEdits !== null
        ? expectArray(body.savePlan.visualEdits, 'savePlan.visualEdits', expectWorkbookVisualEdit)
        : undefined
    if (visualEdits !== undefined && visualEdits.length > MAX_VISUAL_EDITS) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.visualEdits exceeds ${MAX_VISUAL_EDITS} entries`,
      )
    }
    // EXCEL-023: surgical semantic edits to file-native charts. Strictly
    // validated — nothing unvalidated reaches the engine.
    const chartEdits =
      body.savePlan.chartEdits !== undefined && body.savePlan.chartEdits !== null
        ? expectArray(body.savePlan.chartEdits, 'savePlan.chartEdits', expectWorkbookChartEdit)
        : undefined
    if (chartEdits !== undefined && chartEdits.length > MAX_CHART_EDITS) {
      throw new OfficeValidationError(
        'validation',
        `savePlan.chartEdits exceeds ${MAX_CHART_EDITS} entries`,
      )
    }
    return {
      fileName,
      fileBytes,
      savePlan: {
        edits,
        ...(structuralOps ? { structuralOps } : {}),
        ...(pageSetupStates ? { pageSetupStates } : {}),
        ...(filterStates ? { filterStates } : {}),
        ...(dvStates ? { dvStates } : {}),
        ...(noteStates ? { noteStates } : {}),
        ...(sheetProtections ? { sheetProtections } : {}),
        ...(workbookProtectionState !== null ? { workbookProtectionState } : {}),
        ...(tableAdditions && tableAdditions.length > 0 ? { tableAdditions } : {}),
        ...(visualAdditions && visualAdditions.length > 0 ? { visualAdditions } : {}),
        ...(visualEdits && visualEdits.length > 0 ? { visualEdits } : {}),
        ...(chartEdits && chartEdits.length > 0 ? { chartEdits } : {}),
      },
    }
  }

  // Legacy path: top-level cellEdits.
  if (body.cellEdits !== undefined) {
    const edits = expectArray(body.cellEdits, 'cellEdits', expectCellEdit)
    return { fileName, fileBytes, savePlan: { edits } }
  }

  throw new OfficeValidationError('validation', 'savePlan.edits (or legacy cellEdits) is required')
}

function parseOpenDocumentRequest(
  body: unknown,
  codec: OfficeBinaryCodec,
): ParsedOpenDocumentRequest {
  if (!isRecord(body)) {
    throw new OfficeValidationError('validation', 'Request body must be a JSON object')
  }
  const fileName = validateFileName(body.fileName)
  const fileBytes = decodeFileBytes(body.fileBytes, codec)
  return { fileName, fileBytes }
}

function parseSaveDocumentRequest(
  body: unknown,
  codec: OfficeBinaryCodec,
): ParsedSaveDocumentRequest {
  if (!isRecord(body)) {
    throw new OfficeValidationError('validation', 'Request body must be a JSON object')
  }
  const fileName = validateFileName(body.fileName)
  const fileBytes = decodeFileBytes(body.fileBytes, codec)
  const blocks = expectArray(body.blocks, 'blocks', expectSerializedBlock)
  return { fileName, fileBytes, blocks }
}

// ── Workbook (XLSX) handlers ────────────────────────────────────────────────

async function handleOpenWorkbook(
  body: unknown,
  codec: OfficeBinaryCodec,
): Promise<OfficeApiResponse> {
  const { fileBytes } = parseOpenWorkbookRequest(body, codec)
  let imported
  try {
    imported = await readBasicWorkbook(Buffer.from(fileBytes))
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to read workbook',
    )
  }
  const res: OpenWorkbookResponse = {
    snapshot: imported.snapshot,
    sheetNamesById: imported.sheetNamesById,
  }
  return { status: 200, body: res }
}

async function handleSaveWorkbook(
  body: unknown,
  codec: OfficeBinaryCodec,
): Promise<OfficeApiResponse> {
  const req = parseSaveWorkbookRequest(body, codec)
  const buf = Buffer.from(req.fileBytes)
  const edits = req.savePlan.edits
  const structuralOps = req.savePlan.structuralOps ?? []
  const pageSetupStates = req.savePlan.pageSetupStates ?? []
  const filterStates = req.savePlan.filterStates ?? []
  const dvStates = req.savePlan.dvStates ?? []
  const noteStates = req.savePlan.noteStates ?? []
  const sheetProtections = req.savePlan.sheetProtections ?? []
  const workbookProtectionState = req.savePlan.workbookProtectionState ?? null
  const tableAdditions = req.savePlan.tableAdditions ?? []
  const visualAdditions = req.savePlan.visualAdditions ?? []
  const visualEdits = req.savePlan.visualEdits ?? []
  const chartEdits = req.savePlan.chartEdits ?? []
  let mutation
  try {
    // filterStates is argument 6 (after chartEdits/sheetPlan): the canonical
    // engine applies filter snapshots AFTER structural replay, cell edits,
    // and page setup, so their coordinates and row set match the sheet's
    // final content. dvStates is argument 9 (after hyperlinkEdits/cfStates):
    // validation rules likewise replay after structural + cell changes.
    // noteStates is argument 13 (after pageSetupStates): note snapshots run
    // after the worksheet flush so the legacyDrawing element lands on the
    // final sheet XML. sheetProtections is argument 10 (after dvStates) and
    // workbookProtectionState is the next-to-trailing argument — both write
    // their OOXML protection elements after the content flush.
    // tableAdditions is the trailing argument (EXCEL-021): new tables run
    // on the flushed worksheet XML — the <tableParts> element and overlap
    // checks see the final sheet content. visualAdditions/visualEdits are
    // the two trailing arguments (EXCEL-022): surgical image edits run
    // after the structural flush (post-shift coordinates) and new image
    // anchors append after that. See planCellEditsToXlsx in
    // @genoffice/xlsx-gateway.
    mutation = await applyCellEditsToXlsx(
      buf,
      edits,
      structuralOps,
      chartEdits,
      undefined,
      filterStates,
      [],
      [],
      dvStates,
      sheetProtections,
      null,
      pageSetupStates,
      noteStates,
      [],
      workbookProtectionState,
      tableAdditions,
      visualAdditions,
      visualEdits,
    )
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to apply cell edits',
    )
  }
  const res: SaveWorkbookResponse = {
    fileBytes: encodeFileBytes(mutation.buffer, codec),
    ...(mutation.addedVisuals !== undefined ? { addedVisuals: mutation.addedVisuals } : {}),
  }
  return { status: 200, body: res }
}

// ── Document (DOCX) handlers ────────────────────────────────────────────────

function serializeRun(run: Run): SerializedRun {
  const out: {
    text: string
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strike?: boolean
    link?: { href: string; tooltip?: string }
  } = { text: run.text }
  if (run.bold) out.bold = true
  if (run.italic) out.italic = true
  if (run.underline) out.underline = true
  if (run.strike) out.strike = true
  if (run.link)
    out.link = { href: run.link.href, ...(run.link.tooltip ? { tooltip: run.link.tooltip } : {}) }
  return out
}

// ── Table model ↔ wire conversion (canonical TableModel is the source of truth) ──

/** docx-engine CellBorder → wire */
function serializeCellBorder(b: CellBorder): SerializedCellBorder {
  return {
    style: b.style,
    ...(b.szEighths !== undefined ? { szEighths: b.szEighths } : {}),
    ...(b.color !== undefined ? { color: b.color } : {}),
  }
}

function serializeCellBorders(b: CellBorders): SerializedCellBorders {
  const out: { -readonly [K in keyof SerializedCellBorders]: SerializedCellBorders[K] } = {}
  if (b.top) out.top = serializeCellBorder(b.top)
  if (b.left) out.left = serializeCellBorder(b.left)
  if (b.bottom) out.bottom = serializeCellBorder(b.bottom)
  if (b.right) out.right = serializeCellBorder(b.right)
  return out
}

/**
 * Canonical TableModel → wire SerializedTable.
 *
 * Tables the browser cannot safely regenerate (nested tables, anchored
 * shapes inside cells) return undefined — the block stays byte-preserved
 * read-only instead of getting a lossy editable payload.
 */
function serializeTableModel(model: TableModel): SerializedTable | undefined {
  // Nested tables / anchored shapes are display-only model extensions whose
  // content lives in cell paragraphs; regenerating from an editable payload
  // would silently drop them. Keep such tables on the passthrough path.
  for (const row of model.rows) {
    for (const cell of row) {
      if ((cell.nestedTables?.length ?? 0) > 0 || (cell.anchoredBoxes?.length ?? 0) > 0) {
        return undefined
      }
    }
  }
  const rows = model.rows.map((row) =>
    row.map((cell): SerializedTableCell => {
      const richParas = cell.richParas?.map((p): SerializedTableParagraph => ({
        runs: p.runs.map(serializeRun),
        ...(p.align ? { align: p.align } : {}),
        ...(p.styleId ? { styleId: p.styleId } : {}),
      }))
      const borders = cell.borders ? serializeCellBorders(cell.borders) : undefined
      return {
        paras: cell.paras,
        ...(richParas ? { richParas } : {}),
        ...(cell.colSpan && cell.colSpan > 1 ? { colSpan: cell.colSpan } : {}),
        ...(cell.vMerge ? { vMerge: cell.vMerge } : {}),
        ...(cell.fill ? { fill: cell.fill } : {}),
        ...(cell.color ? { color: cell.color } : {}),
        ...(cell.bold ? { bold: cell.bold } : {}),
        ...(cell.align ? { align: cell.align } : {}),
        ...(cell.vAlign && cell.vAlign !== 'top' ? { vAlign: cell.vAlign } : {}),
        ...(borders ? { borders } : {}),
        ...(cell.rawTcPr ? { rawTcPr: cell.rawTcPr } : {}),
      }
    }),
  )
  const headerRows = model.rawTrPrs?.map((trPr) => !!trPr && /<w:tblHeader[\s/>]/.test(trPr))
  const borders = model.borders
    ? (() => {
        const out = serializeCellBorders(model.borders)
        return {
          ...out,
          ...(model.borders.insideH ? { insideH: serializeCellBorder(model.borders.insideH) } : {}),
          ...(model.borders.insideV ? { insideV: serializeCellBorder(model.borders.insideV) } : {}),
        }
      })()
    : undefined
  return {
    rows,
    ...(model.colWidthsPct ? { colWidthsPct: model.colWidthsPct } : {}),
    ...(model.colWidthsTwips ? { colWidthsTwips: model.colWidthsTwips } : {}),
    ...(model.widthPct !== undefined ? { widthPct: model.widthPct } : {}),
    ...(model.autoLayout ? { autoLayout: model.autoLayout } : {}),
    ...(model.cellMarTwips ? { cellMarTwips: model.cellMarTwips } : {}),
    ...(borders ? { borders } : {}),
    ...(model.align ? { align: model.align } : {}),
    ...(model.indentTwips !== undefined ? { indentTwips: model.indentTwips } : {}),
    ...(model.rowHeightsTwips ? { rowHeightsTwips: model.rowHeightsTwips } : {}),
    ...(model.rowHeightRules ? { rowHeightRules: model.rowHeightRules } : {}),
    ...(model.rawTrPrs ? { rawTrPrs: model.rawTrPrs } : {}),
    ...(model.tblStyleId !== undefined ? { tblStyleId: model.tblStyleId } : {}),
    ...(model.bidiVisual ? { bidiVisual: model.bidiVisual } : {}),
    ...(headerRows?.some(Boolean) ? { headerRows } : {}),
  }
}

/** wire border → docx-engine CellBorder */
function toCellBorder(b: SerializedCellBorder): CellBorder {
  return {
    style: b.style,
    ...(b.szEighths !== undefined ? { szEighths: b.szEighths } : {}),
    ...(b.color !== undefined ? { color: b.color } : {}),
  }
}

/**
 * Patch the wire's headerRows flag into a row's raw <w:trPr> bytes:
 * add <w:tblHeader/> when the flag is set, strip it when clear.
 * (The canonical model keeps tblHeader inside rawTrPrs; the wire exposes it
 * as a boolean so the browser can toggle header rows without XML surgery.)
 */
function applyHeaderRowToTrPr(
  trPr: string | null | undefined,
  header: boolean | undefined,
): string | null | undefined {
  if (header === undefined) return trPr ?? undefined
  const has = !!trPr && /<w:tblHeader[\s/>]/.test(trPr)
  if (header === has) return trPr ?? undefined
  if (header) {
    const tag = '<w:tblHeader/>'
    if (!trPr) return `<w:trPr>${tag}</w:trPr>`
    // tblHeader is schema-valid anywhere inside trPr in practice; Word writes it early.
    return trPr.replace('</w:trPr>', `${tag}</w:trPr>`)
  }
  if (!trPr) return undefined
  const stripped = trPr.replace(/<w:tblHeader[^>]*\/>/, '')
  return /^<w:trPr(?:\s[^>]*)?>\s*<\/w:trPr>$/.test(stripped) ? undefined : stripped
}

/**
 * wire SerializedTable → canonical TableModel for the engine generator.
 * All XML generation stays inside the engine (generateTableModelXml).
 */
function toTableModel(t: SerializedTable): TableModel {
  const rows = t.rows.map((row) =>
    row.map((cell) => ({
      paras: [...cell.paras],
      ...(cell.richParas
        ? {
            richParas: cell.richParas.map((p) => ({
              runs: p.runs.map((r): Run => ({
                text: r.text,
                ...(r.bold ? { bold: true } : {}),
                ...(r.italic ? { italic: true } : {}),
                ...(r.underline ? { underline: true } : {}),
                ...(r.strike ? { strike: true } : {}),
                ...(r.link
                  ? {
                      link: {
                        href: r.link.href,
                        ...(r.link.tooltip ? { tooltip: r.link.tooltip } : {}),
                      },
                    }
                  : {}),
              })),
              ...(p.align ? { align: p.align } : {}),
              ...(p.styleId ? { styleId: p.styleId } : {}),
            })),
          }
        : {}),
      ...(cell.colSpan && cell.colSpan > 1 ? { colSpan: cell.colSpan } : {}),
      ...(cell.vMerge ? { vMerge: cell.vMerge } : {}),
      ...(cell.fill ? { fill: cell.fill } : {}),
      ...(cell.color ? { color: cell.color } : {}),
      ...(cell.bold ? { bold: cell.bold } : {}),
      ...(cell.align ? { align: cell.align } : {}),
      ...(cell.vAlign ? { vAlign: cell.vAlign } : {}),
      ...(cell.borders
        ? {
            borders: {
              ...(cell.borders.top ? { top: toCellBorder(cell.borders.top) } : {}),
              ...(cell.borders.left ? { left: toCellBorder(cell.borders.left) } : {}),
              ...(cell.borders.bottom ? { bottom: toCellBorder(cell.borders.bottom) } : {}),
              ...(cell.borders.right ? { right: toCellBorder(cell.borders.right) } : {}),
            },
          }
        : {}),
      ...(cell.rawTcPr ? { rawTcPr: cell.rawTcPr } : {}),
    })),
  )
  const rawTrPrs =
    t.rawTrPrs || t.headerRows
      ? t.rows.map((_, ri) => {
          const patched = applyHeaderRowToTrPr(t.rawTrPrs?.[ri], t.headerRows?.[ri])
          return patched ?? null
        })
      : undefined
  const model: TableModel = { rows }
  if (t.colWidthsPct) model.colWidthsPct = [...t.colWidthsPct]
  if (t.colWidthsTwips) model.colWidthsTwips = [...t.colWidthsTwips]
  if (t.widthPct !== undefined) model.widthPct = t.widthPct
  if (t.autoLayout) model.autoLayout = t.autoLayout
  if (t.cellMarTwips) model.cellMarTwips = { ...t.cellMarTwips }
  if (t.borders) {
    model.borders = {
      ...(t.borders.top ? { top: toCellBorder(t.borders.top) } : {}),
      ...(t.borders.left ? { left: toCellBorder(t.borders.left) } : {}),
      ...(t.borders.bottom ? { bottom: toCellBorder(t.borders.bottom) } : {}),
      ...(t.borders.right ? { right: toCellBorder(t.borders.right) } : {}),
      ...(t.borders.insideH ? { insideH: toCellBorder(t.borders.insideH) } : {}),
      ...(t.borders.insideV ? { insideV: toCellBorder(t.borders.insideV) } : {}),
    }
  }
  if (t.align) model.align = t.align
  if (t.indentTwips !== undefined) model.indentTwips = t.indentTwips
  if (t.rowHeightsTwips) model.rowHeightsTwips = [...t.rowHeightsTwips]
  if (t.rowHeightRules) model.rowHeightRules = [...t.rowHeightRules]
  if (rawTrPrs) model.rawTrPrs = rawTrPrs
  if (t.tblStyleId !== undefined) model.tblStyleId = t.tblStyleId
  if (t.bidiVisual) model.bidiVisual = t.bidiVisual
  return model
}

// ── Image model ↔ wire conversion (canonical Block image fields are the source) ──

/** Canonical ImageWrap ↔ wire (canonical model has no 'inline' member: undefined = inline). */
function wireWrapOf(block: Block): SerializedImageWrap | undefined {
  return block.imageWrap as SerializedImageWrap | undefined
}

/** Canonical Block image fields → wire SerializedImage (null when not editable). */
function serializeImage(block: Block): SerializedImage | undefined {
  // Only pure-image paragraphs with readable media are editable; broken
  // images and OLE previews stay byte-preserved read-only blocks.
  if (!block.imageDataUrl) return undefined
  return {
    imageDataUrl: block.imageDataUrl,
    ...(block.imageWidthPx !== undefined ? { widthPx: block.imageWidthPx } : {}),
    ...(block.imageHeightPx !== undefined ? { heightPx: block.imageHeightPx } : {}),
    ...(block.imageCrop ? { crop: { ...block.imageCrop } } : {}),
    ...(block.imageFillRect ? { fillRect: { ...block.imageFillRect } } : {}),
    ...(block.imageAlign ? { align: block.imageAlign } : {}),
    ...(wireWrapOf(block) ? { wrap: wireWrapOf(block) } : {}),
    ...(block.imageOffsetXEmu !== undefined ? { offsetXEmu: block.imageOffsetXEmu } : {}),
    ...(block.imageOffsetYEmu !== undefined ? { offsetYEmu: block.imageOffsetYEmu } : {}),
    ...(block.imagePosH ? { posH: block.imagePosH } : {}),
    ...(block.imagePosV ? { posV: block.imagePosV } : {}),
    ...(block.imagePosHRel ? { posHRel: block.imagePosHRel } : {}),
    ...(block.imagePosVRel ? { posVRel: block.imagePosVRel } : {}),
    ...(block.imageRotDeg !== undefined ? { rotDeg: block.imageRotDeg } : {}),
    ...(block.imageFlipH ? { flipH: true } : {}),
    ...(block.imageFlipV ? { flipV: true } : {}),
    ...(block.imageAlt ? { alt: block.imageAlt } : {}),
  }
}

/** Normalize a wire/canonical crop for diffing (undefined → all-zero). */
function cropOf(crop: { l: number; t: number; r: number; b: number } | undefined | null): {
  l: number
  t: number
  r: number
  b: number
} {
  return { l: crop?.l ?? 0, t: crop?.t ?? 0, r: crop?.r ?? 0, b: crop?.b ?? 0 }
}

function cropsEqual(
  a: { l: number; t: number; r: number; b: number },
  b: { l: number; t: number; r: number; b: number },
): boolean {
  return a.l === b.l && a.t === b.t && a.r === b.r && a.b === b.b
}

/**
 * Diff the wire's edited image state against the parsed original block and
 * build the canonical ImagePatch (mirrors the desktop editor's imagePatchOf).
 * `wrap`/`posH`/`posV` ride along for the applyImageWrap orchestration — the
 * canonical ImagePatch keeps them out because applyImageWrap is their writer
 * (same split the desktop editor makes with its ImageBlockPatch).
 * Returns null when nothing browser-editable changed.
 */
interface ImageBlockPatch extends ImagePatch {
  /** wrap mode change; undefined keeps (written by applyImageWrap) */
  wrap?: ImageWrap | null
  /** margin-relative align pair (Word position-gallery presets) */
  posH?: 'left' | 'center' | 'right'
  posV?: 'top' | 'center' | 'bottom'
}

function imagePatchFromWire(wire: SerializedImage, original: Block): ImageBlockPatch | null {
  const patch: ImageBlockPatch = {}
  const w = wire.widthPx ?? null
  const h = wire.heightPx ?? null
  if (
    w !== null &&
    h !== null &&
    (w !== (original.imageWidthPx ?? null) || h !== (original.imageHeightPx ?? null))
  ) {
    patch.widthPx = w
    patch.heightPx = h
  }
  const align = wire.align ?? null
  if (align !== (original.imageAlign ?? null)) patch.align = align
  const wrap = (wire.wrap === 'inline' ? null : wire.wrap) as ImageWrap | null
  const origWrap = (original.imageWrap as ImageWrap | null) ?? null
  if (wrap !== origWrap) patch.wrap = wrap
  const offX = wire.offsetXEmu
  if (offX !== undefined && offX !== (original.imageOffsetXEmu ?? undefined)) {
    patch.posOffsetX = offX
  }
  const offY = wire.offsetYEmu
  if (offY !== undefined && offY !== (original.imageOffsetYEmu ?? undefined)) {
    patch.posOffsetY = offY
  }
  const rot = wire.rotDeg ?? 0
  if (rot !== (original.imageRotDeg ?? 0)) patch.rotDeg = rot
  const flipH = wire.flipH === true
  if (flipH !== (original.imageFlipH ?? false)) patch.flipH = flipH
  const flipV = wire.flipV === true
  if (flipV !== (original.imageFlipV ?? false)) patch.flipV = flipV
  const crop = cropOf(wire.crop)
  if (!cropsEqual(crop, cropOf(original.imageCrop))) patch.crop = crop
  // alt text diff (tri-state): undefined = keep; null = clear descr;
  // non-empty = set descr. The wire validator strips control chars + bounds
  // length; the canonical generator escapes for the wp:docPr descr attribute.
  if (wire.alt !== undefined) {
    const wireAlt = wire.alt === null || wire.alt.length === 0 ? null : wire.alt
    const origAlt = original.imageAlt ?? null
    if (wireAlt !== origAlt) patch.alt = wireAlt
  }
  // margin-relative position presets are written by applyImageWrap only when
  // the patch carries a wrap; force it in when the preset changed
  const posH = wire.posH ?? null
  const posV = wire.posV ?? null
  if (
    posH &&
    posV &&
    (posH !== (original.imagePosH ?? null) || posV !== (original.imagePosV ?? null))
  ) {
    patch.posH = posH
    patch.posV = posV
    if (patch.wrap === undefined) patch.wrap = wrap
  }
  return Object.keys(patch).length > 0 ? patch : null
}

/** Wire SerializedNewImage → canonical NewImage for the engine embed path. */
function toNewImage(n: SerializedNewImage): NewImage {
  return {
    base64: n.base64,
    mime: n.mime,
    widthPx: n.widthPx,
    heightPx: n.heightPx,
    ...(n.align ? { align: n.align } : {}),
    ...(n.wrap && n.wrap !== 'inline' ? { wrap: n.wrap as ImageWrap } : {}),
    ...(n.rotDeg ? { rotDeg: n.rotDeg } : {}),
    ...(n.flipH ? { flipH: true } : {}),
    ...(n.flipV ? { flipV: true } : {}),
  }
}

function serializeBlock(block: Block): SerializedBlock {
  const runs = block.runs ?? []
  const text = runs.map((r) => r.text).join('')
  // Collapse docx-engine's rich BlockType into our wire set. Passthrough
  // covers charts/SmartArt/OLE, broken images and non-editable tables — the
  // browser shows their label. Editable tables/images carry typed payloads.
  const type: SerializedBlock['type'] = block.hidden
    ? 'hidden'
    : block.type === 'heading'
      ? 'heading'
      : block.type === 'listItem'
        ? 'listItem'
        : block.type === 'paragraph'
          ? 'paragraph'
          : block.type === 'table'
            ? 'table'
            : block.type === 'image'
              ? 'image'
              : 'passthrough'
  // Serialize runs with inline formatting marks (bold, italic, underline,
  // strike, link) so the browser editor can render faithful rich text.
  const serializedRuns = runs.length > 0 ? runs.map(serializeRun) : undefined
  // Editable table payload from the canonical TableModel (absent for tables
  // the browser cannot safely regenerate — those stay read-only passthrough).
  const table = type === 'table' && block.table ? serializeTableModel(block.table) : undefined
  // Editable image payload from the canonical image fields (absent for
  // images with unreadable media — those stay read-only passthrough).
  const image = type === 'image' ? serializeImage(block) : undefined
  return {
    docxIndex: block.docxIndex,
    type,
    text,
    ...(serializedRuns ? { runs: serializedRuns } : {}),
    ...(table ? { table } : {}),
    ...(image ? { image } : {}),
    level: block.level,
    listKind: block.list?.kind,
    edited: false,
    hidden: block.hidden === true,
  }
}

async function handleOpenDocument(
  body: unknown,
  codec: OfficeBinaryCodec,
): Promise<OfficeApiResponse> {
  const { fileBytes } = parseOpenDocumentRequest(body, codec)
  let parsed: ParsedDocFull
  try {
    parsed = await parseDocx(fileBytes)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to parse document',
    )
  }
  const blocks = parsed.blocks.map(serializeBlock)
  const res: OpenDocumentResponse = { blocks }
  return { status: 200, body: res }
}

/**
 * Convert the editor's serialized blocks back into the docx-engine's
 * SaveBlock[] shape.
 *
 * Strategy:
 *  - A block with `docxIndex` and `edited !== true` becomes
 *    `{ kind: 'original', docxIndex }` — the engine copies its bytes verbatim.
 *  - A block with `edited === true` becomes
 *    `{ kind: 'generated', block: { type, runs: [{ text }] } }` — the engine
 *    emits a fresh paragraph/heading/listItem using an existing style.
 *  - An edited table becomes `{ kind: 'xml', xml: generateTableModelXml(model, originalXml) }`
 *    — the CANONICAL engine generator emits the OOXML (tblPr preserved from
 *    the original bytes); no table XML is ever built in the browser.
 *  - A block without `docxIndex` (newly inserted in the browser) is also
 *    generated.
 *
 * Hidden trailing blocks (sectPr) are re-appended by the engine automatically.
 */
function toSaveBlocks(blocks: readonly SerializedBlock[], parsed: ParsedDocFull): SaveBlock[] {
  const out: SaveBlock[] = []
  for (const b of blocks) {
    if (b.hidden) continue // hidden trailing blocks are re-appended by the engine
    const edited = b.edited === true || b.docxIndex === null
    if (!edited && b.docxIndex !== null) {
      out.push({ kind: 'original', docxIndex: b.docxIndex })
      continue
    }
    // ── Tables: regenerate through the canonical engine generator ──────
    if (b.type === 'table') {
      if (!b.table) {
        // A table block without a payload cannot be regenerated; the wire
        // validation already rejects edited tables without payloads, so this
        // is only reachable for non-edited tables (kept by the branch above).
        // Byte-preserve defensively rather than lose the block.
        if (b.docxIndex !== null) {
          out.push({ kind: 'original', docxIndex: b.docxIndex })
          continue
        }
        throw new OfficeValidationError('validation', 'new table block carries no table payload')
      }
      const model = toTableModel(b.table)
      // tblPr (borders, cell margins, style, width, bidi…) is preserved by
      // feeding the original table's bytes to the generator — same rule the
      // desktop editor uses.
      const original =
        b.docxIndex !== null ? (parsed.blocks[b.docxIndex]?.originalXml ?? null) : null
      out.push({
        kind: 'xml',
        xml: generateTableModelXml(model, original ?? undefined),
      })
      continue
    }
    // ── Images: canonical patch/new-image paths (browser never builds XML) ──
    if (b.type === 'image') {
      // New image (editor-inserted): the engine creates the media part +
      // relationship + drawing from the canonical NewImage.
      if (b.docxIndex === null) {
        if (!b.newImage) {
          // Unreachable: wire validation rejects new image blocks without
          // a newImage payload.
          throw new OfficeValidationError(
            'validation',
            'new image block carries no newImage payload',
          )
        }
        out.push({ kind: 'image', image: toNewImage(b.newImage) })
        continue
      }
      // Existing image, edited: diff the wire state against the parsed
      // original and surgically patch only what changed — the untouched
      // drawing bytes (and any unmodeled properties) survive verbatim.
      const originalBlock = parsed.blocks[b.docxIndex]
      if (!originalBlock || originalBlock.type !== 'image' || !originalBlock.originalXml) {
        throw new OfficeValidationError(
          'validation',
          `edited image block references docxIndex ${b.docxIndex} which is not an image`,
        )
      }
      if (!b.image) {
        // Unreachable: wire validation rejects edited images without payloads.
        throw new OfficeValidationError('validation', 'edited image block carries no image payload')
      }
      const patch = imagePatchFromWire(b.image, originalBlock)
      if (!patch) {
        // Nothing actually changed (fingerprint false positive): keep the
        // original bytes rather than emit a no-op patch.
        out.push({ kind: 'original', docxIndex: b.docxIndex })
        continue
      }
      let xml = patchImageParagraphXml(originalBlock.originalXml, patch)
      if (patch.wrap !== undefined) {
        const posOffset =
          patch.posOffsetX !== undefined && patch.posOffsetY !== undefined
            ? { x: patch.posOffsetX, y: patch.posOffsetY }
            : undefined
        const marginAlign =
          posOffset === undefined && patch.posH && patch.posV
            ? { h: patch.posH, v: patch.posV }
            : undefined
        xml = applyImageWrap(xml, patch.wrap, posOffset, marginAlign)
      }
      out.push({ kind: 'xml', xml })
      continue
    }
    // Regenerate with run-level formatting (bold, italic, underline, strike, link).
    // If the browser sent `runs`, use them; otherwise fall back to plain text.
    const runs: Array<{
      text: string
      bold?: boolean
      italic?: boolean
      underline?: boolean
      strike?: boolean
      link?: { href: string; tooltip?: string }
    }> =
      b.runs && b.runs.length > 0
        ? b.runs.map((r) => ({
            text: r.text,
            ...(r.bold ? { bold: true } : {}),
            ...(r.italic ? { italic: true } : {}),
            ...(r.underline ? { underline: true } : {}),
            ...(r.strike ? { strike: true } : {}),
            ...(r.link ? { link: r.link } : {}),
          }))
        : [{ text: b.text }]
    if (b.type === 'heading') {
      out.push({
        kind: 'generated',
        block: {
          type: 'heading',
          level: b.level ?? 1,
          runs,
        },
      })
    } else if (b.type === 'listItem') {
      out.push({
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: b.listKind ?? 'bullet', numId: '1', ilvl: 0 },
          runs,
        },
      })
    } else {
      out.push({
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs,
        },
      })
    }
  }
  return out
}

async function handleSaveDocument(
  body: unknown,
  codec: OfficeBinaryCodec,
): Promise<OfficeApiResponse> {
  const req = parseSaveDocumentRequest(body, codec)
  let parsed: ParsedDocFull
  try {
    parsed = await parseDocx(req.fileBytes)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to parse document',
    )
  }
  const saveBlocks = toSaveBlocks(req.blocks, parsed)
  let saved: Uint8Array
  try {
    saved = await saveDocx(parsed, saveBlocks)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to save document',
    )
  }
  const res: SaveDocumentResponse = { fileBytes: encodeFileBytes(saved, codec) }
  return { status: 200, body: res }
}

// ── Router ──────────────────────────────────────────────────────────────────

/**
 * Optional services injected at the route boundary. All fields default to
 * the in-process base64 implementation so callers that don't care about
 * the codec get the original behavior.
 */
export interface OfficeRouteServices {
  /**
   * Byte↔wire codec used to encode/decode the fileBytes field. Defaults
   * to Base64Codec (Node Buffer-backed). Provide a custom implementation
   * to swap in a different transport (multipart, MessagePack, …) without
   * touching the route handlers.
   */
  readonly codec?: OfficeBinaryCodec
}

/**
 * Route an office API request. Returns null for non-office paths so the
 * caller (vercel-handler) can fall through to other handlers.
 *
 * Accepts an optional services object so tests (and future hosts) can
 * inject a custom OfficeBinaryCodec. When omitted, the default Base64Codec
 * is used.
 */
export async function routeOffice(
  req: OfficeApiRequest,
  services?: OfficeRouteServices,
): Promise<OfficeApiResponse | null> {
  const path = req.path
  const method = req.method
  const codec = services?.codec ?? DEFAULT_CODEC
  // Match the four canonical office routes. Anything else returns null so the
  // host can fall through to CoreApi / 404.
  try {
    if (method === 'POST' && path === '/office/workbooks/open') {
      return await handleOpenWorkbook(req.body, codec)
    }
    if (method === 'POST' && path === '/office/workbooks/save') {
      return await handleSaveWorkbook(req.body, codec)
    }
    if (method === 'POST' && path === '/office/documents/open') {
      return await handleOpenDocument(req.body, codec)
    }
    if (method === 'POST' && path === '/office/documents/save') {
      return await handleSaveDocument(req.body, codec)
    }
    return null
  } catch (e) {
    if (e instanceof OfficeValidationError) {
      const status = e.code === 'validation' ? 400 : 400
      return {
        status,
        body: { error: e.code, message: e.message },
      }
    }
    // Unexpected error — do not leak details
    return {
      status: 500,
      body: { error: 'internal', message: 'Office operation failed' },
    }
  }
}
