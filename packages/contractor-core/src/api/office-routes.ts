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
  type CellEdit,
  type EditableBorderStyle,
  type SheetPageSetupState,
  type SheetStructuralOps,
  type StructuralOp,
  type WorkbookSnapshot,
  type WorkbookStyleEdit,
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

const ROWCOL_STRUCTURAL_KINDS = ['insert-rows', 'remove-rows', 'insert-cols', 'remove-cols'] as const
const MERGE_STRUCTURAL_KINDS = ['merge-cells', 'unmerge-cells'] as const
const STRUCTURAL_KINDS = [...ROWCOL_STRUCTURAL_KINDS, ...MERGE_STRUCTURAL_KINDS]
const MAX_STRUCTURAL_COUNT = 10_000

/**
 * Validate a StructuralOp from the wire. The browser emits
 * insert/remove row/column ops (with index/count) AND merge/unmerge ops
 * (with range). Malformed ops throw OfficeValidationError → the existing
 * 400 validation error shape.
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
    return {
      fileName,
      fileBytes,
      savePlan: {
        edits,
        ...(structuralOps ? { structuralOps } : {}),
        ...(pageSetupStates ? { pageSetupStates } : {}),
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
  let mutation
  try {
    mutation = await applyCellEditsToXlsx(buf, edits, structuralOps, [], undefined, [], [], [], [], [], null, pageSetupStates)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to apply cell edits',
    )
  }
  const res: SaveWorkbookResponse = { fileBytes: encodeFileBytes(mutation.buffer, codec) }
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
