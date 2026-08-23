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
  type WorkbookSnapshot,
} from '@genoffice/xlsx-gateway'
import {
  parseDocx,
  saveDocx,
  type Block,
  type ParsedDocFull,
  type SaveBlock,
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
  // Extensibility seam — future mutation families land here as optional
  // readonly fields (structuralOps?, chartEdits?, hyperlinkEdits?, …).
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
 * Simplified Tiptap-compatible block representation.
 *
 * The browser renders blocks in Tiptap; the docx-engine's rich Block type
 * carries display-only fields (charts, ink, image data URLs…) that we do not
 * round-trip through JSON. SerializedBlock keeps just enough to:
 *   - render the block (type + text + level)
 *   - save the block back: `docxIndex` lets the server re-emit the original
 *     XML byte-identically when the block is unchanged; `edited` lets the
 *     server regenerate the paragraph from `text` when the user typed into it.
 *
 * Stable identity rule: `docxIndex` is the index of the original
 * <w:body> child the block came from. New blocks inserted by the editor
 * carry `docxIndex: null`. Deleted blocks are simply absent from the save
 * request. The browser must NOT renumber blocks by their visible position —
 * that would corrupt patch anchors when an edit changes the block order.
 */
export interface SerializedBlock {
  readonly docxIndex: number | null
  readonly type: 'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'passthrough' | 'hidden'
  readonly text: string
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
  constructor(public readonly code: string, message: string) {
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
function expectCellState(value: unknown, field: string): {
  readonly value: string | number | boolean | null
  readonly formula?: string
} {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `${field} must be an object`)
  }
  const v = value.value
  // `value` may be string | number | boolean | null. Anything else (object,
  // array, undefined) is rejected.
  if (
    v !== null &&
    typeof v !== 'string' &&
    typeof v !== 'number' &&
    typeof v !== 'boolean'
  ) {
    throw new OfficeValidationError('validation', `${field}.value must be a string, number, boolean, or null`)
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
function expectCellEdit(value: unknown, index: number): CellEdit {
  if (!isRecord(value)) {
    throw new OfficeValidationError('validation', `cellEdits[${index}] must be an object`)
  }
  const sheetName = expectString(value.sheetName, `cellEdits[${index}].sheetName`)
  const row = expectNumber(value.row, `cellEdits[${index}].row`)
  const column = expectNumber(value.column, `cellEdits[${index}].column`)
  const writeValue = expectBoolean(value.writeValue, `cellEdits[${index}].writeValue`)
  const cell = expectCellState(value.cell, `cellEdits[${index}].cell`)
  // Optional pass-through fields — keep them when the browser sends them.
  const edit: CellEdit = {
    sheetName,
    row,
    column,
    writeValue,
    cell,
    ...(typeof value.style === 'object' && value.style !== null
      ? { style: value.style as CellEdit['style'] }
      : {}),
    ...(Array.isArray(value.rich) ? { rich: value.rich as CellEdit['rich'] } : {}),
    ...(typeof value.styleReset === 'boolean' ? { styleReset: value.styleReset } : {}),
  }
  return edit
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
  const block: SerializedBlock = {
    docxIndex,
    type: type as SerializedBlock['type'],
    text,
    ...(value.level !== undefined ? { level: expectNumber(value.level, `blocks[${index}].level`) } : {}),
    ...(value.listKind !== undefined
      ? (() => {
          const k = value.listKind
          if (k !== 'bullet' && k !== 'ordered') {
            throw new OfficeValidationError('validation', `blocks[${index}].listKind must be 'bullet' or 'ordered'`)
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
  if (typeof name !== 'string') throw new OfficeValidationError('validation', 'fileName is required') // narrowed above
  if (name.length > 255) {
    throw new OfficeValidationError('validation', 'fileName exceeds 255 characters')
  }
  // Reject any path separator or traversal attempt — even percent-encoded.
  if (name.includes('/') || name.includes('\\') || name.includes('..') || name.includes('\0')) {
    throw new OfficeValidationError('validation', 'fileName may not contain path separators or traversal sequences')
  }
  if (!SAFE_NAME_RE.test(name)) {
    throw new OfficeValidationError('validation', 'fileName contains invalid characters (allowed: A-Z a-z 0-9 . - _)')
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
  if (typeof b64 !== 'string') throw new OfficeValidationError('validation', 'fileBytes (base64) is required') // narrowed above
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
function parseSaveWorkbookRequest(
  body: unknown,
  codec: OfficeBinaryCodec,
): ParsedSaveWorkbookRequest {
  if (!isRecord(body)) {
    throw new OfficeValidationError('validation', 'Request body must be a JSON object')
  }
  const fileName = validateFileName(body.fileName)
  const fileBytes = decodeFileBytes(body.fileBytes, codec)

  // Canonical path: savePlan.edits.
  if (body.savePlan !== undefined) {
    if (!isRecord(body.savePlan)) {
      throw new OfficeValidationError('validation', 'savePlan must be an object')
    }
    const edits = expectArray(body.savePlan.edits, 'savePlan.edits', expectCellEdit)
    return { fileName, fileBytes, savePlan: { edits } }
  }

  // Legacy path: top-level cellEdits.
  if (body.cellEdits !== undefined) {
    const edits = expectArray(body.cellEdits, 'cellEdits', expectCellEdit)
    return { fileName, fileBytes, savePlan: { edits } }
  }

  throw new OfficeValidationError(
    'validation',
    'savePlan.edits (or legacy cellEdits) is required',
  )
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
  let mutation
  try {
    mutation = await applyCellEditsToXlsx(buf, edits)
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

function serializeBlock(block: Block): SerializedBlock {
  const runs = block.runs ?? []
  const text = runs.map((r) => r.text).join('')
  // Collapse docx-engine's rich BlockType into our wire set. Passthrough
  // covers tables/images/charts/SmartArt/OLE — the browser shows their label.
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
  return {
    docxIndex: block.docxIndex,
    type,
    text,
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
 *  - A block without `docxIndex` (newly inserted in the browser) is also
 *    generated.
 *
 * Hidden trailing blocks (sectPr) are re-appended by the engine automatically.
 */
function toSaveBlocks(blocks: readonly SerializedBlock[]): SaveBlock[] {
  const out: SaveBlock[] = []
  for (const b of blocks) {
    if (b.hidden) continue // hidden trailing blocks are re-appended by the engine
    const edited = b.edited === true || b.docxIndex === null
    if (!edited && b.docxIndex !== null) {
      out.push({ kind: 'original', docxIndex: b.docxIndex })
      continue
    }
    // Regenerate as a paragraph/heading/listItem with plain-text runs.
    if (b.type === 'heading') {
      out.push({
        kind: 'generated',
        block: {
          type: 'heading',
          level: b.level ?? 1,
          runs: [{ text: b.text }],
        },
      })
    } else if (b.type === 'listItem') {
      out.push({
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: b.listKind ?? 'bullet', numId: '1', ilvl: 0 },
          runs: [{ text: b.text }],
        },
      })
    } else {
      out.push({
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: b.text }],
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
  const saveBlocks = toSaveBlocks(req.blocks)
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
