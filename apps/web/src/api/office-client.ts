/**
 * Office API client — typed fetch wrapper for /api/office/* routes.
 *
 * The browser uses this to round-trip .xlsx and .docx files through the
 * pure office engines on the server. Bytes are base64-encoded inside JSON
 * envelopes (acceptable for fixture-sized files per the Phase-2 spec), but
 * the byte transport is abstracted behind `OfficeBinaryCodec` so a future
 * binary transport (multipart, MessagePack, …) can be swapped in without
 * touching the call sites.
 *
 * Browser-only. Uses ONLY `fetch` (no Node APIs, no Electron).
 *
 * Purity: zero imports of `electron`, `node:*`, `fs`, `child_process`, or
 * `@genoffice/platform-electron`. Type-only imports from
 * `@genoffice/xlsx-gateway` are erased at compile time and pull in no
 * runtime code.
 */

// Type-only imports — erased at build time. The browser never bundles the
// engine packages; these types exist purely so the client's wire shapes line
// up with the server-side route handlers.
import type { CellEdit, WorkbookSnapshot } from '@genoffice/xlsx-gateway'

// ── Wire types (mirror the @contractor/core/api office-routes module) ────────

export interface OpenWorkbookRequest {
  readonly fileName: string
  readonly fileBytes: string
}
export interface OpenWorkbookResponse {
  readonly snapshot: WorkbookSnapshot
  readonly sheetNamesById: Readonly<Record<string, string>>
}

/**
 * Save-mutation plan for a workbook. Mirrors the server-side
 * `BrowserWorkbookSavePlan` shape: `edits` carries the CellEdit[] today,
 * and the open index signature lets future mutation families land without
 * a wire-breaking change.
 */
export interface BrowserWorkbookSavePlan {
  readonly edits: readonly CellEdit[]
  readonly [key: string]: unknown
}

export interface BrowserWorkbookSaveRequest {
  readonly fileName: string
  readonly fileBytes: string
  readonly savePlan: BrowserWorkbookSavePlan
}

export interface SaveWorkbookResponse {
  readonly fileBytes: string
}

export type SerializedBlockType =
  'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'passthrough' | 'hidden'

export interface SerializedRun {
  readonly text: string
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  readonly link?: { readonly href: string; readonly tooltip?: string }
}

// ── Serialized table (browser mirror of the server wire type) ──────────────

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

export interface SerializedTableParagraph {
  readonly runs: readonly SerializedRun[]
  readonly align?: 'left' | 'center' | 'right' | 'justify' | 'distribute'
  readonly styleId?: string
}

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
  readonly headerRows?: readonly boolean[]
}

export interface SerializedBlock {
  readonly docxIndex: number | null
  readonly type: SerializedBlockType
  readonly text: string
  readonly runs?: readonly SerializedRun[]
  readonly table?: SerializedTable
  readonly level?: number
  readonly listKind?: 'bullet' | 'ordered'
  readonly edited?: boolean
  readonly hidden?: boolean
}

export interface OpenDocumentRequest {
  readonly fileName: string
  readonly fileBytes: string
}
export interface OpenDocumentResponse {
  readonly blocks: readonly SerializedBlock[]
}
export interface SaveDocumentRequest {
  readonly fileName: string
  readonly fileBytes: string
  readonly blocks: readonly SerializedBlock[]
}
export interface SaveDocumentResponse {
  readonly fileBytes: string
}

// ── Errors ───────────────────────────────────────────────────────────────────

export interface OfficeApiError {
  readonly status: number
  readonly error: string
  readonly message: string
}

export class OfficeApiRequestError extends Error {
  readonly status: number
  readonly error: string
  constructor(e: OfficeApiError) {
    super(e.message)
    this.name = 'OfficeApiRequestError'
    this.status = e.status
    this.error = e.error
  }
}

// ── Byte transport codec ────────────────────────────────────────────────────

/**
 * Browser-side mirror of the server-side `OfficeBinaryCodec` interface.
 * The route handlers depend only on the interface; the default
 * `BrowserBase64Codec` uses `atob`/`btoa` so the browser bundle never
 * pulls in a Node Buffer polyfill.
 *
 * To swap in a different transport (e.g. a binary protocol with custom
 * framing), provide a custom `OfficeBinaryCodec` implementation to the
 * API functions.
 */
export interface OfficeBinaryCodec {
  encode(bytes: Uint8Array): string
  decode(encoded: string): Uint8Array
}

/**
 * Default base64 codec using the browser-native `atob`/`btoa` globals.
 * This is the browser counterpart to `Base64Codec` in office-routes.ts —
 * both produce/consume the same wire bytes, but this one avoids any Node
 * Buffer dependency so the browser bundle stays Node-free.
 */
export class BrowserBase64Codec implements OfficeBinaryCodec {
  encode(bytes: Uint8Array): string {
    let binary = ''
    const CHUNK = 0x8000 // 32k chunks — avoids call-stack limits on large files
    for (let i = 0; i < bytes.length; i += CHUNK) {
      const end = Math.min(i + CHUNK, bytes.length)
      let chunk = ''
      for (let j = i; j < end; j++) chunk += String.fromCharCode(bytes[j])
      binary += chunk
    }
    return btoa(binary)
  }
  decode(encoded: string): Uint8Array {
    const binary = atob(encoded)
    const len = binary.length
    const out = new Uint8Array(len)
    for (let i = 0; i < len; i++) out[i] = binary.charCodeAt(i)
    return out
  }
}

const DEFAULT_CODEC: OfficeBinaryCodec = new BrowserBase64Codec()

// ── Document/workbook handles ───────────────────────────────────────────────

/**
 * Stable handle to an open XLSX workbook. Replaces the editor's ad-hoc
 * `sourceBytesRef` + `fileNameRef` pair: every save returns a new handle
 * whose `revision` is incremented, so callers can tell at a glance whether
 * they're holding the latest persisted state.
 *
 * The handle is browser-only — it holds the source bytes (the canonical
 * engine input) and the display name. It carries no engine state.
 */
export interface OfficeWorkbookHandle {
  readonly fileName: string
  readonly sourceBytes: Uint8Array
  /** Monotonically incremented on each successful save. 0 = freshly opened. */
  readonly revision: number
}

/**
 * Stable handle to an open DOCX document. Like `OfficeWorkbookHandle`
 * minus the revision counter — the DOCX engine's patch-save is keyed by
 * `docxIndex` anchors rather than by a revision counter, so a single
 * (fileName, sourceBytes) pair is sufficient identity.
 */
export interface OfficeDocumentHandle {
  readonly fileName: string
  readonly sourceBytes: Uint8Array
}

/** Construct a workbook handle from a freshly opened file. */
export function createWorkbookHandle(
  fileName: string,
  sourceBytes: Uint8Array,
): OfficeWorkbookHandle {
  return { fileName, sourceBytes, revision: 0 }
}

/** Construct a document handle from a freshly opened file. */
export function createDocumentHandle(
  fileName: string,
  sourceBytes: Uint8Array,
): OfficeDocumentHandle {
  return { fileName, sourceBytes }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Decode a base64 string into a Uint8Array (browser-friendly; no Node Buffer). */
export function decodeBase64ToBytes(b64: string): Uint8Array {
  return DEFAULT_CODEC.decode(b64)
}

/** Encode a Uint8Array (or ArrayBuffer) into a base64 string. */
export function encodeBytesToBase64(input: Uint8Array | ArrayBuffer): string {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  return DEFAULT_CODEC.encode(bytes)
}

/** Trigger a browser download for a blob of bytes. */
export function downloadBytes(
  bytes: Uint8Array,
  fileName: string,
  mime = 'application/octet-stream',
): void {
  // Copy into a fresh ArrayBuffer so TS is satisfied that the Blob part is a
  // real ArrayBuffer (not a SharedArrayBuffer view).
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  const blob = new Blob([copy.buffer], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  // Defer revocation so the browser has time to start the download.
  setTimeout(() => URL.revokeObjectURL(url), 0)
}

// ── Core request helper ──────────────────────────────────────────────────────

/**
 * Runtime type guard for an OfficeApiError-shaped JSON body. Replaces the
 * previous `parsed as { error?: string; message?: string } | null` cast —
 * the cast silently trusted the server, this guard verifies the shape.
 */
function isOfficeApiErrorBody(value: unknown): value is { error: string; message: string } {
  if (typeof value !== 'object' || value === null) return false
  const v = value as Record<string, unknown>
  return typeof v.error === 'string' && typeof v.message === 'string'
}

// ── Response type guards ────────────────────────────────────────────────────

function isString(v: unknown): v is string {
  return typeof v === 'string'
}
function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

function isSerializedRun(v: unknown): v is SerializedRun {
  if (!isObject(v)) return false
  if (!isString(v.text)) return false
  if (v.bold !== undefined && typeof v.bold !== 'boolean') return false
  if (v.italic !== undefined && typeof v.italic !== 'boolean') return false
  if (v.underline !== undefined && typeof v.underline !== 'boolean') return false
  if (v.strike !== undefined && typeof v.strike !== 'boolean') return false
  if (v.link !== undefined) {
    if (!isObject(v.link) || !isString(v.link.href)) return false
  }
  return true
}

// ── Table payload guards (mirror the server-side validation shape) ─────────

const TABLE_PARA_ALIGNS = ['left', 'center', 'right', 'justify', 'distribute']
const TABLE_ALIGNS = ['left', 'center', 'right']
const TABLE_VALIGNS = ['top', 'center', 'bottom']
const MAX_TABLE_ROWS = 1000
const MAX_TABLE_COLS = 63

function isCellBorder(v: unknown): v is SerializedCellBorder {
  if (!isObject(v)) return false
  if (!isString(v.style)) return false
  if (v.szEighths !== undefined && typeof v.szEighths !== 'number') return false
  if (v.color !== undefined && !isString(v.color)) return false
  return true
}

function isCellBorders(v: unknown): v is SerializedCellBorders {
  if (!isObject(v)) return false
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    if (v[edge] !== undefined && !isCellBorder(v[edge])) return false
  }
  return true
}

function isTableBorders(v: unknown): v is SerializedTableBorders {
  if (!isCellBorders(v)) return false
  if ((v as Record<string, unknown>).insideH !== undefined) {
    if (!isCellBorder((v as Record<string, unknown>).insideH)) return false
  }
  if ((v as Record<string, unknown>).insideV !== undefined) {
    if (!isCellBorder((v as Record<string, unknown>).insideV)) return false
  }
  return true
}

function isCellMargins(v: unknown): v is SerializedCellMargins {
  if (!isObject(v)) return false
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    if (v[side] !== undefined && typeof v[side] !== 'number') return false
  }
  return true
}

function isTableParagraph(v: unknown): v is SerializedTableParagraph {
  if (!isObject(v)) return false
  if (!Array.isArray(v.runs) || !v.runs.every(isSerializedRun)) return false
  if (v.align !== undefined && (!isString(v.align) || !TABLE_PARA_ALIGNS.includes(v.align))) {
    return false
  }
  if (v.styleId !== undefined && !isString(v.styleId)) return false
  return true
}

function isTableCell(v: unknown): v is SerializedTableCell {
  if (!isObject(v)) return false
  if (!Array.isArray(v.paras) || !v.paras.every((p) => typeof p === 'string')) return false
  if (v.richParas !== undefined) {
    if (!Array.isArray(v.richParas) || !v.richParas.every(isTableParagraph)) return false
  }
  if (
    v.colSpan !== undefined &&
    (!Number.isInteger(v.colSpan) ||
      (v.colSpan as number) < 1 ||
      (v.colSpan as number) > MAX_TABLE_COLS)
  ) {
    return false
  }
  if (v.vMerge !== undefined && v.vMerge !== 'restart' && v.vMerge !== 'continue') return false
  if (v.fill !== undefined && !isString(v.fill)) return false
  if (v.color !== undefined && !isString(v.color)) return false
  if (v.bold !== undefined && typeof v.bold !== 'boolean') return false
  if (v.align !== undefined && (!isString(v.align) || !TABLE_PARA_ALIGNS.includes(v.align))) {
    return false
  }
  if (v.vAlign !== undefined && (!isString(v.vAlign) || !TABLE_VALIGNS.includes(v.vAlign))) {
    return false
  }
  if (v.borders !== undefined && !isCellBorders(v.borders)) return false
  if (v.rawTcPr !== undefined && !isString(v.rawTcPr)) return false
  return true
}

function isSerializedTable(v: unknown): v is SerializedTable {
  if (!isObject(v)) return false
  if (!Array.isArray(v.rows) || v.rows.length === 0 || v.rows.length > MAX_TABLE_ROWS) return false
  for (const row of v.rows) {
    if (!Array.isArray(row) || row.length === 0 || row.length > MAX_TABLE_COLS) return false
    if (!row.every(isTableCell)) return false
    const width = row.reduce(
      (sum: number, cell: SerializedTableCell) => sum + (cell.colSpan ?? 1),
      0,
    )
    if (width > MAX_TABLE_COLS) return false
  }
  for (const key of ['colWidthsPct', 'colWidthsTwips'] as const) {
    if (v[key] !== undefined) {
      if (!Array.isArray(v[key]) || !(v[key] as unknown[]).every((n) => typeof n === 'number')) {
        return false
      }
    }
  }
  if (v.widthPct !== undefined && typeof v.widthPct !== 'number') return false
  if (v.autoLayout !== undefined && typeof v.autoLayout !== 'boolean') return false
  if (v.cellMarTwips !== undefined && !isCellMargins(v.cellMarTwips)) return false
  if (v.borders !== undefined && !isTableBorders(v.borders)) return false
  if (v.align !== undefined && (!isString(v.align) || !TABLE_ALIGNS.includes(v.align))) {
    return false
  }
  if (v.indentTwips !== undefined && typeof v.indentTwips !== 'number') return false
  if (v.rowHeightsTwips !== undefined) {
    if (
      !Array.isArray(v.rowHeightsTwips) ||
      !v.rowHeightsTwips.every((h) => h === null || typeof h === 'number')
    ) {
      return false
    }
  }
  if (v.rowHeightRules !== undefined) {
    if (
      !Array.isArray(v.rowHeightRules) ||
      !v.rowHeightRules.every((r) => r === null || r === 'atLeast' || r === 'exact')
    ) {
      return false
    }
  }
  if (v.rawTrPrs !== undefined) {
    if (!Array.isArray(v.rawTrPrs) || !v.rawTrPrs.every((r) => r === null || isString(r))) {
      return false
    }
  }
  if (v.tblStyleId !== undefined && !isString(v.tblStyleId)) return false
  if (v.bidiVisual !== undefined && typeof v.bidiVisual !== 'boolean') return false
  if (v.headerRows !== undefined) {
    if (!Array.isArray(v.headerRows) || !v.headerRows.every((h) => typeof h === 'boolean')) {
      return false
    }
  }
  return true
}

function isSerializedBlock(v: unknown): v is SerializedBlock {
  if (!isObject(v)) return false
  if (v.docxIndex !== null && typeof v.docxIndex !== 'number') return false
  if (!isString(v.type)) return false
  if (!isString(v.text)) return false
  if (v.runs !== undefined) {
    if (!Array.isArray(v.runs)) return false
    if (!v.runs.every(isSerializedRun)) return false
  }
  if (v.table !== undefined) {
    if (v.type !== 'table') return false
    if (!isSerializedTable(v.table)) return false
  }
  if (v.level !== undefined && typeof v.level !== 'number') return false
  if (v.listKind !== undefined && v.listKind !== 'bullet' && v.listKind !== 'ordered') return false
  if (v.edited !== undefined && typeof v.edited !== 'boolean') return false
  if (v.hidden !== undefined && typeof v.hidden !== 'boolean') return false
  return true
}

function isOpenWorkbookResponse(v: unknown): v is OpenWorkbookResponse {
  if (!isObject(v)) return false
  if (!isObject(v.snapshot)) return false
  if (!Array.isArray(v.snapshot.sheets)) return false
  for (const sheet of v.snapshot.sheets) {
    if (!isObject(sheet)) return false
    if (!isString(sheet.id)) return false
    if (!isString(sheet.name)) return false
    if (!isObject(sheet.cells)) return false
  }
  if (!isObject(v.sheetNamesById)) return false
  return true
}

function isSaveWorkbookResponse(v: unknown): v is SaveWorkbookResponse {
  if (!isObject(v)) return false
  return isString(v.fileBytes)
}

function isOpenDocumentResponse(v: unknown): v is OpenDocumentResponse {
  if (!isObject(v)) return false
  if (!Array.isArray(v.blocks)) return false
  return v.blocks.every(isSerializedBlock)
}

function isSaveDocumentResponse(v: unknown): v is SaveDocumentResponse {
  if (!isObject(v)) return false
  return isString(v.fileBytes)
}

async function postJson<T>(path: string, body: unknown, guard: (v: unknown) => v is T): Promise<T> {
  const res = await fetch(`/api/office${path}`, {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  const text = await res.text()
  let parsed: unknown
  try {
    parsed = text.length > 0 ? JSON.parse(text) : null
  } catch {
    throw new OfficeApiRequestError({
      status: res.status,
      error: 'internal',
      message: 'Invalid JSON response from server',
    })
  }
  if (!res.ok) {
    if (isOfficeApiErrorBody(parsed)) {
      throw new OfficeApiRequestError({
        status: res.status,
        error: parsed.error,
        message: parsed.message,
      })
    }
    throw new OfficeApiRequestError({
      status: res.status,
      error: 'internal',
      message: `Office API request failed (${res.status})`,
    })
  }
  // Runtime-verify the response shape — guard is mandatory.
  if (!guard(parsed)) {
    throw new OfficeApiRequestError({
      status: res.status,
      error: 'internal',
      message: 'Office API returned a malformed response (failed type guard)',
    })
  }
  return parsed
}

// ── Public API ──────────────────────────────────────────────────────────────

/**
 * Open an XLSX file: upload its bytes (base64-encoded) and receive the
 * server-side WorkbookSnapshot. Returns the snapshot — the caller pairs
 * it with the source bytes to construct an `OfficeWorkbookHandle`.
 */
export async function openWorkbook(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
  codec?: OfficeBinaryCodec
}): Promise<OpenWorkbookResponse> {
  const codec = input.codec ?? DEFAULT_CODEC
  const fileBytes = codec.encode(
    input.fileBytes instanceof Uint8Array ? input.fileBytes : new Uint8Array(input.fileBytes),
  )
  const req: OpenWorkbookRequest = { fileName: input.fileName, fileBytes }
  return postJson<OpenWorkbookResponse>('/workbooks/open', req, isOpenWorkbookResponse)
}

/**
 * Save an XLSX file: send the source bytes plus the user's save plan, and
 * receive the mutated bytes (base64-encoded). The caller decodes them.
 *
 * The request body uses the canonical `BrowserWorkbookSaveRequest` shape
 * (`{ fileName, fileBytes, savePlan: { edits } }`). The server accepts
 * both this shape and the legacy `{ fileName, fileBytes, cellEdits }`
 * shape, so older callers keep working, but the browser always sends
 * the canonical form.
 */
export async function saveWorkbook(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
  savePlan: BrowserWorkbookSavePlan
  codec?: OfficeBinaryCodec
}): Promise<Uint8Array> {
  const codec = input.codec ?? DEFAULT_CODEC
  const fileBytes = codec.encode(
    input.fileBytes instanceof Uint8Array ? input.fileBytes : new Uint8Array(input.fileBytes),
  )
  const req: BrowserWorkbookSaveRequest = {
    fileName: input.fileName,
    fileBytes,
    savePlan: input.savePlan,
  }
  const res = await postJson<SaveWorkbookResponse>('/workbooks/save', req, isSaveWorkbookResponse)
  return codec.decode(res.fileBytes)
}

/**
 * Open a DOCX file: upload its bytes (base64-encoded) and receive the
 * simplified SerializedBlock[] representation for Tiptap. The caller pairs
 * the result with the source bytes to construct an `OfficeDocumentHandle`.
 */
export async function openDocument(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
  codec?: OfficeBinaryCodec
}): Promise<OpenDocumentResponse> {
  const codec = input.codec ?? DEFAULT_CODEC
  const fileBytes = codec.encode(
    input.fileBytes instanceof Uint8Array ? input.fileBytes : new Uint8Array(input.fileBytes),
  )
  const req: OpenDocumentRequest = { fileName: input.fileName, fileBytes }
  return postJson<OpenDocumentResponse>('/documents/open', req, isOpenDocumentResponse)
}

/**
 * Save a DOCX file: send the source bytes plus the editor's blocks, and
 * receive the patched DOCX bytes. The caller decodes them.
 */
export async function saveDocument(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
  blocks: readonly SerializedBlock[]
  codec?: OfficeBinaryCodec
}): Promise<Uint8Array> {
  const codec = input.codec ?? DEFAULT_CODEC
  const fileBytes = codec.encode(
    input.fileBytes instanceof Uint8Array ? input.fileBytes : new Uint8Array(input.fileBytes),
  )
  const req: SaveDocumentRequest = {
    fileName: input.fileName,
    fileBytes,
    blocks: input.blocks,
  }
  const res = await postJson<SaveDocumentResponse>('/documents/save', req, isSaveDocumentResponse)
  return codec.decode(res.fileBytes)
}

/** Read a browser File into a Uint8Array. */
export async function readFileBytes(file: File): Promise<Uint8Array> {
  const buf = await file.arrayBuffer()
  return new Uint8Array(buf)
}
