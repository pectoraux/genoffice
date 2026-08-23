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
import type {
  CellEdit,
  WorkbookSnapshot,
} from '@genoffice/xlsx-gateway'

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
  | 'paragraph'
  | 'heading'
  | 'listItem'
  | 'table'
  | 'image'
  | 'passthrough'
  | 'hidden'

export interface SerializedBlock {
  readonly docxIndex: number | null
  readonly type: SerializedBlockType
  readonly text: string
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
  const v = value as { error?: unknown; message?: unknown }
  return typeof v.error === 'string' && typeof v.message === 'string'
}

async function postJson<T>(path: string, body: unknown): Promise<T> {
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
  // The caller is responsible for shaping `parsed` into the expected
  // response type. We return `parsed` cast to T because the wire-shape
  // validation already happened server-side; doing a full client-side
  // runtime check on every response would duplicate the route's
  // validators. Callers that want full type safety can wrap the result
  // in their own type guard.
  return parsed as T
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
  return postJson<OpenWorkbookResponse>('/workbooks/open', req)
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
  const res = await postJson<SaveWorkbookResponse>('/workbooks/save', req)
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
  return postJson<OpenDocumentResponse>('/documents/open', req)
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
  const res = await postJson<SaveDocumentResponse>('/documents/save', req)
  return codec.decode(res.fileBytes)
}

/** Read a browser File into a Uint8Array. */
export async function readFileBytes(file: File): Promise<Uint8Array> {
  const buf = await file.arrayBuffer()
  return new Uint8Array(buf)
}
