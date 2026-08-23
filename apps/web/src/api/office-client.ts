/**
 * Office API client — typed fetch wrapper for /api/office/* routes.
 *
 * The browser uses this to round-trip .xlsx and .docx files through the
 * pure office engines on the server. Bytes are base64-encoded inside JSON
 * envelopes (acceptable for fixture-sized files per the Phase-2 spec).
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
export interface SaveWorkbookRequest {
  readonly fileName: string
  readonly fileBytes: string
  readonly cellEdits: readonly CellEdit[]
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

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Decode a base64 string into a Uint8Array (browser-friendly; no Node Buffer). */
export function decodeBase64ToBytes(b64: string): Uint8Array {
  // Use the browser-native atob + Uint8Array. We avoid Buffer because the
  // browser must not pull in any Node polyfill.
  const binary = atob(b64)
  const len = binary.length
  const out = new Uint8Array(len)
  for (let i = 0; i < len; i++) out[i] = binary.charCodeAt(i)
  return out
}

/** Encode a Uint8Array (or ArrayBuffer) into a base64 string. */
export function encodeBytesToBase64(input: Uint8Array | ArrayBuffer): string {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  let binary = ''
  const CHUNK = 0x8000 // 32k chunks — avoids call-stack limits on large files
  for (let i = 0; i < bytes.length; i += CHUNK) {
    const end = Math.min(i + CHUNK, bytes.length)
    // Build a binary string from the byte slice. We avoid the TS-incompatible
    // `String.fromCharCode.apply(null, slice as unknown as number[])` cast by
    // pushing one char at a time within the chunk — modern engines optimize
    // this fine for the chunk sizes we use here.
    let chunk = ''
    for (let j = i; j < end; j++) chunk += String.fromCharCode(bytes[j])
    binary += chunk
  }
  return btoa(binary)
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
    const err = parsed as { error?: string; message?: string } | null
    throw new OfficeApiRequestError({
      status: res.status,
      error: err?.error ?? 'internal',
      message: err?.message ?? `Office API request failed (${res.status})`,
    })
  }
  return parsed as T
}

// ── Public API ──────────────────────────────────────────────────────────────

/**
 * Open an XLSX file: upload its bytes (base64-encoded) and receive the
 * server-side WorkbookSnapshot.
 */
export async function openWorkbook(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
}): Promise<OpenWorkbookResponse> {
  const fileBytes = encodeBytesToBase64(input.fileBytes)
  const req: OpenWorkbookRequest = { fileName: input.fileName, fileBytes }
  return postJson<OpenWorkbookResponse>('/workbooks/open', req)
}

/**
 * Save an XLSX file: send the source bytes plus the user's cell edits, and
 * receive the mutated bytes (base64-encoded). The caller decodes them.
 */
export async function saveWorkbook(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
  cellEdits: readonly CellEdit[]
}): Promise<Uint8Array> {
  const fileBytes = encodeBytesToBase64(input.fileBytes)
  const req: SaveWorkbookRequest = {
    fileName: input.fileName,
    fileBytes,
    cellEdits: input.cellEdits,
  }
  const res = await postJson<SaveWorkbookResponse>('/workbooks/save', req)
  return decodeBase64ToBytes(res.fileBytes)
}

/**
 * Open a DOCX file: upload its bytes (base64-encoded) and receive the
 * simplified SerializedBlock[] representation for Tiptap.
 */
export async function openDocument(input: {
  fileName: string
  fileBytes: Uint8Array | ArrayBuffer
}): Promise<OpenDocumentResponse> {
  const fileBytes = encodeBytesToBase64(input.fileBytes)
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
}): Promise<Uint8Array> {
  const fileBytes = encodeBytesToBase64(input.fileBytes)
  const req: SaveDocumentRequest = {
    fileName: input.fileName,
    fileBytes,
    blocks: input.blocks,
  }
  const res = await postJson<SaveDocumentResponse>('/documents/save', req)
  return decodeBase64ToBytes(res.fileBytes)
}

/** Read a browser File into a Uint8Array. */
export async function readFileBytes(file: File): Promise<Uint8Array> {
  const buf = await file.arrayBuffer()
  return new Uint8Array(buf)
}
