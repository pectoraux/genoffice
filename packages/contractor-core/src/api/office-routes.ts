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
 *
 * Purity: zero Electron imports, zero node:* imports except for the Buffer
 * global (used by @genoffice/xlsx-gateway). The host owns the HTTP transport;
 * this module owns request validation + engine delegation.
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
export interface SaveWorkbookRequest {
  readonly fileName: string
  readonly fileBytes: string // base64-encoded source XLSX
  readonly cellEdits: readonly CellEdit[]
}
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

function validateFileName(name: unknown): string {
  if (typeof name !== 'string' || name.length === 0) {
    throw new OfficeValidationError('validation', 'fileName is required')
  }
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

function decodeFileBytes(b64: unknown): Buffer {
  if (typeof b64 !== 'string' || b64.length === 0) {
    throw new OfficeValidationError('validation', 'fileBytes (base64) is required')
  }
  let buf: Buffer
  try {
    buf = Buffer.from(b64, 'base64')
  } catch {
    throw new OfficeValidationError('validation', 'fileBytes is not valid base64')
  }
  if (buf.length === 0) {
    throw new OfficeValidationError('validation', 'fileBytes decoded to an empty buffer')
  }
  if (buf.length > MAX_FILE_BYTES) {
    throw new OfficeValidationError(
      'validation',
      `fileBytes exceeds the 10MB limit (decoded to ${buf.length} bytes)`,
    )
  }
  return buf
}

function encodeFileBytes(buf: Uint8Array): string {
  // Normalize to a Node Buffer so base64 encoding is deterministic across runtimes.
  const out = buf instanceof Buffer ? buf : Buffer.from(buf)
  return out.toString('base64')
}

function asObject(body: unknown): Record<string, unknown> {
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    throw new OfficeValidationError('validation', 'Request body must be a JSON object')
  }
  return body as Record<string, unknown>
}

function asArray<T = unknown>(value: unknown, field: string): readonly T[] {
  if (!Array.isArray(value)) {
    throw new OfficeValidationError('validation', `${field} must be an array`)
  }
  return value as readonly T[]
}

// ── Workbook (XLSX) handlers ────────────────────────────────────────────────

async function handleOpenWorkbook(body: unknown): Promise<OfficeApiResponse> {
  const obj = asObject(body)
  validateFileName(obj.fileName)
  const buf = decodeFileBytes(obj.fileBytes)
  let imported
  try {
    imported = await readBasicWorkbook(buf)
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

async function handleSaveWorkbook(body: unknown): Promise<OfficeApiResponse> {
  const obj = asObject(body)
  validateFileName(obj.fileName)
  const buf = decodeFileBytes(obj.fileBytes)
  const edits = asArray<CellEdit>(obj.cellEdits, 'cellEdits')
  // Validate each edit has the required shape so an attacker can't smuggle
  // unexpected keys past the engine. The engine itself also validates, but
  // we fail closed here with a clearer message.
  for (const edit of edits) {
    if (!edit || typeof edit !== 'object') {
      throw new OfficeValidationError('validation', 'cellEdits contains a non-object entry')
    }
    if (typeof edit.sheetName !== 'string' || edit.sheetName.length === 0) {
      throw new OfficeValidationError('validation', 'cellEdits entry is missing sheetName')
    }
    if (typeof edit.row !== 'number' || typeof edit.column !== 'number') {
      throw new OfficeValidationError('validation', 'cellEdits entry is missing row/column')
    }
    if (typeof edit.writeValue !== 'boolean') {
      throw new OfficeValidationError('validation', 'cellEdits entry is missing writeValue')
    }
    if (!edit.cell || typeof edit.cell !== 'object') {
      throw new OfficeValidationError('validation', 'cellEdits entry is missing cell')
    }
  }
  let mutation
  try {
    mutation = await applyCellEditsToXlsx(buf, edits)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to apply cell edits',
    )
  }
  const res: SaveWorkbookResponse = { fileBytes: encodeFileBytes(mutation.buffer) }
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

async function handleOpenDocument(body: unknown): Promise<OfficeApiResponse> {
  const obj = asObject(body)
  validateFileName(obj.fileName)
  const buf = decodeFileBytes(obj.fileBytes)
  let parsed: ParsedDocFull
  try {
    parsed = await parseDocx(buf)
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

async function handleSaveDocument(body: unknown): Promise<OfficeApiResponse> {
  const obj = asObject(body)
  validateFileName(obj.fileName)
  const buf = decodeFileBytes(obj.fileBytes)
  const blocks = asArray<SerializedBlock>(obj.blocks, 'blocks')
  let parsed: ParsedDocFull
  try {
    parsed = await parseDocx(buf)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to parse document',
    )
  }
  const saveBlocks = toSaveBlocks(blocks)
  let saved: Uint8Array
  try {
    saved = await saveDocx(parsed, saveBlocks)
  } catch (e) {
    throw new OfficeValidationError(
      'malformed',
      e instanceof Error ? e.message : 'Failed to save document',
    )
  }
  const res: SaveDocumentResponse = { fileBytes: encodeFileBytes(saved) }
  return { status: 200, body: res }
}

// ── Router ──────────────────────────────────────────────────────────────────

/**
 * Route an office API request. Returns null for non-office paths so the
 * caller (vercel-handler) can fall through to other handlers.
 */
export async function routeOffice(req: OfficeApiRequest): Promise<OfficeApiResponse | null> {
  const path = req.path
  const method = req.method
  // Match the four canonical office routes. Anything else returns null so the
  // host can fall through to CoreApi / 404.
  try {
    if (method === 'POST' && path === '/office/workbooks/open') {
      return await handleOpenWorkbook(req.body)
    }
    if (method === 'POST' && path === '/office/workbooks/save') {
      return await handleSaveWorkbook(req.body)
    }
    if (method === 'POST' && path === '/office/documents/open') {
      return await handleOpenDocument(req.body)
    }
    if (method === 'POST' && path === '/office/documents/save') {
      return await handleSaveDocument(req.body)
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
