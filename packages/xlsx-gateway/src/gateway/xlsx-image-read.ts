/// EXCEL-022 canonical image reader: resolves a worksheet's drawing
/// relationship chain (worksheet → drawing part → image relationships →
/// xl/media/*) into typed SheetImageInfo entries carrying the inline media
/// bytes as a data URL, so the browser renders and edits pictures without
/// any OOXML knowledge.
///
/// Fail-closed semantics mirror the other per-sheet read families
/// (filters / validations / notes / tables): unreadable drawing wiring
/// surfaces NO images for that sheet — the workbook still opens and a
/// no-op save preserves the file's parts byte-for-byte. Individual
/// pictures with unsupported media types, oversized bytes, or missing
/// parts are skipped; their anchors still count toward drawingIndex
/// parity with the desktop sidecar (the index counts EVERY anchor
/// element in document order, exactly like visuals.rs and
/// xlsx-drawing-edit.ts).
///
/// ARCHITECT REVIEW (PR #20, blocker 2): absolute-anchored pictures are
/// OMITTED from the browser model. Their fixed-sheet geometry (xdr:pos +
/// xdr:ext, EMU from the sheet origin) cannot be represented in the
/// two-cell wire model, and approximating it with a zero marker would
/// silently relocate the picture. Fail closed: the picture never
/// surfaces, stays untouched in the file, and a no-op save preserves
/// its drawing XML and media byte-for-byte. The anchor still counts
/// toward drawingIndex parity (omitted anchors keep later locators
/// stable).

import type { DrawingAnchor } from './xlsx-drawing-add'

export class ImageReadError extends Error {}

/// The writer-supported media set (xlsx-drawing-add IMAGE_EXTENSIONS) —
/// reads accept exactly the same extensions so a re-save of an imported
/// image never changes its part type.
const SUPPORTED_MEDIA_TYPES: Record<string, 'image/png' | 'image/jpeg' | 'image/gif'> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
}

const DRAWING_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/// Anchors count in document order — identical pattern to
/// xlsx-drawing-edit.ts, so a visualEdit's drawingIndex from this reader
/// addresses the exact same anchor the editor patches.
const ANCHOR_PATTERN = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g

/// One picture on one sheet — the typed wire model the browser consumes.
/// `drawingPath` + `drawingIndex` are the canonical edit locator (the same
/// pair the desktop sidecar emits); the browser echoes them back in
/// visualEdits. `anchor` carries raw OOXML marker geometry (EMU offsets);
/// two-cell images get their pixel size from the live grid, one-cell
/// images carry the explicit `widthPx`/`heightPx` derived from a:ext
/// (documented conversion: 1 px = 9525 EMU at 96 dpi). Absolute-anchored
/// pictures are never surfaced (see the module docs) — the union has no
/// 'absolute' member by design.
export interface SheetImageInfo {
  readonly drawingPath: string
  readonly drawingIndex: number
  readonly anchorType: 'two-cell' | 'one-cell'
  readonly anchor: DrawingAnchor
  /// One-cell pictures: explicit a:ext size in px (rounded).
  readonly widthPx?: number | undefined
  readonly heightPx?: number | undefined
  /// a:xfrm/@rot in degrees clockwise, when present.
  readonly rotationDeg?: number | undefined
  /// xdr:cNvPr/@name — display label / alt text echo.
  readonly name?: string | undefined
  readonly mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly dataUrl: string
}

/// Entry sources that can hand out raw bytes (the in-memory buffer
/// source always can; the platform archive adapter never calls this
/// reader, so the optional member keeps its surface unchanged).
export interface BinaryEntrySource {
  readText(path: string): Promise<string>
  has(path: string): Promise<boolean>
  readBinary?(path: string): Promise<Uint8Array>
}

/// Per-image cap: a picture whose base64 exceeds this is skipped (its
/// bytes are unrepresentable on the wire). 2M chars ≈ 1.5 MB binary.
const MAX_IMAGE_BASE64_CHARS = 2_000_000
/// Per-sheet cap: the whole inline media set must stay under this or the
/// sheet surfaces NO images (fail closed, like every other read family).
const MAX_SHEET_IMAGE_BASE64_CHARS = 4_000_000

const EMU_PER_PX = 9525

export async function parseSheetImages(
  reader: BinaryEntrySource,
  worksheetPath: string,
  worksheetXml: string,
): Promise<readonly SheetImageInfo[]> {
  const drawingRelId = /<drawing\b[^>]*\br:id="([^"]+)"/.exec(worksheetXml)?.[1]
  if (drawingRelId === undefined) return []
  const drawingPath = await resolveDrawingPath(reader, worksheetPath, drawingRelId)
  if (drawingPath === null) {
    throw new ImageReadError(
      `The worksheet's drawing relationship "${drawingRelId}" could not be resolved.`,
    )
  }
  if (!(await reader.has(drawingPath))) {
    throw new ImageReadError(`Workbook is missing ${drawingPath}.`)
  }
  const drawingXml = await reader.readText(drawingPath)
  const imageRels = await parseImageRelationships(reader, drawingPath)
  const images: SheetImageInfo[] = []
  let totalBase64Chars = 0
  let anchorIndex = 0
  for (const match of drawingXml.matchAll(ANCHOR_PATTERN)) {
    const index = anchorIndex
    anchorIndex += 1
    const anchorXml = match[0]
    const kind = match[1]
    if (kind === 'absoluteAnchor') {
      // Fail closed (architect review, PR #20 blocker 2): an absolute
      // picture's fixed-sheet geometry has no two-cell representation —
      // omit it rather than relocate it to a zero marker. The index
      // already counted above, so later anchors keep their locators.
      continue
    }
    if (!anchorXml.includes('<xdr:pic')) continue
    const parsed = parsePictureAnchor(anchorXml, kind as 'twoCellAnchor' | 'oneCellAnchor')
    if (parsed === null) {
      // An <xdr:pic> without a readable r:embed or markers — skip the
      // picture, but the anchor index still counted above.
      continue
    }
    const rel = imageRels.get(parsed.embedId)
    if (rel === undefined) continue
    const mediaType = SUPPORTED_MEDIA_TYPES[extensionOf(rel.mediaPath)] ?? null
    if (mediaType === null) continue
    if (reader.readBinary === undefined) {
      throw new ImageReadError('This package source cannot read binary media parts.')
    }
    if (!(await reader.has(rel.mediaPath))) continue
    let bytes: Uint8Array
    try {
      bytes = await reader.readBinary(rel.mediaPath)
    } catch {
      continue
    }
    const base64 = Buffer.from(bytes).toString('base64')
    if (base64.length > MAX_IMAGE_BASE64_CHARS) continue
    totalBase64Chars += base64.length
    if (totalBase64Chars > MAX_SHEET_IMAGE_BASE64_CHARS) {
      throw new ImageReadError(
        'The sheet inline media exceeds the supported size — images are not available.',
      )
    }
    images.push({
      drawingPath,
      drawingIndex: index,
      anchorType: parsed.anchorType,
      anchor: parsed.anchor,
      ...(parsed.widthPx !== undefined ? { widthPx: parsed.widthPx } : {}),
      ...(parsed.heightPx !== undefined ? { heightPx: parsed.heightPx } : {}),
      ...(parsed.rotationDeg !== undefined ? { rotationDeg: parsed.rotationDeg } : {}),
      ...(parsed.name !== undefined ? { name: parsed.name } : {}),
      mediaType,
      dataUrl: `data:${mediaType};base64,${base64}`,
    })
  }
  return images
}

/// worksheet rels → drawing part path (the same two-step lookup the
/// comments and table readers use).
async function resolveDrawingPath(
  reader: BinaryEntrySource,
  worksheetPath: string,
  relId: string,
): Promise<string | null> {
  const relsPath = worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
  if (!(await reader.has(relsPath))) return null
  const relsXml = await reader.readText(relsPath)
  const relationship = new RegExp(
    `<Relationship\\b[^>]*\\bId="${escapeRegExp(relId)}"[^>]*/?>`,
  ).exec(relsXml)?.[0]
  if (relationship === undefined) return null
  if (!relationship.includes(`Type="${DRAWING_REL_TYPE}"`)) return null
  const target = /\bTarget="([^"]+)"/.exec(relationship)?.[1]
  if (target === undefined) return null
  if (target.startsWith('/')) return target.slice(1)
  return resolveRelativePart(worksheetPath.split('/').slice(0, -1), target)
}

/// drawing rels → map of image relationship id → resolved media path.
async function parseImageRelationships(
  reader: BinaryEntrySource,
  drawingPath: string,
): Promise<Map<string, { mediaPath: string }>> {
  const relsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
  const map = new Map<string, { mediaPath: string }>()
  if (!(await reader.has(relsPath))) return map
  const relsXml = await reader.readText(relsPath)
  const base = drawingPath.split('/').slice(0, -1)
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const element = match[0]
    if (!element.includes(`Type="${IMAGE_REL_TYPE}"`)) continue
    const id = /\bId="([^"]+)"/.exec(element)?.[1]
    const target = /\bTarget="([^"]+)"/.exec(element)?.[1]
    if (id === undefined || target === undefined) continue
    if (/\bTargetMode="External"/.test(element)) continue
    map.set(id, { mediaPath: resolveRelativePart(base, target) })
  }
  return map
}

interface ParsedPictureAnchor {
  readonly anchorType: 'two-cell' | 'one-cell'
  readonly anchor: DrawingAnchor
  readonly widthPx?: number | undefined
  readonly heightPx?: number | undefined
  readonly rotationDeg?: number | undefined
  readonly name?: string | undefined
  readonly embedId: string
}

function parsePictureAnchor(
  anchorXml: string,
  kind: 'twoCellAnchor' | 'oneCellAnchor',
): ParsedPictureAnchor | null {
  const embedId = /<a:blip\b[^>]*\br:embed="([^"]+)"/.exec(anchorXml)?.[1]
  if (embedId === undefined) return null
  const name = /<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(anchorXml)?.[1]
  const rot = /\brot="(-?\d+)"/.exec(anchorXml)?.[1]
  const extMatch = /<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(anchorXml)
  const widthPx = extMatch !== null ? Math.round(Number(extMatch[1]) / EMU_PER_PX) : undefined
  const heightPx = extMatch !== null ? Math.round(Number(extMatch[2]) / EMU_PER_PX) : undefined
  const from = parseMarker(anchorXml, 'from')
  if (from === null) return null
  if (kind === 'oneCellAnchor') {
    // One-cell pictures move with their from marker but keep the a:ext
    // size; resize is not representable by the canonical edit family
    // (only the from marker is rewritten), so the browser treats them as
    // read-only.
    return {
      anchorType: 'one-cell',
      anchor: {
        ...from,
        toRow: from.fromRow,
        toColumn: from.fromColumn,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      ...(widthPx !== undefined ? { widthPx } : {}),
      ...(heightPx !== undefined ? { heightPx } : {}),
      ...(rot !== undefined ? { rotationDeg: Number(rot) / 60000 } : {}),
      ...(name !== undefined && name !== '' ? { name } : {}),
      embedId,
    }
  }
  const to = parseMarker(anchorXml, 'to')
  if (to === null) return null
  return {
    anchorType: 'two-cell',
    anchor: {
      fromRow: from.fromRow,
      fromColumn: from.fromColumn,
      fromRowOffset: from.fromRowOffset,
      fromColumnOffset: from.fromColumnOffset,
      toRow: to.fromRow,
      toColumn: to.fromColumn,
      toRowOffset: to.fromRowOffset,
      toColumnOffset: to.fromColumnOffset,
    },
    ...(rot !== undefined ? { rotationDeg: Number(rot) / 60000 } : {}),
    ...(name !== undefined && name !== '' ? { name } : {}),
    embedId,
  }
}

/// Parses an <xdr:from> or <xdr:to> marker (col, colOff, row, rowOff).
function parseMarker(
  anchorXml: string,
  tag: 'from' | 'to',
): Omit<DrawingAnchor, 'toRow' | 'toColumn' | 'toRowOffset' | 'toColumnOffset'> | null {
  const section = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(anchorXml)
  if (section === null) return null
  const body = section[1] ?? ''
  const col = Number(/<xdr:col>(\d+)<\/xdr:col>/.exec(body)?.[1])
  const colOff = Number(/<xdr:colOff>(\d+)<\/xdr:colOff>/.exec(body)?.[1])
  const row = Number(/<xdr:row>(\d+)<\/xdr:row>/.exec(body)?.[1])
  const rowOff = Number(/<xdr:rowOff>(\d+)<\/xdr:rowOff>/.exec(body)?.[1])
  if (![col, colOff, row, rowOff].every((value) => Number.isInteger(value) && value >= 0)) {
    return null
  }
  return { fromRow: row, fromColumn: col, fromRowOffset: rowOff, fromColumnOffset: colOff }
}

function extensionOf(path: string): string {
  const dot = path.lastIndexOf('.')
  return dot < 0 ? '' : path.slice(dot + 1).toLowerCase()
}

function resolveRelativePart(base: readonly string[], target: string): string {
  const segments = [...base]
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.') segments.push(part)
  }
  return segments.join('/')
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
