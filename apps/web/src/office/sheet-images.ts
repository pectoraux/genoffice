/**
 * EXCEL-022 — worksheet image state mapping between the canonical wire
 * model (SheetImageInfo from /api/office/workbooks/open) and Univer's
 * over-grid image model (@univerjs/preset-sheets-drawing).
 *
 * PURITY: this module performs NO OOXML work whatsoever. All drawing XML
 * (anchors, pictures, relationships, media parts, content types) is
 * produced exclusively by the canonical xlsx-gateway
 * (applyVisualAdditions / applyVisualEdits) on the server. The browser
 * only maps typed fields ↔ Univer image state.
 *
 * Geometry contract (documented EMU ↔ pixel conversion): OOXML anchors
 * carry EMU offsets; Univer's over-grid positions carry pixel offsets.
 * 1 CSS px = 9525 EMU (96 dpi) — the same constant the desktop renderer
 * and the Word image pipeline use. The browser is NOT authoritative for
 * OOXML geometry: it renders the typed anchor and derives edit anchors
 * from Univer's live model, and the gateway serializes them verbatim.
 */

import { ImageSourceType } from '@univerjs/core'
import type { SheetImageInfo, WorkbookVisualEdit } from '@genoffice/xlsx-gateway'
import type { DrawingAnchor, SheetVisualAddition } from '@genoffice/xlsx-gateway'
import type { ISheetImage } from '@univerjs/sheets-drawing'

/** 1 CSS px = 9525 EMU at 96 dpi (ECMA-376 default). */
export const EMU_PER_PX = 9525

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX)
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX
}

/** Deterministic Univer drawing id for a file-native image locator. */
export function fileImageId(drawingPath: string, drawingIndex: number): string {
  return `file-img:${drawingPath}#${drawingIndex}`
}

/** Grid geometry read surface (FWorksheet satisfies this). */
export interface GridGeometry {
  getColumnWidth(column: number): number
  getRowHeight(row: number): number
}

/** A from-position + explicit size, in pixels (Univer's native units). */
export interface OverGridPlacement {
  readonly column: number
  readonly columnOffsetPx: number
  readonly row: number
  readonly rowOffsetPx: number
  readonly widthPx: number
  readonly heightPx: number
}

/**
 * Walks the grid forward from (index, offsetPx) consuming `sizePx` and
 * returns the end marker's index + offset. Guards against zero-width
 * columns with a hard iteration cap.
 */
function walkForward(
  geometry: GridGeometry,
  startIndex: number,
  startOffsetPx: number,
  sizePx: number,
  sizeOf: (index: number) => number,
): { index: number; offsetPx: number } {
  let index = startIndex
  let offsetPx = startOffsetPx + sizePx
  let guard = 0
  while (offsetPx > 0 && guard < 5_000) {
    const size = Math.max(1, sizeOf(index))
    if (offsetPx < size) return { index, offsetPx }
    offsetPx -= size
    index += 1
    guard += 1
  }
  return { index, offsetPx: Math.max(0, offsetPx) }
}

/**
 * Sums the grid span between two markers on one axis (start inclusive of
 * its offset, end exclusive). Same-axis markers return the offset delta.
 */
function spanBetween(
  geometry: GridGeometry,
  from: { index: number; offsetPx: number },
  to: { index: number; offsetPx: number },
  sizeOf: (index: number) => number,
): number {
  if (to.index <= from.index) {
    return Math.max(1, to.offsetPx - from.offsetPx)
  }
  let span = sizeOf(from.index) - from.offsetPx
  for (let index = from.index + 1; index < to.index; index += 1) {
    span += sizeOf(index)
  }
  span += to.offsetPx
  return Math.max(1, span)
}

/**
 * Converts a canonical two-cell anchor into the Univer placement the
 * editor installs: the from marker becomes the position (px offsets), and
 * the width/height are derived by summing the live grid between the from
 * and to markers — deterministic within a session, and exactly how the
 * desktop renders the same anchor.
 */
export function placementFromAnchor(
  geometry: GridGeometry,
  anchor: DrawingAnchor,
): OverGridPlacement {
  const columnOffsetPx = emuToPx(anchor.fromColumnOffset)
  const rowOffsetPx = emuToPx(anchor.fromRowOffset)
  const widthPx = spanBetween(
    geometry,
    { index: anchor.fromColumn, offsetPx: columnOffsetPx },
    { index: anchor.toColumn, offsetPx: emuToPx(anchor.toColumnOffset) },
    (index) => geometry.getColumnWidth(index),
  )
  const heightPx = spanBetween(
    geometry,
    { index: anchor.fromRow, offsetPx: rowOffsetPx },
    { index: anchor.toRow, offsetPx: emuToPx(anchor.toRowOffset) },
    (index) => geometry.getRowHeight(index),
  )
  return {
    column: anchor.fromColumn,
    columnOffsetPx,
    row: anchor.fromRow,
    rowOffsetPx,
    widthPx,
    heightPx,
  }
}

/**
 * Converts a live Univer placement into the canonical two-cell anchor —
 * the to marker is derived by walking the live grid forward. This is the
 * ONLY place edit anchors are produced, and it emits plain typed data
 * the gateway serializes.
 */
export function anchorFromPlacement(
  geometry: GridGeometry,
  placement: OverGridPlacement,
): DrawingAnchor {
  const toColumn = walkForward(
    geometry,
    placement.column,
    placement.columnOffsetPx,
    placement.widthPx,
    (index) => geometry.getColumnWidth(index),
  )
  const toRow = walkForward(
    geometry,
    placement.row,
    placement.rowOffsetPx,
    placement.heightPx,
    (index) => geometry.getRowHeight(index),
  )
  return {
    fromRow: placement.row,
    fromColumn: placement.column,
    fromRowOffset: pxToEmu(placement.rowOffsetPx),
    fromColumnOffset: pxToEmu(placement.columnOffsetPx),
    toRow: toRow.index,
    toColumn: toColumn.index,
    toRowOffset: pxToEmu(toRow.offsetPx),
    toColumnOffset: pxToEmu(toColumn.offsetPx),
  }
}

/** Minimal structural view of a live over-grid image (FOverGridImage). */
export interface OverGridImageFacade {
  getId(): string
  /** Public removal (the facade's documented remove()). */
  remove(): boolean
  /** Public geometry read surface — see BuiltImageGeometry. */
  toBuilder(): { buildAsync(): Promise<BuiltImageGeometry> }
}

/**
 * EXPLICIT ADAPTER (architect review, PR #20 blocker 1) over the PUBLIC
 * builder return: `FOverGridImage.toBuilder().buildAsync()` hands back the
 * live image data — the same public surface the facade's own
 * setPositionAsync / setSizeAsync build their commands on. The narrow
 * typed view below reads only what the public type surface guarantees:
 * the maintained two-cell transform (pixel offsets). No private Univer
 * internals (`_image`, casts) are touched anywhere in this module.
 */
export interface BuiltImageGeometry {
  readonly sheetTransform?: {
    readonly from?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number }
    readonly to?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number }
  }
}

/** Minimal structural view of a Univer facade worksheet. */
export interface ImageWorksheetFacade {
  getSheetId(): string
  getSheetName(): string
  getColumnWidth(column: number): number
  getRowHeight(row: number): number
  newOverGridImage(): OverGridImageBuilderFacade
  insertImages(images: unknown[]): unknown
  getImages(): OverGridImageFacade[]
}

/**
 * The install payload for the builder's PUBLIC setImage(): the identity
 * and source fields the install path provides. `Partial<ISheetImage>`
 * because Univer's own documented example omits the derived transform
 * fields (the builder computes them); the `Pick` set is exactly what the
 * install path must supply.
 */
export type OverGridImageParam = Partial<ISheetImage> &
  Pick<
    ISheetImage,
    'drawingId' | 'drawingType' | 'imageSourceType' | 'source' | 'unitId' | 'subUnitId'
  >

/** The builder surface the install path uses. */
export interface OverGridImageBuilderFacade {
  setImage(image: OverGridImageParam): unknown
  setColumn(column: number): unknown
  setRow(row: number): unknown
  setColumnOffset(offset: number): unknown
  setRowOffset(offset: number): unknown
  setWidth(width: number): unknown
  setHeight(height: number): unknown
  setAnchorType(anchorType: unknown): unknown
  buildAsync(): Promise<unknown>
}

/** Univer SheetDrawingAnchorType.Both — two-cell semantics. */
const TWO_CELL_ANCHOR_TYPE = '1'

/**
 * Installs one image into the live Univer sheet. The drawing id is the
 * caller's stable identity (locator-based for file images, sequence-based
 * for session adds), so later mutation events map straight back to the
 * canonical model. Anchors are two-cell (`Both`), matching the OOXML the
 * gateway writes.
 */
export async function insertOverGridImage(
  worksheet: ImageWorksheetFacade,
  options: {
    readonly id: string
    readonly unitId: string
    readonly dataUrl: string
    readonly placement: OverGridPlacement
  },
): Promise<void> {
  const builder = worksheet.newOverGridImage()
  builder.setImage({
    drawingId: options.id,
    // DrawingTypeEnum.DRAWING_IMAGE — 0.
    drawingType: 0,
    // ImageSourceType.BASE64 — the data-URL source form.
    imageSourceType: ImageSourceType.BASE64,
    source: options.dataUrl,
    unitId: options.unitId,
    subUnitId: worksheet.getSheetId(),
  })
  builder.setColumn(options.placement.column)
  builder.setRow(options.placement.row)
  builder.setColumnOffset(Math.round(options.placement.columnOffsetPx))
  builder.setRowOffset(Math.round(options.placement.rowOffsetPx))
  builder.setWidth(Math.round(options.placement.widthPx))
  builder.setHeight(Math.round(options.placement.heightPx))
  builder.setAnchorType(TWO_CELL_ANCHOR_TYPE)
  const image = await builder.buildAsync()
  worksheet.insertImages([image])
}

/**
 * Reads a live image's placement through the PUBLIC facade surface:
 * `toBuilder().buildAsync()` returns the live image data, whose
 * sheetTransform carries the maintained from/to markers (pixel
 * offsets). The width/height derive from the live grid between the
 * markers — deterministic and exactly how the install path sized the
 * image. Returns null when the image is no longer on the sheet or its
 * geometry is not fully readable (nothing journals — fail closed).
 */
export async function readLivePlacement(
  worksheet: ImageWorksheetFacade,
  id: string,
): Promise<OverGridPlacement | null> {
  const image = worksheet.getImages().find((entry) => entry.getId() === id)
  if (image === undefined) return null
  const data = await image.toBuilder().buildAsync()
  const from = data.sheetTransform?.from
  const to = data.sheetTransform?.to
  const column = from?.column
  const columnOffsetPx = from?.columnOffset
  const row = from?.row
  const rowOffsetPx = from?.rowOffset
  if (
    column === undefined ||
    columnOffsetPx === undefined ||
    row === undefined ||
    rowOffsetPx === undefined
  ) {
    return null
  }
  let widthPx: number | undefined
  let heightPx: number | undefined
  if (to !== undefined && to.column !== undefined && to.row !== undefined) {
    const toColumnOffsetPx = to.columnOffset ?? 0
    const toRowOffsetPx = to.rowOffset ?? 0
    if (to.column > column) {
      let span = -columnOffsetPx
      for (let index = column; index < to.column; index += 1) {
        span += worksheet.getColumnWidth(index)
      }
      widthPx = span + toColumnOffsetPx
    } else {
      widthPx = toColumnOffsetPx - columnOffsetPx
    }
    if (to.row > row) {
      let span = -rowOffsetPx
      for (let index = row; index < to.row; index += 1) {
        span += worksheet.getRowHeight(index)
      }
      heightPx = span + toRowOffsetPx
    } else {
      heightPx = toRowOffsetPx - rowOffsetPx
    }
  }
  if (widthPx === undefined || heightPx === undefined) return null
  return {
    column,
    columnOffsetPx,
    row,
    rowOffsetPx,
    widthPx: Math.max(1, widthPx),
    heightPx: Math.max(1, heightPx),
  }
}

/**
 * Reads the canonical edit anchor for a live image, or null when the
 * image is gone. THE save-side adapter: plain typed data out.
 */
export async function readLiveAnchor(
  worksheet: ImageWorksheetFacade,
  id: string,
): Promise<DrawingAnchor | null> {
  const placement = await readLivePlacement(worksheet, id)
  if (placement === null) return null
  return anchorFromPlacement(worksheet, placement)
}

/** Natural dimensions of a data URL image (fallback matches the desktop picker). */
export function measureImage(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const image = new Image()
    image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight })
    image.onerror = () => resolve({ width: 480, height: 320 })
    image.src = dataUrl
  })
}

/** Normalizes a File.type to the writer-supported media set. */
export function supportedMediaType(
  fileType: string,
): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  const normalized = fileType === 'image/jpg' ? 'image/jpeg' : fileType
  if (normalized === 'image/png' || normalized === 'image/jpeg' || normalized === 'image/gif') {
    return normalized
  }
  return null
}

/** A file-native image's tracked browser state. */
export interface FileImageEntry {
  readonly sheetName: string
  readonly info: SheetImageInfo
}

/** A session-created image's journal entry (pre-persist). */
export interface SessionImageAdd {
  readonly id: string
  readonly sheetName: string
  readonly anchor: DrawingAnchor
  readonly mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly base64: string
}

/** Locator the gateway returns for a persisted visual addition. */
export interface AddedVisualLocator {
  readonly worksheetPath?: string
  readonly drawingPath: string
  readonly drawingIndex: number
}

/**
 * Builds the wire visualEdits for the journaled image state: removals
 * first (any order — the gateway sorts), then anchor edits for moved or
 * resized file images. Anchors are read from the LIVE Univer model so
 * multiple consecutive interactions collapse into one final-state edit.
 */
export async function collectImageVisualEdits(
  worksheetByName: (sheetName: string) => ImageWorksheetFacade | null,
  fileImages: ReadonlyMap<string, FileImageEntry>,
  dirty: ReadonlySet<string>,
  removals: ReadonlySet<string>,
): Promise<WorkbookVisualEdit[]> {
  const edits: WorkbookVisualEdit[] = []
  for (const id of removals) {
    const entry = fileImages.get(id)
    if (entry === undefined) continue
    edits.push({
      drawingPath: entry.info.drawingPath,
      drawingIndex: entry.info.drawingIndex,
      remove: true,
    })
  }
  for (const id of dirty) {
    if (removals.has(id)) continue
    const entry = fileImages.get(id)
    if (entry === undefined) continue
    const worksheet = worksheetByName(entry.sheetName)
    if (worksheet === null) continue
    const anchor = await readLiveAnchor(worksheet, id)
    if (anchor === null) continue
    edits.push({
      drawingPath: entry.info.drawingPath,
      drawingIndex: entry.info.drawingIndex,
      anchor,
    })
  }
  return edits
}

/**
 * Builds the wire visualAdditions for session images still on the sheet
 * (deleted session images drop out — never persisted). Anchors refresh
 * from the LIVE model so post-insert moves persist at their final
 * position.
 */
export async function collectImageVisualAdditions(
  worksheetByName: (sheetName: string) => ImageWorksheetFacade | null,
  adds: readonly SessionImageAdd[],
  removals: ReadonlySet<string>,
): Promise<SheetVisualAddition[]> {
  const additions: SheetVisualAddition[] = []
  for (const add of adds) {
    if (removals.has(add.id)) continue
    const worksheet = worksheetByName(add.sheetName)
    const anchor = worksheet === null ? null : await readLiveAnchor(worksheet, add.id)
    additions.push({
      sheetName: add.sheetName,
      anchor: anchor ?? add.anchor,
      image: { mediaType: add.mediaType, base64: add.base64 },
    })
  }
  return additions
}
