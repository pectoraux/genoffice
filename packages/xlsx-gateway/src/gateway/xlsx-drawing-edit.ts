/// Surgical edits to visuals that already live in the file: remove an
/// anchor, or move/resize one by rewriting its from/to markers. Visuals are
/// located by the sidecar's (drawingPath, drawingIndex) pair — the index
/// counts every anchor element in document order, matching visuals.rs.

import type { WorkbookVisualEdit } from '../types.js'
import type { MutablePackage } from './xlsx-drawing-add'

export class VisualEditError extends Error {}

const ANCHOR_PATTERN = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g

export async function applyVisualEdits(
  pkg: MutablePackage,
  edits: readonly WorkbookVisualEdit[],
  touchedEntries: Set<string>,
): Promise<void> {
  const byPath = new Map<string, WorkbookVisualEdit[]>()
  for (const edit of edits) {
    const group = byPath.get(edit.drawingPath) ?? []
    group.push(edit)
    byPath.set(edit.drawingPath, group)
  }
  for (const [drawingPath, group] of byPath) {
    if (!(await pkg.has(drawingPath))) {
      throw new VisualEditError(`Workbook is missing ${drawingPath}.`)
    }
    let xml = await pkg.readText(drawingPath)
    // Removals splice text out, so process high indexes first — lower
    // indexes keep their document positions.
    const ordered = [...group].sort((left, right) => right.drawingIndex - left.drawingIndex)
    const seen = new Set<number>()
    const removedChartRelIds: string[] = []
    const removedImageRelIds: string[] = []
    for (const edit of ordered) {
      if (seen.has(edit.drawingIndex)) {
        throw new VisualEditError('Duplicate edits target the same drawing anchor.')
      }
      seen.add(edit.drawingIndex)
      xml = applyOneEdit(xml, edit, removedChartRelIds, removedImageRelIds)
    }
    pkg.write(drawingPath, xml)
    touchedEntries.add(drawingPath)
    if (removedChartRelIds.length > 0) {
      await cascadeChartRemovals(pkg, drawingPath, xml, removedChartRelIds, touchedEntries)
    }
    if (removedImageRelIds.length > 0) {
      await cascadeImageRemovals(pkg, drawingPath, xml, removedImageRelIds, touchedEntries)
    }
  }
}

/// A deleted picture anchor cascades differently from a chart: the image
/// relationship is dropped from the drawing rels, and the media part is
/// removed ONLY when no remaining relationship ANYWHERE in the package
/// still references it (two pictures can share one xl/media entry, and a
/// media part can be referenced from another drawing's rels). The
/// [Content_Types].xml Default for the extension goes only when no other
/// part with that extension remains. Never remove a media part merely
/// because one picture was deleted.
async function cascadeImageRemovals(
  pkg: MutablePackage,
  drawingPath: string,
  remainingDrawingXml: string,
  relIds: readonly string[],
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
  if (!(await pkg.has(relsPath))) {
    throw new VisualEditError('The drawing is missing its relationships part.')
  }
  let relsXml = await pkg.readText(relsPath)
  const removedMediaPaths: string[] = []
  for (const relId of relIds) {
    // Another anchor still embeds this relationship (shared media) — the
    // relationship and the media part both stay.
    if (remainingDrawingXml.includes(`r:embed="${relId}"`)) continue
    const relMatch = new RegExp(`<Relationship\\b[^>]*\\bId="${escapeRegExp(relId)}"[^>]*/>`).exec(
      relsXml,
    )
    if (!relMatch) {
      // Already removed by an earlier edit in this save (two deleted
      // pictures shared one relationship) — nothing left to do.
      continue
    }
    const target = /\bTarget="([^"]+)"/.exec(relMatch[0])?.[1]
    const type = /\bType="([^"]+)"/.exec(relMatch[0])?.[1] ?? ''
    if (!target || !/\/image$/.test(type)) {
      throw new VisualEditError('The deleted picture is not a plain image — not supported.')
    }
    const mediaPath = resolveRelativePart(drawingPath.replace(/\/[^/]+$/, ''), target)
    relsXml = relsXml.replace(relMatch[0], '')
    if (!removedMediaPaths.includes(mediaPath)) removedMediaPaths.push(mediaPath)
  }
  pkg.write(relsPath, relsXml)
  touchedEntries.add(relsPath)
  for (const mediaPath of removedMediaPaths) {
    if (!(await pkg.has(mediaPath))) continue
    if (await anyRelationshipReferences(pkg, mediaPath)) continue
    pkg.remove(mediaPath)
    await removeContentTypeDefaultIfUnused(pkg, mediaPath, touchedEntries)
  }
}

/// Scans every *.rels part in the package for a relationship whose
/// resolved target is `mediaPath`. External targets are skipped.
async function anyRelationshipReferences(pkg: MutablePackage, mediaPath: string): Promise<boolean> {
  for (const path of await pkg.paths()) {
    if (!/_rels\/[^/]+\.rels$/.test(path)) continue
    const relsXml = await pkg.readText(path)
    const base = path.split('/').slice(0, -2)
    for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
      const element = match[0]
      if (/\bTargetMode="External"/.test(element)) continue
      const target = /\bTarget="([^"]+)"/.exec(element)?.[1]
      if (target === undefined) continue
      if (resolveRelativePart(base.join('/'), target) === mediaPath) return true
    }
  }
  return false
}

/// Drops the [Content_Types].xml Default for the removed part's extension
/// only when NO other package entry still carries that extension.
async function removeContentTypeDefaultIfUnused(
  pkg: MutablePackage,
  removedPath: string,
  touchedEntries: Set<string>,
): Promise<void> {
  const dot = removedPath.lastIndexOf('.')
  if (dot < 0) return
  const extension = removedPath.slice(dot + 1).toLowerCase()
  const paths = await pkg.paths()
  const stillUsed = paths.some(
    (path) =>
      path !== removedPath && path.slice(path.lastIndexOf('.') + 1).toLowerCase() === extension,
  )
  if (stillUsed) return
  const contentTypes = await pkg.readText('[Content_Types].xml')
  const stripped = contentTypes.replace(
    new RegExp(`<Default\\b[^>]*\\bExtension="${escapeRegExp(extension)}"[^>]*/>`),
    '',
  )
  if (stripped !== contentTypes) {
    pkg.write('[Content_Types].xml', stripped)
    touchedEntries.add('[Content_Types].xml')
  }
}

/// A deleted chart anchor cascades: the drawing relationship, the chart part,
/// its own rels, and its content-type override all go. Chart-owned colors/
/// style parts stay behind as unreferenced-but-present entries (harmless).
async function cascadeChartRemovals(
  pkg: MutablePackage,
  drawingPath: string,
  remainingDrawingXml: string,
  relIds: readonly string[],
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
  if (!(await pkg.has(relsPath))) {
    throw new VisualEditError('The drawing is missing its relationships part.')
  }
  let relsXml = await pkg.readText(relsPath)
  for (const relId of relIds) {
    if (remainingDrawingXml.includes(`r:id="${relId}"`)) {
      throw new VisualEditError('Another anchor still references the deleted chart.')
    }
    const relMatch = new RegExp(`<Relationship\\b[^>]*\\bId="${escapeRegExp(relId)}"[^>]*/>`).exec(
      relsXml,
    )
    if (!relMatch) {
      throw new VisualEditError('The deleted chart has no drawing relationship.')
    }
    const target = /\bTarget="([^"]+)"/.exec(relMatch[0])?.[1]
    const type = /\bType="([^"]+)"/.exec(relMatch[0])?.[1] ?? ''
    if (!target || !/\/chart$/.test(type)) {
      throw new VisualEditError('The deleted graphic frame is not a plain chart — not supported.')
    }
    const chartPath = resolveRelativePart(drawingPath.replace(/\/[^/]+$/, ''), target)
    relsXml = relsXml.replace(relMatch[0], '')
    pkg.remove(chartPath)
    const chartRelsPath = chartPath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
    if (await pkg.has(chartRelsPath)) pkg.remove(chartRelsPath)
    const contentTypes = await pkg.readText('[Content_Types].xml')
    const stripped = contentTypes.replace(
      new RegExp(`<Override\\b[^>]*PartName="/${escapeRegExp(chartPath)}"[^>]*/>`),
      '',
    )
    if (stripped !== contentTypes) {
      pkg.write('[Content_Types].xml', stripped)
      touchedEntries.add('[Content_Types].xml')
    }
  }
  pkg.write(relsPath, relsXml)
  touchedEntries.add(relsPath)
}

function resolveRelativePart(baseDir: string, target: string): string {
  const segments = baseDir.split('/')
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.') segments.push(part)
  }
  return segments.join('/')
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function applyOneEdit(
  xml: string,
  edit: WorkbookVisualEdit,
  removedChartRelIds: string[],
  removedImageRelIds: string[],
): string {
  const anchors = [...xml.matchAll(ANCHOR_PATTERN)]
  const match = anchors[edit.drawingIndex]
  if (!match) {
    throw new VisualEditError(
      `Drawing anchor #${edit.drawingIndex} was not found — the file may have changed.`,
    )
  }
  const anchorXml = match[0]
  const kind = match[1]
  if (edit.remove) {
    // A chart's graphicFrame anchor cascades: its rel, part, and override
    // are collected here and removed after the drawing XML is final.
    if (anchorXml.includes('<xdr:graphicFrame')) {
      const relId = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(anchorXml)?.[1]
      if (!relId) {
        throw new VisualEditError(
          'This graphic frame is not a chart — deleting it is not supported.',
        )
      }
      removedChartRelIds.push(relId)
    }
    // A picture anchor collects its image relationship id — the cascade
    // after the drawing XML is final decides whether the media part goes
    // (only when nothing else references it anymore). A picture without
    // an r:embed references no media part, so its removal is a plain
    // splice (pre-existing behavior for synthetic/degenerate pics).
    if (anchorXml.includes('<xdr:pic')) {
      const relId = /<a:blip\b[^>]*\br:embed="([^"]+)"/.exec(anchorXml)?.[1]
      if (relId !== undefined) {
        removedImageRelIds.push(relId)
      }
    }
    return xml.slice(0, match.index) + xml.slice(match.index + anchorXml.length)
  }
  const anchor = edit.anchor
  if (!anchor) throw new VisualEditError('A visual edit needs a removal or a new anchor.')
  if (kind === 'absoluteAnchor') {
    throw new VisualEditError('This visual uses an absolute anchor — moving it is not supported.')
  }
  const from =
    `<xdr:from><xdr:col>${anchor.fromColumn}</xdr:col>` +
    `<xdr:colOff>${anchor.fromColumnOffset}</xdr:colOff>` +
    `<xdr:row>${anchor.fromRow}</xdr:row>` +
    `<xdr:rowOff>${anchor.fromRowOffset}</xdr:rowOff></xdr:from>`
  let patched = anchorXml.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, () => from)
  if (patched === anchorXml && !anchorXml.includes('<xdr:from>')) {
    throw new VisualEditError('Drawing anchor has no from marker — moving it is not supported.')
  }
  if (kind === 'twoCellAnchor') {
    const to =
      `<xdr:to><xdr:col>${anchor.toColumn}</xdr:col>` +
      `<xdr:colOff>${anchor.toColumnOffset}</xdr:colOff>` +
      `<xdr:row>${anchor.toRow}</xdr:row>` +
      `<xdr:rowOff>${anchor.toRowOffset}</xdr:rowOff></xdr:to>`
    // An unchanged edge replaces to an identical string, so presence must be
    // checked directly (an NW resize touches only the from marker).
    const withTo = patched.replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, () => to)
    if (withTo === patched && !patched.includes('<xdr:to>')) {
      throw new VisualEditError('Drawing anchor has no to marker — moving it is not supported.')
    }
    patched = withTo
  }
  return xml.slice(0, match.index) + patched + xml.slice(match.index + anchorXml.length)
}
