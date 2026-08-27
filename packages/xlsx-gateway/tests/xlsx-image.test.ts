import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assembleWithJsZip,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'
import type { WorkbookVisualEdit } from '../src/types.js'

/// ── Fixture kit: a workbook whose sheet carries a drawing part with
///    configurable picture anchors. PNG bytes are a deterministic 1×1
///    solid-color PNG (structure-valid, tiny).
const PNG_BYTES = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5c, 0xcd, 0xff, 0x69, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82,
])
const PNG_BASE64 = PNG_BYTES.toString('base64')

interface AnchorSpec {
  kind: 'twoCellAnchor' | 'oneCellAnchor' | 'absoluteAnchor'
  embedId?: string
  name?: string
  from?: { col: number; colOff: number; row: number; rowOff: number }
  to?: { col: number; colOff: number; row: number; rowOff: number }
  ext?: { cx: number; cy: number }
  rot?: number
  shape?: boolean
}

function marker(m: { col: number; colOff: number; row: number; rowOff: number }): string {
  return (
    `<xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff>` +
    `<xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff>`
  )
}

function anchorXml(spec: AnchorSpec): string {
  const from = spec.from ?? { col: 1, colOff: 0, row: 2, rowOff: 0 }
  const embed = spec.embedId ?? 'rId1'
  const name = spec.name ?? 'Picture'
  const rotAttr = spec.rot !== undefined ? ` rot="${spec.rot}"` : ''
  const ext = spec.ext ?? { cx: 190500, cy: 95250 }
  if (spec.shape) {
    return (
      `<xdr:${spec.kind}>` +
      (spec.kind === 'absoluteAnchor'
        ? `<xdr:pos x="0" y="0"/><xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>`
        : `<xdr:from>${marker(from)}</xdr:from>` +
          (spec.kind === 'twoCellAnchor'
            ? `<xdr:to>${marker(spec.to ?? { col: 6, colOff: 0, row: 12, rowOff: 0 })}</xdr:to>`
            : '') +
          (spec.kind === 'oneCellAnchor' ? `<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>` : '')) +
      '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="9" name="Shape"/></xdr:nvSpPr></xdr:sp>' +
      '<xdr:clientData/></xdr:' +
      spec.kind +
      '>'
    )
  }
  const pic =
    '<xdr:pic><xdr:nvPicPr>' +
    `<xdr:cNvPr id="2" name="${name}"/>` +
    '<xdr:cNvPicPr/></xdr:nvPicPr>' +
    `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${embed}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    `<xdr:spPr><a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
  return (
    `<xdr:${spec.kind}>` +
    (spec.kind === 'absoluteAnchor'
      ? `<xdr:pos x="47625" y="9525"/><xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>`
      : `<xdr:from>${marker(from)}</xdr:from>` +
        (spec.kind === 'twoCellAnchor'
          ? `<xdr:to>${marker(spec.to ?? { col: 6, colOff: 0, row: 12, rowOff: 0 })}</xdr:to>`
          : '') +
        (spec.kind === 'oneCellAnchor' ? `<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>` : '')) +
    pic +
    '<xdr:clientData/></xdr:' +
    spec.kind +
    '>'
  )
}

interface ImageFixtureOptions {
  anchors?: AnchorSpec[]
  rels?: string[]
  media?: Record<string, Buffer | string>
  secondSheetDrawing?: boolean
  sheetDrawingElement?: string
  missingDrawingRels?: boolean
}

async function buildImageFixture(options: ImageFixtureOptions = {}): Promise<Buffer> {
  const anchors = options.anchors ?? [
    { kind: 'twoCellAnchor' as const, embedId: 'rId1', name: 'Red dot' },
  ]
  const rels = options.rels ?? [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
  ]
  const media = options.media ?? { 'xl/media/image1.png': PNG_BYTES }
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  const sheetCount = options.secondSheetDrawing ? 2 : 1
  const sheetEntries = [0, 1]
    .slice(0, sheetCount)
    .map(
      (index) => `<sheet name="Sheet${index + 1}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
    )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetEntries.join('')}</sheets>
</workbook>`,
  )
  const workbookRels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
  ]
  if (options.secondSheetDrawing) {
    workbookRels.push(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>',
    )
  }
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${workbookRels.join('')}</Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData>
  ${options.sheetDrawingElement ?? '<drawing r:id="rIdDrawing1"/>'}
</worksheet>`,
  )
  if (options.secondSheetDrawing) {
    zip.file(
      'xl/worksheets/sheet2.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>20</v></c></row></sheetData>
  <drawing r:id="rIdDrawing1"/>
</worksheet>`,
    )
    zip.file(
      'xl/worksheets/_rels/sheet2.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/drawings/drawing2.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchorXml({ kind: 'twoCellAnchor', embedId: 'rIdShared' })}</xdr:wsDr>`,
    )
    zip.file(
      'xl/drawings/_rels/drawing2.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
    )
  }
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/drawings/drawing1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors
      .map((spec) => anchorXml(spec))
      .join('')}</xdr:wsDr>`,
  )
  if (!options.missingDrawingRels) {
    zip.file(
      'xl/drawings/_rels/drawing1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
    )
  }
  for (const [path, content] of Object.entries(media)) {
    zip.file(path, content)
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/// applyCellEditsToXlsx with every family empty except visualEdits.
async function saveWithVisualEdits(
  buffer: Buffer,
  visualEdits: readonly WorkbookVisualEdit[],
): Promise<XlsxMutation> {
  return applyCellEditsToXlsx(
    buffer,
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    null,
    [],
    [],
    visualEdits,
  )
}

async function readDrawingEntry(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  return zip.file('xl/drawings/drawing1.xml')?.async('string') ?? ''
}

async function readEntry(buffer: Buffer, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  return entry === null ? null : await entry.async('string')
}

async function readEntryBytes(buffer: Buffer, path: string): Promise<Buffer | null> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  return entry === null ? null : await entry.async('nodebuffer')
}

async function listEntries(buffer: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer)
  return Object.entries(zip.files)
    .filter(([, file]) => !file.dir)
    .map(([path]) => path)
}

/// ── Read path ────────────────────────────────────────────────────────

describe('image read: canonical parseSheetImages', () => {
  it('surfaces a two-cell picture with anchor, locator, and inline media', async () => {
    const buffer = await buildImageFixture()
    const imported = await readBasicWorkbook(buffer)
    const images = imported.snapshot.sheets[0]?.images
    expect(images).toBeDefined()
    expect(images).toHaveLength(1)
    const image = images![0]!
    expect(image.drawingPath).toBe('xl/drawings/drawing1.xml')
    expect(image.drawingIndex).toBe(0)
    expect(image.anchorType).toBe('two-cell')
    expect(image.anchor).toEqual({
      fromRow: 2,
      fromColumn: 1,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 12,
      toColumn: 6,
      toRowOffset: 0,
      toColumnOffset: 0,
    })
    expect(image.mediaType).toBe('image/png')
    expect(image.dataUrl).toBe(`data:image/png;base64,${PNG_BASE64}`)
    expect(image.name).toBe('Red dot')
    expect(image.widthPx).toBeUndefined()
  })

  it('counts non-picture anchors toward drawingIndex parity', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', shape: true },
        { kind: 'twoCellAnchor', embedId: 'rId1', name: 'First' },
        { kind: 'twoCellAnchor', embedId: 'rId2', name: 'Second' },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>',
      ],
      media: {
        'xl/media/image1.png': PNG_BYTES,
        'xl/media/image2.png': PNG_BYTES,
      },
    })
    const imported = await readBasicWorkbook(buffer)
    const images = imported.snapshot.sheets[0]?.images ?? []
    expect(images).toHaveLength(2)
    expect(images[0]?.drawingIndex).toBe(1)
    expect(images[1]?.drawingIndex).toBe(2)
  })

  it('models one-cell anchors read-only with explicit a:ext size', async () => {
    const buffer = await buildImageFixture({
      anchors: [{ kind: 'oneCellAnchor', embedId: 'rId1', ext: { cx: 190500, cy: 95250 } }],
    })
    const images = (await readBasicWorkbook(buffer)).snapshot.sheets[0]?.images ?? []
    expect(images).toHaveLength(1)
    expect(images[0]?.anchorType).toBe('one-cell')
    expect(images[0]?.widthPx).toBe(20)
    expect(images[0]?.heightPx).toBe(10)
  })

  it('omits absolute-anchored pictures (fail closed — never relocated)', async () => {
    const buffer = await buildImageFixture({
      anchors: [{ kind: 'absoluteAnchor', embedId: 'rId1', ext: { cx: 95250, cy: 95250 } }],
    })
    const images = (await readBasicWorkbook(buffer)).snapshot.sheets[0]?.images ?? []
    // Architect review (PR #20, blocker 2): absolute geometry has no
    // two-cell representation — the picture must NOT surface (a zero
    // marker would silently relocate it). It stays untouched in the file.
    expect(images).toHaveLength(0)
  })

  it('keeps drawingIndex parity across an omitted absolute anchor', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'absoluteAnchor', embedId: 'rId1', ext: { cx: 95250, cy: 95250 } },
        { kind: 'twoCellAnchor', embedId: 'rId2' },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>',
      ],
      media: {
        'xl/media/image1.png': PNG_BYTES,
        'xl/media/image2.png': PNG_BYTES,
      },
    })
    const images = (await readBasicWorkbook(buffer)).snapshot.sheets[0]?.images ?? []
    // The absolute anchor still counts toward the document-order index,
    // so the following two-cell picture keeps its edit locator.
    expect(images).toHaveLength(1)
    expect(images[0]?.anchorType).toBe('two-cell')
    expect(images[0]?.drawingIndex).toBe(1)
  })

  it('preserves an absolute-anchored picture byte-for-byte through a no-op save', async () => {
    const buffer = await buildImageFixture({
      anchors: [{ kind: 'absoluteAnchor', embedId: 'rId1', ext: { cx: 95250, cy: 95250 } }],
    })
    const saved = await saveWithVisualEdits(buffer, [])
    const before = await readEntry(buffer, 'xl/drawings/drawing1.xml')
    const after = await readEntry(saved.buffer, 'xl/drawings/drawing1.xml')
    expect(after).toBe(before)
    expect(after).toContain('<xdr:pos x="47625" y="9525"/>')
    const mediaBefore = await readEntryBytes(buffer, 'xl/media/image1.png')
    const mediaAfter = await readEntryBytes(saved.buffer, 'xl/media/image1.png')
    expect(mediaAfter?.equals(mediaBefore ?? Buffer.alloc(0))).toBe(true)
  })

  it('reads rotation in degrees when a:xfrm carries rot', async () => {
    const buffer = await buildImageFixture({
      anchors: [{ kind: 'twoCellAnchor', embedId: 'rId1', rot: 5400000 }],
    })
    const images = (await readBasicWorkbook(buffer)).snapshot.sheets[0]?.images ?? []
    expect(images[0]?.rotationDeg).toBe(90)
  })

  it('skips pictures with unsupported media types', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', embedId: 'rId1' },
        { kind: 'twoCellAnchor', embedId: 'rId2' },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.wmf"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>',
      ],
      media: {
        'xl/media/image1.wmf': Buffer.from([0xd7, 0xcd, 0xc6, 0x9a]),
        'xl/media/image2.png': PNG_BYTES,
      },
    })
    const images = (await readBasicWorkbook(buffer)).snapshot.sheets[0]?.images ?? []
    expect(images).toHaveLength(1)
    expect(images[0]?.mediaType).toBe('image/png')
  })

  it('skips pictures whose media part is missing', async () => {
    const buffer = await buildImageFixture({ media: {} })
    const images = (await readBasicWorkbook(buffer)).snapshot.sheets[0]?.images
    expect(images).toBeUndefined()
  })

  it('fails closed per sheet when the drawing wiring is unreadable', async () => {
    const buffer = await buildImageFixture({ sheetDrawingElement: '<drawing r:id="rIdMissing"/>' })
    const imported = await readBasicWorkbook(buffer)
    // The workbook still opens with its cells; only the images are absent.
    expect(imported.snapshot.sheets[0]?.images).toBeUndefined()
    expect(imported.snapshot.sheets[0]?.cells['A1']?.value).toBe(10)
  })

  it('surfaces images on multiple sheets independently', async () => {
    const buffer = await buildImageFixture({ secondSheetDrawing: true })
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]?.images).toHaveLength(1)
    expect(imported.snapshot.sheets[1]?.images).toHaveLength(1)
    expect(imported.snapshot.sheets[1]?.images?.[0]?.drawingPath).toBe('xl/drawings/drawing2.xml')
  })

  it('no-op save preserves media, drawing, rels, worksheet, and content types byte-for-byte', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', embedId: 'rId1' },
        { kind: 'twoCellAnchor', embedId: 'rId2' },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>',
      ],
      media: { 'xl/media/image1.png': PNG_BYTES, 'xl/media/image2.png': PNG_BYTES },
    })
    const mutation = await applyCellEditsToXlsx(buffer, [])
    // The engine's ensureFullCalcOnLoad may touch xl/workbook.xml (adds
    // calcPr when absent — standard save behavior, unrelated to images);
    // the IMAGE-carrying parts must stay untouched and byte-identical.
    for (const path of [
      'xl/media/image1.png',
      'xl/media/image2.png',
      'xl/drawings/drawing1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
      '[Content_Types].xml',
    ]) {
      expect(mutation.touchedEntries).not.toContain(path)
      expect(await readEntryBytes(mutation.buffer, path)).toEqual(
        await readEntryBytes(buffer, path),
      )
    }
  })
})

/// ── Delete cascade ───────────────────────────────────────────────────

describe('image delete: relationship and media cascade', () => {
  const ONE: WorkbookVisualEdit = {
    drawingPath: 'xl/drawings/drawing1.xml',
    drawingIndex: 0,
    remove: true,
  }

  it('deleting the only image removes anchor, rel, media part, and the png Default', async () => {
    const buffer = await buildImageFixture()
    const mutation = await saveWithVisualEdits(buffer, [ONE])
    const drawing = await readDrawingEntry(mutation.buffer)
    expect(drawing).not.toContain('<xdr:pic>')
    const rels = await readEntry(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).not.toContain('rId1')
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image1.png')).toBeNull()
    const contentTypes = await readEntry(mutation.buffer, '[Content_Types].xml')
    expect(contentTypes).not.toContain('Extension="png"')
  })

  it('deleting one of two images preserves the other image exactly', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', embedId: 'rId1', name: 'First' },
        {
          kind: 'twoCellAnchor',
          embedId: 'rId2',
          name: 'Second',
          from: { col: 8, colOff: 0, row: 4, rowOff: 0 },
        },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>',
      ],
      media: { 'xl/media/image1.png': PNG_BYTES, 'xl/media/image2.png': PNG_BYTES },
    })
    const mutation = await saveWithVisualEdits(buffer, [ONE])
    const drawing = await readDrawingEntry(mutation.buffer)
    expect(drawing).not.toContain('r:embed="rId1"')
    expect(drawing).toContain('r:embed="rId2"')
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image2.png')).toEqual(PNG_BYTES)
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image1.png')).toBeNull()
    const contentTypes = await readEntry(mutation.buffer, '[Content_Types].xml')
    expect(contentTypes).toContain('Extension="png"')
  })

  it('never removes media shared through one relationship by two pictures', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', embedId: 'rId1' },
        { kind: 'twoCellAnchor', embedId: 'rId1' },
      ],
    })
    const mutation = await saveWithVisualEdits(buffer, [ONE])
    const drawing = await readDrawingEntry(mutation.buffer)
    expect(drawing).toContain('r:embed="rId1"')
    const rels = await readEntry(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).toContain('rId1')
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image1.png')).toEqual(PNG_BYTES)
  })

  it('never removes media shared through two relationships', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', embedId: 'rId1' },
        { kind: 'twoCellAnchor', embedId: 'rId2' },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
      ],
      media: { 'xl/media/image1.png': PNG_BYTES },
    })
    const mutation = await saveWithVisualEdits(buffer, [ONE])
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image1.png')).toEqual(PNG_BYTES)
    const rels = await readEntry(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).not.toContain('Id="rId1"')
    expect(rels).toContain('Id="rId2"')
  })

  it('never removes media referenced by another sheet’s drawing', async () => {
    const buffer = await buildImageFixture({ secondSheetDrawing: true })
    const mutation = await saveWithVisualEdits(buffer, [ONE])
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image1.png')).toEqual(PNG_BYTES)
    const sheet2Drawing = await readEntry(mutation.buffer, 'xl/drawings/drawing2.xml')
    expect(sheet2Drawing).toContain('r:embed="rIdShared"')
  })

  it('deleting the final image leaves a valid (empty) drawing part wired to the sheet', async () => {
    const buffer = await buildImageFixture()
    const mutation = await saveWithVisualEdits(buffer, [ONE])
    const entries = await listEntries(mutation.buffer)
    expect(entries).toContain('xl/drawings/drawing1.xml')
    const worksheet = await readEntry(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<drawing r:id="rIdDrawing1"/>')
    // Round-trip: the emptied workbook still parses with no images.
    const reopened = await readBasicWorkbook(mutation.buffer)
    expect(reopened.snapshot.sheets[0]?.images).toBeUndefined()
    expect(reopened.snapshot.sheets[0]?.cells['A1']?.value).toBe(10)
  })

  it('move edits rewrite only anchor geometry — media and rels untouched', async () => {
    const buffer = await buildImageFixture()
    const mutation = await saveWithVisualEdits(buffer, [
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
        anchor: {
          fromRow: 5,
          fromColumn: 3,
          fromRowOffset: 12345,
          fromColumnOffset: 6789,
          toRow: 15,
          toColumn: 9,
          toRowOffset: 1,
          toColumnOffset: 2,
        },
      },
    ])
    const drawing = await readDrawingEntry(mutation.buffer)
    expect(drawing).toContain('<xdr:col>3</xdr:col>')
    expect(drawing).toContain('<xdr:colOff>6789</xdr:colOff>')
    expect(drawing).toContain('<xdr:row>5</xdr:row>')
    expect(drawing).toContain('<xdr:rowOff>12345</xdr:rowOff>')
    expect(drawing).toContain('<xdr:to><xdr:col>9</xdr:col>')
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image1.png')).toEqual(PNG_BYTES)
    const rels = await readEntry(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).toContain('rId1')
  })

  it('delete and move compose in one save (descending order keeps indexes stable)', async () => {
    const buffer = await buildImageFixture({
      anchors: [
        { kind: 'twoCellAnchor', embedId: 'rId1' },
        { kind: 'twoCellAnchor', embedId: 'rId2', name: 'Second' },
      ],
      rels: [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>',
      ],
      media: { 'xl/media/image1.png': PNG_BYTES, 'xl/media/image2.png': PNG_BYTES },
    })
    const mutation = await saveWithVisualEdits(buffer, [
      { drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, remove: true },
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 1,
        anchor: {
          fromRow: 7,
          fromColumn: 2,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 17,
          toColumn: 8,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
      },
    ])
    const drawing = await readDrawingEntry(mutation.buffer)
    expect(drawing).not.toContain('rId1')
    expect(drawing).toContain('<xdr:row>7</xdr:row>')
    expect(await readEntryBytes(mutation.buffer, 'xl/media/image2.png')).toEqual(PNG_BYTES)
  })
})
