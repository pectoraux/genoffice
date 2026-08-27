/**
 * Unit tests — /office/workbooks/save request validation for the
 * visualAdditions / visualEdits families (EXCEL-022 — Insert → Picture,
 * image move/resize/delete).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical visual payloads
 * (image-only additions with bounded anchors and supported media types;
 * surgical edits carrying exactly one of remove|anchor with a valid
 * drawing locator) and rejects chart/shape additions, malformed base64
 * payloads, out-of-bounds coordinates, unknown fields, both/neither edit
 * kinds, bad drawing paths, and excessive counts with 400s — nothing
 * unvalidated reaches the engine.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(
  savePlanExtras: Record<string, unknown>,
): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: { edits: [], ...savePlanExtras },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

const ANCHOR = {
  fromRow: 2,
  fromColumn: 1,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRow: 12,
  toColumn: 6,
  toRowOffset: 0,
  toColumnOffset: 0,
}

function canonicalAddition(): Record<string, unknown> {
  return {
    sheetName: 'Data',
    anchor: { ...ANCHOR },
    image: { mediaType: 'image/png', base64: 'aGVsbG8=' },
  }
}

function canonicalEdit(): Record<string, unknown> {
  return {
    drawingPath: 'xl/drawings/drawing1.xml',
    drawingIndex: 0,
    anchor: { ...ANCHOR },
  }
}

describe('workbooks/save visualAdditions validation', () => {
  it('accepts a canonical image addition', async () => {
    const res = await save({ visualAdditions: [canonicalAddition()] })
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts an empty visualAdditions array (no visual change)', async () => {
    const res = await save({ visualAdditions: [] })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects shape additions and image+chart double payloads', async () => {
    const shape = await save({
      visualAdditions: [{ ...canonicalAddition(), shape: { shapeType: 'rect' } }],
    })
    expect(shape.status).toBe(400)
    expect(shape.body.message).toContain('unsupported visual kind')

    const both = await save({
      visualAdditions: [
        {
          ...canonicalAddition(),
          chart: {
            chartType: 'column',
            title: 'X',
            series: [{ name: 'S', categories: ['a'], values: [1] }],
          },
        },
      ],
    })
    expect(both.status).toBe(400)
    expect(both.body.message).toContain('both image and chart')
  })

  it('rejects a missing image payload', async () => {
    const res = await save({ visualAdditions: [{ sheetName: 'Data', anchor: ANCHOR }] })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('image')
  })

  it('rejects unsupported media types', async () => {
    for (const mediaType of ['image/wmf', 'image/tiff', 'image/svg+xml', 'png']) {
      const res = await save({
        visualAdditions: [{ ...canonicalAddition(), image: { mediaType, base64: 'aGVsbG8=' } }],
      })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('mediaType')
    }
  })

  it('rejects missing or non-string base64', async () => {
    const missing = await save({
      visualAdditions: [{ ...canonicalAddition(), image: { mediaType: 'image/png' } }],
    })
    expect(missing.status).toBe(400)
    expect(missing.body.message).toContain('base64')

    const notString = await save({
      visualAdditions: [{ ...canonicalAddition(), image: { mediaType: 'image/png', base64: 42 } }],
    })
    expect(notString.status).toBe(400)
    expect(notString.body.message).toContain('base64')
  })

  it('rejects missing sheet names and anchors', async () => {
    const noSheet = await save({
      visualAdditions: [{ anchor: ANCHOR, image: { mediaType: 'image/png', base64: 'aGk=' } }],
    })
    expect(noSheet.status).toBe(400)
    expect(noSheet.body.message).toContain('sheetName')

    const noAnchor = await save({
      visualAdditions: [{ sheetName: 'Data', image: { mediaType: 'image/png', base64: 'aGk=' } }],
    })
    expect(noAnchor.status).toBe(400)
    expect(noAnchor.body.message).toContain('anchor')
  })

  it('rejects fractional, negative, or out-of-bounds anchor coordinates', async () => {
    for (const bad of [
      { ...ANCHOR, fromRow: 1.5 },
      { ...ANCHOR, fromColumn: -1 },
      { ...ANCHOR, toRow: 1_048_576 },
      { ...ANCHOR, toColumn: 16_384 },
      { ...ANCHOR, fromRowOffset: -5 },
      { ...ANCHOR, toRowOffset: 50_000_001 },
    ]) {
      const res = await save({ visualAdditions: [{ ...canonicalAddition(), anchor: bad }] })
      expect(res.status).toBe(400)
      expect(res.body.message).toMatch(/anchor\./)
    }
  })

  it('rejects unknown fields on the addition and its nested objects', async () => {
    const extra = await save({
      visualAdditions: [{ ...canonicalAddition(), rotation: 90 }],
    })
    expect(extra.status).toBe(400)
    expect(extra.body.message).toContain('unknown field')

    const nested = await save({
      visualAdditions: [
        {
          ...canonicalAddition(),
          image: { mediaType: 'image/png', base64: 'aGk=', alt: 'nope' },
        },
      ],
    })
    expect(nested.status).toBe(400)
    expect(nested.body.message).toContain('unknown field')
  })

  it('rejects excessive entry counts', async () => {
    const many = Array.from({ length: 51 }, () => canonicalAddition())
    const res = await save({ visualAdditions: many })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('50')
  })
})

describe('workbooks/save visualEdits validation', () => {
  it('accepts a canonical anchor edit and a canonical removal', async () => {
    const move = await save({ visualEdits: [canonicalEdit()] })
    expect(move.body.error).toBe('malformed')
    const remove = await save({
      visualEdits: [{ drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 3, remove: true }],
    })
    expect(remove.body.error).toBe('malformed')
  })

  it('accepts an empty visualEdits array', async () => {
    const res = await save({ visualEdits: [] })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects both remove and anchor together', async () => {
    const res = await save({
      visualEdits: [{ ...canonicalEdit(), remove: true }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exactly one')
  })

  it('rejects neither remove nor anchor', async () => {
    const res = await save({
      visualEdits: [{ drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0 }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('either remove or anchor')
  })

  it('rejects a non-true remove flag', async () => {
    const res = await save({
      visualEdits: [{ drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, remove: false }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('remove')
  })

  it('rejects drawing paths outside xl/drawings', async () => {
    for (const drawingPath of [
      'xl/worksheets/sheet1.xml',
      'xl/media/image1.png',
      '../drawings/drawing1.xml',
      'xl/drawings/drawing1',
    ]) {
      const res = await save({ visualEdits: [{ ...canonicalEdit(), drawingPath }] })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('drawingPath')
    }
  })

  it('rejects fractional, negative, or oversized drawing indexes', async () => {
    for (const drawingIndex of [1.5, -1, 10_001]) {
      const res = await save({ visualEdits: [{ ...canonicalEdit(), drawingIndex }] })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('drawingIndex')
    }
  })

  it('rejects unknown fields on the edit', async () => {
    const res = await save({
      visualEdits: [{ ...canonicalEdit(), sheetName: 'Data' }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field')
  })

  it('rejects excessive edit counts', async () => {
    const many = Array.from({ length: 201 }, () => canonicalEdit())
    const res = await save({ visualEdits: many })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('200')
  })
})
