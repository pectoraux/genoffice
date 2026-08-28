/**
 * Unit tests — /office/workbooks/save request validation for the
 * WorkbookStyleEdit advanced-formatting fields (EXCEL-027 — borders, line
 * styles, border colors, text rotation, and indentation).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical style payloads
 * (the same shape the web shell journals) and rejects wrong types,
 * non-integer numbers, out-of-range values, and unknown enum members with
 * 400s — nothing unvalidated reaches the engine. Placeholder bytes make
 * every ACCEPTED case fail at the engine stage with `malformed`, proving
 * validation passed without depending on a real workbook.
 *
 * Also proves the border null-clear SURVIVES validation (EXCEL-027 defect
 * fix): `borderTop: null` must reach the engine as a real clear — dropping
 * it would silently keep the file's border, the same defect class the
 * EXCEL-026 freeze-clear fix closed for panes.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(style: unknown): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: {
        edits: [
          {
            sheetName: 'Data',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style,
          },
        ],
      },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

/// A valid payload reaches the engine stage and fails there on the
/// placeholder bytes (MALFORMED, not validation).
const ACCEPTED = 'malformed'

describe('workbooks/save style validation — EXCEL-027 border edges', () => {
  it('accepts every border side with every line style and a color', async () => {
    const styles = [
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
      'hair',
      'dashDot',
      'dashDotDot',
      'mediumDashed',
      'mediumDashDot',
      'mediumDashDotDot',
      'slantDashDot',
    ]
    for (const style of styles) {
      const res = await save({
        borderTop: { style, color: '#C00000' },
        borderBottom: { style },
        borderLeft: { style, color: '#00FF00' },
        borderRight: { style },
      })
      expect(res.body.error, `style ${style} must be accepted`).toBe(ACCEPTED)
    }
  })

  it('passes the null border clear through to the engine (defect fix)', async () => {
    // borderTop: null is a REAL clear — validation must not drop it, or the
    // file's border would silently survive the save.
    const res = await save({ borderTop: null, borderBottom: null })
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('rejects a non-object border edge', async () => {
    for (const bad of ['thin', 7, true, []]) {
      const res = await save({ borderTop: bad })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('borderTop')
    }
  })

  it('rejects an unknown line-style keyword', async () => {
    const res = await save({ borderLeft: { style: 'wavy' } })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('border style keyword')
  })

  it('rejects a malformed border color', async () => {
    for (const bad of ['red', 'C00000', '#C0000', '#ZZZZZZ', 42]) {
      const res = await save({ borderBottom: { style: 'thin', color: bad } })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('color')
    }
  })
})

describe('workbooks/save style validation — EXCEL-027 text rotation', () => {
  it('accepts every OOXML rotation form (1..90, 91..180, 255, and the 0 clear)', async () => {
    for (const textRotation of [0, 45, 90, 135, 180, 255]) {
      const res = await save({ textRotation })
      expect(res.body.error, `rotation ${textRotation} must be accepted`).toBe(ACCEPTED)
    }
  })

  it('rejects out-of-range and non-integer rotations', async () => {
    // (null is the route's established "absent" convention for optional
    // numbers — the browser journal never emits it for rotation.)
    for (const bad of [-1, 256, 45.5, '45']) {
      const res = await save({ textRotation: bad })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('textRotation')
    }
  })
})

describe('workbooks/save style validation — EXCEL-027 indentation', () => {
  it('accepts the full OOXML indent domain 0..250 (incl. the 0 clear)', async () => {
    for (const indent of [0, 1, 15, 16, 100, 250]) {
      const res = await save({ indent })
      expect(res.body.error, `indent ${indent} must be accepted`).toBe(ACCEPTED)
    }
  })

  it('rejects out-of-range and non-integer indents', async () => {
    for (const bad of [-1, 251, 2.5, '2']) {
      const res = await save({ indent: bad })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('indent')
    }
  })
})

describe('workbooks/save style validation — combined payload and seam', () => {
  it('accepts borders + rotation + indent + basic families in one delta', async () => {
    const res = await save({
      bold: true,
      fillColor: '#FFF2CC',
      numberFormat: '#,##0.000',
      horizontalAlignment: 'center',
      wrapText: true,
      textRotation: 60,
      indent: 2,
      borderTop: { style: 'double', color: '#00B050' },
      borderBottom: { style: 'double', color: '#00B050' },
      borderLeft: { style: 'double', color: '#00B050' },
      borderRight: { style: 'double', color: '#00B050' },
    })
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('ignores unknown style keys but still validates the wired fields', async () => {
    const res = await save({ futureField: { nested: true }, borderLeft: { style: 'wavy' } })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('border style keyword')
  })
})
