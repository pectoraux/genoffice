/**
 * Unit tests — /office/workbooks/save request validation for the
 * pageSetupStates family (EXCEL-026 — View / Page Layout persistence:
 * freeze panes, gridline visibility, formula-view state, and the print
 * page-setup family).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical page-setup
 * payloads (the same shape the web shell journals) and rejects wrong
 * types, out-of-range values, unknown enum members, and excessive counts
 * with 400s — nothing unvalidated reaches the engine. Placeholder bytes
 * make every ACCEPTED case fail at the engine stage with `malformed`,
 * proving validation passed without depending on a real workbook.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(
  pageSetupStates: unknown,
): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: { edits: [], pageSetupStates },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

/// A valid payload reaches the engine stage and fails there on the
/// placeholder bytes (MALFORMED, not validation).
const ACCEPTED = 'malformed'

describe('workbooks/save pageSetupStates validation — freeze (regression)', () => {
  it('accepts frozen row/column counts including the 0/0 clear', async () => {
    const res = await save([
      { sheetName: 'Data', frozenRows: 3, frozenColumns: 2 },
      { sheetName: 'Other', frozenRows: 0, frozenColumns: 0 },
    ])
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('rejects non-integer or negative freeze counts', async () => {
    for (const bad of [1.5, -1, '2', null]) {
      const res = await save([{ sheetName: 'S', frozenRows: bad }])
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('frozenRows')
    }
    const res = await save([{ sheetName: 'S', frozenColumns: -3 }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('frozenColumns')
  })

  it('rejects freeze counts beyond the OOXML sheet dimensions', async () => {
    const rows = await save([{ sheetName: 'S', frozenRows: 1_048_577 }])
    expect(rows.status).toBe(400)
    expect(rows.body.message).toContain('frozenRows')
    const cols = await save([{ sheetName: 'S', frozenColumns: 16_385 }])
    expect(cols.status).toBe(400)
    expect(cols.body.message).toContain('frozenColumns')
  })
})

describe('workbooks/save pageSetupStates validation — view flags (EXCEL-026)', () => {
  it('accepts boolean showGridlines / showFormulas / showHeadings', async () => {
    const res = await save([
      {
        sheetName: 'Data',
        showGridlines: false,
        showFormulas: true,
        showHeadings: false,
      },
    ])
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('rejects non-boolean view flags', async () => {
    for (const field of ['showGridlines', 'showFormulas', 'showHeadings']) {
      for (const bad of ['1', 1, 0, null, 'true']) {
        const res = await save([{ sheetName: 'S', [field]: bad }])
        expect(res.status).toBe(400)
        expect(res.body.message).toContain(field)
      }
    }
  })

  it('rejects a missing or empty sheet name', async () => {
    const missing = await save([{ showGridlines: false }])
    expect(missing.status).toBe(400)
    expect(missing.body.message).toContain('sheetName')
    const empty = await save([{ sheetName: '', showGridlines: false }])
    expect(empty.status).toBe(400)
    expect(empty.body.message).toContain('sheetName')
  })

  it('rejects a non-object entry', async () => {
    const res = await save(['nope'])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('pageSetupStates[0]')
  })
})

describe('workbooks/save pageSetupStates validation — print family (EXCEL-026)', () => {
  it('accepts a canonical print page-setup state', async () => {
    const res = await save([
      {
        sheetName: 'Data',
        orientation: 'landscape',
        margins: 'narrow',
        paperSize: 9,
        scale: 75,
        fitToWidth: 2,
        fitToHeight: 3,
        fitToPage: true,
      },
    ])
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('rejects an unknown orientation', async () => {
    const res = await save([{ sheetName: 'S', orientation: 'diagonal' }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('orientation')
  })

  it('rejects a non-string orientation', async () => {
    const res = await save([{ sheetName: 'S', orientation: 1 }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('orientation')
  })

  it('rejects an unknown margins preset', async () => {
    const res = await save([{ sheetName: 'S', margins: 'custom' }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('margins')
  })

  it('rejects paperSize outside the OOXML code range 1..118', async () => {
    for (const bad of [0, 119, 9.5, '9', null, -1]) {
      const res = await save([{ sheetName: 'S', paperSize: bad }])
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('paperSize')
    }
  })

  it("rejects scale outside Excel's 10..400 range", async () => {
    for (const bad of [9, 401, 0, -10, 55.5, '100']) {
      const res = await save([{ sheetName: 'S', scale: bad }])
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('scale')
    }
  })

  it('rejects fit axes outside 0..1000', async () => {
    for (const bad of [-1, 1001, 2.5, '2']) {
      const res = await save([{ sheetName: 'S', fitToWidth: bad }])
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('fitToWidth')
      const res2 = await save([{ sheetName: 'S', fitToHeight: bad }])
      expect(res2.status).toBe(400)
      expect(res2.body.message).toContain('fitToHeight')
    }
  })

  it('rejects a non-boolean fitToPage', async () => {
    const res = await save([{ sheetName: 'S', fitToPage: 'yes' }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('fitToPage')
  })
})

describe('workbooks/save pageSetupStates validation — seam behavior', () => {
  it('ignores unknown keys without error (documented forward seam)', async () => {
    const res = await save([
      {
        sheetName: 'Data',
        showGridlines: false,
        printArea: 'A1:C10',
        printTitles: '1:2',
        rowBreaks: [15],
        somethingNew: { nested: true },
      },
    ])
    // Unknown keys ride the documented seam: validation passes (the typed
    // wired fields are all valid) and the request proceeds to the engine
    // stage, which fails on the placeholder bytes.
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('validates a wired field even when unknown keys are present', async () => {
    const res = await save([{ sheetName: 'Data', printArea: 'A1:C10', scale: 999 }])
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('scale')
  })

  it('rejects more than 100 page-setup states (cap)', async () => {
    const states = Array.from({ length: 101 }, (_, index) => ({
      sheetName: `Sheet${index}`,
      showGridlines: false,
    }))
    const res = await save(states)
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('pageSetupStates exceeds 100 entries')
  })

  it('accepts exactly 100 page-setup states (at the cap)', async () => {
    const states = Array.from({ length: 100 }, (_, index) => ({
      sheetName: `Sheet${index}`,
      showGridlines: false,
    }))
    const res = await save(states)
    expect(res.body.error).toBe(ACCEPTED)
  })

  it('rejects a non-array pageSetupStates', async () => {
    const res = await save({ sheetName: 'S', showGridlines: false })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('pageSetupStates')
  })
})
