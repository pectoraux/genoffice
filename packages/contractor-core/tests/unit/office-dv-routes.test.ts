/**
 * Unit tests — /office/workbooks/save request validation for the dvStates
 * family (Phase 4 Increment 5 — Data → Data Validation).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical SheetDvState
 * payloads (the same shape the desktop ships) and rejects unknown DV types,
 * unknown operators, invalid ranges, malformed rule objects, excessive rule
 * counts, oversized prompt/error strings, and unknown rule fields with 400s
 * — nothing unvalidated reaches the engine.
 *
 * The XLSX application itself is covered by @genoffice/xlsx-gateway tests;
 * these tests pin the WIRE contract.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(plan: {
  edits?: unknown[]
  dvStates?: unknown[]
}): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: { edits: [], ...plan },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

const AREA = { startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }

describe('workbooks/save dvStates validation', () => {
  it('accepts a canonical whole-number between rule', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              rule: {
                type: 'whole',
                operator: 'between',
                formula1: '1',
                formula2: '100',
                allowBlank: true,
                showErrorMessage: true,
                errorTitle: 'Bad',
                error: '1-100 only',
                errorStyle: 2,
              },
            },
          ],
        },
      ],
    })
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error — proving the DV state
    // itself cleared validation.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts a canonical list rule with dropdown + renderMode chrome', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              rule: {
                type: 'list',
                formula1: 'Fruit,Vegetable,Grain',
                showDropDown: true,
                uid: 'rule-1',
                renderMode: 1,
              },
            },
          ],
        },
      ],
    })
    expect(res.body.error).toBe('malformed')
  })

  it('accepts a custom formula rule and an empty cleared rules list', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'Data',
          rules: [{ ranges: [AREA], rule: { type: 'custom', formula1: '=ISNUMBER(A2)' } }],
        },
        { sheetName: 'Cleared', rules: [] },
      ],
    })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects an unknown DV type', async () => {
    const res = await save({
      dvStates: [{ sheetName: 'S', rules: [{ ranges: [AREA], rule: { type: 'iconSet' } }] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('not a supported data-validation type')
  })

  it('rejects an unknown operator', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [
            { ranges: [AREA], rule: { type: 'whole', operator: 'beginsWith', formula1: 'x' } },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('not a supported data-validation operator')
  })

  it('rejects negative row/column', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [
            {
              ranges: [{ startRow: -1, endRow: 5, startColumn: 0, endColumn: 0 }],
              rule: { type: 'whole', formula1: '1' },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('non-negative integer')
  })

  it('rejects reversed ranges', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [
            {
              ranges: [{ startRow: 5, endRow: 1, startColumn: 0, endColumn: 0 }],
              rule: { type: 'whole', formula1: '1' },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('end must be >= start')
  })

  it('rejects too many ranges on one rule', async () => {
    const ranges = Array.from({ length: 101 }, () => AREA)
    const res = await save({
      dvStates: [{ sheetName: 'S', rules: [{ ranges, rule: { type: 'whole', formula1: '1' } }] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('1..100 areas')
  })

  it('rejects a non-string/non-number formula type', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [{ ranges: [AREA], rule: { type: 'whole', formula1: { ref: 'A1' } as unknown } }],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('formula1 must be a string or number')
  })

  it('rejects oversized prompt strings', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [
            {
              ranges: [AREA],
              rule: { type: 'whole', formula1: '1', prompt: 'x'.repeat(256) },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 255 characters')
  })

  it('rejects oversized error messages', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [
            {
              ranges: [AREA],
              rule: { type: 'whole', formula1: '1', error: 'y'.repeat(300) },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 255 characters')
  })

  it('rejects an unknown rule field (extension markers do not slip through)', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [
            {
              ranges: [AREA],
              rule: { type: 'whole', formula1: '1', x14: 'extension' } as unknown,
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field "x14"')
  })

  it('rejects a non-object rule', async () => {
    const res = await save({
      dvStates: [{ sheetName: 'S', rules: ['not-a-rule' as unknown] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('must be an object')
  })

  it('rejects a range-less rule', async () => {
    const res = await save({
      dvStates: [{ sheetName: 'S', rules: [{ ranges: [], rule: { type: 'whole' } }] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('1..100 areas')
  })

  it('rejects too many rules on one sheet', async () => {
    const rules = Array.from({ length: 501 }, () => ({
      ranges: [AREA],
      rule: { type: 'whole', formula1: '1' },
    }))
    const res = await save({ dvStates: [{ sheetName: 'S', rules }] })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 500 entries')
  })

  it('rejects an absurd number of dv states', async () => {
    const states = Array.from({ length: 1001 }, () => ({ sheetName: 'S', rules: [] }))
    const res = await save({ dvStates: states })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 1000 entries')
  })

  it('rejects a non-boolean errorStyle with the 400 (type check, not whitelist alone)', async () => {
    const res = await save({
      dvStates: [
        {
          sheetName: 'S',
          rules: [{ ranges: [AREA], rule: { type: 'whole', formula1: '1', errorStyle: 'stop' } }],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('errorStyle')
  })
})
