/**
 * Unit tests — /office/workbooks/save request validation for the
 * definedNamesState family (EXCEL-025 — Formulas → Name Manager).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical DefinedNamesState
 * payloads and rejects unknown fields, unsaveable names (the gateway's own
 * predicate — cell-ref lookalikes, reserved _xlnm prefixes, invalid starts,
 * over-length names), over-length formulas, non-integer / out-of-bounds
 * sheet scopes, in-scope duplicates, the modeled∩preserve collision, and
 * guard-rail overruns with 400s — nothing unvalidated reaches the engine.
 *
 * The XLSX application itself is covered by @genoffice/xlsx-gateway tests;
 * these tests pin the WIRE contract.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach a real engine —
 * a 400 'malformed' with the zip error means validation PASSED and the
 * engine stage refused the placeholder. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')
const ENGINE_STAGE_REACHED = "Can't find end of central directory"

async function save(plan: {
  edits?: unknown[]
  definedNamesState?: unknown
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

describe('workbooks/save definedNamesState validation', () => {
  it('accepts a canonical workbook- and sheet-scoped name', async () => {
    const res = await save({
      definedNamesState: {
        names: [
          { name: 'Revenue', formula: 'Data!$A$1:$A$9' },
          { name: 'LocalTotal', formula: 'Other!$A$2', sheetIndex: 1 },
        ],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(ENGINE_STAGE_REACHED)
  })

  it('accepts an empty model (all names cleared)', async () => {
    const res = await save({
      definedNamesState: { names: [], preserveNames: [] },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(ENGINE_STAGE_REACHED)
  })

  it('rejects a non-object payload', async () => {
    const res = await save({ definedNamesState: 'nope' })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('savePlan.definedNamesState must be an object')
  })

  it('rejects unknown fields at the state level', async () => {
    const res = await save({
      definedNamesState: { names: [], preserveNames: [], extra: 1 },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(
      'savePlan.definedNamesState carries an unknown field',
    )
  })

  it('rejects unknown fields on a name entry', async () => {
    const res = await save({
      definedNamesState: {
        names: [{ name: 'Revenue', formula: 'Data!$A$1', hidden: false }],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('names[0] carries an unknown field')
  })

  it('rejects unsaveable names (the canonical writer predicate)', async () => {
    for (const name of ['A1', 'TRUE', '_xlnm.Print_Area', '1Starts', 'Has Space']) {
      const res = await save({
        definedNamesState: { names: [{ name, formula: 'Data!$A$1' }], preserveNames: [] },
      })
      expect(res.status).toBe(400)
      expect(JSON.stringify(res.body)).toContain('not a saveable defined name')
    }
  })

  it('rejects a missing name or formula', async () => {
    const missingName = await save({
      definedNamesState: { names: [{ formula: 'Data!$A$1' }], preserveNames: [] },
    })
    expect(missingName.status).toBe(400)
    expect(JSON.stringify(missingName.body)).toContain('.name')
    const missingFormula = await save({
      definedNamesState: { names: [{ name: 'Revenue' }], preserveNames: [] },
    })
    expect(missingFormula.status).toBe(400)
    expect(JSON.stringify(missingFormula.body)).toContain('.formula')
  })

  it('rejects an over-length formula body', async () => {
    const res = await save({
      definedNamesState: {
        names: [{ name: 'Revenue', formula: 'X'.repeat(1_001) }],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('exceeds 1000 characters')
  })

  it('rejects a non-integer sheet scope', async () => {
    const res = await save({
      definedNamesState: {
        names: [{ name: 'Revenue', formula: 'Data!$A$1', sheetIndex: 1.5 }],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('sheetIndex must be an integer')
  })

  it('rejects an out-of-bounds sheet scope', async () => {
    const negative = await save({
      definedNamesState: {
        names: [{ name: 'Revenue', formula: 'Data!$A$1', sheetIndex: -1 }],
        preserveNames: [],
      },
    })
    expect(negative.status).toBe(400)
    expect(JSON.stringify(negative.body)).toContain('sheetIndex must be 0..255')
    const huge = await save({
      definedNamesState: {
        names: [{ name: 'Revenue', formula: 'Data!$A$1', sheetIndex: 900 }],
        preserveNames: [],
      },
    })
    expect(huge.status).toBe(400)
    expect(JSON.stringify(huge.body)).toContain('sheetIndex must be 0..255')
  })

  it('rejects a duplicate name within one scope', async () => {
    const res = await save({
      definedNamesState: {
        names: [
          { name: 'Revenue', formula: 'Data!$A$1' },
          { name: 'Revenue', formula: 'Data!$A$2' },
        ],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('defined twice')
  })

  it('rejects a case-insensitive duplicate within one scope', async () => {
    // Excel resolves names case-insensitively: 'Revenue' and 'REVENUE' at
    // one scope are the same name — the uniqueness key is
    // (case-insensitive name, scope).
    const res = await save({
      definedNamesState: {
        names: [
          { name: 'Revenue', formula: 'Data!$A$1' },
          { name: 'REVENUE', formula: 'Data!$A$2' },
        ],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('defined twice')
  })

  it('accepts the same name across different scopes', async () => {
    const res = await save({
      definedNamesState: {
        names: [
          { name: 'Revenue', formula: 'Data!$A$1' },
          { name: 'Revenue', formula: 'Other!$A$1', sheetIndex: 1 },
        ],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(ENGINE_STAGE_REACHED)
  })

  it('accepts case-variant names across different scopes', async () => {
    const res = await save({
      definedNamesState: {
        names: [
          { name: 'Revenue', formula: 'Data!$A$1' },
          { name: 'revenue', formula: 'Other!$A$1', sheetIndex: 1 },
        ],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(ENGINE_STAGE_REACHED)
  })

  it('rejects the modeled-preserve collision up front (400 before bytes)', async () => {
    const res = await save({
      definedNamesState: {
        names: [{ name: 'Revenue', formula: 'Data!$A$1' }],
        preserveNames: ['Revenue'],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('saving would duplicate it')
  })

  it('rejects a case-variant modeled-preserve collision up front (every case combination)', async () => {
    // The architect's fail-closed regression: 'Foo' + 'foo' at the same
    // scope. The reader models the winner and preserves the loser; Excel
    // resolves names case-insensitively, so the writer's collision guard
    // matches the preserve list case-insensitively — mirrored here as a
    // 400 BEFORE the engine touches bytes, in every case combination.
    const cases: Array<[string, string]> = [
      ['Foo', 'foo'],
      ['FOO', 'foo'],
      ['foo', 'FOO'],
      ['Foo', 'FOO'],
    ]
    for (const [modeled, preserved] of cases) {
      const res = await save({
        definedNamesState: {
          names: [{ name: modeled, formula: 'Data!$A$1' }],
          preserveNames: [preserved],
        },
      })
      expect(res.status).toBe(400)
      expect(JSON.stringify(res.body)).toContain('saving would duplicate it')
    }
  })

  it('accepts the same name at workbook and sheet scope with an empty preserve list (valid round-trip shape)', async () => {
    // The architect's positive regression: 'Total' at workbook scope +
    // 'Total' at sheet scope is a LEGAL Excel pair — the case-insensitive
    // collision guard must not over-block it. Validation passes and the
    // request reaches the engine stage.
    const res = await save({
      definedNamesState: {
        names: [
          { name: 'Total', formula: 'Data!$B$2:$B$4' },
          { name: 'Total', formula: 'Data!$C$7:$C$9', sheetIndex: 0 },
        ],
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(ENGINE_STAGE_REACHED)
  })

  it('rejects malformed preserve entries', async () => {
    const empty = await save({
      definedNamesState: { names: [], preserveNames: [''] },
    })
    expect(empty.status).toBe(400)
    expect(JSON.stringify(empty.body)).toContain('1..255 character string')
    const numeric = await save({
      definedNamesState: { names: [], preserveNames: [42] },
    })
    expect(numeric.status).toBe(400)
    expect(JSON.stringify(numeric.body)).toContain('1..255 character string')
  })

  it('rejects guard-rail overruns', async () => {
    const res = await save({
      definedNamesState: {
        names: Array.from({ length: 1_001 }, (_, i) => ({
          name: `Name${i}`,
          formula: 'Data!$A$1',
        })),
        preserveNames: [],
      },
    })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('exceeds 1000 entries')
  })

  it('names is required (a names-less object is malformed)', async () => {
    const res = await save({ definedNamesState: { preserveNames: [] } })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain('names must be an array')
  })

  it('null definedNamesState is the untouched no-op (accepted shape)', async () => {
    // null = the family is absent — validation passes, the engine sees no
    // names family at all (the placeholder bytes still fail the engine
    // stage, which is the expected outcome for validation-only fixtures).
    const res = await save({ definedNamesState: null })
    expect(res.status).toBe(400)
    expect(JSON.stringify(res.body)).toContain(ENGINE_STAGE_REACHED)
  })
})
