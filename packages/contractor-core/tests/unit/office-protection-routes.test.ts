/**
 * Unit tests — /office/workbooks/save request validation for the
 * sheetProtections + workbookProtectionState families (EXCEL-020 —
 * Review → Sheet/Workbook Protection).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical protection
 * payloads (the same shape the desktop ships) and rejects non-boolean
 * flags, unknown fields (including password-bearing ones — the engine
 * takes no passwords by design), missing sheet names, non-object
 * workbook states, and excessive counts with 400s — nothing unvalidated
 * reaches the engine.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(plan: {
  edits?: unknown[]
  sheetProtections?: unknown
  workbookProtectionState?: unknown
}): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: {
        edits: [],
        ...(plan.sheetProtections !== undefined ? { sheetProtections: plan.sheetProtections } : {}),
        ...(plan.workbookProtectionState !== undefined
          ? { workbookProtectionState: plan.workbookProtectionState }
          : {}),
      },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

describe('workbooks/save sheetProtections validation', () => {
  it('accepts a canonical sheet protection state (protect + unprotect)', async () => {
    const res = await save({
      sheetProtections: [
        { sheetName: 'Data', protected: true },
        { sheetName: 'Other', protected: false },
      ],
    })
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts an empty sheetProtections array (no protection change)', async () => {
    const res = await save({ sheetProtections: [] })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects a non-boolean protected flag', async () => {
    for (const bad of ['1', 1, null, 'true', undefined]) {
      const res = await save({ sheetProtections: [{ sheetName: 'S', protected: bad }] })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('protected')
    }
  })

  it('rejects a missing or empty sheet name', async () => {
    const missing = await save({ sheetProtections: [{ protected: true }] })
    expect(missing.status).toBe(400)
    expect(missing.body.message).toContain('sheetName')

    const empty = await save({ sheetProtections: [{ sheetName: '', protected: true }] })
    expect(empty.status).toBe(400)
    expect(empty.body.message).toContain('sheetName')
  })

  it('rejects unknown fields — passwords must not slip through the wire', async () => {
    const res = await save({
      sheetProtections: [{ sheetName: 'S', protected: true, password: 'hunter2' }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field')
    expect(res.body.message).toContain('password')
  })

  it('rejects a non-array sheetProtections payload', async () => {
    const res = await save({ sheetProtections: { sheetName: 'S', protected: true } })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('sheetProtections')
  })

  it('rejects an oversized sheetProtections payload', async () => {
    const entries = Array.from({ length: 1_001 }, (_, i) => ({
      sheetName: `S${i}`,
      protected: true,
    }))
    const res = await save({ sheetProtections: entries })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds')
  })
})

describe('workbooks/save workbookProtectionState validation', () => {
  it('accepts a canonical lock state', async () => {
    const res = await save({ workbookProtectionState: { lockStructure: true } })
    expect(res.body.error).toBe('malformed')
  })

  it('accepts null (untouched)', async () => {
    const res = await save({ workbookProtectionState: null })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects a non-boolean lockStructure', async () => {
    for (const bad of ['1', 1, null, 'true']) {
      const res = await save({ workbookProtectionState: { lockStructure: bad } })
      expect(res.status).toBe(400)
      expect(res.body.message).toContain('lockStructure')
    }
  })

  it('rejects unknown fields on the workbook state', async () => {
    const res = await save({
      workbookProtectionState: { lockStructure: true, workbookPassword: 'x' },
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field')
    expect(res.body.message).toContain('workbookPassword')
  })

  it('rejects a non-object workbook state', async () => {
    const res = await save({ workbookProtectionState: 'locked' })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('workbookProtectionState')
  })
})
