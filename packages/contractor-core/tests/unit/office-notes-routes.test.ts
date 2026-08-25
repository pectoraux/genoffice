/**
 * Unit tests — /office/workbooks/save request validation for the noteStates
 * family (Phase 4 Increment 6 — Review → Notes/Comments).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical SheetNoteState
 * payloads (the same shape the desktop ships) and rejects invalid
 * coordinates, missing text, oversized strings, invalid authors, unknown
 * fields, and excessive counts with 400s — nothing unvalidated reaches the
 * engine.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(plan: {
  edits?: unknown[]
  noteStates?: unknown[]
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

describe('workbooks/save noteStates validation', () => {
  it('accepts a canonical note state with an empty notes list (cleared)', async () => {
    const res = await save({ noteStates: [{ sheetName: 'Data', notes: [] }] })
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts canonical notes with authors and special text', async () => {
    const res = await save({
      noteStates: [
        {
          sheetName: 'Data',
          notes: [
            { row: 1, column: 1, author: 'Alice', text: 'Check <this> & that' },
            { row: 4, column: 0, author: '', text: 'second' },
          ],
        },
      ],
    })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects an invalid row coordinate (negative)', async () => {
    const res = await save({
      noteStates: [{ sheetName: 'S', notes: [{ row: -1, column: 0, author: '', text: 'x' }] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('row')
  })

  it('rejects an invalid column coordinate (out of range)', async () => {
    const res = await save({
      noteStates: [{ sheetName: 'S', notes: [{ row: 0, column: 99_999, author: '', text: 'x' }] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('column')
  })

  it('rejects missing text', async () => {
    const res = await save({
      noteStates: [
        {
          sheetName: 'S',
          notes: [{ row: 0, column: 0, author: '' }] as unknown[],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('text')
  })

  it('rejects empty text', async () => {
    const res = await save({
      noteStates: [{ sheetName: 'S', notes: [{ row: 0, column: 0, author: '', text: '' }] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('text')
  })

  it('rejects oversized text', async () => {
    const res = await save({
      noteStates: [
        {
          sheetName: 'S',
          notes: [{ row: 0, column: 0, author: '', text: 'x'.repeat(32_768) }],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('text')
  })

  it('rejects an invalid author (non-string)', async () => {
    const res = await save({
      noteStates: [
        {
          sheetName: 'S',
          notes: [{ row: 0, column: 0, author: 42, text: 'x' }] as unknown[],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('author')
  })

  it('rejects an oversized author', async () => {
    const res = await save({
      noteStates: [
        {
          sheetName: 'S',
          notes: [{ row: 0, column: 0, author: 'a'.repeat(256), text: 'x' }],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('author')
  })

  it('rejects unknown fields on a note', async () => {
    const res = await save({
      noteStates: [
        {
          sheetName: 'S',
          notes: [{ row: 0, column: 0, author: '', text: 'x', threadId: 't1' } as unknown],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('unknown field "threadId"')
  })

  it('rejects too many notes on one sheet', async () => {
    const notes = Array.from({ length: 1001 }, () => ({
      row: 0,
      column: 0,
      author: '',
      text: 'x',
    }))
    const res = await save({ noteStates: [{ sheetName: 'S', notes }] })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 1000 entries')
  })

  it('rejects an absurd number of note states', async () => {
    const states = Array.from({ length: 1001 }, () => ({ sheetName: 'S', notes: [] }))
    const res = await save({ noteStates: states })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('exceeds 1000 entries')
  })

  it('rejects a non-object note', async () => {
    const res = await save({
      noteStates: [{ sheetName: 'S', notes: ['not-a-note' as unknown] }],
    })
    expect(res.status).toBe(400)
    expect(res.body.message).toContain('must be an object')
  })
})
