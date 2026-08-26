import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  createBufferEntrySource,
  planCellEditsToXlsx,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'
import type { SheetNoteState } from '../src/gateway/xlsx-gateway'
import { NoteReadError, parseCommentsPart, type SheetNote } from '../src/gateway/xlsx-notes'
import { buildEditFixture } from './fixture-builder'

async function planNotes(noteStates: SheetNoteState[]) {
  const source = await createBufferEntrySource(await buildEditFixture())
  return planCellEditsToXlsx(
    source,
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
    noteStates,
  )
}

const NOTES: SheetNoteState[] = [
  {
    sheetName: 'Data',
    notes: [
      { row: 0, column: 1, author: 'Reviewer', text: 'Tax inclusive <confirm>' },
      { row: 4, column: 0, author: '', text: 'second note' },
    ],
  },
]

describe('note snapshots', () => {
  it('creates the comments part with authors, refs, and escaped text', async () => {
    const plan = await planNotes(NOTES)
    const comments = [...plan.added.entries()].find(([path]) => /xl\/comments\d+\.xml/.test(path))
    expect(comments).toBeDefined()
    const xml = comments![1]
    expect(xml).toContain('<author>Reviewer</author>')
    expect(xml).toContain('<comment ref="B1" authorId="0">')
    expect(xml).toContain('Tax inclusive &lt;confirm&gt;')
    expect(xml).toContain('<comment ref="A5" authorId="1">')
  })

  it('creates the VML drawing with one Note shape per comment', async () => {
    const plan = await planNotes(NOTES)
    const vml = [...plan.added.entries()].find(([path]) =>
      /xl\/drawings\/vmlDrawing\d+\.vml/.test(path),
    )
    expect(vml).toBeDefined()
    expect(vml![1].match(/ObjectType="Note"/g)).toHaveLength(2)
    expect(vml![1]).toContain('<x:Row>0</x:Row><x:Column>1</x:Column>')
    expect(vml![1]).toContain('<x:Row>4</x:Row><x:Column>0</x:Column>')
  })

  it('registers rels, content types, and the legacyDrawing element', async () => {
    const plan = await planNotes(NOTES)
    const rels =
      plan.replaced.get('xl/worksheets/_rels/sheet1.xml.rels') ??
      plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(rels).toContain('relationships/comments')
    expect(rels).toContain('relationships/vmlDrawing')
    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('spreadsheetml.comments+xml')
    expect(contentTypes).toContain('Extension="vml"')
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<legacyDrawing r:id="')
  })

  it('is a no-op when clearing notes on a sheet that never had any', async () => {
    const plan = await planNotes([{ sheetName: 'Data', notes: [] }])
    expect([...plan.added.keys()].some((path) => path.includes('comments'))).toBe(false)
    expect(plan.replaced.has('xl/worksheets/sheet1.xml')).toBe(false)
  })
})

// ── Read path (Phase 4 Increment 6 — Review → Notes/Comments) ────────────────

const commentsPart = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice</author><author>Bob &amp; Co</author></authors>
  <commentList>${body}</commentList>
</comments>`

describe('parseCommentsPart (read path)', () => {
  it('1. parses an existing note with author and text', () => {
    const notes = parseCommentsPart(
      commentsPart('<comment ref="B2" authorId="0"><text><t>Check this value</t></text></comment>'),
    )
    expect(notes).toEqual([{ row: 1, column: 1, author: 'Alice', text: 'Check this value' }])
  })

  it('2. parses multiple notes with authorId deduplication', () => {
    const notes = parseCommentsPart(
      commentsPart(
        '<comment ref="A1" authorId="0"><text><t>first</t></text></comment>' +
          '<comment ref="B2" authorId="1"><text><t>second</t></text></comment>' +
          '<comment ref="C3" authorId="0"><text><t>third</t></text></comment>',
      ),
    )
    expect(notes).toHaveLength(3)
    expect(notes[0]).toEqual({ row: 0, column: 0, author: 'Alice', text: 'first' })
    expect(notes[1]).toEqual({ row: 1, column: 1, author: 'Bob & Co', text: 'second' })
    expect(notes[2]).toEqual({ row: 2, column: 2, author: 'Alice', text: 'third' })
  })

  it('3. preserves the author verbatim (XML-decoded)', () => {
    const notes = parseCommentsPart(
      commentsPart('<comment ref="A1" authorId="1"><text><t>x</t></text></comment>'),
    )
    expect(notes[0]!.author).toBe('Bob & Co')
  })

  it('4. preserves text with special XML characters', () => {
    const notes = parseCommentsPart(
      commentsPart(
        '<comment ref="A1" authorId="0"><text><t xml:space="preserve">a &lt; b &amp; c &gt; d &quot;e&quot;</t></text></comment>',
      ),
    )
    expect(notes[0]!.text).toBe('a < b & c > d "e"')
  })

  it('5. multi-line text survives (the author-split convention depends on it)', () => {
    const notes = parseCommentsPart(
      commentsPart(
        '<comment ref="A1" authorId="0"><text><t>line one&#10;line two</t></text></comment>',
      ),
    )
    expect(notes[0]!.text).toBe('line one\nline two')
  })

  it('6. write → reopen: applySheetNotes output parses back to the same notes', async () => {
    const original: SheetNote[] = [
      { row: 1, column: 1, author: 'Reviewer', text: 'Tax inclusive <confirm>' },
      { row: 4, column: 0, author: '', text: 'second note' },
    ]
    const mutation = await applyCellEditsToXlsx(
      await buildEditFixture(),
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
      [{ sheetName: 'Data', notes: original }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const comments = [...Object.keys(zip.files)].find((p) => /xl\/comments\d+\.xml/.test(p))
    expect(comments).toBeDefined()
    const reparsed = parseCommentsPart(await zip.file(comments!)!.async('text'))
    // Author order in the written part: 'Reviewer' then '' (first-seen).
    expect(reparsed).toEqual([
      { row: 1, column: 1, author: 'Reviewer', text: 'Tax inclusive <confirm>' },
      { row: 4, column: 0, author: '', text: 'second note' },
    ])
  })

  it('7. edit one note: the snapshot replaces the set; the other rides along', async () => {
    const edited: SheetNote[] = [
      { row: 1, column: 1, author: 'Reviewer', text: 'Tax inclusive <confirmed>' },
      { row: 4, column: 0, author: '', text: 'second note' },
    ]
    const mutation = await applyCellEditsToXlsx(
      await buildEditFixture(),
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
      [{ sheetName: 'Data', notes: edited }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const comments = [...Object.keys(zip.files)].find((p) => /xl\/comments\d+\.xml/.test(p))
    const reparsed = parseCommentsPart(await zip.file(comments!)!.async('text'))
    expect(reparsed[0]!.text).toBe('Tax inclusive <confirmed>')
    expect(reparsed[1]!.text).toBe('second note')
  })

  it('8. delete one note: the remaining note survives', async () => {
    const remaining: SheetNote[] = [{ row: 4, column: 0, author: '', text: 'second note' }]
    const mutation = await applyCellEditsToXlsx(
      await buildEditFixture(),
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
      [{ sheetName: 'Data', notes: remaining }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const comments = [...Object.keys(zip.files)].find((p) => /xl\/comments\d+\.xml/.test(p))
    const reparsed = parseCommentsPart(await zip.file(comments!)!.async('text'))
    expect(reparsed).toEqual([{ row: 4, column: 0, author: '', text: 'second note' }])
  })

  it('9. delete all notes: the comments part is removed entirely', async () => {
    // First add notes, then clear them from the SAME bytes.
    const withNotes = await applyCellEditsToXlsx(
      await buildEditFixture(),
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
      [{ sheetName: 'Data', notes: [{ row: 1, column: 1, author: 'A', text: 'x' }] }],
    )
    const cleared = await applyCellEditsToXlsx(
      withNotes.buffer,
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
      [{ sheetName: 'Data', notes: [] }],
    )
    const zip = await JSZip.loadAsync(cleared.buffer)
    expect([...Object.keys(zip.files)].some((p) => /xl\/comments\d+\.xml/.test(p))).toBe(false)
    // The worksheet's legacyDrawing element was stripped too.
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).not.toContain('legacyDrawing')
  })

  it('10. no-op save preserves the note XML byte-identically', async () => {
    const withNotes = await applyCellEditsToXlsx(
      await buildEditFixture(),
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
      [
        {
          sheetName: 'Data',
          notes: [{ row: 1, column: 1, author: 'Reviewer', text: 'keep me' }],
        },
      ],
    )
    // Empty plan (no noteStates): the comments part must survive untouched.
    const noop = await applyCellEditsToXlsx(withNotes.buffer, [])
    expect(noop.touchedEntries.some((p) => /xl\/comments\d+\.xml/.test(p))).toBe(false)
    const zip = await JSZip.loadAsync(noop.buffer)
    const comments = [...Object.keys(zip.files)].find((p) => /xl\/comments\d+\.xml/.test(p))
    const xml = await zip.file(comments!)!.async('text')
    expect(xml).toContain('<comment ref="B2" authorId="0">')
    expect(xml).toContain('keep me')
  })

  it('11. malformed note rejection: no ref, bad ref, missing text', () => {
    expect(() =>
      parseCommentsPart(commentsPart('<comment authorId="0"><text><t>x</t></text></comment>')),
    ).toThrow(NoteReadError)
    expect(() =>
      parseCommentsPart(
        commentsPart('<comment ref="not-a-cell" authorId="0"><text><t>x</t></text></comment>'),
      ),
    ).toThrow(NoteReadError)
    expect(() =>
      parseCommentsPart(commentsPart('<comment ref="A1" authorId="0"><text></text></comment>')),
    ).toThrow(NoteReadError)
  })

  it('12. oversized note-set rejection: more than 1000 comments fails closed', () => {
    const many = Array.from(
      { length: 1001 },
      (_, i) => `<comment ref="A${i + 1}" authorId="0"><text><t>note ${i}</t></text></comment>`,
    ).join('')
    expect(() => parseCommentsPart(commentsPart(many))).toThrow(NoteReadError)
  })

  it('13. unsupported legacy form: an out-of-sheet ref fails closed', () => {
    expect(() =>
      parseCommentsPart(
        commentsPart('<comment ref="ZZZZ1" authorId="0"><text><t>x</t></text></comment>'),
      ),
    ).toThrow(NoteReadError)
  })

  it('14. readBasicWorkbook integration: a fixture with notes surfaces them', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/comments1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice</author></authors>
  <commentList>
    <comment ref="A1" authorId="0"><text><t>first cell note</t></text></comment>
    <comment ref="B1" authorId="0"><text><t>second cell note &amp; more</t></text></comment>
  </commentList>
</comments>`,
    )
    const imported = await readBasicWorkbook(
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
    const sheet = imported.snapshot.sheets[0]!
    expect(sheet.cells.A1?.value).toBe(1)
    expect(sheet.notes).toEqual([
      { row: 0, column: 0, author: 'Alice', text: 'first cell note' },
      { row: 0, column: 1, author: 'Alice', text: 'second cell note & more' },
    ])
  })

  it('15. fail-closed per sheet: an unparseable comments part surfaces no notes but the workbook opens', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row>
  </sheetData>
</worksheet>`,
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>`,
    )
    // Malformed: a comment with no ref attribute.
    zip.file(
      'xl/comments1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>A</author></authors>
  <commentList><comment authorId="0"><text><t>x</t></text></comment></commentList>
</comments>`,
    )
    const imported = await readBasicWorkbook(
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
    expect(imported.snapshot.sheets[0]!.cells.A1?.value).toBe('ok')
    expect(imported.snapshot.sheets[0]!.notes).toBeUndefined()
  })
})
