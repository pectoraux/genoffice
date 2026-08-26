/**
 * REAL browser E2E — Word editor.
 *
 * Exercises the full production-shaped path:
 *
 *   browser (Tiptap) → Vite proxy → HTTP /api/office/documents/* →
 *   vercel-handler → routeOffice → @genoffice/docx-engine → DOCX bytes →
 *   browser download
 *
 * Never calls routeOffice() directly — every assertion observes the real
 * HTTP boundary (network interception) or the real file bytes (zip parsing).
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute } from './helpers'
import { buildWordFixture, readZipEntry } from './fixtures'
import { writeFileSync } from 'node:fs'

interface WireRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  link?: { href: string }
}
interface WireBlock {
  docxIndex: number | null
  type: string
  text: string
  runs?: WireRun[]
  level?: number
  listKind?: string
  edited?: boolean
  hidden?: boolean
}

test.describe('Word browser E2E (real HTTP + real engine)', () => {
  test('upload → render → edit paragraph → toggle bold → save → reopen → verify fidelity', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    // ── 1. Launch browser → navigate /office/word (real login flow) ───────
    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    // ── 2. Upload fixture.docx through the hidden file input ──────────────
    const fixture = await buildWordFixture()
    const fixturePath = '/tmp/e2e-word-fixture.docx'
    writeFileSync(fixturePath, fixture)

    const openResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', fixturePath)
    await expect(page.getByText('Opened e2e-word-fixture.docx')).toBeVisible({ timeout: 30_000 })
    const openResponse = await openResponsePromise
    expect(openResponse.status()).toBe(200)
    const openBody = await openResponse.json()
    const blocks = openBody.blocks as WireBlock[]

    // ── 3. Verify the blocks the browser received over HTTP ───────────────
    expect(blocks.length).toBe(9) // 8 content blocks + trailing sectPr (hidden)
    expect(blocks[0]).toMatchObject({
      docxIndex: 0,
      type: 'paragraph',
      text: 'This is a plain paragraph.',
    })
    expect(blocks[1]).toMatchObject({
      docxIndex: 1,
      type: 'heading',
      level: 1,
      text: 'Fixture Document Heading',
    })
    // Bold + italic runs survived the wire.
    expect(blocks[2].runs).toEqual([
      { text: 'This paragraph has ' },
      { text: 'bold text', bold: true },
      { text: ' and ' },
      { text: 'italic text', italic: true },
      { text: '.' },
    ])
    // Nested marks: bold + bold-italic runs.
    expect(blocks[3].runs).toEqual([
      { text: 'Nested marks: ' },
      { text: 'outer bold ', bold: true },
      { text: 'bold italic', bold: true, italic: true },
    ])
    expect(blocks[4]).toMatchObject({
      type: 'listItem',
      listKind: 'bullet',
      text: 'First bullet item',
    })
    expect(blocks[5]).toMatchObject({
      type: 'listItem',
      listKind: 'bullet',
      text: 'Second bullet item',
    })
    expect(blocks[6]).toMatchObject({
      type: 'listItem',
      listKind: 'ordered',
      text: 'First ordered item',
    })
    expect(blocks[7]).toMatchObject({
      type: 'listItem',
      listKind: 'ordered',
      text: 'Second ordered item',
    })
    expect(blocks[8]).toMatchObject({ type: 'hidden' })

    // ── 4. Verify content rendered in the REAL Tiptap editor ──────────────
    const para0 = page.locator('.ProseMirror p[data-docx-index="0"]')
    await expect(para0).toHaveText('This is a plain paragraph.')
    await expect(page.locator('.ProseMirror h1[data-docx-index="1"]')).toHaveText(
      'Fixture Document Heading',
    )
    // Bold and italic runs render as real strong/em elements.
    await expect(page.locator('.ProseMirror p[data-docx-index="2"] strong')).toHaveText('bold text')
    await expect(page.locator('.ProseMirror p[data-docx-index="2"] em')).toHaveText('italic text')
    // Nested marks render as nested strong>em.
    await expect(page.locator('.ProseMirror p[data-docx-index="3"] strong')).toContainText(
      'outer bold',
    )
    await expect(page.locator('.ProseMirror p[data-docx-index="3"] strong em')).toHaveText(
      'bold italic',
    )
    // Lists render as ul/ol with schema-backed docxIndex on each li.
    await expect(page.locator('.ProseMirror ul li[data-docx-index="4"]')).toHaveText(
      'First bullet item',
    )
    await expect(page.locator('.ProseMirror ul li[data-docx-index="5"]')).toHaveText(
      'Second bullet item',
    )
    await expect(page.locator('.ProseMirror ol li[data-docx-index="6"]')).toHaveText(
      'First ordered item',
    )
    await expect(page.locator('.ProseMirror ol li[data-docx-index="7"]')).toHaveText(
      'Second ordered item',
    )

    // Save is disabled before any edit (nothing dirty).
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeDisabled()

    // ── 5. Edit paragraph 0: triple-click selects the text block ──────────
    await para0.click({ clickCount: 3 })
    await page.keyboard.type('E2E edited paragraph text.')

    // ── 6. Toggle bold on the word "italic" in paragraph 2 ────────────────
    // Double-click selects the word; the Bold toolbar button toggles the mark.
    // The click position is offset into the first word — the em's center
    // falls on the space between "italic" and "text".
    await page
      .locator('.ProseMirror p[data-docx-index="2"] em')
      .dblclick({ position: { x: 12, y: 8 } })
    await page.getByRole('button', { name: 'Bold', exact: true }).click()

    // Dirty-state flipped (fingerprint changed for both blocks).
    await expect(page.getByText('● Unsaved')).toBeVisible()
    await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeEnabled()

    // The DOM now shows the toggled mark nested inside the italic run
    // (strong wrapping em — Tiptap's canonical mark order).
    await expect(page.locator('.ProseMirror p[data-docx-index="2"] strong em')).toHaveText('italic')

    // ── 7. Save → capture request payload + download ──────────────────────
    const saveRequestPromise = page.waitForRequest(
      (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
    )
    const saveResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/save') && r.request().method() === 'POST',
    )
    const downloadPromise = page.waitForEvent('download', { timeout: 30_000 })
    await page.getByRole('button', { name: 'Save', exact: true }).click()

    const saveRequest = await saveRequestPromise
    const saveResponse = await saveResponsePromise
    const download = await downloadPromise
    await expect(page.getByText('Saved e2e-word-fixture.docx')).toBeVisible({ timeout: 15_000 })
    expect(saveResponse.status()).toBe(200)

    // Dirty-state fingerprint tracking: ONLY the two edited blocks carry
    // edited=true; every untouched block is sent as edited=false (the server
    // copies their original bytes byte-identically).
    const saveBody = JSON.parse(saveRequest.postData() ?? '{}') as { blocks: WireBlock[] }
    const editedFlags = saveBody.blocks
      .filter((b) => b.type !== 'hidden')
      .map((b) => ({ index: b.docxIndex, edited: b.edited === true }))
    expect(editedFlags).toEqual([
      { index: 0, edited: true }, // typed into
      { index: 1, edited: false }, // heading untouched
      { index: 2, edited: true }, // bold toggled
      { index: 3, edited: false }, // nested marks untouched
      { index: 4, edited: false },
      { index: 5, edited: false },
      { index: 6, edited: false },
      { index: 7, edited: false },
    ])

    // Run-level fidelity in the edited blocks: the split run (bold toggled on
    // one word of an italic run) arrives as separate runs with correct marks.
    const savedPara2 = saveBody.blocks.find((b) => b.docxIndex === 2)
    expect(savedPara2?.runs).toEqual([
      { text: 'This paragraph has ' },
      { text: 'bold text', bold: true },
      { text: ' and ' },
      { text: 'italic', italic: true, bold: true },
      { text: ' text', italic: true },
      { text: '.' },
    ])

    const stream = await download.createReadStream()
    const chunks: Buffer[] = []
    for await (const chunk of stream) chunks.push(chunk as Buffer)
    const saved = Buffer.concat(chunks)
    expect(saved.subarray(0, 2).toString('latin1')).toBe('PK') // real zip bytes

    // ── 8. Verify the saved DOCX bytes (engine output fidelity) ───────────
    const docXml = await readZipEntry(saved, 'word/document.xml')
    // Edited paragraph 0 regenerated with the new text.
    expect(docXml).toContain('E2E edited paragraph text.')
    expect(docXml).not.toContain('This is a plain paragraph.')
    // Edited paragraph 2: "italic" run now carries BOTH bold and italic (the
    // generator also emits the complex-script twins w:bCs / w:iCs); the
    // remainder " text" keeps only italic. The run was split by the
    // browser's parseRuns text-leaf walking.
    expect(docXml).toContain(
      '<w:r><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">italic</w:t></w:r>',
    )
    expect(docXml).toContain(
      '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve"> text</w:t></w:r>',
    )
    expect(docXml).toContain(
      '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">bold text</w:t></w:r>',
    )
    // Original structures survived: heading, nested marks, lists.
    expect(docXml).toContain('Fixture Document Heading')
    expect(docXml).toContain(
      '<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">bold italic</w:t></w:r>',
    )
    expect(docXml).toContain('First bullet item')
    expect(docXml).toContain('Second bullet item')
    expect(docXml).toContain('First ordered item')
    expect(docXml).toContain('Second ordered item')
    // No phantom empty paragraph was appended after the trailing list.
    expect(docXml).not.toMatch(/Second ordered item[\s\S]*?<\/w:p><w:p\/>/)
    // Both list numbering kinds are still referenced.
    expect(docXml).toContain('<w:numId w:val="1"/>')
    expect(docXml).toContain('<w:numId w:val="2"/>')

    // ── 9. Reopen the downloaded DOCX in the browser ──────────────────────
    const savedPath = '/tmp/e2e-word-saved.docx'
    writeFileSync(savedPath, saved)
    const reopenResponsePromise = page.waitForResponse(
      (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
    )
    await page.setInputFiles('input[type="file"]', savedPath)
    await expect(page.getByText('Opened e2e-word-saved.docx')).toBeVisible({ timeout: 30_000 })
    const reopenResponse = await reopenResponsePromise
    expect(reopenResponse.status()).toBe(200)
    const reopenBody = await reopenResponse.json()
    const reopened = reopenBody.blocks as WireBlock[]

    // Edited text survived the round-trip.
    expect(reopened[0]).toMatchObject({ type: 'paragraph', text: 'E2E edited paragraph text.' })
    // The bold-toggled run survived as bold+italic.
    const reopenedRuns = reopened[2].runs ?? []
    expect(reopenedRuns).toContainEqual({ text: 'italic', bold: true, italic: true })
    expect(reopenedRuns).toContainEqual({ text: ' text', italic: true })
    // Original structures survived.
    expect(reopened[1]).toMatchObject({ type: 'heading', text: 'Fixture Document Heading' })
    expect(reopened[3].runs).toContainEqual({ text: 'bold italic', bold: true, italic: true })
    expect(reopened[4]).toMatchObject({
      type: 'listItem',
      listKind: 'bullet',
      text: 'First bullet item',
    })
    expect(reopened[6]).toMatchObject({
      type: 'listItem',
      listKind: 'ordered',
      text: 'First ordered item',
    })

    // And the reopened content renders in the real editor DOM.
    await expect(page.locator('.ProseMirror p[data-docx-index="0"]')).toHaveText(
      'E2E edited paragraph text.',
    )
    await expect(page.locator('.ProseMirror p strong em').first()).toHaveText('italic')

    // ── 10. No unexpected page errors during the whole flow ───────────────
    expect(pageErrors).toEqual([])
  })
})
