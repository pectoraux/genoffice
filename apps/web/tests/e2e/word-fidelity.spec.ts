/**
 * REAL browser E2E — Word editor fidelity matrix, part 2.
 *
 * Covers the cases word-browser.spec.ts does not:
 *   - underline / strike / hyperlink round-trip fidelity
 *   - NEW-block creation (typed paragraph → docxIndex:null → generated block)
 *   - DELETED-block behavior (block removed in the editor → absent from the
 *     save payload → absent from the saved bytes → absent on reopen)
 *
 * Same production-shaped path as the other specs: browser → Vite proxy →
 * HTTP /api/office/documents/* → vercel-handler → routeOffice →
 * @genoffice/docx-engine → DOCX bytes → browser download.
 */
import { test, expect } from '@playwright/test'
import { loginAsDemoOwner, gotoHashRoute } from './helpers'
import { buildWordFixture, buildWordMarksFixture, readZipEntry } from './fixtures'
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

/** Compute the viewport center of a word inside a ProseMirror paragraph. */
async function wordCenter(
  page: import('@playwright/test').Page,
  paragraphSelector: string,
  word: string,
): Promise<{ x: number; y: number }> {
  const rect = await page.evaluate(
    ({ selector, w }) => {
      const p = document.querySelector(selector)
      if (!p) return null
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
      while (walker.nextNode()) {
        const text = walker.currentNode.textContent ?? ''
        const idx = text.indexOf(w)
        if (idx >= 0) {
          const range = document.createRange()
          range.setStart(walker.currentNode, idx)
          range.setEnd(walker.currentNode, idx + w.length)
          const r = range.getBoundingClientRect()
          return { x: r.x + r.width / 2, y: r.y + r.height / 2 }
        }
      }
      return null
    },
    { selector: paragraphSelector, w: word },
  )
  expect(rect, `word "${word}" should exist in ${paragraphSelector}`).not.toBeNull()
  return rect!
}

/** Open a fixture in the Word editor and wait for the API round-trip. */
async function openWordFixture(
  page: import('@playwright/test').Page,
  bytes: Buffer,
  fileName: string,
  path: string,
): Promise<WireBlock[]> {
  writeFileSync(path, bytes)
  const openResponsePromise = page.waitForResponse(
    (r) => r.url().includes('/api/office/documents/open') && r.request().method() === 'POST',
  )
  await page.setInputFiles('input[type="file"]', path)
  await expect(page.getByText(`Opened ${fileName}`)).toBeVisible({ timeout: 30_000 })
  const res = await openResponsePromise
  expect(res.status()).toBe(200)
  return (await res.json()).blocks as WireBlock[]
}

/** Click Save, capture request payload + downloaded bytes. */
async function saveAndCapture(
  page: import('@playwright/test').Page,
  fileName: string,
): Promise<{ blocks: WireBlock[]; saved: Buffer }> {
  const saveRequestPromise = page.waitForRequest(
    (r) => r.url().includes('/api/office/documents/save') && r.method() === 'POST',
  )
  const saveResponsePromise = page.waitForResponse(
    (r) => r.url().includes('/api/office/documents/save') && r.request().method() === 'POST',
  )
  const downloadPromise = page.waitForEvent('download', { timeout: 30_000 })
  await page.getByRole('button', { name: 'Save', exact: true }).click()
  const [saveRequest, saveResponse, download] = await Promise.all([
    saveRequestPromise,
    saveResponsePromise,
    downloadPromise,
  ])
  await expect(page.getByText(`Saved ${fileName}`)).toBeVisible({ timeout: 15_000 })
  expect(saveResponse.status()).toBe(200)
  const blocks = JSON.parse(saveRequest.postData() ?? '{}').blocks as WireBlock[]
  const stream = await download.createReadStream()
  const chunks: Buffer[] = []
  for await (const chunk of stream) chunks.push(chunk as Buffer)
  return { blocks, saved: Buffer.concat(chunks) }
}

/** Reopen a saved file in the browser and return the parsed blocks. */
async function reopenWord(
  page: import('@playwright/test').Page,
  bytes: Buffer,
  fileName: string,
  path: string,
): Promise<WireBlock[]> {
  return openWordFixture(page, bytes, fileName, path)
}

test.describe('Word fidelity part 2 (real HTTP + real engine)', () => {
  test('underline / strike / hyperlink survive open → edit → save → reopen', async ({ page }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordMarksFixture()
    const blocks = await openWordFixture(page, fixture, 'e2e-marks.docx', '/tmp/e2e-marks.docx')

    // ── Wire: parser resolved every mark over HTTP ─────────────────────────
    expect(blocks[0].runs).toEqual([
      { text: 'Underlined words ', underline: true },
      { text: 'struck words ', strike: true },
      { text: 'plain words.' },
    ])
    expect(blocks[1].runs).toEqual([
      { text: 'Before link ' },
      { text: 'linked text', link: { href: 'https://example.com/fixture' } },
      { text: ' after link.' },
    ])
    expect(blocks[2]).toMatchObject({ type: 'heading', level: 1, text: 'Marks Fixture Heading' })

    // ── Real Tiptap DOM renders u / s / a ──────────────────────────────────
    const para0 = page.locator('.ProseMirror p[data-docx-index="0"]')
    await expect(para0.locator('u')).toHaveText('Underlined words ')
    await expect(para0.locator('s')).toHaveText('struck words ')
    const link = page.locator('.ProseMirror p[data-docx-index="1"] a')
    await expect(link).toHaveText('linked text')
    await expect(link).toHaveAttribute('href', 'https://example.com/fixture')

    // ── Toggle Underline on the plain word "plain" via the toolbar ─────────
    const center = await wordCenter(page, '.ProseMirror p[data-docx-index="0"]', 'plain')
    await page.mouse.dblclick(center.x, center.y)
    await page.getByRole('button', { name: 'Underline', exact: true }).click()
    await expect(page.getByText('● Unsaved')).toBeVisible()
    // The toggled word now renders inside a <u> element.
    await expect(para0.locator('u', { hasText: 'plain' })).toBeVisible()

    // ── Save: only block 0 edited; others byte-preserved ───────────────────
    const { blocks: savedBlocks, saved } = await saveAndCapture(page, 'e2e-marks.docx')
    const flags = savedBlocks
      .filter((b) => b.type !== 'hidden')
      .map((b) => ({ index: b.docxIndex, edited: b.edited === true }))
    expect(flags).toEqual([
      { index: 0, edited: true }, // underline toggled
      { index: 1, edited: false }, // hyperlink paragraph untouched
      { index: 2, edited: false }, // heading untouched
    ])
    // The toggle split block 0's plain run: "plain" now carries underline.
    expect(savedBlocks[0].runs).toEqual([
      { text: 'Underlined words ', underline: true },
      { text: 'struck words ', strike: true },
      { text: 'plain', underline: true },
      { text: ' words.' },
    ])

    // ── Saved bytes: regenerated block 0 has the split runs with w:u;
    //    untouched blocks keep their original XML bytes (byte-identical copy) ──
    const docXml = await readZipEntry(saved, 'word/document.xml')
    // Regenerated block 0: "Underlined words" and the toggled "plain" carry
    // w:u; " words." (split from "plain words." by the browser's parseRuns)
    // stays plain; "struck words" keeps w:strike.
    expect(docXml).toContain('Underlined words ')
    expect(docXml).toContain('struck words ')
    expect(docXml).toMatch(
      /<w:r><w:rPr><w:u w:val="single"\/><\/w:rPr><w:t xml:space="preserve">plain<\/w:t><\/w:r>/,
    )
    expect(docXml).toMatch(/<w:r><w:t xml:space="preserve"> words\.<\/w:t><\/w:r>/)
    expect(docXml).toMatch(
      /<w:r><w:rPr><w:strike\/><\/w:rPr><w:t xml:space="preserve">struck words <\/w:t><\/w:r>/,
    )
    // Untouched hyperlink paragraph: original bytes copied verbatim.
    expect(docXml).toContain(
      '<w:hyperlink r:id="rId10"><w:r><w:t xml:space="preserve">linked text</w:t></w:r></w:hyperlink>',
    )
    expect(docXml).toContain('Marks Fixture Heading')

    // ── Reopen: every mark survives the round-trip ─────────────────────────
    const reopened = await reopenWord(
      page,
      saved,
      'e2e-marks-saved.docx',
      '/tmp/e2e-marks-saved.docx',
    )
    expect(reopened[0].runs).toEqual([
      { text: 'Underlined words ', underline: true },
      { text: 'struck words ', strike: true },
      { text: 'plain', underline: true },
      { text: ' words.' },
    ])
    expect(reopened[1].runs).toEqual([
      { text: 'Before link ' },
      { text: 'linked text', link: { href: 'https://example.com/fixture' } },
      { text: ' after link.' },
    ])
    expect(reopened[2]).toMatchObject({ type: 'heading', text: 'Marks Fixture Heading' })

    expect(pageErrors).toEqual([])
  })

  test('new block created in the browser is regenerated and persisted in position', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordFixture()
    await openWordFixture(page, fixture, 'e2e-new-block.docx', '/tmp/e2e-new-block.docx')

    // ── Create a new paragraph after block 0: click into it, End, Enter ────
    const para0 = page.locator('.ProseMirror p[data-docx-index="0"]')
    await para0.click()
    await page.keyboard.press('End')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Brand new browser paragraph.')
    await expect(page.getByText('● Unsaved')).toBeVisible()

    // ── Save: the new block is generated (edited=true). Known behavior:
    // Tiptap's Enter-split clones the source paragraph's docxIndex (0), so
    // the new block carries a duplicated anchor — safe today because the
    // fingerprint comparison forces edited=true (content differs), which
    // routes it to regeneration instead of original-byte copying. The
    // original paragraph 0 remains edited=false. ────────────────────────────
    const { blocks: savedBlocks, saved } = await saveAndCapture(page, 'e2e-new-block.docx')
    const contentBlocks = savedBlocks.filter((b) => b.type !== 'hidden')
    expect(contentBlocks.map((b) => b.docxIndex)).toEqual([0, 0, 1, 2, 3, 4, 5, 6, 7])
    const newBlock = contentBlocks[1]
    expect(newBlock).toMatchObject({
      docxIndex: 0, // cloned by the Enter-split — see note above
      type: 'paragraph',
      text: 'Brand new browser paragraph.',
      edited: true,
    })
    // The paragraph it was typed after is unchanged.
    expect(contentBlocks[0]).toMatchObject({ docxIndex: 0, edited: false })

    // ── Saved bytes: new paragraph present between block 0 and the heading ─
    const docXml = await readZipEntry(saved, 'word/document.xml')
    expect(docXml).toContain('Brand new browser paragraph.')
    expect(docXml).toContain('This is a plain paragraph.')
    expect(docXml.indexOf('This is a plain paragraph.')).toBeLessThan(
      docXml.indexOf('Brand new browser paragraph.'),
    )
    expect(docXml.indexOf('Brand new browser paragraph.')).toBeLessThan(
      docXml.indexOf('Fixture Document Heading'),
    )

    // ── Reopen: the new paragraph is persisted in position ─────────────────
    const reopened = await reopenWord(
      page,
      saved,
      'e2e-new-block-saved.docx',
      '/tmp/e2e-new-block-saved.docx',
    )
    const reopenedContent = reopened.filter((b) => b.type !== 'hidden')
    // Indices are renumbered by the re-parse (the new paragraph becomes 1).
    expect(reopenedContent[1]).toMatchObject({
      type: 'paragraph',
      text: 'Brand new browser paragraph.',
    })
    expect(reopenedContent[2]).toMatchObject({ type: 'heading', text: 'Fixture Document Heading' })
    expect(reopenedContent.length).toBe(9)
    // And it renders in the real editor.
    await expect(
      page.locator('.ProseMirror p', { hasText: 'Brand new browser paragraph.' }),
    ).toBeVisible()

    expect(pageErrors).toEqual([])
  })

  test('deleted block is absent from the save payload, saved bytes, and reopen', async ({
    page,
  }) => {
    test.setTimeout(120_000)
    const pageErrors: string[] = []
    page.on('pageerror', (err) => pageErrors.push(String(err)))

    await loginAsDemoOwner(page)
    await gotoHashRoute(page, '/office/word')
    await page.waitForSelector('.ProseMirror', { timeout: 30_000 })

    const fixture = await buildWordFixture()
    await openWordFixture(page, fixture, 'e2e-delete.docx', '/tmp/e2e-delete.docx')

    // ── Delete block 2 (the bold/italic paragraph): select its text, then
    //    remove the emptied paragraph node (Backspace joins it away) ────────
    const para2 = page.locator('.ProseMirror p[data-docx-index="2"]')
    await para2.click({ clickCount: 3 }) // select the whole text block
    await page.keyboard.press('Backspace') // clear the text (empty paragraph remains)
    await page.keyboard.press('Backspace') // delete the empty paragraph node
    await expect(page.getByText('● Unsaved')).toBeVisible()
    // Block 2 is gone from the live editor DOM.
    await expect(page.locator('.ProseMirror p[data-docx-index="2"]')).toHaveCount(0)

    // ── Save: no block with docxIndex 2 in the payload ─────────────────────
    const { blocks: savedBlocks, saved } = await saveAndCapture(page, 'e2e-delete.docx')
    const contentBlocks = savedBlocks.filter((b) => b.type !== 'hidden')
    expect(contentBlocks.map((b) => b.docxIndex)).toEqual([0, 1, 3, 4, 5, 6, 7])
    // Remaining blocks are all sent as unchanged.
    expect(contentBlocks.every((b) => b.edited === false)).toBe(true)

    // ── Saved bytes: block 2's text is gone, everything else survived ──────
    const docXml = await readZipEntry(saved, 'word/document.xml')
    expect(docXml).not.toContain('This paragraph has')
    expect(docXml).not.toContain('bold text')
    expect(docXml).toContain('This is a plain paragraph.')
    expect(docXml).toContain('Fixture Document Heading')
    expect(docXml).toContain('bold italic')
    expect(docXml).toContain('First bullet item')
    expect(docXml).toContain('Second ordered item')

    // ── Reopen: block absent, indices renumbered, order intact ─────────────
    const reopened = await reopenWord(
      page,
      saved,
      'e2e-delete-saved.docx',
      '/tmp/e2e-delete-saved.docx',
    )
    const reopenedContent = reopened.filter((b) => b.type !== 'hidden')
    expect(reopenedContent.length).toBe(7)
    expect(reopenedContent.map((b) => b.docxIndex)).toEqual([0, 1, 2, 3, 4, 5, 6])
    expect(reopenedContent[0]).toMatchObject({
      type: 'paragraph',
      text: 'This is a plain paragraph.',
    })
    expect(reopenedContent[1]).toMatchObject({ type: 'heading', text: 'Fixture Document Heading' })
    // The old block 3 (nested marks) is now index 2.
    expect(reopenedContent[2]).toMatchObject({
      type: 'paragraph',
      text: 'Nested marks: outer bold bold italic',
    })
    expect(reopenedContent[3]).toMatchObject({ type: 'listItem', listKind: 'bullet' })

    expect(pageErrors).toEqual([])
  })
})
