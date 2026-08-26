/**
 * REAL browser DOM test — nested run fidelity for parseRuns().
 *
 * parseRuns() (src/office/parse-runs.ts) uses document.createTreeWalker to
 * walk DOM text leaves — code that cannot run in vitest's node environment
 * (the old suite only verified the algorithm conceptually by counting text
 * nodes with a regex).
 *
 * This test runs the REAL function against the REAL DOM of a REAL Tiptap
 * editor (mounted with the same extensions the WordEditor uses, imported
 * through the Vite dev server so the exact app modules execute):
 *
 *   <strong>bold <em>bold+italic</em></strong>
 *
 * must produce two runs with correct mark sets.
 */
import { test, expect } from '@playwright/test'

/** Shape of the Vite-served test host module (see nested-marks-host.ts). */
interface TiptapHostModule {
  mountTiptapEditor(parent: HTMLElement): {
    element: HTMLElement
    setContent(html: string): void
    getHTML(): string
    destroy(): void
  }
}
/** Shape of the Vite-served parse-runs module (src/office/parse-runs.ts). */
interface ParseRunsModule {
  parseRuns(element: Element): unknown[]
}

test.describe('nested run fidelity (real DOM, real parseRuns)', () => {
  test('<strong>bold <em>bold+italic</em></strong> produces two runs with correct marks', async ({
    page,
  }) => {
    test.setTimeout(60_000)

    // Any page on the Vite origin serves modules — the login screen is fine.
    await page.goto('/')

    const result = await page.evaluate(async () => {
      // Vite dev-server module URLs — resolved at runtime by the browser,
      // not statically by TypeScript (hence the indirect specifiers).
      const hostModuleUrl = '/tests/e2e/nested-marks-host.ts'
      const parseRunsModuleUrl = '/src/office/parse-runs.ts'

      // Mount a real Tiptap editor with the WordEditor's real extensions.
      // These modules are served + transformed by the Vite dev server, so
      // the exact app code (tiptap-docx-extensions) executes in the browser.
      const host = (await import(hostModuleUrl)) as TiptapHostModule
      const mountPoint = document.createElement('div')
      mountPoint.setAttribute('data-e2e-tiptap-host', 'true')
      document.body.appendChild(mountPoint)
      const editor = host.mountTiptapEditor(mountPoint)

      try {
        // Nested marks: bold wrapping an italic span.
        editor.setContent('<p><strong>bold <em>bold+italic</em></strong></p>')

        // The editor's actual DOM (what the save path walks in WordEditor).
        const editorDom = editor.element
        const html = editor.getHTML()

        // Run the REAL parseRuns (the same function WordEditor's buildBlocks
        // calls on every save) against the real editor DOM paragraph.
        const parseRunsModule = (await import(parseRunsModuleUrl)) as ParseRunsModule
        const paragraph = editorDom.querySelector('p')
        const runs = paragraph ? parseRunsModule.parseRuns(paragraph) : null

        return { html, domHasNestedMarks: paragraph?.querySelector('strong em') !== null, runs }
      } finally {
        editor.destroy()
        mountPoint.remove()
      }
    })

    // The real Tiptap document model preserves the nested marks.
    expect(result.html).toContain('<strong>bold <em>bold+italic</em></strong>')
    expect(result.domHasNestedMarks).toBe(true)

    // The REAL parseRuns walks the DOM text leaves and collects marks from
    // every ancestor — producing two runs with distinct mark sets.
    expect(result.runs).toEqual([
      { text: 'bold ', bold: true },
      { text: 'bold+italic', bold: true, italic: true },
    ])
  })

  test('deeper nesting and adjacent-leaf merging behave correctly in the real DOM', async ({
    page,
  }) => {
    test.setTimeout(60_000)
    await page.goto('/')

    const result = await page.evaluate(async () => {
      const hostModuleUrl = '/tests/e2e/nested-marks-host.ts'
      const parseRunsModuleUrl = '/src/office/parse-runs.ts'
      const host = (await import(hostModuleUrl)) as TiptapHostModule
      const mountPoint = document.createElement('div')
      document.body.appendChild(mountPoint)
      const editor = host.mountTiptapEditor(mountPoint)
      try {
        editor.setContent(
          '<p><em>only-italic</em> plain <strong><em><u>all-three</u></em></strong> <em>italic-again</em></p>',
        )
        const paragraph = editor.element.querySelector('p')
        const parseRunsModule = (await import(parseRunsModuleUrl)) as ParseRunsModule
        return paragraph ? parseRunsModule.parseRuns(paragraph) : null
      } finally {
        editor.destroy()
        mountPoint.remove()
      }
    })

    // Five leaves; none share adjacent identical mark sets, so no merging.
    expect(result).toEqual([
      { text: 'only-italic', italic: true },
      { text: ' plain ' },
      { text: 'all-three', bold: true, italic: true, underline: true },
      { text: ' ' },
      { text: 'italic-again', italic: true },
    ])
  })
})
