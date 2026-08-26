/**
 * Browser-side host for the nested-run fidelity Playwright test.
 *
 * Served through the Vite dev server (imported via its URL from
 * page.evaluate) so the test can mount a REAL Tiptap editor using the REAL
 * app extensions (DocxParagraph / DocxHeading / DocxListItem from
 * src/office/tiptap-docx-extensions) inside the real browser, then hand the
 * editor's DOM to the real parseRuns() from src/office/parse-runs.
 *
 * This file is test infrastructure — it is NOT imported by the app and is
 * excluded from the vitest node suite (vitest only picks up "*.test.ts").
 */

import { Editor } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import Link from '@tiptap/extension-link'
import {
  DocxParagraph,
  DocxHeading,
  DocxListItem,
  PassthroughBlock,
} from '../../src/office/tiptap-docx-extensions'

export interface MountedEditor {
  /** The contenteditable element Tiptap renders into. */
  readonly element: HTMLElement
  /** Replace the editor content. */
  setContent(html: string): void
  /** Current inner HTML of the editor. */
  getHTML(): string
  destroy(): void
}

/** Mount a real Tiptap editor with the WordEditor extension set. */
export function mountTiptapEditor(parent: HTMLElement): MountedEditor {
  const mount = document.createElement('div')
  parent.appendChild(mount)
  const editor = new Editor({
    element: mount,
    extensions: [
      StarterKit.configure({
        paragraph: false,
        heading: false,
        listItem: false,
      }),
      DocxParagraph,
      DocxHeading,
      DocxListItem,
      Underline,
      Link.configure({ openOnClick: false }),
      PassthroughBlock,
    ],
    content: '<p></p>',
  })
  return {
    element: editor.view.dom,
    setContent(html: string) {
      editor.commands.setContent(html)
    },
    getHTML() {
      return editor.getHTML()
    },
    destroy() {
      editor.destroy()
      mount.remove()
    },
  }
}
