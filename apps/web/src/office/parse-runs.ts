/**
 * DOM run parsing for the Word editor's save path.
 *
 * Extracted from WordEditor.tsx so the exact function that walks the
 * browser DOM (document.createTreeWalker over text leaves) can be exercised
 * by browser-level tests (Playwright) — vitest's node environment has no
 * DOM, so this code path was previously untestable.
 *
 * Purity: browser-only DOM code. Zero imports (no Node, no Electron, no
 * engine packages) — operates on standard DOM types only.
 */

import type { SerializedRun } from '../api/office-client'

function isTextNode(node: Node): node is Text {
  return node.nodeType === Node.TEXT_NODE
}

function isElementNode(node: Node): node is Element {
  return node.nodeType === Node.ELEMENT_NODE
}

/**
 * Parse a Tiptap HTML element's content into SerializedRun[].
 * Walks DOM text leaves (not container textContent) to correctly handle
 * nested marks. Adjacent runs with identical mark sets are merged.
 *
 * Example: <strong>bold <em>bold+italic</em></strong> produces
 *   [ { text: "bold ", bold: true },
 *     { text: "bold+italic", bold: true, italic: true } ]
 */
export function parseRuns(element: Element): SerializedRun[] {
  interface LeafRun {
    text: string
    bold: boolean
    italic: boolean
    underline: boolean
    strike: boolean
    linkHref: string | undefined
  }
  const leafRuns: LeafRun[] = []
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
  while (walker.nextNode()) {
    const currentNode = walker.currentNode
    if (!isTextNode(currentNode)) continue
    const text = currentNode.textContent ?? ''
    if (text.length === 0) continue
    let bold = false,
      italic = false,
      underline = false,
      strike = false
    let linkHref: string | undefined
    let current: Node | null = currentNode.parentElement
    while (current && current !== element) {
      if (isElementNode(current)) {
        const ct = current.tagName.toLowerCase()
        if (ct === 'strong' || ct === 'b') bold = true
        else if (ct === 'em' || ct === 'i') italic = true
        else if (ct === 'u') underline = true
        else if (ct === 's' || ct === 'del' || ct === 'strike') strike = true
        else if (ct === 'a') {
          const href = current.getAttribute('href')
          if (href) linkHref = href
        }
      }
      current = current.parentElement
    }
    leafRuns.push({ text, bold, italic, underline, strike, linkHref })
  }
  // Merge adjacent leaf runs with identical mark sets.
  const merged: LeafRun[] = []
  for (const lr of leafRuns) {
    const last = merged[merged.length - 1]
    if (
      last &&
      last.bold === lr.bold &&
      last.italic === lr.italic &&
      last.underline === lr.underline &&
      last.strike === lr.strike &&
      last.linkHref === lr.linkHref
    ) {
      last.text += lr.text
    } else {
      merged.push({ ...lr })
    }
  }
  return merged.map((lr) => {
    const r: {
      text: string
      bold?: boolean
      italic?: boolean
      underline?: boolean
      strike?: boolean
      link?: { href: string; tooltip?: string }
    } = { text: lr.text }
    if (lr.bold) r.bold = true
    if (lr.italic) r.italic = true
    if (lr.underline) r.underline = true
    if (lr.strike) r.strike = true
    if (lr.linkHref) r.link = { href: lr.linkHref }
    return r
  })
}
