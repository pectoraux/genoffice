import { useCallback, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import Link from '@tiptap/extension-link'
import { styles } from '../styles'
import { openDocument, saveDocument, readFileBytes } from '../api/office-client'
import type { SerializedBlock, SerializedRun, OfficeDocumentHandle } from '../api/office-client'

/**
 * WordEditor — real DOCX I/O backed by the GenOffice office API.
 *
 * Phase 3 Increment 4:
 *  - Inline formatting: bold, italic, underline, strike, links via Tiptap marks.
 *    SerializedBlock now carries SerializedRun[] with run-level marks.
 *  - Lists: proper <ul>/<ol> Tiptap structures, not paragraphs.
 *  - Tables/images: typed passthrough (preserved on no-op save).
 *  - Stable docxIndex: original blocks keep their canonical index via
 *    data-docx-index attribute; new blocks get docxIndex: null.
 *
 * Flow:
 *   New: creates a blank Tiptap document
 *   Open: user picks .docx → upload bytes → API parses DOCX → receive blocks → render in Tiptap
 *   Save: read Tiptap content → build SerializedBlocks with runs → API patches DOCX → download
 *
 * The browser is a thin client — the actual DOCX parsing/mutation happens
 * server-side via @genoffice/docx-engine.
 */

const DOCX_INDEX_ATTR = 'data-docx-index'

/** Escape HTML special characters in text. */
function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Convert a SerializedRun to Tiptap HTML with inline marks. */
function runToHtml(run: SerializedRun): string {
  let html = escapeHtml(run.text)
  if (run.link) {
    html = `<a href="${escapeHtml(run.link.href)}">${html}</a>`
  }
  // Wrap in marks — Tiptap supports nested marks.
  if (run.bold) html = `<strong>${html}</strong>`
  if (run.italic) html = `<em>${html}</em>`
  if (run.underline) html = `<u>${html}</u>`
  if (run.strike) html = `<s>${html}</s>`
  return html
}

/** Convert runs to Tiptap HTML. */
function runsToHtml(runs: readonly SerializedRun[] | undefined, fallbackText: string): string {
  if (!runs || runs.length === 0) return escapeHtml(fallbackText)
  return runs.map(runToHtml).join('')
}

/** Parse a Tiptap HTML element's children into SerializedRun[]. */
function parseRuns(element: Element): SerializedRun[] {
  const runs: SerializedRun[] = []
  for (const child of element.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = child.textContent ?? ''
      if (text.length > 0) runs.push({ text })
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const el = child as Element
      const text = el.textContent ?? ''
      if (text.length === 0) continue
      // Collect marks by walking up the DOM tree.
      let bold = false
      let italic = false
      let underline = false
      let strike = false
      let link: { href: string; tooltip?: string } | undefined
      let current: Element | null = el
      while (current && current !== element) {
        const ct = current.tagName.toLowerCase()
        if (ct === 'strong' || ct === 'b') bold = true
        else if (ct === 'em' || ct === 'i') italic = true
        else if (ct === 'u') underline = true
        else if (ct === 's' || ct === 'del' || ct === 'strike') strike = true
        else if (ct === 'a') {
          const href = current.getAttribute('href')
          if (href) link = { href }
        }
        current = current.parentElement
      }
      const run: { text: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; link?: { href: string; tooltip?: string } } = { text }
      if (bold) run.bold = true
      if (italic) run.italic = true
      if (underline) run.underline = true
      if (strike) run.strike = true
      if (link) run.link = link
      runs.push(run)
    }
  }
  return runs
}

export function WordEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const [title, setTitle] = useState('Document')
  const [saved, setSaved] = useState(true)
  const [status, setStatus] = useState('Ready')
  const handleRef = useRef<OfficeDocumentHandle | null>(null)
  const fileInputRef = useRef<HTMLInputElement | null>(null)

  const editor = useEditor({
    extensions: [StarterKit, Underline, Link.configure({ openOnClick: false })],
    content: '<h1>Untitled document</h1><p>Start writing your document here.</p>',
    onUpdate: () => setSaved(false),
    immediatelyRender: false,
  })

  /**
   * Convert Tiptap HTML content into SerializedBlock[] for the API.
   * P6 fix: docxIndex is read from data-docx-index attribute.
   * P4 fix: runs carry inline marks (bold, italic, underline, strike, link).
   */
  const buildBlocks = useCallback((): SerializedBlock[] => {
    if (!editor) return []
    const html = editor.getHTML()
    const doc = new DOMParser().parseFromString(html, 'text/html')
    const blocks: SerializedBlock[] = []
    const body = doc.body

    for (const node of body.children) {
      const tag = node.tagName.toLowerCase()
      const rawIndex = node.getAttribute(DOCX_INDEX_ATTR)
      const docxIndex = rawIndex !== null ? parseInt(rawIndex, 10) : null
      const hasOriginal = docxIndex !== null
      const runs = parseRuns(node)
      const text = node.textContent ?? ''
      // Mark as edited if the node has text (user may have changed it) or is new.
      const edited = hasOriginal ? text.length > 0 : true

      if (tag.match(/^h[1-6]$/)) {
        blocks.push({ docxIndex, type: 'heading', text, runs, level: parseInt(tag.slice(1), 10), edited })
      } else if (tag === 'ul' || tag === 'ol') {
        for (const li of node.querySelectorAll(':scope > li')) {
          const liRawIndex = li.getAttribute(DOCX_INDEX_ATTR)
          const liDocxIndex = liRawIndex !== null ? parseInt(liRawIndex, 10) : null
          const liRuns = parseRuns(li)
          const liText = li.textContent ?? ''
          blocks.push({
            docxIndex: liDocxIndex,
            type: 'listItem',
            text: liText,
            runs: liRuns,
            listKind: tag === 'ol' ? 'ordered' : 'bullet',
            edited: liDocxIndex !== null ? liText.length > 0 : true,
          })
        }
      } else if (tag === 'div' && node.getAttribute('data-passthrough') === 'true') {
        const ptType = node.getAttribute('data-passthrough-type') ?? 'passthrough'
        blocks.push({ docxIndex, type: ptType as SerializedBlock['type'], text, edited: false })
      } else {
        blocks.push({ docxIndex, type: 'paragraph', text, runs, edited })
      }
    }
    return blocks
  }, [editor])

  /**
   * Render API SerializedBlock[] into Tiptap HTML.
   * P4+5 fix: lists as <ul>/<ol>, tables/images as typed passthrough.
   * P4 fix: runs produce inline marks (bold, italic, underline, strike, links).
   */
  const renderBlocks = useCallback((blocks: readonly SerializedBlock[]): string => {
    if (blocks.length === 0) return '<p></p>'
    const parts: string[] = []
    let currentListTag: 'ul' | 'ol' | null = null
    let listItems: string[] = []

    const flushList = () => {
      if (currentListTag && listItems.length > 0) {
        parts.push(`<${currentListTag}>${listItems.join('')}</${currentListTag}>`)
        listItems = []
        currentListTag = null
      }
    }

    for (const block of blocks) {
      if (block.hidden) continue
      const indexAttr = block.docxIndex !== null ? ` ${DOCX_INDEX_ATTR}="${block.docxIndex}"` : ''
      const innerHtml = runsToHtml(block.runs, block.text || '')

      switch (block.type) {
        case 'heading':
          flushList()
          parts.push(`<h${block.level ?? 1}${indexAttr}>${innerHtml}</h${block.level ?? 1}>`)
          break
        case 'listItem': {
          const expectedTag = block.listKind === 'ordered' ? 'ol' : 'ul'
          if (currentListTag !== expectedTag) {
            flushList()
            currentListTag = expectedTag
          }
          listItems.push(`<li${indexAttr}>${innerHtml}</li>`)
          break
        }
        case 'table':
          flushList()
          parts.push(`<div${indexAttr} data-passthrough="true" data-passthrough-type="table">${escapeHtml(block.text || '[Table — edit in desktop app]')}</div>`)
          break
        case 'image':
          flushList()
          parts.push(`<div${indexAttr} data-passthrough="true" data-passthrough-type="image">${escapeHtml(block.text || '[Image — edit in desktop app]')}</div>`)
          break
        case 'passthrough':
          flushList()
          parts.push(`<div${indexAttr} data-passthrough="true" data-passthrough-type="passthrough">${escapeHtml(block.text || '[Embedded content — edit in desktop app]')}</div>`)
          break
        default:
          flushList()
          parts.push(`<p${indexAttr}>${innerHtml}</p>`)
      }
    }
    flushList()
    return parts.join('')
  }, [])

  const handleOpenFile = useCallback(async (file: File) => {
    setStatus('Opening...')
    try {
      const bytes = await readFileBytes(file)
      const res = await openDocument({ fileName: file.name, fileBytes: bytes })
      handleRef.current = { fileName: file.name, sourceBytes: bytes }
      const html = renderBlocks(res.blocks)
      editor?.commands.setContent(html)
      setTitle(file.name.replace(/\.[^.]+$/, ''))
      setStatus(`Opened ${file.name}`)
      setSaved(true)
    } catch (e) {
      setStatus(`Open failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }, [editor, renderBlocks])

  const handleSave = useCallback(async (saveAs: boolean) => {
    const handle = handleRef.current
    if (!handle) { setStatus('Nothing to save — open a file first'); return }
    setStatus('Saving...')
    try {
      const blocks = buildBlocks()
      let fileName = handle.fileName
      if (saveAs) {
        const newName = window.prompt('Save as:', fileName)
        if (!newName) { setStatus('Save cancelled'); return }
        fileName = newName.endsWith('.docx') ? newName : `${newName}.docx`
      }
      const savedBytes = await saveDocument({ fileName, fileBytes: handle.sourceBytes, blocks })
      handleRef.current = { fileName, sourceBytes: savedBytes }
      const blob = new Blob([savedBytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = fileName
      a.click()
      URL.revokeObjectURL(url)
      setStatus(`Saved ${fileName}`)
      setSaved(true)
    } catch (e) {
      setStatus(`Save failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }, [buildBlocks])

  const toolbar = useMemo(() => [
    ['Bold', () => editor?.chain().focus().toggleBold().run()],
    ['Italic', () => editor?.chain().focus().toggleItalic().run()],
    ['Underline', () => editor?.chain().focus().toggleUnderline().run()],
    ['Strikethrough', () => editor?.chain().focus().toggleStrike().run()],
    ['Link', () => {
      const url = window.prompt('URL:')
      if (url) editor?.chain().focus().setLink({ href: url }).run()
    }],
    ['H1', () => editor?.chain().focus().toggleHeading({ level: 1 }).run()],
    ['H2', () => editor?.chain().focus().toggleHeading({ level: 2 }).run()],
    ['• List', () => editor?.chain().focus().toggleBulletList().run()],
    ['1. List', () => editor?.chain().focus().toggleOrderedList().run()],
    ['Undo', () => editor?.chain().focus().undo().run()],
    ['Redo', () => editor?.chain().focus().redo().run()],
  ] as const, [editor])

  if (!editor) return null

  return (
    <div style={{ minHeight: 'calc(100vh - 64px)', background: '#eef1f5' }}>
      <header style={{ ...styles.header, position: 'sticky', top: 0, zIndex: 5 }}>
        <button style={styles.button} onClick={() => onRoute('/office')}>← Office</button>
        <input value={title} onChange={(e) => { setTitle(e.target.value); setSaved(false) }} style={{ marginLeft: 12, border: 0, background: 'transparent', fontSize: 18, fontWeight: 700, flex: 1 }} />
        <span style={{ opacity: .65, marginRight: 12 }}>{saved ? '✓ Saved' : '● Unsaved'}</span>
        <button style={styles.button} onClick={() => fileInputRef.current?.click()}>Open</button>
        <button style={styles.button} onClick={() => handleSave(false)} disabled={saved}>Save</button>
        <button style={styles.button} onClick={() => handleSave(true)}>Save As</button>
        <input
          ref={fileInputRef}
          hidden
          type="file"
          accept=".docx"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleOpenFile(f); e.target.value = '' }}
        />
      </header>
      {status !== 'Ready' && (
        <div style={{ padding: '4px 18px', background: '#e8f0fe', fontSize: 13, color: '#1a56c4' }}>{status}</div>
      )}
      <div style={{ background: '#fff', borderBottom: '1px solid #d9dee7', padding: '8px 18px', display: 'flex', gap: 6, flexWrap: 'wrap' }}>
        {toolbar.map(([label, action]) => <button key={label} style={styles.button} onClick={action}>{label}</button>)}
      </div>
      <main style={{ padding: '32px 20px 80px' }}>
        <div style={{ maxWidth: 850, minHeight: 1000, margin: '0 auto', background: '#fff', padding: '72px 82px', boxShadow: '0 2px 12px rgba(0,0,0,.10)' }}>
          <EditorContent editor={editor} />
        </div>
      </main>
    </div>
  )
}
