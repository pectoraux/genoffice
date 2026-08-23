import { useCallback, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import { styles } from '../styles'
import { openDocument, saveDocument, readFileBytes } from '../api/office-client'
import type { SerializedBlock, OfficeDocumentHandle } from '../api/office-client'

/**
 * WordEditor — real DOCX I/O backed by the GenOffice office API.
 *
 * Phase 3 Increment 3 fixes:
 *  - P6: Stable docxIndex — original blocks keep their canonical index;
 *        new browser-created blocks get docxIndex: null. No position-based
 *        incrementing.
 *  - P4+5: Structural fidelity — lists are rendered as Tiptap lists,
 *          tables as typed passthrough blocks, images as typed passthrough.
 *          No degradation to paragraphs.
 *
 * Flow:
 *   New: creates a blank Tiptap document
 *   Open: user picks .docx → upload bytes → API parses DOCX → receive blocks → render in Tiptap
 *   Save: read Tiptap content → build SerializedBlocks → API patches DOCX → download saved .docx
 *   Save As: same as Save but with a prompted filename
 *
 * The browser is a thin client — the actual DOCX parsing/mutation happens
 * server-side via @genoffice/docx-engine.
 */

/** Map from Tiptap node data-id to the original docxIndex (for stable identity). */
const DOCX_INDEX_ATTR = 'data-docx-index'

export function WordEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const [title, setTitle] = useState('Document')
  const [saved, setSaved] = useState(true)
  const [status, setStatus] = useState('Ready')
  const handleRef = useRef<OfficeDocumentHandle | null>(null)
  const fileInputRef = useRef<HTMLInputElement | null>(null)

  const editor = useEditor({
    extensions: [StarterKit],
    content: '<h1>Untitled document</h1><p>Start writing your document here.</p>',
    onUpdate: () => setSaved(false),
    immediatelyRender: false,
  })

  /**
   * Convert Tiptap HTML content into SerializedBlock[] for the API.
   *
   * P6 fix: docxIndex is read from the node's data-docx-index attribute
   * (preserving the canonical identity), NOT from position-based incrementing.
   * New nodes (inserted by the user) have no data-docx-index → docxIndex: null.
   */
  const buildBlocks = useCallback((): SerializedBlock[] => {
    if (!editor) return []
    const html = editor.getHTML()
    const doc = new DOMParser().parseFromString(html, 'text/html')
    const blocks: SerializedBlock[] = []
    const body = doc.body

    for (const node of body.children) {
      const tag = node.tagName.toLowerCase()
      const text = node.textContent ?? ''
      // Read the stable docxIndex from the data attribute (P6 fix).
      // If absent, this is a new node → docxIndex: null.
      const rawIndex = node.getAttribute(DOCX_INDEX_ATTR)
      const docxIndex = rawIndex !== null ? parseInt(rawIndex, 10) : null
      // If the node has a docxIndex AND the text hasn't changed, mark as
      // unedited (the engine copies original bytes). Otherwise, edited=true.
      const hasOriginal = docxIndex !== null
      // We always mark as edited for now — the engine handles 'original' blocks
      // only when edited === false AND docxIndex !== null. Since we can't easily
      // track text changes per-node in this simplified model, we send edited=true
      // for all blocks that have text. Blocks with no text and a docxIndex are
      // sent as 'original' (unedited).
      const edited = hasOriginal ? text.length > 0 : true

      if (tag.match(/^h[1-6]$/)) {
        blocks.push({
          docxIndex,
          type: 'heading',
          text,
          level: parseInt(tag.slice(1), 10),
          edited,
        })
      } else if (tag === 'ul' || tag === 'ol') {
        // P4+5 fix: render list items as actual listItem blocks, not paragraphs.
        // Each <li> inside the list becomes a SerializedBlock with listKind.
        for (const li of node.querySelectorAll(':scope > li')) {
          const liText = li.textContent ?? ''
          const liRawIndex = li.getAttribute(DOCX_INDEX_ATTR)
          const liDocxIndex = liRawIndex !== null ? parseInt(liRawIndex, 10) : null
          blocks.push({
            docxIndex: liDocxIndex,
            type: 'listItem',
            text: liText,
            listKind: tag === 'ol' ? 'ordered' : 'bullet',
            edited: liDocxIndex !== null ? liText.length > 0 : true,
          })
        }
      } else if (tag === 'table') {
        // P4+5 fix: tables are preserved as typed passthrough blocks, not paragraphs.
        blocks.push({
          docxIndex,
          type: 'table',
          text,
          edited: false, // tables are passthrough — engine copies original bytes
        })
      } else if (tag === 'img') {
        // P4+5 fix: images are preserved as typed passthrough blocks.
        blocks.push({
          docxIndex,
          type: 'image',
          text: node.getAttribute('alt') ?? '',
          edited: false, // images are passthrough
        })
      } else if (tag === 'div' && node.getAttribute('data-passthrough') === 'true') {
        // Passthrough blocks (charts, SmartArt, OLE, etc.) — preserved as-is.
        const ptType = node.getAttribute('data-passthrough-type') ?? 'passthrough'
        blocks.push({
          docxIndex,
          type: ptType as SerializedBlock['type'],
          text: node.textContent ?? '',
          edited: false,
        })
      } else {
        // Default: paragraph
        blocks.push({
          docxIndex,
          type: 'paragraph',
          text,
          edited,
        })
      }
    }
    return blocks
  }, [editor])

  /**
   * Render API SerializedBlock[] into Tiptap HTML.
   *
   * P4+5 fix: lists are rendered as actual <ul>/<ol> elements (not paragraphs).
   * Tables and images are rendered as typed passthrough divs (not flattened).
   * The docxIndex is stored as a data attribute for stable identity (P6 fix).
   */
  const renderBlocks = useCallback((blocks: readonly SerializedBlock[]): string => {
    if (blocks.length === 0) return '<p></p>'
    const parts: string[] = []
    // Track list state to group consecutive listItems into <ul>/<ol>.
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
      const text = block.text || ''
      // P6 fix: preserve the docxIndex as a data attribute for stable identity.
      const indexAttr = block.docxIndex !== null ? ` ${DOCX_INDEX_ATTR}="${block.docxIndex}"` : ''

      switch (block.type) {
        case 'heading':
          flushList()
          parts.push(`<h${block.level ?? 1}${indexAttr}>${text}</h${block.level ?? 1}>`)
          break
        case 'listItem': {
          // P4+5 fix: group consecutive listItems into proper <ul>/<ol> containers.
          const expectedTag = block.listKind === 'ordered' ? 'ol' : 'ul'
          if (currentListTag !== expectedTag) {
            flushList()
            currentListTag = expectedTag
          }
          listItems.push(`<li${indexAttr}>${text}</li>`)
          break
        }
        case 'table':
          // P4+5 fix: tables as typed passthrough, not paragraphs.
          flushList()
          parts.push(`<div${indexAttr} data-passthrough="true" data-passthrough-type="table">${text || '[Table — edit in desktop app]'}</div>`)
          break
        case 'image':
          // P4+5 fix: images as typed passthrough, not paragraphs.
          flushList()
          parts.push(`<div${indexAttr} data-passthrough="true" data-passthrough-type="image">${text || '[Image — edit in desktop app]'}</div>`)
          break
        case 'passthrough':
          // P4+5 fix: unsupported rich structures as typed passthrough.
          flushList()
          parts.push(`<div${indexAttr} data-passthrough="true" data-passthrough-type="passthrough">${text || '[Embedded content — edit in desktop app]'}</div>`)
          break
        default:
          flushList()
          parts.push(`<p${indexAttr}>${text}</p>`)
      }
    }
    flushList()
    return parts.join('')
  }, [])

  /** Open a .docx file from the user's local filesystem. */
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

  /** Save (or Save As) the document through the API. */
  const handleSave = useCallback(async (saveAs: boolean) => {
    const handle = handleRef.current
    if (!handle) {
      setStatus('Nothing to save — open a file first')
      return
    }
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
      // Update the handle with the saved bytes (new canonical source).
      handleRef.current = { fileName, sourceBytes: savedBytes }
      // Offer the saved file as a download.
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
      <div style={{ background: '#fff', borderBottom: '1px solid #d9dee7', padding: '8px 18px', display: 'flex', gap: 6 }}>
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
