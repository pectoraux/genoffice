import { useCallback, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import { styles } from '../styles'
import { openDocument, saveDocument, readFileBytes } from '../api/office-client'
import type { SerializedBlock } from '../api/office-client'

/**
 * WordEditor — real DOCX I/O backed by the GenOffice office API.
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
export function WordEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const [title, setTitle] = useState('Document')
  const [saved, setSaved] = useState(true)
  const [status, setStatus] = useState('Ready')
  const sourceBytesRef = useRef<Uint8Array | null>(null)
  const fileNameRef = useRef<string>('document.docx')
  const fileInputRef = useRef<HTMLInputElement | null>(null)

  const editor = useEditor({
    extensions: [StarterKit],
    content: '<h1>Untitled document</h1><p>Start writing your document here.</p>',
    onUpdate: () => setSaved(false),
    immediatelyRender: false,
  })

  /** Convert Tiptap HTML content into SerializedBlock[] for the API. */
  const buildBlocks = useCallback((): SerializedBlock[] => {
    if (!editor) return []
    const html = editor.getHTML()
    // Parse the Tiptap HTML into a simple block list.
    // Each top-level child of the ProseMirror editor is a block.
    const doc = new DOMParser().parseFromString(html, 'text/html')
    const blocks: SerializedBlock[] = []
    const body = doc.body
    let docxIndex = 0
    for (const node of body.children) {
      const tag = node.tagName.toLowerCase()
      const text = node.textContent ?? ''
      if (tag.match(/^h[1-6]$/)) {
        blocks.push({
          docxIndex: sourceBytesRef.current ? docxIndex++ : null,
          type: 'heading',
          text,
          level: parseInt(tag.slice(1), 10),
          edited: true,
        })
      } else if (tag === 'ul' || tag === 'ol') {
        // Unwrap list items as individual blocks.
        for (const li of node.querySelectorAll(':scope > li')) {
          blocks.push({
            docxIndex: sourceBytesRef.current ? docxIndex++ : null,
            type: 'listItem',
            text: li.textContent ?? '',
            listKind: tag === 'ol' ? 'ordered' : 'bullet',
            edited: true,
          })
        }
      } else {
        blocks.push({
          docxIndex: sourceBytesRef.current ? docxIndex++ : null,
          type: 'paragraph',
          text,
          edited: true,
        })
      }
    }
    return blocks
  }, [editor])

  /** Render API SerializedBlock[] into Tiptap HTML. */
  const renderBlocks = useCallback((blocks: readonly SerializedBlock[]): string => {
    if (blocks.length === 0) return '<p></p>'
    const parts: string[] = []
    for (const block of blocks) {
      if (block.hidden) continue
      const text = block.text || ''
      switch (block.type) {
        case 'heading':
          parts.push(`<h${block.level ?? 1}>${text}</h${block.level ?? 1}>`)
          break
        case 'listItem':
          // We don't track list containers — render as paragraphs for now.
          parts.push(`<p>${text}</p>`)
          break
        case 'table':
          parts.push(`<p>${text || '[Table]'}</p>`)
          break
        case 'image':
          parts.push(`<p>${text || '[Image]'}</p>`)
          break
        case 'passthrough':
          parts.push(`<p>${text || '[Embedded content]'}</p>`)
          break
        default:
          parts.push(`<p>${text}</p>`)
      }
    }
    return parts.join('')
  }, [])

  /** Open a .docx file from the user's local filesystem. */
  const handleOpenFile = useCallback(async (file: File) => {
    setStatus('Opening...')
    try {
      const bytes = await readFileBytes(file)
      const res = await openDocument({ fileName: file.name, fileBytes: bytes })
      sourceBytesRef.current = bytes
      fileNameRef.current = file.name
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
    const sourceBytes = sourceBytesRef.current
    if (!sourceBytes) {
      setStatus('Nothing to save — open a file first')
      return
    }
    setStatus('Saving...')
    try {
      const blocks = buildBlocks()
      let fileName = fileNameRef.current
      if (saveAs) {
        const newName = window.prompt('Save as:', fileName)
        if (!newName) { setStatus('Save cancelled'); return }
        fileName = newName.endsWith('.docx') ? newName : `${newName}.docx`
      }
      const savedBytes = await saveDocument({ fileName, fileBytes: sourceBytes, blocks })
      sourceBytesRef.current = savedBytes
      fileNameRef.current = fileName
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
