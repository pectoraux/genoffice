import { useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import { styles } from '../styles'

export function WordEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const [title, setTitle] = useState('Document')
  const [saved, setSaved] = useState(true)
  const editor = useEditor({
    extensions: [StarterKit],
    content: '<h1>Untitled document</h1><p>Start writing your document here.</p>',
    onUpdate: () => setSaved(false),
    immediatelyRender: false,
  })
  const fileInput = useRef<HTMLInputElement>(null)

  const save = () => {
    if (!editor) return
    const payload = { title, html: editor.getHTML() }
    localStorage.setItem('genoffice-web-word', JSON.stringify(payload))
    setSaved(true)
  }

  const openLocal = async (file: File) => {
    const text = await file.text()
    if (!editor) return
    editor.commands.setContent(text.startsWith('{') ? JSON.parse(text).html ?? text : text)
    setTitle(file.name.replace(/\.[^.]+$/, ''))
    setSaved(true)
  }

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
        <span style={{ opacity: .65, marginRight: 12 }}>{saved ? 'Saved' : 'Unsaved changes'}</span>
        <button style={styles.button} onClick={() => fileInput.current?.click()}>Open</button>
        <button style={styles.button} onClick={save}>Save</button>
        <input ref={fileInput} hidden type="file" accept=".html,.txt,.json" onChange={(e) => e.target.files?.[0] && void openLocal(e.target.files[0])} />
      </header>
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
