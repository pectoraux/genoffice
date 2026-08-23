import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import Link from '@tiptap/extension-link'
import { styles } from '../styles'
import { openDocument, saveDocument, readFileBytes } from '../api/office-client'
import type {
  SerializedBlock,
  SerializedImage,
  SerializedRun,
  SerializedTable,
  OfficeDocumentHandle,
} from '../api/office-client'
import {
  DocxParagraph,
  DocxHeading,
  DocxListItem,
  PassthroughBlock,
} from '../office/tiptap-docx-extensions'
import {
  DocxTable,
  DocxTableRow,
  DocxTableCell,
  DocxTableHeader,
} from '../office/tiptap-table-extensions'
import { DocxImage } from '../office/tiptap-image-extensions'
import { parseRuns } from '../office/parse-runs'
import {
  tableToHtml,
  tableFromHtml,
  tableGridFingerprint,
  setTableParseRuns,
} from '../office/table-conversion'
import {
  imageToHtml,
  imageAttrsFromElement,
  imageAttrsToWire,
  imageFingerprint,
  imageAttrsFingerprint,
  newImageFromAttrs,
} from '../office/image-conversion'

// Wire the DOM run parser into the table conversion module (lazy injection
// avoids a module cycle; parse-runs has no dependencies).
setTableParseRuns(parseRuns)

/**
 * WordEditor — real DOCX I/O backed by the GenOffice office API.
 *
 * Phase 3 Increment 6:
 *  - Schema-backed source identity: docxIndex is a proper Tiptap node attribute
 *    on DocxParagraph, DocxHeading, DocxListItem, and PassthroughBlock — not
 *    an arbitrary HTML attribute. It survives setContent() → editing → getHTML().
 *  - PassthroughBlock: a custom atomic Tiptap node for tables/images/embedded
 *    content. Not flattened to a paragraph. Preserves docxIndex + passthroughType.
 *  - Dirty-state: fingerprint-based — unchanged blocks are sent as edited=false.
 *  - Run fidelity: text-leaf walking with mark collection + adjacent merge.
 *  - Safe DOM narrowing: no `as Text` / `as Element` casts.
 */

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

// parseRuns (DOM text-leaf walking for the save path) lives in
// ../office/parse-runs — extracted so browser-level tests can exercise the
// real document.createTreeWalker code path (vitest's node env has no DOM).

/** Compute a deterministic fingerprint for a block's editable content. */
function blockFingerprint(runs: readonly SerializedRun[] | undefined, text: string): string {
  return JSON.stringify({ runs: runs ?? [{ text }], text })
}

/** Joined plain text of a table (display-only; the payload is authoritative). */
function tableDisplayText(table: SerializedTable): string {
  return table.rows
    .flatMap((row) => row.map((cell) => cell.paras.join(' ')))
    .join(' ')
    .trim()
}

/** Minimal editor CSS for tables (borders, header shading, cell selection). */
const TABLE_EDITOR_CSS = `
.ProseMirror table { border-collapse: collapse; table-layout: fixed; width: 100%; margin: 8px 0; }
.ProseMirror th, .ProseMirror td { border: 1px solid #9aa0a6; padding: 4px 8px; vertical-align: top; min-width: 2em; }
.ProseMirror th { background: #eef1f4; font-weight: 700; text-align: left; }
.ProseMirror th p, .ProseMirror td p { margin: 0; }
.ProseMirror .selectedCell::after { content: ''; position: absolute; inset: 0; background: rgba(35, 131, 226, 0.14); pointer-events: none; }
.ProseMirror .column-resize-handle { position: absolute; right: -2px; top: 0; bottom: 0; width: 4px; background: #2383e2; }
`

/** Image editor styles (selection outline, clickable image). */
const IMAGE_EDITOR_CSS = `
.ProseMirror img[data-docx-image] { border: 1px solid #c3c9d1; padding: 2px; background: #fff; cursor: pointer; }
.ProseMirror img[data-docx-image].ProseMirror-selectednode { outline: 2px solid #2383e2; outline-offset: 2px; }
`

export function WordEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const [title, setTitle] = useState('Document')
  const [saved, setSaved] = useState(true)
  const [status, setStatus] = useState('Ready')
  const handleRef = useRef<OfficeDocumentHandle | null>(null)
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const imageInputRef = useRef<HTMLInputElement | null>(null)
  const originalFingerprintsRef = useRef<Map<number, string>>(new Map())
  /** Loaded tables by docxIndex — the echo source for byte-preservation fields
   *  (rawTcPr/borders/colWidths…) and the baseline for grid fingerprints. */
  const loadedTablesRef = useRef<Map<number, SerializedTable>>(new Map())
  /** Loaded images by docxIndex — the echo source for the unchanged-image
   *  payload and the baseline for image fingerprints. */
  const loadedImagesRef = useRef<Map<number, SerializedImage>>(new Map())
  const [inTable, setInTable] = useState(false)
  /** Selection state of the image node (drives the image toolbar). */
  const [imageSelection, setImageSelection] = useState<{
    pos: number
    attrs: Record<string, unknown>
  } | null>(null)

  const editor = useEditor({
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
      // Editable tables (Phase 3 Increment 7): real Tiptap table nodes with
      // schema-backed docxIndex on the table node.
      DocxTable.configure({ resizable: false }),
      DocxTableRow,
      DocxTableCell,
      DocxTableHeader,
      // Editable images (Phase 3 Increment 8): atomic image node with
      // schema-backed docxIndex and the canonical image properties.
      DocxImage,
    ],
    content: '<h1>Untitled document</h1><p>Start writing your document here.</p>',
    onUpdate: () => setSaved(false),
    immediatelyRender: false,
  })

  // Track whether the selection is inside a table (drives the table toolbar)
  // and whether an image node is selected (drives the image toolbar).
  useEffect(() => {
    if (!editor) return
    const update = () => {
      setInTable(editor.isActive('table'))
      // NodeSelection on the DocxImage atom: surface its current attributes.
      const sel = editor.state.selection
      const selNode = sel.$from.nodeAfter
      if (selNode && selNode.type.name === 'docxImage') {
        setImageSelection({ pos: sel.$from.pos, attrs: { ...selNode.attrs } })
      } else {
        setImageSelection(null)
      }
    }
    editor.on('selectionUpdate', update)
    editor.on('transaction', update)
    update()
    // Test hook: lets Playwright drive real editor commands/selections
    // (e.g. building a prosemirror-tables CellSelection for merge/split
    // E2E, which cannot be produced reliably by synthetic mouse drags).
    const w = window as { __genofficeWordEditor?: unknown }
    w.__genofficeWordEditor = editor
    return () => {
      editor.off('selectionUpdate', update)
      editor.off('transaction', update)
      delete w.__genofficeWordEditor
    }
  }, [editor])

  /** Update the selected image node's attributes (image toolbar edits). */
  const updateSelectedImage = useCallback(
    (attrs: Record<string, unknown>) => {
      if (!editor || !imageSelection) return
      editor
        .chain()
        .focus()
        .command(({ tr }) => {
          tr.setNodeMarkup(imageSelection.pos, undefined, {
            ...imageSelection.attrs,
            ...attrs,
          })
          return true
        })
        .run()
    },
    [editor, imageSelection],
  )

  /** Read an inserted image file and insert a new DocxImage node. */
  const handleInsertImageFile = useCallback(
    async (file: File) => {
      if (!editor) return
      if (!/^image\/(png|jpeg|gif)$/.test(file.type)) {
        setStatus(`Unsupported image type: ${file.type}`)
        return
      }
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolve(String(reader.result))
        reader.onerror = () => reject(reader.error)
        reader.readAsDataURL(file)
      }).catch(() => null)
      if (!dataUrl) {
        setStatus('Could not read the image file')
        return
      }
      // Natural pixel size from the decoded image (fallback 400×300).
      const dims = await new Promise<{ w: number; h: number }>((resolve) => {
        const img = new Image()
        img.onload = () => resolve({ w: img.naturalWidth || 400, h: img.naturalHeight || 300 })
        img.onerror = () => resolve({ w: 400, h: 300 })
        img.src = dataUrl
      })
      editor
        .chain()
        .focus()
        .insertDocxImage({
          src: dataUrl,
          widthPx: dims.w,
          heightPx: dims.h,
          wrap: 'inline',
          rotDeg: 0,
          flipH: false,
          flipV: false,
        })
        .run()
      setSaved(false)
    },
    [editor],
  )

  /**
   * Convert Tiptap HTML content into SerializedBlock[] for the API.
   * docxIndex is read from the schema-backed data-docx-index attribute
   * (rendered by DocxParagraph/DocxHeading/DocxListItem/PassthroughBlock).
   */
  const buildBlocks = useCallback((): SerializedBlock[] => {
    if (!editor) return []
    const html = editor.getHTML()
    const doc = new DOMParser().parseFromString(html, 'text/html')
    const blocks: SerializedBlock[] = []
    const body = doc.body
    const fingerprints = originalFingerprintsRef.current

    for (const node of body.children) {
      const tag = node.tagName.toLowerCase()
      const rawIndex = node.getAttribute('data-docx-index')
      const docxIndex = rawIndex !== null ? parseInt(rawIndex, 10) : null

      // ── Tables: DOM grid → SerializedTable (identity = docxIndex) ──────
      if (node instanceof HTMLTableElement) {
        const loaded = docxIndex !== null ? loadedTablesRef.current.get(docxIndex) : undefined
        const reconstructed = tableFromHtml(node, loaded)
        const edited =
          docxIndex === null ||
          loaded === undefined ||
          tableGridFingerprint(reconstructed) !== tableGridFingerprint(loaded)
        // Unchanged tables echo the loaded payload verbatim (round-trip
        // stability); edited tables send the reconstruction.
        blocks.push({
          docxIndex,
          type: 'table',
          text: tableDisplayText(loaded ?? reconstructed),
          table: edited ? reconstructed : loaded,
          edited,
        })
        continue
      }

      // ── Images: DOM attrs → SerializedImage / newImage ────────────────
      if (node instanceof HTMLImageElement && node.getAttribute('data-docx-image') === 'true') {
        const attrs = imageAttrsFromElement(node)
        if (docxIndex === null) {
          // Editor-inserted image: embed as a new media part.
          const newImage = newImageFromAttrs(attrs)
          if (newImage) {
            blocks.push({ docxIndex: null, type: 'image', text: '', newImage, edited: true })
          }
          // A non-data-URL new image cannot be embedded; drop it rather
          // than fabricate a docxIndex.
          continue
        }
        // Existing image: fingerprint the browser-editable state against
        // the loaded original (same dirty model as tables/paragraphs).
        const loaded = loadedImagesRef.current.get(docxIndex)
        const edited =
          loaded === undefined || imageAttrsFingerprint(attrs) !== imageFingerprint(loaded)
        // Unchanged images echo the loaded payload verbatim (round-trip
        // stability + server-side diff confirmation); edited images send
        // the current state for the server to patch canonically.
        blocks.push({
          docxIndex,
          type: 'image',
          text: '',
          image: edited ? imageAttrsToWire(attrs) : loaded,
          edited,
        })
        continue
      }

      const runs = parseRuns(node)
      const text = node.textContent ?? ''

      // Determine edited state via fingerprint comparison.
      let edited: boolean
      if (docxIndex !== null) {
        const original = fingerprints.get(docxIndex)
        edited = original !== undefined ? original !== blockFingerprint(runs, text) : true
      } else {
        edited = true
      }

      if (tag.match(/^h[1-6]$/)) {
        blocks.push({
          docxIndex,
          type: 'heading',
          text,
          runs,
          level: parseInt(tag.slice(1), 10),
          edited,
        })
      } else if (tag === 'ul' || tag === 'ol') {
        for (const li of node.querySelectorAll(':scope > li')) {
          const liRawIndex = li.getAttribute('data-docx-index')
          const liDocxIndex = liRawIndex !== null ? parseInt(liRawIndex, 10) : null
          const liRuns = parseRuns(li)
          const liText = li.textContent ?? ''
          let liEdited: boolean
          if (liDocxIndex !== null) {
            const liOriginal = fingerprints.get(liDocxIndex)
            liEdited =
              liOriginal !== undefined ? liOriginal !== blockFingerprint(liRuns, liText) : true
          } else {
            liEdited = true
          }
          blocks.push({
            docxIndex: liDocxIndex,
            type: 'listItem',
            text: liText,
            runs: liRuns,
            listKind: tag === 'ol' ? 'ordered' : 'bullet',
            edited: liEdited,
          })
        }
      } else if (tag === 'div' && node.getAttribute('data-passthrough') === 'true') {
        // PassthroughBlock node — preserved without destruction.
        const ptType = node.getAttribute('data-passthrough-type') ?? 'passthrough'
        blocks.push({ docxIndex, type: ptType as SerializedBlock['type'], text, edited: false })
      } else {
        blocks.push({ docxIndex, type: 'paragraph', text, runs, edited })
      }
    }
    // Tiptap appends an empty trailing paragraph when the document's last
    // block is a list (cursor placement). That paragraph is editor chrome,
    // not document content — persisting it would grow the document by a
    // phantom empty paragraph on every save. Strip trailing blocks that
    // (a) were not in the source document (docxIndex === null) and
    // (b) carry no content. A trailing paragraph the user actually typed
    // in keeps its runs and is preserved.
    while (blocks.length > 0) {
      const last = blocks[blocks.length - 1]
      const noRuns =
        !last.runs || last.runs.length === 0 || last.runs.every((r) => r.text.length === 0)
      if (
        last.docxIndex === null &&
        last.type === 'paragraph' &&
        noRuns &&
        last.text.trim().length === 0
      ) {
        blocks.pop()
      } else {
        break
      }
    }
    return blocks
  }, [editor])

  /**
   * Render API SerializedBlock[] into Tiptap HTML.
   * Uses data-docx-index on all node types (now schema-backed).
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
      const indexAttr = block.docxIndex !== null ? ` data-docx-index="${block.docxIndex}"` : ''
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
          if (block.table) {
            // Editable table: real Tiptap table nodes with schema-backed
            // docxIndex; cell content is real paragraph/text nodes.
            parts.push(tableToHtml(block.table, block.docxIndex))
          } else {
            // Table the editor cannot safely regenerate (nested tables /
            // anchored shapes in cells): read-only passthrough, byte-preserved.
            parts.push(
              `<div${indexAttr} data-passthrough="true" data-passthrough-type="table">${escapeHtml(block.text || '[Table — edit in desktop app]')}</div>`,
            )
          }
          break
        case 'image':
          flushList()
          if (block.image) {
            // Editable image: real <img> with the canonical properties as
            // schema-backed attributes; pixels render from the data URL.
            parts.push(imageToHtml(block.image, block.docxIndex))
          } else {
            // Image with unreadable media (or OLE preview): read-only
            // byte-preserved passthrough.
            parts.push(
              `<div${indexAttr} data-passthrough="true" data-passthrough-type="image">${escapeHtml(block.text || '[Image — edit in desktop app]')}</div>`,
            )
          }
          break
        case 'passthrough':
          flushList()
          parts.push(
            `<div${indexAttr} data-passthrough="true" data-passthrough-type="passthrough">${escapeHtml(block.text || '[Embedded content — edit in desktop app]')}</div>`,
          )
          break
        default:
          flushList()
          parts.push(`<p${indexAttr}>${innerHtml}</p>`)
      }
    }
    flushList()
    return parts.join('')
  }, [])

  const handleOpenFile = useCallback(
    async (file: File) => {
      setStatus('Opening...')
      try {
        const bytes = await readFileBytes(file)
        const res = await openDocument({ fileName: file.name, fileBytes: bytes })
        handleRef.current = { fileName: file.name, sourceBytes: bytes }
        const fingerprints = originalFingerprintsRef.current
        fingerprints.clear()
        loadedTablesRef.current.clear()
        loadedImagesRef.current.clear()
        for (const block of res.blocks) {
          if (block.docxIndex !== null) {
            if (block.type === 'table' && block.table) {
              // Tables participate in the same fingerprint-based dirty
              // tracking as paragraphs, over their editable grid surface.
              loadedTablesRef.current.set(block.docxIndex, block.table)
              fingerprints.set(block.docxIndex, tableGridFingerprint(block.table))
            } else if (block.type === 'image' && block.image) {
              // Images: fingerprint-based dirty tracking over the
              // browser-editable image properties (same model as tables).
              loadedImagesRef.current.set(block.docxIndex, block.image)
              fingerprints.set(block.docxIndex, imageFingerprint(block.image))
            } else {
              fingerprints.set(block.docxIndex, blockFingerprint(block.runs, block.text))
            }
          }
        }
        const html = renderBlocks(res.blocks)
        editor?.commands.setContent(html)
        setTitle(file.name.replace(/\.[^.]+$/, ''))
        setStatus(`Opened ${file.name}`)
        setSaved(true)
      } catch (e) {
        setStatus(`Open failed: ${e instanceof Error ? e.message : String(e)}`)
      }
    },
    [editor, renderBlocks],
  )

  const handleSave = useCallback(
    async (saveAs: boolean) => {
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
          if (!newName) {
            setStatus('Save cancelled')
            return
          }
          fileName = newName.endsWith('.docx') ? newName : `${newName}.docx`
        }
        const savedBytes = await saveDocument({ fileName, fileBytes: handle.sourceBytes, blocks })
        handleRef.current = { fileName, sourceBytes: savedBytes }
        const blob = new Blob([savedBytes.buffer as ArrayBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })
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
    },
    [buildBlocks],
  )

  const toolbar = useMemo(
    () =>
      [
        ['Bold', () => editor?.chain().focus().toggleBold().run()],
        ['Italic', () => editor?.chain().focus().toggleItalic().run()],
        ['Underline', () => editor?.chain().focus().toggleUnderline().run()],
        ['Strikethrough', () => editor?.chain().focus().toggleStrike().run()],
        [
          'Link',
          () => {
            const url = window.prompt('URL:')
            if (url) editor?.chain().focus().setLink({ href: url }).run()
          },
        ],
        ['H1', () => editor?.chain().focus().toggleHeading({ level: 1 }).run()],
        ['H2', () => editor?.chain().focus().toggleHeading({ level: 2 }).run()],
        ['• List', () => editor?.chain().focus().toggleBulletList().run()],
        ['1. List', () => editor?.chain().focus().toggleOrderedList().run()],
        [
          'Table',
          () =>
            editor?.chain().focus().insertTable({ rows: 2, cols: 3, withHeaderRow: false }).run(),
        ],
        ['Image', () => imageInputRef.current?.click()],
        ['Undo', () => editor?.chain().focus().undo().run()],
        ['Redo', () => editor?.chain().focus().redo().run()],
      ] as const,
    [editor],
  )

  /** Table actions — shown while the selection is inside a table. */
  const tableToolbar = useMemo(
    () =>
      [
        ['+ Row', () => editor?.chain().focus().addRowAfter().run()],
        ['- Row', () => editor?.chain().focus().deleteRow().run()],
        ['+ Col', () => editor?.chain().focus().addColumnAfter().run()],
        ['- Col', () => editor?.chain().focus().deleteColumn().run()],
        ['Header Row', () => editor?.chain().focus().toggleHeaderRow().run()],
        ['Merge Cells', () => editor?.chain().focus().mergeCells().run()],
        ['Split Cell', () => editor?.chain().focus().splitCell().run()],
        ['Delete Table', () => editor?.chain().focus().deleteTable().run()],
      ] as const,
    [editor],
  )

  /** Current attributes of the selected image (typed view for the toolbar). */
  const selImage = imageSelection
    ? {
        widthPx: Number(imageSelection.attrs['widthPx'] ?? 0) || null,
        heightPx: Number(imageSelection.attrs['heightPx'] ?? 0) || null,
        aspect:
          Number(imageSelection.attrs['widthPx'] ?? 0) /
          Math.max(1, Number(imageSelection.attrs['heightPx'] ?? 1)),
        rotDeg: Number(imageSelection.attrs['rotDeg'] ?? 0),
        flipH: imageSelection.attrs['flipH'] === true,
        flipV: imageSelection.attrs['flipV'] === true,
        wrap: String(imageSelection.attrs['wrap'] ?? 'inline'),
        crop: imageSelection.attrs['crop'] as { l: number; t: number; r: number; b: number } | null,
      }
    : null

  /** Resize preserving the aspect ratio (default) or freely. */
  const resizeSelectedImage = useCallback(
    (widthPx: number, keepAspect: boolean) => {
      if (!selImage || !selImage.widthPx || !selImage.heightPx) return
      const w = Math.max(1, Math.min(10_000, Math.round(widthPx)))
      const h = keepAspect
        ? Math.max(1, Math.min(10_000, Math.round(w / selImage.aspect)))
        : selImage.heightPx
      updateSelectedImage({ widthPx: w, heightPx: h })
    },
    [selImage, updateSelectedImage],
  )

  if (!editor) return null

  return (
    <div style={{ minHeight: 'calc(100vh - 64px)', background: '#eef1f5' }}>
      {/* Table editor styles (borders, header shading, cell selection) + image styles. */}
      <style>{TABLE_EDITOR_CSS}</style>
      <style>{IMAGE_EDITOR_CSS}</style>
      <header style={{ ...styles.header, position: 'sticky', top: 0, zIndex: 5 }}>
        <button style={styles.button} onClick={() => onRoute('/office')}>
          ← Office
        </button>
        <input
          value={title}
          onChange={(e) => {
            setTitle(e.target.value)
            setSaved(false)
          }}
          style={{
            marginLeft: 12,
            border: 0,
            background: 'transparent',
            fontSize: 18,
            fontWeight: 700,
            flex: 1,
          }}
        />
        <span style={{ opacity: 0.65, marginRight: 12 }}>{saved ? '✓ Saved' : '● Unsaved'}</span>
        <button style={styles.button} onClick={() => fileInputRef.current?.click()}>
          Open
        </button>
        <button style={styles.button} onClick={() => handleSave(false)} disabled={saved}>
          Save
        </button>
        <button style={styles.button} onClick={() => handleSave(true)}>
          Save As
        </button>
        <input
          ref={fileInputRef}
          hidden
          type="file"
          accept=".docx"
          onChange={(e) => {
            const f = e.target.files?.[0]
            if (f) void handleOpenFile(f)
            e.target.value = ''
          }}
        />
      </header>
      {status !== 'Ready' && (
        <div style={{ padding: '4px 18px', background: '#e8f0fe', fontSize: 13, color: '#1a56c4' }}>
          {status}
        </div>
      )}
      <div
        style={{
          background: '#fff',
          borderBottom: '1px solid #d9dee7',
          padding: '8px 18px',
          display: 'flex',
          gap: 6,
          flexWrap: 'wrap',
        }}
      >
        {toolbar.map(([label, action]) => (
          <button key={label} style={styles.button} onClick={action}>
            {label}
          </button>
        ))}
      </div>
      {inTable && (
        <div
          style={{
            background: '#f6f8fa',
            borderBottom: '1px solid #d9dee7',
            padding: '6px 18px',
            display: 'flex',
            gap: 6,
            flexWrap: 'wrap',
          }}
        >
          <span style={{ fontSize: 12, color: '#5b6470', alignSelf: 'center', marginRight: 4 }}>
            Table:
          </span>
          {tableToolbar.map(([label, action]) => (
            <button
              key={label}
              style={{ ...styles.button, padding: '4px 10px', fontSize: 13 }}
              onClick={action}
            >
              {label}
            </button>
          ))}
        </div>
      )}
      {selImage && (
        <div
          style={{
            background: '#f6f8fa',
            borderBottom: '1px solid #d9dee7',
            padding: '6px 18px',
            display: 'flex',
            gap: 6,
            flexWrap: 'wrap',
            alignItems: 'center',
          }}
          data-testid="image-toolbar"
        >
          <span style={{ fontSize: 12, color: '#5b6470', marginRight: 4 }}>Image:</span>
          {/* Width + height display (resize preserves aspect by default). */}
          <label style={{ fontSize: 12, display: 'flex', alignItems: 'center', gap: 4 }}>
            W×H
            <input
              type="number"
              min={1}
              max={10000}
              value={selImage.widthPx ?? ''}
              style={{ width: 70, ...styles.input, padding: '3px 6px', fontSize: 13 }}
              data-testid="image-width"
              onChange={(e) => resizeSelectedImage(Number(e.target.value), true)}
            />
            <span style={{ opacity: 0.6 }}>×</span>
            <span data-testid="image-height" style={{ minWidth: 40 }}>
              {selImage.heightPx ?? '—'}
            </span>
          </label>
          {/* Alignment. */}
          {(['left', 'center', 'right'] as const).map((a) => (
            <button
              key={`align-${a}`}
              style={{ ...styles.button, padding: '4px 8px', fontSize: 13 }}
              onClick={() => updateSelectedImage({ align: a })}
            >
              {a === 'left' ? '⯇' : a === 'center' ? '↔' : '⯈'} {a}
            </button>
          ))}
          {/* Rotate 90° / flips. */}
          <button
            style={{ ...styles.button, padding: '4px 10px', fontSize: 13 }}
            data-testid="image-rotate"
            onClick={() => updateSelectedImage({ rotDeg: (selImage.rotDeg + 90) % 360 })}
          >
            Rotate 90°
          </button>
          <button
            style={{ ...styles.button, padding: '4px 10px', fontSize: 13 }}
            data-testid="image-flip-h"
            onClick={() => updateSelectedImage({ flipH: !selImage.flipH })}
          >
            Flip H
          </button>
          <button
            style={{ ...styles.button, padding: '4px 10px', fontSize: 13 }}
            data-testid="image-flip-v"
            onClick={() => updateSelectedImage({ flipV: !selImage.flipV })}
          >
            Flip V
          </button>
          {/* Wrap mode. */}
          <label style={{ fontSize: 12, display: 'flex', alignItems: 'center', gap: 4 }}>
            Wrap
            <select
              value={selImage.wrap}
              style={{ ...styles.input, padding: '3px 6px', fontSize: 13 }}
              data-testid="image-wrap"
              onChange={(e) => updateSelectedImage({ wrap: e.target.value })}
            >
              <option value="inline">Inline</option>
              <option value="square-left">Square left</option>
              <option value="square-right">Square right</option>
              <option value="topBottom">Top &amp; bottom</option>
              <option value="behind">Behind text</option>
              <option value="front">In front</option>
            </select>
          </label>
          {/* Crop metadata (per-side fractions). */}
          <label style={{ fontSize: 12, display: 'flex', alignItems: 'center', gap: 4 }}>
            Crop L/T/R/B %
            {(['l', 't', 'r', 'b'] as const).map((side) => (
              <input
                key={`crop-${side}`}
                type="number"
                min={0}
                max={100}
                value={Math.round((selImage.crop?.[side] ?? 0) * 100)}
                style={{ width: 52, ...styles.input, padding: '3px 4px', fontSize: 13 }}
                data-testid={`image-crop-${side}`}
                onChange={(e) => {
                  const v = Math.max(0, Math.min(100, Number(e.target.value))) / 100
                  const base = selImage.crop ?? { l: 0, t: 0, r: 0, b: 0 }
                  updateSelectedImage({ crop: { ...base, [side]: v } })
                }}
              />
            ))}
          </label>
          <button
            style={{ ...styles.button, padding: '4px 10px', fontSize: 13, color: '#b3261e' }}
            data-testid="image-delete"
            onClick={() =>
              editor
                ?.chain()
                .focus()
                .command(({ state, tr }) => {
                  const sel = state.selection
                  const node = sel.$from.nodeAfter
                  if (node && node.type.name === 'docxImage') {
                    tr.delete(sel.$from.pos, sel.$from.pos + node.nodeSize)
                    return true
                  }
                  return false
                })
                .run()
            }
          >
            Delete
          </button>
        </div>
      )}
      <input
        ref={imageInputRef}
        hidden
        type="file"
        accept="image/png,image/jpeg,image/gif"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (f) void handleInsertImageFile(f)
          e.target.value = ''
        }}
      />
      <main style={{ padding: '32px 20px 80px' }}>
        <div
          style={{
            maxWidth: 850,
            minHeight: 1000,
            margin: '0 auto',
            background: '#fff',
            padding: '72px 82px',
            boxShadow: '0 2px 12px rgba(0,0,0,.10)',
          }}
        >
          <EditorContent editor={editor} />
        </div>
      </main>
    </div>
  )
}
