import { useCallback, useEffect, useRef, useState } from 'react'
import type { WorkbookSnapshot, CellEdit, CellState } from '@genoffice/xlsx-gateway'
import { createBrowserUniver, type BrowserUniverRuntime } from '../office/create-browser-univer'
import { openWorkbook, saveWorkbook, readFileBytes } from '../api/office-client'
import { styles } from '../styles'

/**
 * ExcelEditor — real XLSX I/O backed by the GenOffice office API.
 *
 * Flow:
 *   New: creates a blank Univer workbook
 *   Open: user picks .xlsx → upload bytes → API reads workbook → load snapshot into Univer
 *   Save: read Univer cells → build CellEdits → API applies edits → download saved .xlsx
 *   Save As: same as Save but with a prompted filename
 *
 * The browser is a thin client — the actual XLSX parsing/mutation happens
 * server-side via @genoffice/xlsx-gateway.
 */
export function ExcelEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const runtimeRef = useRef<BrowserUniverRuntime | null>(null)
  const sourceBytesRef = useRef<Uint8Array | null>(null)
  const fileNameRef = useRef<string>('workbook.xlsx')
  const [dirty, setDirty] = useState(false)
  const [status, setStatus] = useState<string>('Ready')
  const fileInputRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    if (!containerRef.current) return
    const runtime = createBrowserUniver('genoffice-web-excel')
    runtimeRef.current = runtime
    runtime.univerAPI.createWorkbook({
      id: 'genoffice-web-workbook',
      name: 'Workbook',
      sheets: { sheet1: { name: 'Sheet1' } },
    })

    const sub = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      () => setDirty(true),
    )
    return () => {
      sub.dispose()
      runtime.univer.dispose()
      runtimeRef.current = null
    }
  }, [])

  /** Load a WorkbookSnapshot (from the API) into the Univer workbook. */
  const loadSnapshot = useCallback((snapshot: WorkbookSnapshot, _sheetNamesById: Readonly<Record<string, string>>) => {
    const rt = runtimeRef.current
    if (!rt) return
    // Dispose the old workbook and create a fresh one with the snapshot's sheets.
    try { rt.univer.dispose() } catch { /* */ }
    const fresh = createBrowserUniver('genoffice-web-excel')
    runtimeRef.current = fresh
    const sheetsConfig: Record<string, { name: string }> = {}
    for (const sheet of snapshot.sheets) {
      sheetsConfig[sheet.id] = { name: sheet.name }
    }
    const newWb = fresh.univerAPI.createWorkbook({
      id: 'genoffice-web-workbook',
      name: fileNameRef.current.replace(/\.[^.]+$/, ''),
      sheets: sheetsConfig,
    })
    // Populate cells from the snapshot.
    for (const sheet of snapshot.sheets) {
      const ws = newWb.getSheetByName(sheet.name)
      if (!ws) continue
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        const value = cell.formula ? `=${cell.formula}` : cell.value
        ws.getRange(addr).setValue(String(value))
      }
    }
    const sub = fresh.univerAPI.addEvent(
      fresh.univerAPI.Event.CommandExecuted,
      () => setDirty(true),
    )
    void sub
    setDirty(false)
  }, [])

  /** Read the current Univer cell state and build CellEdits for save. */
  const buildCellEdits = useCallback((): CellEdit[] => {
    const rt = runtimeRef.current
    if (!rt) return []
    const wb = rt.univerAPI.getActiveWorkbook()
    if (!wb) return []
    const edits: CellEdit[] = []
    for (const sheet of wb.getSheets()) {
      const sheetName = sheet.getSheetName()
      const rowCount = sheet.getMaxRows()
      const colCount = sheet.getMaxColumns()
      for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < colCount; col++) {
          const value = sheet.getRange(row, col, 1, 1).getValue()
          if (value === '' || value === null || value === undefined) continue
          // Build a CellState — the value as a string, formula if it starts with =
          const formula = typeof value === 'string' && value.startsWith('=') ? value.slice(1) : undefined
          const cellState: CellState = {
            value: formula ? '' : value,
            ...(formula ? { formula } : {}),
          }
          edits.push({
            sheetName,
            row,
            column: col,
            writeValue: true,
            cell: cellState,
          })
        }
      }
    }
    return edits
  }, [])

  /** Open a file from the user's local filesystem. */
  const handleOpenFile = useCallback(async (file: File) => {
    setStatus('Opening...')
    try {
      const bytes = await readFileBytes(file)
      const res = await openWorkbook({ fileName: file.name, fileBytes: bytes })
      sourceBytesRef.current = bytes
      fileNameRef.current = file.name
      loadSnapshot(res.snapshot, res.sheetNamesById)
      setStatus(`Opened ${file.name}`)
      setDirty(false)
    } catch (e) {
      setStatus(`Open failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }, [loadSnapshot])

  /** Save (or Save As) the workbook through the API. */
  const handleSave = useCallback(async (saveAs: boolean) => {
    const sourceBytes = sourceBytesRef.current
    if (!sourceBytes) {
      // New workbook — need to build from scratch. For now, show a message.
      setStatus('Nothing to save — open a file first')
      return
    }
    setStatus('Saving...')
    try {
      const edits = buildCellEdits()
      let fileName = fileNameRef.current
      if (saveAs) {
        const newName = window.prompt('Save as:', fileName)
        if (!newName) { setStatus('Save cancelled'); return }
        fileName = newName.endsWith('.xlsx') ? newName : `${newName}.xlsx`
      }
      const savedBytes = await saveWorkbook({ fileName, fileBytes: sourceBytes, cellEdits: edits })
      // Update the source bytes so subsequent saves are based on the saved file.
      sourceBytesRef.current = savedBytes
      fileNameRef.current = fileName
      // Offer the saved file as a download.
      const blob = new Blob([savedBytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = fileName
      a.click()
      URL.revokeObjectURL(url)
      setStatus(`Saved ${fileName}`)
      setDirty(false)
    } catch (e) {
      setStatus(`Save failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }, [buildCellEdits])

  return (
    <div style={{ height: 'calc(100vh - 64px)', display: 'flex', flexDirection: 'column', background: '#f5f6f8' }}>
      <header style={{ minHeight: 56, display: 'flex', alignItems: 'center', gap: 10, padding: '0 16px', background: '#fff', borderBottom: '1px solid #d8dde6' }}>
        <button onClick={() => onRoute('/office')} style={{ ...styles.button, padding: '7px 12px' }}>← Office</button>
        <strong style={{ flex: 1 }}>GenOffice Excel — {fileNameRef.current}</strong>
        <span style={{ opacity: 0.65 }}>{dirty ? '● Unsaved changes' : '✓ Saved'}</span>
        <button style={styles.button} onClick={() => fileInputRef.current?.click()}>Open</button>
        <button style={styles.button} onClick={() => handleSave(false)} disabled={!dirty}>Save</button>
        <button style={styles.button} onClick={() => handleSave(true)}>Save As</button>
        <input
          ref={fileInputRef}
          hidden
          type="file"
          accept=".xlsx,.csv,.xls"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleOpenFile(f); e.target.value = '' }}
        />
      </header>
      {status !== 'Ready' && (
        <div style={{ padding: '4px 16px', background: '#e8f0fe', fontSize: 13, color: '#1a56c4' }}>{status}</div>
      )}
      <div id="genoffice-web-excel" ref={containerRef} style={{ flex: 1, minHeight: 0 }} />
    </div>
  )
}
