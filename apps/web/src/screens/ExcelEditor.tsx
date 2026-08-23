import { useEffect, useRef, useState } from 'react'
import type { FWorkbook } from '@univerjs/core/facade'
import { createBrowserUniver, type BrowserUniverRuntime } from '../office/create-browser-univer'

export function ExcelEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const runtimeRef = useRef<BrowserUniverRuntime | null>(null)
  const workbookRef = useRef<FWorkbook | null>(null)
  const [dirty, setDirty] = useState(false)

  useEffect(() => {
    if (!containerRef.current) return
    const runtime = createBrowserUniver('genoffice-web-excel')
    runtimeRef.current = runtime
    const workbook = runtime.univerAPI.createWorkbook({
      id: 'genoffice-web-workbook',
      name: 'Workbook',
    })
    workbookRef.current = workbook

    const sheet = workbook.getActiveSheet()
    if (sheet) {
      sheet.getRange('A1:B4').setValues([
        ['GenOffice', 'Excel'],
        ['Browser editor', 'Univer runtime'],
        ['Formula', '=SUM(B1:B3)'],
        ['Ready', 'Edit me'],
      ])
      sheet.getRange('A1:B1').setFontWeight('bold')
      sheet.getRange('A1:B4').setWrap(true)
    }

    const subscription = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      () => setDirty(true),
    )

    return () => {
      subscription.dispose()
      runtime.univer.dispose()
      runtimeRef.current = null
      workbookRef.current = null
    }
  }, [])

  return (
    <div style={{ height: 'calc(100vh - 64px)', display: 'flex', flexDirection: 'column', background: '#f5f6f8' }}>
      <header style={{ minHeight: 56, display: 'flex', alignItems: 'center', gap: 10, padding: '0 16px', background: '#fff', borderBottom: '1px solid #d8dde6' }}>
        <button onClick={() => onRoute('/office')} style={{ padding: '7px 12px', border: '1px solid #d8dde6', borderRadius: 6, background: '#fff', cursor: 'pointer' }}>← Office</button>
        <strong style={{ flex: 1 }}>GenOffice Excel</strong>
        <span style={{ opacity: 0.65 }}>{dirty ? 'Unsaved changes' : 'Saved'}</span>
      </header>
      <div id="genoffice-web-excel" ref={containerRef} style={{ flex: 1, minHeight: 0 }} />
    </div>
  )
}
