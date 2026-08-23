import { useMemo, useState } from 'react'
import { styles } from '../styles'

type Cell = string
const COLS = 18
const ROWS = 40
const key = (r: number, c: number) => `${r}:${c}`
const colName = (c: number) => { let n = c + 1; let s = ''; while (n) { const x = (n - 1) % 26; s = String.fromCharCode(65 + x) + s; n = Math.floor((n - 1) / 26) } return s }

function evaluate(value: string, cells: Record<string, Cell>): string {
  if (!value.startsWith('=')) return value
  const m = value.match(/^=SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
  if (!m) return value
  const parse = (ref: string) => { const x = ref.match(/^([A-Z]+)(\d+)$/i)!; let c = 0; for (const ch of x[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; return { c: c - 1, r: Number(x[2]) - 1 } }
  const a = parse(m[1]); const b = parse(m[2]); let total = 0
  for (let r = a.r; r <= b.r; r++) for (let c = a.c; c <= b.c; c++) total += Number(cells[key(r, c)] ?? 0) || 0
  return String(total)
}

export function ExcelEditor({ onRoute }: { onRoute: (route: string) => void }) {
  const [cells, setCells] = useState<Record<string, Cell>>(() => {
    try { return JSON.parse(localStorage.getItem('genoffice-web-excel') ?? '{}') } catch { return {} }
  })
  const [selected, setSelected] = useState(key(0, 0))
  const [formula, setFormula] = useState('')
  const [saved, setSaved] = useState(true)

  const selectedValue = cells[selected] ?? ''
  const display = useMemo(() => evaluate(selectedValue, cells), [selectedValue, cells])

  const select = (r: number, c: number) => { const k = key(r, c); setSelected(k); setFormula(cells[k] ?? '') }
  const commit = (value: string) => { setCells((old) => ({ ...old, [selected]: value })); setSaved(false) }
  const save = () => { localStorage.setItem('genoffice-web-excel', JSON.stringify(cells)); setSaved(true) }
  const clear = () => { setCells({}); setFormula(''); setSaved(false) }

  return (
    <div style={{ height: 'calc(100vh - 64px)', display: 'flex', flexDirection: 'column', background: '#f5f6f8' }}>
      <header style={{ ...styles.header, flexShrink: 0 }}>
        <button style={styles.button} onClick={() => onRoute('/office')}>← Office</button>
        <strong style={{ marginLeft: 14, flex: 1 }}>Workbook</strong>
        <span style={{ opacity: .65, marginRight: 12 }}>{saved ? 'Saved' : 'Unsaved changes'}</span>
        <button style={styles.button} onClick={clear}>New</button>
        <button style={styles.button} onClick={save}>Save</button>
      </header>
      <div style={{ background: '#fff', borderBottom: '1px solid #d8dde6', padding: 8, display: 'flex', gap: 8 }}>
        <div style={{ minWidth: 72, padding: '7px 9px', border: '1px solid #d8dde6', borderRadius: 5, fontFamily: 'monospace' }}>{selected.split(':').map(Number).map((x, i) => i === 0 ? colName(x) : x + 1).join('')}</div>
        <input value={formula} onChange={(e) => setFormula(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') commit(formula) }} placeholder="Enter value or formula (e.g. =SUM(A1:A5))" style={{ flex: 1, border: '1px solid #d8dde6', borderRadius: 5, padding: '7px 10px' }} />
        <div style={{ minWidth: 100, padding: '7px 10px', border: '1px solid #d8dde6', borderRadius: 5, background: '#fafbfc' }}>={display}</div>
      </div>
      <div style={{ overflow: 'auto', flex: 1 }}>
        <table style={{ borderCollapse: 'collapse', background: '#fff', minWidth: 900 }}>
          <thead><tr><th style={{ position: 'sticky', top: 0, left: 0, zIndex: 3, width: 44, background: '#f0f2f5', border: '1px solid #d9dde5' }} />{Array.from({ length: COLS }, (_, c) => <th key={c} style={{ position: 'sticky', top: 0, zIndex: 2, width: 100, background: '#f0f2f5', border: '1px solid #d9dde5', fontWeight: 600 }}>{colName(c)}</th>)}</tr></thead>
          <tbody>{Array.from({ length: ROWS }, (_, r) => <tr key={r}><th style={{ position: 'sticky', left: 0, zIndex: 1, background: '#f0f2f5', border: '1px solid #d9dde5', fontWeight: 600 }}>{r + 1}</th>{Array.from({ length: COLS }, (_, c) => { const k = key(r, c); const active = k === selected; return <td key={c} onClick={() => select(r, c)} style={{ border: '1px solid #e1e4e9', padding: 0, minWidth: 100, height: 26, outline: active ? '2px solid #1677ff' : 'none', outlineOffset: -2 }}><input value={cells[k] ?? ''} onChange={(e) => { const v = e.target.value; setCells((old) => ({ ...old, [k]: v })); setSaved(false); if (k === selected) setFormula(v) }} onFocus={() => select(r, c)} style={{ width: '100%', height: 26, boxSizing: 'border-box', border: 0, padding: '3px 6px', outline: 0, background: 'transparent' }} /></td> })}</tr>)}</tbody>
        </table>
      </div>
    </div>
  )
}
