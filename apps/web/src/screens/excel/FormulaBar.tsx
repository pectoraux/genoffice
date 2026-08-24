/**
 * GenOffice web Sheets — Formula Bar.
 *
 * Displays the active cell's formula (when present) or its value. Editing
 * and pressing Enter (or blur) commits through the runtime's commitFormula,
 * which writes via FRange.setValueForCell — firing `sheet.mutation.
 * set-range-values`, journaled by ExcelEditor's existing subscription through
 * cell-mutation-merge.ts. The browser never runs a second formula engine.
 *
 * The display mirrors the formula-priority invariant: a leading '=' is
 * stored as the formula body (without the '='), exactly like buildCellDataMatrix
 * in ExcelEditor. The recalc-echo merge in cell-mutation-merge.ts guarantees
 * the displayed formula cannot be silently converted back into a literal.
 */
import { useEffect, useRef, useState } from 'react'
import type { BrowserUniverRuntime } from '../../office/create-browser-univer'

interface FormulaBarProps {
  runtime: BrowserUniverRuntime | null
  /** Stamp that bumps whenever the selection changes — triggers a re-read. */
  selectionStamp: number
  onCommit: (text: string) => void
}

export function FormulaBar({ runtime, selectionStamp, onCommit }: FormulaBarProps) {
  const [text, setText] = useState<string>('')
  const [editing, setEditing] = useState(false)
  const inputRef = useRef<HTMLInputElement | null>(null)

  // Read the active cell's formula/value whenever the selection moves (and
  // never while the user is mid-edit, to avoid clobbering their draft).
  useEffect(() => {
    if (editing) return
    const wb = runtime?.univerAPI.getActiveWorkbook()
    const cell = wb?.getActiveSheet()?.getActiveCell() ?? null
    if (!cell) {
      setText('')
      return
    }
    const data = cell.getCellData()
    if (data?.f) {
      setText(`=${data.f}`)
    } else if (data?.v !== undefined && data?.v !== null) {
      setText(String(data.v))
    } else {
      setText('')
    }
  }, [runtime, selectionStamp, editing])

  const commit = () => {
    onCommit(text)
    setEditing(false)
  }

  return (
    <input
      ref={inputRef}
      className="excel-formula-bar"
      aria-label="Formula Bar"
      data-testid="excel-formula-bar"
      spellCheck={false}
      value={text}
      onFocus={() => setEditing(true)}
      onChange={(event) => setText(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault()
          commit()
          event.currentTarget.blur()
        } else if (event.key === 'Escape') {
          // Re-read to discard the draft.
          setEditing(false)
          event.currentTarget.blur()
        }
      }}
    />
  )
}
