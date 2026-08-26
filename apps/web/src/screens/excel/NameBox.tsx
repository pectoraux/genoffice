/**
 * GenOffice web Sheets — Name Box.
 *
 * Echoes the active cell/range while idle; focusing it starts a draft, Enter
 * jumps to the typed A1/range (selecting it), Esc or blur cancels back to the
 * echo. Mirrors the desktop's NameBox in ExcelShell.tsx.
 */
import { useEffect, useState } from 'react'

export function NameBox({
  activeCellA1,
  onGoTo,
}: {
  activeCellA1: string
  onGoTo: (ref: string) => string | null
}) {
  const [draft, setDraft] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  // Keep the echoed value current when the selection moves while not editing.
  useEffect(() => {
    if (draft === null) setError(null)
  }, [activeCellA1, draft])

  return (
    <input
      className={`excel-name-box${error === null ? '' : ' invalid'}`}
      aria-label="Name Box"
      data-testid="excel-name-box"
      placeholder="A1"
      spellCheck={false}
      value={draft ?? activeCellA1}
      onFocus={(event) => {
        setDraft(activeCellA1 || 'A1')
        setError(null)
        event.target.select()
      }}
      onChange={(event) => {
        setDraft(event.target.value)
        setError(null)
      }}
      onBlur={() => {
        setDraft(null)
        setError(null)
      }}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          const failure = onGoTo(draft ?? activeCellA1)
          setError(failure)
          if (failure === null) {
            setDraft(null)
            event.currentTarget.blur()
          }
        } else if (event.key === 'Escape') {
          setDraft(null)
          setError(null)
          event.currentTarget.blur()
        }
      }}
    />
  )
}
