/**
 * GenOffice web Sheets — Name Manager dialog (EXCEL-025).
 *
 * Minimal desktop parity (apps/sheets NameManagerDialog.tsx is the frozen
 * reference): a list of the modeled names (name / refers-to / scope), an
 * add/edit form (name + refers-to; scope is chosen at CREATION only —
 * desktop semantics), and per-row delete. All actions route through the
 * PUBLIC Univer defined-name facade in ExcelEditor (the builder validates:
 * duplicates, reference lookalikes, sheet/table/function-name conflicts),
 * never through XML. The dialog itself is pure presentation — it renders
 * the rows it is given and reports actions upward.
 */
import { useState } from 'react'

export interface DefinedNameRow {
  readonly name: string
  readonly ref: string
  /** null = workbook scope. */
  readonly scopeSheetId: string | null
  readonly scopeLabel: string
}

export type DefinedNameAction =
  | { kind: 'add'; name: string; ref: string; sheetId: string | null }
  | {
      kind: 'update'
      originalName: string
      scopeSheetId: string | null
      name: string
      ref: string
    }
  | { kind: 'remove'; name: string; scopeSheetId: string | null }

export function NameManagerDialog({
  names,
  sheets,
  locked,
  lockedReason,
  onAction,
  onClose,
}: {
  readonly names: readonly DefinedNameRow[]
  readonly sheets: readonly { id: string; name: string }[]
  /// namesLocked: the workbook's <definedNames> could not be parsed — every
  /// edit is refused (a declarative save would drop the unparseable names).
  readonly locked: boolean
  readonly lockedReason: string | null
  /// Returns an error message, or null on success.
  readonly onAction: (action: DefinedNameAction) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const [selected, setSelected] = useState<DefinedNameRow | null>(null)
  const [name, setName] = useState('')
  const [ref, setRef] = useState('')
  const [scope, setScope] = useState('')
  const [error, setError] = useState<string | null>(null)

  const pick = (row: DefinedNameRow): void => {
    setSelected(row)
    setName(row.name)
    setRef(row.ref)
    setError(null)
  }
  const run = (action: DefinedNameAction): void => {
    const failure = onAction(action)
    setError(failure)
    if (failure === null) {
      setSelected(null)
      setName('')
      setRef('')
    }
  }

  return (
    <div className="name-manager-backdrop" onClick={onClose}>
      <div
        className="name-manager-dialog"
        role="dialog"
        aria-label="Name Manager"
        data-testid="name-manager-dialog"
        onClick={(event) => event.stopPropagation()}
      >
        <header>
          <span>Name Manager</span>
          <button onClick={onClose} aria-label="Close Name Manager">
            ✕
          </button>
        </header>
        <section className="name-manager-body">
          {locked && (
            <p className="name-manager-error" data-testid="name-manager-locked">
              {lockedReason ?? 'This workbook\u2019s defined names cannot be edited here.'}
            </p>
          )}
          <table className="name-manager-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Refers To</th>
                <th>Scope</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {names.length === 0 && (
                <tr>
                  <td colSpan={4} className="name-manager-empty">
                    The workbook defines no editable names.
                  </td>
                </tr>
              )}
              {names.map((row) => (
                <tr
                  key={`${row.scopeSheetId ?? ''}!${row.name}`}
                  data-testid="name-manager-row"
                  className={
                    selected?.name === row.name && selected.scopeSheetId === row.scopeSheetId
                      ? 'selected'
                      : ''
                  }
                  onClick={() => pick(row)}
                >
                  <td>{row.name}</td>
                  <td>{row.ref}</td>
                  <td>{row.scopeLabel}</td>
                  <td>
                    <button
                      data-tip="Delete name"
                      aria-label={`Delete ${row.name}`}
                      disabled={locked}
                      onClick={(event) => {
                        event.stopPropagation()
                        run({ kind: 'remove', name: row.name, scopeSheetId: row.scopeSheetId })
                      }}
                    >
                      🗑
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="name-manager-grid">
            <label>
              Name
              <input
                data-testid="name-manager-name"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="MyRange"
                disabled={locked}
              />
            </label>
            <label>
              Refers To
              <input
                data-testid="name-manager-ref"
                value={ref}
                onChange={(e) => setRef(e.target.value)}
                placeholder="=Sheet1!$A$1:$B$9"
                disabled={locked}
              />
            </label>
            {selected === null && (
              <label>
                Scope
                <select
                  data-testid="name-manager-scope"
                  value={scope}
                  onChange={(e) => setScope(e.target.value)}
                  disabled={locked}
                >
                  <option value="">Workbook</option>
                  {sheets.map((sheet) => (
                    <option key={sheet.id} value={sheet.id}>
                      {sheet.name}
                    </option>
                  ))}
                </select>
              </label>
            )}
          </div>
          {error && <p className="name-manager-error">{error}</p>}
        </section>
        <footer className="name-manager-actions">
          <button className="secondary" onClick={onClose}>
            Close
          </button>
          {selected !== null && (
            <button
              className="secondary"
              onClick={() => {
                setSelected(null)
                setName('')
                setRef('')
                setError(null)
              }}
            >
              New
            </button>
          )}
          <button
            className="primary-action"
            data-testid="name-manager-apply"
            disabled={locked || name.trim() === '' || ref.trim() === ''}
            onClick={() =>
              run(
                selected === null
                  ? { kind: 'add', name: name.trim(), ref: ref.trim(), sheetId: scope || null }
                  : {
                      kind: 'update',
                      originalName: selected.name,
                      scopeSheetId: selected.scopeSheetId,
                      name: name.trim(),
                      ref: ref.trim(),
                    },
              )
            }
          >
            {selected === null ? 'Add' : 'Update'}
          </button>
        </footer>
      </div>
    </div>
  )
}
