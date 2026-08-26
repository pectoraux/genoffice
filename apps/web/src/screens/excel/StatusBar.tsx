/**
 * GenOffice web Sheets — Status Bar.
 *
 * Mirrors the desktop's custom status bar (ExcelShell.tsx:535-567): status
 * message on the left, zoom controls (− / slider / + / percent) on the right.
 * Sheet tabs themselves render inside the Univer grid's own footer
 * (footer.sheetBar: true) — the desktop does the same.
 */
import type { ExcelRuntimeApi, ExcelRuntimeState } from './useExcelRuntime'

const NULL_STATE: ExcelRuntimeState = {
  ready: false,
  activeCellA1: '',
  selectionFormat: {
    bold: false,
    italic: false,
    underline: false,
    fontFamily: null,
    fontSize: null,
    fontColor: null,
    fillColor: null,
    hAlign: null,
    vAlign: null,
    wrap: false,
  },
  zoomPercent: 100,
  gridlinesVisible: true,
  canUndo: false,
  canRedo: false,
  seq: 0,
}

export function StatusBar({
  api,
  status,
  isError,
}: {
  api: ExcelRuntimeApi | null
  status: string
  isError: boolean
}) {
  const s = api?.state ?? NULL_STATE
  const z = Math.min(400, Math.max(50, s.zoomPercent))
  return (
    <footer className="excel-statusbar" data-testid="excel-statusbar">
      <span
        className={`excel-status-msg${isError ? ' error' : ''}`}
        role="status"
        aria-live="polite"
      >
        {status}
      </span>
      <span className="stat">{s.ready ? 'Ready' : ''}</span>
      <div className="excel-zoom" data-testid="excel-zoom">
        <button
          type="button"
          aria-label="Zoom out"
          title="Zoom out"
          onClick={() => api?.zoomOut()}
          disabled={!s.ready}
        >
          −
        </button>
        <input
          type="range"
          min={50}
          max={400}
          step={10}
          aria-label="Zoom slider"
          value={z}
          disabled={!s.ready}
          onChange={(e) => api?.setZoom(Number(e.target.value) / 100)}
        />
        <button
          type="button"
          aria-label="Zoom in"
          title="Zoom in"
          onClick={() => api?.zoomIn()}
          disabled={!s.ready}
        >
          +
        </button>
        <span className="excel-zoom-value">{s.zoomPercent}%</span>
      </div>
    </footer>
  )
}
