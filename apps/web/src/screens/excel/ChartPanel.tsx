/**
 * EXCEL-023 — the web Chart Design pane (desktop ChartPanels.tsx parity,
 * compact web subset). Create mode: pick a supported chart type for the
 * current selection (the gateway's shared chartDataFromValues parses the
 * range — header/category detection identical to the desktop ribbon).
 * Edit mode: edit the SELECTED chart's supported properties (title, the
 * six convertible types, legend, data labels, per-series colors) and
 * display its typed source ranges. Every change journals through the
 * canonical chartEdits family via recordChartEdit — never a direct
 * workbook mutation.
 */
import React, { useState } from 'react'

import type { ChartAdd } from '@genoffice/xlsx-gateway'
import { chartDataFromValues, type ChartGridValue } from '../../office/chart-domain'
import {
  convertibleChartType,
  useChartStoreVersion as useChartStore,
  type ChartEditingStore,
  liveChartSpec,
  recordChartEdit,
} from '../../office/sheet-charts'

const CREATE_TYPES: readonly { value: ChartAdd['chartType']; label: string }[] = [
  { value: 'column', label: 'Column' },
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'pie', label: 'Pie' },
  { value: 'doughnut', label: 'Doughnut' },
  { value: 'scatter', label: 'Scatter' },
  { value: 'radar', label: 'Radar' },
  { value: 'combo', label: 'Combo (column + line)' },
]

const CONVERTIBLE_TYPES: readonly { value: string; label: string }[] = [
  { value: 'column', label: 'Column' },
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'pie', label: 'Pie' },
  { value: 'doughnut', label: 'Doughnut' },
]

const LEGEND_POSITIONS: readonly { value: string; label: string }[] = [
  { value: 'right', label: 'Right' },
  { value: 'bottom', label: 'Bottom' },
  { value: 'top', label: 'Top' },
  { value: 'left', label: 'Left' },
  { value: 'none', label: 'None' },
]

const LABEL_MODES: readonly { value: string; label: string }[] = [
  { value: 'none', label: 'None' },
  { value: 'value', label: 'Value' },
  { value: 'percent', label: 'Percent' },
  { value: 'category-percent', label: 'Category + Percent' },
]

export interface ChartPanelSelectionInfo {
  /** The selected chart's live spec (pending edits already overlaid). */
  readonly key: string
  readonly isSession: boolean
}

export function ChartPanel({
  store,
  mode,
  selection,
  selectionValues,
  selectionRangeLabel,
  onInsert,
  onClose,
}: {
  readonly store: ChartEditingStore
  readonly mode: 'create' | 'edit'
  readonly selection: ChartPanelSelectionInfo | null
  /** The active range's values (create mode: parsed for chart data). */
  readonly selectionValues: readonly (readonly ChartGridValue[])[] | null
  readonly selectionRangeLabel: string
  readonly onInsert: (chart: ChartAdd) => void
  readonly onClose: () => void
}): React.JSX.Element {
  // Subscribes the pane to store changes (selection, pending edits) —
  // the returned version isn't rendered directly.
  useChartStore(store)
  const spec = selection !== null ? liveChartSpec(store, selection.key) : null
  const [createType, setCreateType] = useState<ChartAdd['chartType']>('column')
  const parsed =
    mode === 'create' && selectionValues !== null ? chartDataFromValues(selectionValues) : null

  if (mode === 'create') {
    return (
      <div className="chart-panel" data-testid="chart-panel-create">
        <header>
          <span>Insert Chart</span>
          <button aria-label="Close" onClick={onClose}>
            ✕
          </button>
        </header>
        <div className="chart-panel-section">
          <strong>Data range</strong>
          <code data-testid="chart-panel-range">{selectionRangeLabel || 'Select a data range first'}</code>
          {parsed !== null ? (
            <p className="chart-panel-hint" data-testid="chart-panel-parsed">
              {parsed.series.length} series, {parsed.categories.length} categories
            </p>
          ) : (
            <p className="chart-panel-hint chart-panel-error">
              The selection needs at least one numeric column.
            </p>
          )}
        </div>
        <div className="chart-panel-section">
          <strong>Chart type</strong>
          <select
            data-testid="chart-panel-type"
            value={createType}
            onChange={(event) => setCreateType(event.target.value as ChartAdd['chartType'])}
          >
            {CREATE_TYPES.map((entry) => (
              <option key={entry.value} value={entry.value}>
                {entry.label}
              </option>
            ))}
          </select>
        </div>
        <button
          className="chart-panel-primary"
          data-testid="chart-panel-insert"
          disabled={parsed === null}
          onClick={() => {
            if (parsed === null) return
            onInsert({
              chartType: createType,
              title:
                parsed.series.length === 1 && parsed.series[0] !== undefined
                  ? parsed.series[0].name
                  : 'Chart Title',
              series: parsed.series.map((series) => ({
                name: series.name,
                categories: [...parsed.categories],
                values: [...series.values],
              })),
            })
          }}
        >
          Insert Chart
        </button>
      </div>
    )
  }

  if (selection === null || spec === null || spec.isSession) {
    return (
      <div className="chart-panel" data-testid="chart-panel">
        <header>
          <span>Chart Design</span>
          <button aria-label="Close" onClick={onClose}>
            ✕
          </button>
        </header>
        <p className="chart-panel-hint">Select a chart on the sheet to design it.</p>
      </div>
    )
  }

  const chart = spec.chart
  if (!('chartTypes' in chart)) {
    return (
      <div className="chart-panel" data-testid="chart-panel">
        <header>
          <span>Chart Design</span>
          <button aria-label="Close" onClick={onClose}>
            ✕
          </button>
        </header>
        <p className="chart-panel-hint">Select a chart on the sheet to design it.</p>
      </div>
    )
  }
  const convertible = convertibleChartType(chart)
  const pending = store.edits.get(selection.key)
  return (
    <div className="chart-panel" data-testid="chart-panel">
      <header>
        <span>Chart Design</span>
        <button aria-label="Close" onClick={onClose}>
          ✕
        </button>
      </header>
      <div className="chart-panel-section">
        <strong>Chart title</strong>
        <input
          type="text"
          data-testid="chart-panel-title"
          key={chart.title}
          defaultValue={chart.title}
          maxLength={255}
          onBlur={(event) => {
            if (event.target.value !== chart.title) {
              recordChartEdit(store, selection.key, { title: event.target.value })
            }
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') (event.target as HTMLInputElement).blur()
          }}
        />
      </div>
      {convertible !== null && (
        <div className="chart-panel-section">
          <strong>Chart type</strong>
          <select
            data-testid="chart-panel-convert"
            value={pending?.chartType ?? convertible}
            onChange={(event) =>
              recordChartEdit(store, selection.key, {
                chartType: event.target.value as 'column',
              })
            }
          >
            {CONVERTIBLE_TYPES.map((entry) => (
              <option key={entry.value} value={entry.value}>
                {entry.label}
              </option>
            ))}
          </select>
        </div>
      )}
      {convertible === null && (
        <div className="chart-panel-section">
          <strong>Chart type</strong>
          <p className="chart-panel-hint">
            {chart.chartTypes.join(' + ')} — this family does not convert.
          </p>
        </div>
      )}
      <div className="chart-panel-section">
        <strong>Legend</strong>
        <select
          data-testid="chart-panel-legend"
          value={chart.legend ?? 'right'}
          onChange={(event) =>
            recordChartEdit(store, selection.key, {
              legend: event.target.value as 'right',
            })
          }
        >
          {LEGEND_POSITIONS.map((entry) => (
            <option key={entry.value} value={entry.value}>
              {entry.label}
            </option>
          ))}
        </select>
      </div>
      <div className="chart-panel-section">
        <strong>Data labels</strong>
        <select
          data-testid="chart-panel-labels"
          value={chart.dataLabels ?? 'none'}
          onChange={(event) =>
            recordChartEdit(store, selection.key, {
              dataLabels: event.target.value as 'value',
            })
          }
        >
          {LABEL_MODES.map((entry) => (
            <option key={entry.value} value={entry.value}>
              {entry.label}
            </option>
          ))}
        </select>
      </div>
      {chart.series.length > 0 && chart.chartTypes.length === 1 && (
        <div className="chart-panel-section">
          <strong>Series colors</strong>
          {chart.series.map((series, index) => (
            <label key={index} className="chart-panel-color-row">
              <span>{series.name || `Series ${index + 1}`}</span>
              <input
                type="color"
                data-testid={`chart-panel-series-color-${index}`}
                value={series.color ?? '#4472c4'}
                onChange={(event) =>
                  recordChartEdit(store, selection.key, {
                    seriesColors: { [String(index)]: event.target.value },
                  })
                }
              />
            </label>
          ))}
        </div>
      )}
      <div className="chart-panel-section">
        <strong>Source ranges</strong>
        {chart.series.map((series, index) => (
          <div key={index} className="chart-panel-ref" data-testid={`chart-panel-ref-${index}`}>
            <span>{series.name || `Series ${index + 1}`}</span>
            <code>{series.valuesRef ?? 'cached values (no range reference)'}</code>
          </div>
        ))}
      </div>
      <p className="chart-panel-hint">
        Edits journal through the canonical chartEdits save family — the file&apos;s chart XML
        is written by the xlsx-gateway on save.
      </p>
    </div>
  )
}
