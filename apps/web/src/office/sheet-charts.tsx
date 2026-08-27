/**
 * EXCEL-023 — the web chart visual layer (Insert → Chart / chart edit).
 *
 * Desktop parity (apps/sheets WorkbookVisuals.tsx — frozen reference):
 * charts render as THIS APP'S OWN SVG React components floated over the
 * Univer grid through the PUBLIC facade pair registerComponent +
 * addFloatDomToRange (the desktop's exact rendering architecture — Univer
 * 0.25.1 ships no chart plugin). The browser is a thin typed client: it
 * renders the gateway's canonical ChartVisualState, journals semantic
 * changes as typed ChartStateEdit overlays keyed by chartPath (merged at
 * save into the canonical chartEdits wire family), journals geometry
 * through the EXCEL-022 visualEdits family, and session creations through
 * visualAdditions.chart. No OOXML, no relationships, no XML — the
 * xlsx-gateway owns every byte.
 *
 * Rendering semantics are ported from the desktop renderer so both
 * surfaces draw the same chart the same way: Excel-like value-axis
 * auto-scale, gap-width bar layout, stacked/percentStacked shapes, pie
 * slice explosion + leader labels, radar rings, scatter auto-bounds, and
 * the combo bar+line split with a secondary axis.
 */
import React, { useEffect, useRef, useState, useSyncExternalStore } from 'react'

import type {
  ChartAdd,
  DrawingAnchor,
  SheetChartInfo,
  SheetVisualAddition,
  WorkbookChartEdit,
  WorkbookVisualEdit,
} from '@genoffice/xlsx-gateway'
import type { ChartStateEdit, ChartSeriesVisualState, ChartVisualState } from './chart-domain'
import {
  applyChartStateEdit,
  chartCategoryFormat,
  formatCategoryLabel,
  scatterAxisBounds,
  valueAxisScale,
} from './chart-domain'
// Side-effect import — loads @univerjs/sheets-drawing-ui's facade module,
// which declares the PUBLIC addFloatDomToRange/getFloatDomById mixin on
// FWorksheet (the same augmentation pattern the image layer uses for
// FOverGridImage). Without it the augmented worksheet type is not visible.
import '@univerjs/sheets-drawing-ui/facade'

// ── Wire/state types ─────────────────────────────────────────────────

export const EMU_PER_PX = 9525

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX)
}

/// The journal's semantic edit value — the canonical domain's mutable
/// ChartStateEdit (assignable to the wire's readonly WorkbookChartEdit at
/// collect time; the reverse would not typecheck by design).
export type ChartEditData = ChartStateEdit

/** One file-native chart, keyed by its canonical locator id. */
export interface FileChartEntry {
  readonly sheetName: string
  readonly info: SheetChartInfo
}

/** One session-created chart (journal payload — splice on delete). */
export interface SessionChartAdd {
  readonly id: string
  readonly sheetName: string
  readonly anchor: DrawingAnchor
  readonly chart: ChartAdd
}

export function fileChartId(drawingPath: string, drawingIndex: number): string {
  return `file-chart:${drawingPath}#${drawingIndex}`
}

// ── Editing store (shared by the floating frames and the design pane) ──

export interface ChartEditingStore {
  /** File charts by locator id (seeded on open, merged after save). */
  readonly fileCharts: Map<string, FileChartEntry>
  /** Session creations by id. */
  readonly sessionAdds: Map<string, SessionChartAdd>
  /** Pending semantic edits keyed by chartPath (file charts) or session id. */
  readonly edits: Map<string, ChartEditData>
  /** Geometry the user moved/resized this session, keyed by locator id. */
  readonly dirty: Set<string>
  /** Deleted file charts (locator ids) — emitted as visualEdits removes. */
  readonly removals: Set<string>
  /** The selected chart's edit key (chartPath or session id). */
  selection: string | null
  subscribe(listener: () => void): () => void
  getSnapshot(): { version: number }
  bump(): void
}

export function createChartEditingStore(): ChartEditingStore {
  const listeners = new Set<() => void>()
  let snapshot: { version: number } = { version: 0 }
  return {
    fileCharts: new Map(),
    sessionAdds: new Map(),
    edits: new Map(),
    dirty: new Set(),
    removals: new Set(),
    selection: null,
    subscribe(listener: () => void) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    getSnapshot() {
      return snapshot
    },
    bump() {
      snapshot = { version: snapshot.version + 1 }
      for (const listener of listeners) listener()
    },
  }
}

/// Re-renders the caller once per store change (frames, panel, editor).
export function useChartStoreVersion(store: ChartEditingStore): number {
  // The snapshot object stays referentially stable between bumps, so
  // useSyncExternalStore re-renders exactly once per store change.
  return useSyncExternalStore(store.subscribe, store.getSnapshot, store.getSnapshot).version
}

// ── Journal semantics (desktop edit-journal recordChartEdit parity) ──

/// Merges a semantic edit into the pending journal for one chart. A full
/// series replacement invalidates earlier per-index series edits and
/// per-series colors (their indices refer to dropped series) — the same
/// collapse the desktop performs.
export function recordChartEdit(store: ChartEditingStore, key: string, edit: ChartEditData): void {
  let previous = store.edits.get(key)
  if (edit.seriesSet !== undefined && previous !== undefined) {
    const {
      series: _series,
      seriesColors: _colors,
      pointColors: _points,
      pointExplosions: _explosions,
      ...rest
    } = previous
    previous = rest
  }
  store.edits.set(key, {
    ...previous,
    ...edit,
    ...(previous?.seriesColors || edit.seriesColors
      ? { seriesColors: { ...previous?.seriesColors, ...edit.seriesColors } }
      : {}),
    ...(previous?.pointColors || edit.pointColors
      ? {
          pointColors: mergeNestedRecords(
            previous?.pointColors as { [series: string]: { [point: string]: string } } | undefined,
            edit.pointColors as { [series: string]: { [point: string]: string } } | undefined,
          ),
        }
      : {}),
    ...(previous?.pointExplosions || edit.pointExplosions
      ? { pointExplosions: { ...previous?.pointExplosions, ...edit.pointExplosions } }
      : {}),
    ...(previous?.axisTitles || edit.axisTitles
      ? { axisTitles: { ...previous?.axisTitles, ...edit.axisTitles } }
      : {}),
    ...(previous?.valueAxis || edit.valueAxis
      ? { valueAxis: { ...previous?.valueAxis, ...edit.valueAxis } }
      : {}),
    ...(previous?.series || edit.series
      ? { series: mergeSeriesEdits(previous?.series, edit.series) }
      : {}),
  })
  store.bump()
}

function mergeNestedRecords(
  previous: Record<string, Record<string, string>> | undefined,
  patch: Record<string, Record<string, string>> | undefined,
): Record<string, Record<string, string>> | undefined {
  if (previous === undefined && patch === undefined) return undefined
  const merged: Record<string, Record<string, string>> = { ...(previous ?? {}) }
  for (const [seriesKey, points] of Object.entries(patch ?? {})) {
    merged[seriesKey] = { ...(merged[seriesKey] ?? {}), ...points }
  }
  return merged
}

function mergeSeriesEdits(
  previous: readonly { index: number }[] | undefined,
  patch: readonly { index: number }[] | undefined,
): { index: number }[] | undefined {
  if (previous === undefined && patch === undefined) return undefined
  const byIndex = new Map<number, { index: number }>()
  for (const entry of previous ?? []) byIndex.set(entry.index, entry)
  for (const entry of patch ?? []) byIndex.set(entry.index, entry)
  return [...byIndex.values()].sort((left, right) => left.index - right.index)
}

/// The pending semantic edits as the canonical wire family: file charts
/// keyed by their real chartPath; a session chart that was DELETED before
/// save contributes nothing (its journal entry dies with the visual).
export function collectChartEdits(store: ChartEditingStore): WorkbookChartEdit[] {
  const edits: WorkbookChartEdit[] = []
  for (const [key, edit] of store.edits.entries()) {
    const entry = store.fileCharts.get(key)
    if (entry === undefined) continue
    if (store.removals.has(key)) continue
    // Fail closed: a chart without a real chartPath (merged from a locator
    // block that omitted it — older host) is never targeted by a guessed
    // part path.
    if (entry.info.chartPath === '') continue
    edits.push({ chartPath: entry.info.chartPath, ...edit })
  }
  return edits
}

/// Session creations as the canonical wire family (deleted session charts
/// were spliced out of the journal, so they are never persisted). Typed as
/// SheetVisualAddition — the chart-only shape satisfies the family (image
/// is optional, exactly one payload rides).
export function collectChartAdditions(store: ChartEditingStore): SheetVisualAddition[] {
  const additions: SheetVisualAddition[] = []
  for (const add of store.sessionAdds.values()) {
    additions.push({ sheetName: add.sheetName, anchor: add.anchor, chart: add.chart })
  }
  return additions
}

/// Geometry edits for moved/resized/deleted file charts, read against the
/// LIVE anchor state — two moves collapse into one final-state edit.
export function collectChartVisualEdits(store: ChartEditingStore): WorkbookVisualEdit[] {
  const edits: WorkbookVisualEdit[] = []
  for (const id of store.removals) {
    const entry = store.fileCharts.get(id)
    if (entry === undefined) continue
    edits.push({
      drawingPath: entry.info.drawingPath,
      drawingIndex: entry.info.drawingIndex,
      remove: true,
    })
  }
  for (const id of store.dirty) {
    if (store.removals.has(id)) continue
    const entry = store.fileCharts.get(id)
    if (entry === undefined) continue
    edits.push({
      drawingPath: entry.info.drawingPath,
      drawingIndex: entry.info.drawingIndex,
      anchor: entry.info.anchor,
    })
  }
  return edits
}

/// A chart converts only when it has exactly one supported plot type —
/// desktop convertibleType parity (scatter/combos fail closed up front).
export function convertibleChartType(chart: ChartVisualStateLike): string | null {
  if (chart.chartTypes.length !== 1) return null
  const type = chart.chartTypes[0]
  if (type === undefined) return null
  if (type === 'barChart') return chart.barDirection === 'bar' ? 'bar' : 'column'
  if (type === 'lineChart') return 'line'
  if (type === 'areaChart') return 'area'
  if (type === 'pieChart') return 'pie'
  if (type === 'doughnutChart') return 'doughnut'
  return null
}

type ChartVisualStateLike = {
  readonly chartTypes: readonly string[]
  readonly barDirection?: string | undefined
}

// ── Series palette (desktop chartColors parity) ──────────────────────

const CHART_COLORS = [
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
  '#264478',
  '#9e480e',
  '#636363',
  '#997300',
]

type SeriesLike = {
  readonly color?: string | undefined
  readonly pointColors?: readonly { readonly index: number; readonly color: string }[] | undefined
}

function seriesColor(series: SeriesLike | undefined, index: number): string {
  return series?.color ?? CHART_COLORS[index % CHART_COLORS.length] ?? '#4472c4'
}

function pieSliceColor(series: SeriesLike, index: number): string {
  return (
    series.pointColors?.find((point) => point.index === index)?.color ??
    CHART_COLORS[index % CHART_COLORS.length] ??
    '#4472c4'
  )
}

// ── Shared label/format helpers (desktop parity, numfmt via domain) ──

function truncateLabel(value: string, maximumLength = 14): string {
  return value.length > maximumLength ? `${value.slice(0, maximumLength)}…` : value
}

function categoryLabelBudget(count: number): number {
  return Math.max(5, Math.min(16, Math.floor(86 / Math.max(1, count))))
}

function formatAxisValue(value: number, numberFormat: string | undefined): string {
  if (numberFormat && numberFormat !== 'General' && !numberFormat.includes('%')) {
    try {
      const text = formatCategoryLabel(String(value), numberFormat)
      if (text !== String(value)) return text
    } catch {
      // fall through to the plain rendering
    }
  }
  if (numberFormat?.includes('%')) return `${Math.round(value * 100)}%`
  const magnitude = Math.abs(value)
  if (magnitude >= 1e9) return `${(value / 1e9).toFixed(1)}B`
  if (magnitude >= 1e6) return `${(value / 1e6).toFixed(1)}M`
  const clean = Number(value.toPrecision(12))
  return clean.toLocaleString('en-US', { maximumFractionDigits: 4 })
}

function formatLabelValue(
  value: number,
  formatCode: string | undefined,
  numberFormat: string | undefined,
): string {
  if (formatCode === undefined || formatCode === '') return formatAxisValue(value, numberFormat)
  const decimals = /0\.(0+)/.exec(formatCode)?.[1]?.length ?? 0
  if (formatCode.includes('%')) return `${(value * 100).toFixed(decimals)}%`
  const fixed = value.toFixed(decimals)
  if (!formatCode.includes(',')) return fixed
  const [whole, fraction] = fixed.split('.')
  const grouped = Number(whole).toLocaleString('en-US')
  return fraction ? `${grouped}.${fraction}` : grouped
}

function formatScatterTick(value: number, format: string | undefined): string {
  if (format !== undefined && format !== 'General') {
    return formatCategoryLabel(String(value), format)
  }
  if (Number.isInteger(value) || Math.abs(value) >= 1e4) return formatAxisValue(value, undefined)
  return String(Number(value.toPrecision(4)))
}

type AxisScaleInput = {
  min?: number | undefined
  max?: number | undefined
  majorUnit?: number | undefined
  numFmt?: string | undefined
  title?: string | undefined
  majorGridlines?: boolean | undefined
  hidden?: boolean | undefined
}

// ── SVG chart renderers (desktop rendering semantics) ────────────────

/// The canonical domain's series state — the renderers draw the exact
/// typed model the gateway emits (mutable arrays per the domain).
type ChartSeries = ChartSeriesVisualState

type AxisTitles = { category?: string | null; value?: string | null } | undefined
type DataLabels = 'none' | 'value' | 'percent' | 'category-percent' | undefined
type LabelPosition = 'center' | 'inside-end' | 'outside-end' | undefined
type Grouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard' | undefined

function AxisTitleTexts({
  bottom,
  left,
}: {
  readonly bottom?: string | null | undefined
  readonly left?: string | null | undefined
}): React.JSX.Element {
  return (
    <g>
      {bottom ? (
        <text x="320" y="317" textAnchor="middle" className="chart-axis-title">
          {truncateLabel(bottom, 60)}
        </text>
      ) : null}
      {left ? (
        <text
          x="12"
          y="155"
          transform="rotate(-90 12 155)"
          textAnchor="middle"
          className="chart-axis-title"
        >
          {truncateLabel(left, 40)}
        </text>
      ) : null}
    </g>
  )
}

function VerticalAxis({
  minimum = 0,
  maximum,
  ticks,
  numberFormat,
  showGridlines = true,
}: {
  readonly minimum?: number
  readonly maximum: number
  readonly ticks?: readonly number[] | undefined
  readonly numberFormat: string | undefined
  readonly showGridlines?: boolean
}): React.JSX.Element {
  const span = maximum - minimum || 1
  const list =
    ticks && ticks.length >= 2
      ? ticks
      : [0, 0.25, 0.5, 0.75, 1].map((fraction) => minimum + fraction * span)
  return (
    <g>
      {list.map((tick, index) => {
        const y = 280 - ((tick - minimum) / span) * 240
        return (
          <g key={index}>
            {(showGridlines || index === 0) && (
              <line x1="58" y1={y} x2="580" y2={y} stroke="#e3e3e3" strokeWidth="1" />
            )}
            <text x="54" y={y + 4} textAnchor="end" className="chart-axis-label">
              {formatAxisValue(tick, numberFormat)}
            </text>
          </g>
        )
      })}
    </g>
  )
}

function SeriesLegend({
  seriesList,
}: {
  readonly seriesList: readonly ChartSeries[]
}): React.JSX.Element {
  return (
    <div className="chart-legend">
      {seriesList.map((series, index) => (
        <span key={index}>
          <i style={{ background: seriesColor(series, index) }} />
          {truncateLabel(series.name || `Series ${index + 1}`, 18)}
        </span>
      ))}
    </div>
  )
}

function linePoints(values: readonly number[], maximum: number, minimum = 0): string {
  const count = Math.max(1, values.length - 1)
  const span = maximum - minimum || 1
  return values
    .map(
      (value, index) =>
        `${60 + (index / count) * 500},${280 - Math.max(0, Math.min(1, (value - minimum) / span)) * 240}`,
    )
    .join(' ')
}

function BarChartSvg({
  seriesList,
  lineSeries,
  isHorizontal,
  axisTitles,
  dataLabels,
  dataLabelPosition,
  dataLabelFormat,
  grouping,
  gridlines,
  valueAxis,
  secondaryAxis,
  gapWidthPct,
  categoryFormat,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly lineSeries?: ChartSeries | undefined
  readonly isHorizontal: boolean
  readonly axisTitles?: AxisTitles
  readonly dataLabels?: DataLabels
  readonly dataLabelPosition?: LabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly grouping?: Grouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: AxisScaleInput | undefined
  readonly secondaryAxis?: AxisScaleInput | undefined
  readonly gapWidthPct?: number | undefined
  readonly categoryFormat?: string | undefined
}): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="chart-error">Chart has no cached data.</div>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const visibleCount = Math.min(primary.values.length, 48)
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  const categoryTotal = (index: number): number =>
    seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0)
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : valueAxisScale(
        isStacked
          ? Math.max(...Array.from({ length: visibleCount }, (_, index) => categoryTotal(index)), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        valueAxis,
      )
  const span = bounds.max - bounds.min
  const norm = (value: number): number => Math.max(0, Math.min(1, (value - bounds.min) / span))
  const segment = (seriesIndex: number, index: number): number => {
    const value = Math.max(0, seriesList[seriesIndex]?.values[index] ?? 0)
    if (!isPercent) return Math.min(1, value / (bounds.max || 1))
    const total = categoryTotal(index)
    return total === 0 ? 0 : value / total
  }
  const axisNumberFormat = isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)
  const gap = (gapWidthPct ?? 150) / 100

  if (isHorizontal) {
    const rowHeight = 280 / visibleCount
    const barHeight = Math.max(
      3,
      isStacked ? rowHeight / (1 + gap) : rowHeight / (seriesList.length + gap),
    )
    const groupTop = (index: number): number =>
      14 + rowHeight * index + (rowHeight - barHeight * (isStacked ? 1 : seriesList.length)) / 2
    const plotX = (value: number): number => 158 + norm(value) * 390
    return (
      <svg className="chart-svg" viewBox="0 0 600 320" role="img">
        {bounds.ticks.map((tick, index) => (
          <g key={`vt-${index}`}>
            {gridlines !== false && (
              <line
                x1={plotX(tick)}
                y1="12"
                x2={plotX(tick)}
                y2="298"
                stroke="#e3e3e3"
                strokeWidth="1"
              />
            )}
            <text x={plotX(tick)} y="10" textAnchor="middle" className="chart-axis-label">
              {formatAxisValue(tick, axisNumberFormat)}
            </text>
          </g>
        ))}
        {Array.from({ length: visibleCount }, (_, index) => {
          let cursor = 0
          return (
            <g key={`${categories[index] ?? index}-${index}`}>
              <text x="148" y={26 + rowHeight * index} textAnchor="end">
                {truncateLabel(categories[index] ?? String(index + 1))}
              </text>
              {seriesList.map((series, seriesIndex) => {
                const share = isStacked
                  ? segment(seriesIndex, index)
                  : norm(series.values[index] ?? bounds.min)
                const x = 158 + (isStacked ? cursor * 390 : 0)
                if (isStacked) cursor += share
                return (
                  <rect
                    key={seriesIndex}
                    x={x}
                    y={groupTop(index) + (isStacked ? 0 : barHeight * seriesIndex)}
                    width={share * 390}
                    height={barHeight}
                    fill={seriesColor(series, seriesIndex)}
                  />
                )
              })}
              {dataLabels === 'value' &&
                !isStacked &&
                (() => {
                  const width = norm(primary.values[index] ?? bounds.min) * 390
                  const inside =
                    dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
                  const x =
                    dataLabelPosition === 'center'
                      ? 158 + width / 2
                      : dataLabelPosition === 'inside-end'
                        ? 154 + width
                        : 164 + width
                  return (
                    <text
                      x={x}
                      y={groupTop(index) + barHeight / 2 + 3}
                      textAnchor={
                        dataLabelPosition === 'center'
                          ? 'middle'
                          : dataLabelPosition === 'inside-end'
                            ? 'end'
                            : 'start'
                      }
                      className="chart-axis-label"
                      {...(inside ? { fill: '#fff' } : {})}
                    >
                      {formatLabelValue(
                        primary.values[index] ?? 0,
                        dataLabelFormat,
                        primary.numberFormat,
                      )}
                    </text>
                  )
                })()}
            </g>
          )
        })}
        {visibleCount < primary.values.length && (
          <text x="580" y="14" textAnchor="end" className="chart-axis-label" opacity="0.75">
            +{primary.values.length - visibleCount} more
          </text>
        )}
        <AxisTitleTexts bottom={axisTitles?.value} left={axisTitles?.category} />
      </svg>
    )
  }

  const columnWidth = 480 / visibleCount
  const barWidth = Math.max(
    2,
    isStacked ? columnWidth / (1 + gap) : columnWidth / (seriesList.length + gap),
  )
  const groupWidth = barWidth * (isStacked ? 1 : seriesList.length)
  const groupLeft = (index: number): number =>
    62 + columnWidth * index + (columnWidth - groupWidth) / 2
  const lineScale = lineSeries
    ? valueAxisScale(Math.max(...lineSeries.values, 0), secondaryAxis)
    : undefined
  const points =
    lineSeries && lineScale
      ? lineSeries.values
          .slice(0, visibleCount)
          .map(
            (value, index) =>
              `${groupLeft(index) + groupWidth / 2},${
                280 -
                Math.max(
                  0,
                  Math.min(1, (value - lineScale.min) / (lineScale.max - lineScale.min || 1)),
                ) *
                  240
              }`,
          )
          .join(' ')
      : undefined
  const showValueLabels =
    !isStacked &&
    (dataLabels === 'value' ||
      (dataLabels === undefined && visibleCount <= 12 && seriesList.length === 1))
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={axisNumberFormat}
        showGridlines={gridlines !== false}
      />
      {Array.from({ length: visibleCount }, (_, index) => {
        let cursor = 0
        return (
          <g key={`${categories[index] ?? index}-${index}`}>
            {seriesList.map((series, seriesIndex) => {
              const share = isStacked
                ? segment(seriesIndex, index)
                : norm(series.values[index] ?? bounds.min)
              const height = share * 240
              const y = 280 - (isStacked ? (cursor + share) * 240 : height)
              if (isStacked) cursor += share
              return (
                <rect
                  key={seriesIndex}
                  x={groupLeft(index) + (isStacked ? 0 : barWidth * seriesIndex)}
                  y={y}
                  width={barWidth}
                  height={height}
                  fill={seriesColor(series, seriesIndex)}
                />
              )
            })}
            {showValueLabels &&
              (() => {
                const share = norm(primary.values[index] ?? bounds.min)
                const inside = dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
                const y =
                  dataLabelPosition === 'center'
                    ? 283 - share * 120
                    : dataLabelPosition === 'inside-end'
                      ? 292 - share * 240
                      : 272 - share * 240
                return (
                  <text
                    x={groupLeft(index) + groupWidth / 2}
                    y={y}
                    textAnchor="middle"
                    className="chart-axis-label"
                    {...(inside ? { fill: '#fff' } : {})}
                  >
                    {formatLabelValue(
                      primary.values[index] ?? 0,
                      dataLabelFormat,
                      primary.numberFormat,
                    )}
                  </text>
                )
              })()}
            <text x={62 + columnWidth * index + columnWidth / 2} y="298" textAnchor="middle">
              {truncateLabel(
                categories[index] ?? String(index + 1),
                categoryLabelBudget(visibleCount),
              )}
            </text>
          </g>
        )
      })}
      {points && (
        <polyline
          points={points}
          fill="none"
          stroke={seriesColor(lineSeries, seriesList.length)}
          strokeWidth="3"
        />
      )}
      {lineSeries &&
        lineScale &&
        secondaryAxis?.hidden !== true &&
        lineScale.ticks.map((tick, index) => (
          <text
            key={index}
            x="596"
            y={284 - ((tick - lineScale.min) / (lineScale.max - lineScale.min || 1)) * 240}
            textAnchor="end"
            className="chart-axis-label"
          >
            {formatAxisValue(tick, secondaryAxis?.numFmt ?? lineSeries.numberFormat)}
          </text>
        ))}
      {visibleCount < primary.values.length && (
        <text x="580" y="14" textAnchor="end" className="chart-axis-label" opacity="0.75">
          +{primary.values.length - visibleCount} more
        </text>
      )}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

function LineChartSvg({
  seriesList,
  axisTitles,
  dataLabels,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: AxisTitles
  readonly dataLabels?: DataLabels
  readonly grouping?: Grouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: AxisScaleInput | undefined
  readonly categoryFormat?: string | undefined
}): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="chart-error">Chart has no cached data.</div>
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  const stackTotals = isStacked
    ? seriesList.reduce<number[][]>((totals, series) => {
        const previous = totals[totals.length - 1] ?? primary.values.map(() => 0)
        totals.push(previous.map((base, index) => base + Math.max(0, series.values[index] ?? 0)))
        return totals
      }, [])
    : []
  const percentTotals = isPercent
    ? primary.values.map(
        (_, index) =>
          seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0) || 1,
      )
    : []
  const displayValues = (seriesIndex: number): number[] => {
    if (!isStacked) return [...(seriesList[seriesIndex]?.values ?? [])]
    const totals = stackTotals[seriesIndex] ?? []
    return isPercent ? totals.map((value, index) => value / (percentTotals[index] ?? 1)) : totals
  }
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : valueAxisScale(
        isStacked
          ? Math.max(...(stackTotals[stackTotals.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
      )
  const span = bounds.max - bounds.min
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)}
        showGridlines={gridlines !== false}
      />
      {seriesList.map((series, seriesIndex) => (
        <polyline
          key={seriesIndex}
          points={linePoints(displayValues(seriesIndex), bounds.max, bounds.min)}
          fill="none"
          stroke={seriesColor(series, seriesIndex)}
          strokeWidth="3"
        />
      ))}
      {primary.values.map((_, index) => (
        <text key={index} x={60 + (index / count) * 500} y="300" textAnchor="middle">
          {truncateLabel(categories[index] ?? String(index + 1), categoryLabelBudget(count))}
        </text>
      ))}
      {dataLabels === 'value' &&
        displayValues(0).map((displayed, index) => (
          <text
            key={`label-${index}`}
            x={60 + (index / count) * 500}
            y={272 - Math.max(0, Math.min(1, (displayed - bounds.min) / span)) * 240}
            textAnchor="middle"
            className="chart-axis-label"
          >
            {formatAxisValue(primary.values[index] ?? 0, primary.numberFormat)}
          </text>
        ))}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

function AreaChartSvg({
  seriesList,
  axisTitles,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: AxisTitles
  readonly grouping?: Grouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: AxisScaleInput | undefined
  readonly categoryFormat?: string | undefined
}): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="chart-error">Chart has no cached data.</div>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  const stackBounds = isStacked
    ? seriesList.reduce<number[][]>((bounds, series) => {
        const previous = bounds[bounds.length - 1] ?? primary.values.map(() => 0)
        bounds.push(previous.map((base, index) => base + Math.max(0, series.values[index] ?? 0)))
        return bounds
      }, [])
    : []
  const percentTotals = isPercent
    ? primary.values.map(
        (_, index) =>
          seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0) || 1,
      )
    : []
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : valueAxisScale(
        isStacked
          ? Math.max(...(stackBounds[stackBounds.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
      )
  const stackY = (raw: number, index: number): number =>
    280 - (isPercent ? raw / (percentTotals[index] ?? 1) : raw / (bounds.max || 1)) * 240
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)}
        showGridlines={gridlines !== false}
      />
      {seriesList.map((series, seriesIndex) => {
        if (isStacked) {
          const upper = stackBounds[seriesIndex] ?? []
          const lower =
            seriesIndex === 0 ? upper.map(() => 0) : (stackBounds[seriesIndex - 1] ?? [])
          const xFor = (index: number): number => 60 + (index / count) * 500
          const upperPoints = upper.map((value, index) => `${xFor(index)},${stackY(value, index)}`)
          const lowerPoints = lower
            .map((value, index) => `${xFor(index)},${stackY(value, index)}`)
            .reverse()
          return (
            <g key={seriesIndex}>
              <polygon
                points={[...upperPoints, ...lowerPoints].join(' ')}
                fill={seriesColor(series, seriesIndex)}
                opacity="0.75"
              />
              <polyline
                points={upperPoints.join(' ')}
                fill="none"
                stroke={seriesColor(series, seriesIndex)}
                strokeWidth="2"
              />
            </g>
          )
        }
        const xFor = (index: number): number => 60 + (index / count) * 500
        const points = series.values.map(
          (value, index) =>
            `${xFor(index)},${280 - Math.max(0, Math.min(1, (value - bounds.min) / (bounds.max - bounds.min || 1))) * 240}`,
        )
        const base = series.values.map((_, index) => `${xFor(index)},280`)
        return (
          <polygon
            key={seriesIndex}
            points={[...points, ...base].join(' ')}
            fill={seriesColor(series, seriesIndex)}
            opacity="0.75"
          />
        )
      })}
      {primary.values.map((_, index) => (
        <text key={index} x={60 + (index / count) * 500} y="300" textAnchor="middle">
          {truncateLabel(categories[index] ?? String(index + 1), categoryLabelBudget(count))}
        </text>
      ))}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

function RadarChartSvg({
  seriesList,
  dataLabels,
  categoryFormat,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly dataLabels?: DataLabels
  readonly categoryFormat?: string | undefined
}): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="chart-error">Chart has no cached data.</div>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.min(Math.max(primary.values.length, 3), 12)
  const cx = 300
  const cy = 168
  const r = 118
  const maximum = Math.max(...seriesList.flatMap((series) => [...series.values]), 0, 1)
  const vertex = (index: number, radius: number): [number, number] => {
    const theta = (index / count) * 2 * Math.PI
    return [cx + Math.sin(theta) * radius, cy - Math.cos(theta) * radius]
  }
  const ringPoints = (fraction: number): string =>
    Array.from({ length: count }, (_, index) => vertex(index, r * fraction).join(',')).join(' ')
  const seriesPoints = (series: ChartSeries): string =>
    Array.from({ length: count }, (_, index) => {
      const value = Math.max(0, series.values[index] ?? 0)
      return vertex(index, Math.min(1, value / maximum) * r).join(',')
    }).join(' ')
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      {[0.25, 0.5, 0.75, 1].map((fraction) => (
        <polygon key={fraction} points={ringPoints(fraction)} fill="none" stroke="#e3e3e3" />
      ))}
      {Array.from({ length: count }, (_, index) => {
        const [x, y] = vertex(index, r)
        const [lx, ly] = vertex(index, r * 1.12)
        return (
          <g key={index}>
            <line x1={cx} y1={cy} x2={x} y2={y} stroke="#e3e3e3" />
            <text
              x={lx}
              y={ly + 3}
              textAnchor={Math.abs(lx - cx) < 8 ? 'middle' : lx > cx ? 'start' : 'end'}
            >
              {truncateLabel(categories[index] ?? String(index + 1), 8)}
            </text>
          </g>
        )
      })}
      {seriesList.map((series, seriesIndex) => (
        <polygon
          key={seriesIndex}
          points={seriesPoints(series)}
          fill={seriesColor(series, seriesIndex)}
          fillOpacity="0.18"
          stroke={seriesColor(series, seriesIndex)}
          strokeWidth="2.5"
        />
      ))}
      {dataLabels === 'value' &&
        Array.from({ length: count }, (_, index) => {
          const value = primary.values[index] ?? 0
          const [x, y] = vertex(index, Math.min(1, Math.max(0, value) / maximum) * r + 10)
          return (
            <text key={index} x={x} y={y} textAnchor="middle" className="chart-axis-label">
              {formatAxisValue(value, primary.numberFormat)}
            </text>
          )
        })}
    </svg>
  )
}

function ScatterChartSvg({
  seriesList,
  axisTitles,
  gridlines,
  valueAxis,
  xAxis,
  scatterStyle,
  categoryFormat,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: AxisTitles
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: AxisScaleInput | undefined
  readonly xAxis?: AxisScaleInput | undefined
  readonly scatterStyle?: string | undefined
  readonly categoryFormat?: string | undefined
}): React.JSX.Element {
  const points = seriesList.map((series) => {
    const xValues = series.categories.map((value, index) => {
      const parsed = Number.parseFloat(value)
      return Number.isFinite(parsed) ? parsed : index
    })
    return { series, xValues }
  })
  const allX = points.flatMap((entry) => entry.xValues)
  const allY = points.flatMap((entry) => [...entry.series.values])
  if (allY.length === 0) return <div className="chart-error">Chart has no cached data.</div>
  const boundsX = scatterAxisBounds(allX, xAxis)
  const boundsY = scatterAxisBounds(allY, valueAxis)
  const styleWantsLines = scatterStyle !== undefined && scatterStyle !== 'marker'
  const xFormat = xAxis?.numFmt ?? categoryFormat
  const plotX = (value: number): number =>
    60 + Math.max(0, Math.min(1, (value - boundsX.min) / (boundsX.max - boundsX.min))) * 520
  const plotY = (value: number): number =>
    280 - Math.max(0, Math.min(1, (value - boundsY.min) / (boundsY.max - boundsY.min))) * 240
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={boundsY.min}
        maximum={boundsY.max}
        ticks={boundsY.ticks}
        numberFormat={valueAxis?.numFmt ?? seriesList[0]?.numberFormat}
        showGridlines={gridlines !== false}
      />
      {boundsX.ticks.map((tick, index) => (
        <g key={index}>
          {xAxis?.majorGridlines && (
            <line
              x1={plotX(tick)}
              y1="40"
              x2={plotX(tick)}
              y2="280"
              stroke="#e3e3e3"
              strokeWidth="1"
            />
          )}
          <text x={plotX(tick)} y="296" textAnchor="middle" className="chart-axis-label">
            {formatScatterTick(tick, xFormat)}
          </text>
        </g>
      ))}
      {points.map(({ series, xValues }, seriesIndex) => (
        <g key={seriesIndex}>
          {styleWantsLines && series.lineColor !== 'none' && series.values.length > 1 && (
            <polyline
              points={series.values
                .map((value, index) => `${plotX(xValues[index] ?? 0)},${plotY(value)}`)
                .join(' ')}
              fill="none"
              stroke={
                series.lineColor && series.lineColor !== 'none'
                  ? series.lineColor
                  : seriesColor(series, seriesIndex)
              }
              strokeWidth="2"
            />
          )}
          {series.marker !== 'none' &&
            series.values.map((value, index) => (
              <circle
                key={index}
                cx={plotX(xValues[index] ?? 0)}
                cy={plotY(value)}
                r={4}
                fill={seriesColor(series, seriesIndex)}
                opacity="0.85"
              />
            ))}
        </g>
      ))}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

interface PieSliceLabel {
  key: number
  inside: boolean
  x: number
  y: number
  anchor: 'start' | 'middle' | 'end'
  lines: string[]
  leader?: string | undefined
}

function pieSliceLabels(
  values: readonly number[],
  categories: readonly string[],
  total: number,
  mode: 'value' | 'percent' | 'category-percent',
  geometry: { cx: number; cy: number; r: number; inner: number; offsets: readonly number[] },
  position?: LabelPosition,
  formatCode?: string | undefined,
): PieSliceLabel[] {
  const { cx, cy, r, inner, offsets } = geometry
  const labels: PieSliceLabel[] = []
  let cursor = 0
  for (const [index, value] of values.entries()) {
    const share = Math.max(0, value) / total
    const mid = (cursor + share / 2) * 2 * Math.PI
    cursor += share
    if (share < 0.02) continue
    const percentText = `${(share * 100).toFixed(share >= 0.1 ? 0 : 1)}%`
    const lines =
      mode === 'category-percent'
        ? [truncateLabel(categories[index] ?? '', 12), percentText]
        : [mode === 'value' ? formatLabelValue(value, formatCode, undefined) : percentText]
    const sin = Math.sin(mid)
    const cos = Math.cos(mid)
    const offset = offsets[index] ?? 0
    const inside =
      position === 'outside-end' ? false : position !== undefined ? true : share >= 0.09
    if (inside) {
      const baseRadius =
        inner > 0
          ? (r + inner) / 2
          : position === 'inside-end'
            ? r * 0.8
            : position === 'center'
              ? r * 0.5
              : r * 0.62
      const radius = baseRadius + offset
      labels.push({
        key: index,
        inside: true,
        x: cx + sin * radius,
        y: cy - cos * radius,
        anchor: 'middle',
        lines,
      })
      continue
    }
    const elbowR = r * 1.1 + offset
    const tick = 10
    const rightSide = sin >= 0
    const elbowX = cx + sin * elbowR
    const elbowY = cy - cos * elbowR
    const textX = elbowX + (rightSide ? tick : -tick)
    labels.push({
      key: index,
      inside: false,
      x: textX + (rightSide ? 3 : -3),
      y: elbowY,
      anchor: rightSide ? 'start' : 'end',
      lines,
      leader: `${cx + sin * (r + offset)},${cy - cos * (r + offset)} ${elbowX},${elbowY} ${textX},${elbowY}`,
    })
  }
  for (const side of ['start', 'end'] as const) {
    const outside = labels
      .filter((label) => !label.inside && label.anchor === side)
      .sort((a, b) => a.y - b.y)
    const gap = 15
    for (let i = 1; i < outside.length; i += 1) {
      const previous = outside[i - 1]
      const current = outside[i]
      if (previous && current && current.y < previous.y + gap) current.y = previous.y + gap
    }
  }
  return labels
}

function PieChartSvg({
  series,
  isDoughnut,
  legend,
  dataLabels,
  dataLabelPosition,
  dataLabelFormat,
  holeSizePct,
  categoryFormat,
}: {
  readonly series: ChartSeries
  readonly isDoughnut: boolean
  readonly legend?: 'none' | 'right' | 'bottom' | 'top' | 'left' | undefined
  readonly dataLabels?: DataLabels
  readonly dataLabelPosition?: LabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly holeSizePct?: number | undefined
  readonly categoryFormat?: string | undefined
}): React.JSX.Element {
  const { values } = series
  const categories = series.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const total = values.reduce((sum, value) => sum + Math.max(0, value), 0) || 1
  const cx = 210
  const cy = 160
  const r = 100
  const inner = isDoughnut ? (r * Math.min(90, Math.max(10, holeSizePct ?? 50))) / 100 : 0
  const point = (radius: number, theta: number): [number, number] => [
    cx + Math.sin(theta) * radius,
    cy - Math.cos(theta) * radius,
  ]
  const explosionOffset = (index: number): number => {
    const pct =
      series.pointExplosions?.find((entry) => entry.index === index)?.pct ??
      series.explosionPct ??
      0
    return (Math.min(100, Math.max(0, pct)) / 100) * r * 0.4
  }
  const offsets = values.map((_, index) => explosionOffset(index))
  let angle = 0
  const slices = values
    .map((value, index) => {
      const share = Math.max(0, value) / total
      const start = angle
      angle += share * 2 * Math.PI
      return { index, share, start, end: angle, color: pieSliceColor(series, index) }
    })
    .filter((slice) => slice.share > 0.0005)
  const slicePath = (slice: (typeof slices)[number]): string => {
    const end = slice.share > 0.9995 ? slice.start + 2 * Math.PI - 0.001 : slice.end
    const large = end - slice.start > Math.PI ? 1 : 0
    const [x1, y1] = point(r, slice.start)
    const [x2, y2] = point(r, end)
    if (inner === 0) {
      return `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`
    }
    const [ix1, iy1] = point(inner, slice.start)
    const [ix2, iy2] = point(inner, end)
    return (
      `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}` +
      ` L ${ix2} ${iy2} A ${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1} Z`
    )
  }
  const labels =
    dataLabels !== undefined && dataLabels !== 'none'
      ? pieSliceLabels(
          values,
          categories,
          total,
          dataLabels,
          { cx, cy, r, inner, offsets },
          dataLabelPosition,
          dataLabelFormat,
        )
      : []
  return (
    <div className={`chart-pie-layout chart-legend-${legend ?? 'right'}`}>
      <svg className="chart-svg chart-pie-svg" viewBox="0 0 420 320" role="img">
        {slices.map((slice) => {
          const mid = (slice.start + slice.end) / 2
          const offset = offsets[slice.index] ?? 0
          return (
            <path
              key={slice.index}
              d={slicePath(slice)}
              transform={
                offset > 0
                  ? `translate(${Math.sin(mid) * offset} ${-Math.cos(mid) * offset})`
                  : undefined
              }
              fill={slice.color}
              stroke="#fff"
              strokeWidth="1.5"
            />
          )
        })}
        {labels.map((label) => (
          <g key={label.key}>
            {label.leader && (
              <polyline points={label.leader} fill="none" stroke="#9a9a9a" strokeWidth="1" />
            )}
            <text
              x={label.x}
              y={label.y + (label.lines.length > 1 ? -2 : 3)}
              textAnchor={label.anchor}
              className={label.inside ? 'chart-pie-label-inside' : 'chart-pie-label-outside'}
            >
              {label.lines.map((line, lineIndex) => (
                <tspan key={lineIndex} x={label.x} dy={lineIndex === 0 ? 0 : 12}>
                  {line}
                </tspan>
              ))}
            </text>
          </g>
        ))}
      </svg>
      {legend !== 'none' && (
        <div className="chart-legend">
          {categories.slice(0, 12).map((category, index) => {
            const share = (Math.max(0, values[index] ?? 0) / total) * 100
            return (
              <span key={`${category}-${index}`}>
                <i style={{ background: pieSliceColor(series, index) }} />
                {truncateLabel(category, 12)} {share >= 0.05 ? `${share.toFixed(1)}%` : ''}
              </span>
            )
          })}
          {categories.length > 12 && <span>+{categories.length - 12} more</span>}
        </div>
      )}
    </div>
  )
}

/// The chart body dispatcher — desktop ChartVisual family dispatch. The
/// chart parameter is the canonical domain ChartVisualState (file charts
/// carry the pending-edit overlay already; session creations build one).
export function ChartSvg({ chart }: { readonly chart: ChartVisualState }): React.JSX.Element {
  const populated = chart.series.filter((series) => series.values.length > 0)
  const primarySeries = populated[0]
  if (!primarySeries) {
    return <div className="chart-error">Chart has no cached data.</div>
  }
  const types = chart.chartTypes
  const isDoughnut = types.includes('doughnutChart')
  const isPie = types.includes('pieChart') || isDoughnut
  const isScatter = types.includes('scatterChart')
  const isCombo = types.includes('barChart') && types.includes('lineChart')
  const isArea = types.includes('areaChart') && !types.includes('barChart')
  const isLine = types.includes('lineChart') && !types.includes('barChart')
  const isRadar = types.includes('radarChart') && !types.includes('barChart')
  const categoryFormat = chartCategoryFormat(chart)
  const isHorizontalBar = chart.barDirection === 'bar'
  const effectiveAxisTitles = {
    category: (isHorizontalBar ? chart.yAxis : chart.xAxis)?.title ?? chart.axisTitles?.category,
    value: (isHorizontalBar ? chart.xAxis : chart.yAxis)?.title ?? chart.axisTitles?.value,
  }
  const valueScaleInput = {
    min: chart.yAxis?.min ?? chart.valueAxis?.min,
    max: chart.yAxis?.max ?? chart.valueAxis?.max,
    majorUnit: chart.yAxis?.majorUnit,
    numFmt: chart.yAxis?.numFmt,
  }
  return (
    <div className={`chart-body chart-legend-${chart.legend ?? 'default'}`}>
      {isPie ? (
        <PieChartSvg
          series={primarySeries}
          isDoughnut={isDoughnut}
          legend={chart.legend}
          dataLabels={chart.dataLabels}
          dataLabelPosition={chart.dataLabelPosition}
          dataLabelFormat={chart.dataLabelFormat}
          holeSizePct={chart.holeSizePct}
          categoryFormat={categoryFormat}
        />
      ) : isRadar ? (
        <RadarChartSvg
          seriesList={populated}
          dataLabels={chart.dataLabels}
          categoryFormat={categoryFormat}
        />
      ) : isScatter ? (
        <ScatterChartSvg
          seriesList={populated}
          axisTitles={effectiveAxisTitles}
          gridlines={chart.gridlines}
          valueAxis={valueScaleInput}
          xAxis={chart.xAxis}
          scatterStyle={chart.scatterStyle}
          categoryFormat={categoryFormat}
        />
      ) : isArea ? (
        <AreaChartSvg
          seriesList={populated}
          axisTitles={effectiveAxisTitles}
          grouping={chart.grouping}
          gridlines={chart.gridlines}
          valueAxis={valueScaleInput}
          categoryFormat={categoryFormat}
        />
      ) : isLine ? (
        <LineChartSvg
          seriesList={populated}
          axisTitles={effectiveAxisTitles}
          dataLabels={chart.dataLabels}
          grouping={chart.grouping}
          gridlines={chart.gridlines}
          valueAxis={valueScaleInput}
          categoryFormat={categoryFormat}
        />
      ) : (
        <BarChartSvg
          seriesList={isCombo && populated.length > 1 ? populated.slice(0, -1) : populated}
          isHorizontal={chart.barDirection === 'bar'}
          lineSeries={isCombo && populated.length > 1 ? populated[populated.length - 1] : undefined}
          axisTitles={effectiveAxisTitles}
          dataLabels={chart.dataLabels}
          dataLabelPosition={chart.dataLabelPosition}
          dataLabelFormat={chart.dataLabelFormat}
          grouping={chart.grouping}
          gridlines={chart.gridlines}
          valueAxis={valueScaleInput}
          secondaryAxis={chart.secondaryYAxis}
          gapWidthPct={chart.gapWidthPct}
          categoryFormat={categoryFormat}
        />
      )}
      {!isPie &&
        chart.legend !== 'none' &&
        (chart.legend !== undefined || populated.length > 1) && (
          <SeriesLegend seriesList={populated} />
        )}
    </div>
  )
}

// ── Anchor ↔ pixel math over the live grid (desktop walkMarker) ──────

export interface WorksheetGeometry {
  columnWidth(index: number): number
  rowHeight(index: number): number
  maxColumns: number
  maxRows: number
}

interface Marker {
  index: number
  offset: number
}

function walkMarker(
  marker: Marker,
  delta: number,
  sizeOf: (index: number) => number,
  maxIndex: number,
): Marker {
  let index = Math.min(marker.index, maxIndex)
  let offset = marker.offset + delta
  while (offset < 0 && index > 0) {
    index -= 1
    offset += sizeOf(index)
  }
  if (offset < 0) offset = 0
  while (index < maxIndex && offset >= sizeOf(index)) {
    offset -= sizeOf(index)
    index += 1
  }
  if (index >= maxIndex) offset = Math.min(offset, sizeOf(maxIndex))
  return { index, offset }
}

function markerSpan(from: Marker, to: Marker, sizeOf: (index: number) => number): number {
  const span = to.offset - from.offset
  const low = Math.min(from.index, to.index)
  const high = Math.max(from.index, to.index)
  let cells = 0
  for (let index = low; index < high; index += 1) cells += sizeOf(index)
  return span + (from.index <= to.index ? cells : -cells)
}

/// The pixel frame of an anchor against the live grid (two-cell anchors
/// span their markers; one-cell anchors use the explicit size).
export function anchorFrame(
  geometry: WorksheetGeometry,
  anchor: DrawingAnchor,
  size?: { widthPx?: number; heightPx?: number },
): { width: number; height: number; marginX: number; marginY: number } {
  const marginX = Math.max(0, emuToPx(anchor.fromColumnOffset))
  const marginY = Math.max(0, emuToPx(anchor.fromRowOffset))
  const width =
    size?.widthPx ??
    Math.max(
      0,
      markerSpan(
        { index: anchor.fromColumn, offset: marginX },
        { index: anchor.toColumn, offset: emuToPx(anchor.toColumnOffset) },
        geometry.columnWidth,
      ),
    )
  const height =
    size?.heightPx ??
    Math.max(
      0,
      markerSpan(
        { index: anchor.fromRow, offset: marginY },
        { index: anchor.toRow, offset: emuToPx(anchor.toRowOffset) },
        geometry.rowHeight,
      ),
    )
  return { width, height, marginX, marginY }
}

/// Move an anchor by a pixel delta (two-cell: both markers; one-cell:
/// from-marker only, size preserved).
export function anchorMoved(
  geometry: WorksheetGeometry,
  anchor: DrawingAnchor,
  dx: number,
  dy: number,
  fixedSize?: { widthPx?: number; heightPx?: number },
): DrawingAnchor {
  const fromCol = walkMarker(
    { index: anchor.fromColumn, offset: emuToPx(anchor.fromColumnOffset) },
    dx,
    geometry.columnWidth,
    geometry.maxColumns - 1,
  )
  const fromRow = walkMarker(
    { index: anchor.fromRow, offset: emuToPx(anchor.fromRowOffset) },
    dy,
    geometry.rowHeight,
    geometry.maxRows - 1,
  )
  if (fixedSize !== undefined) {
    return {
      fromColumn: fromCol.index,
      fromColumnOffset: pxToEmu(fromCol.offset),
      fromRow: fromRow.index,
      fromRowOffset: pxToEmu(fromRow.offset),
      toColumn: fromCol.index,
      toColumnOffset: anchor.toColumnOffset,
      toRow: fromRow.index,
      toRowOffset: anchor.toRowOffset,
    }
  }
  const toCol = walkMarker(
    { index: anchor.toColumn, offset: emuToPx(anchor.toColumnOffset) },
    dx,
    geometry.columnWidth,
    geometry.maxColumns - 1,
  )
  const toRow = walkMarker(
    { index: anchor.toRow, offset: emuToPx(anchor.toRowOffset) },
    dy,
    geometry.rowHeight,
    geometry.maxRows - 1,
  )
  return {
    fromColumn: fromCol.index,
    fromColumnOffset: pxToEmu(fromCol.offset),
    fromRow: fromRow.index,
    fromRowOffset: pxToEmu(fromRow.offset),
    toColumn: toCol.index,
    toColumnOffset: pxToEmu(toCol.offset),
    toRow: toRow.index,
    toRowOffset: pxToEmu(toRow.offset),
  }
}

/// Resize an anchor by corner-driven pixel deltas, keeping a minimum frame.
export function anchorResized(
  geometry: WorksheetGeometry,
  anchor: DrawingAnchor,
  corner: 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w',
  dx: number,
  dy: number,
): DrawingAnchor {
  const east = corner.includes('e')
  const west = corner.includes('w')
  const north = corner.includes('n')
  const south = corner.includes('s')
  let next: DrawingAnchor = { ...anchor }
  if (east) {
    const to = walkMarker(
      { index: anchor.toColumn, offset: emuToPx(anchor.toColumnOffset) },
      dx,
      geometry.columnWidth,
      geometry.maxColumns - 1,
    )
    next = { ...next, toColumn: to.index, toColumnOffset: pxToEmu(to.offset) }
  }
  if (west) {
    const from = walkMarker(
      { index: anchor.fromColumn, offset: emuToPx(anchor.fromColumnOffset) },
      dx,
      geometry.columnWidth,
      geometry.maxColumns - 1,
    )
    next = { ...next, fromColumn: from.index, fromColumnOffset: pxToEmu(from.offset) }
  }
  if (south) {
    const to = walkMarker(
      { index: anchor.toRow, offset: emuToPx(anchor.toRowOffset) },
      dy,
      geometry.rowHeight,
      geometry.maxRows - 1,
    )
    next = { ...next, toRow: to.index, toRowOffset: pxToEmu(to.offset) }
  }
  if (north) {
    const from = walkMarker(
      { index: anchor.fromRow, offset: emuToPx(anchor.fromRowOffset) },
      dy,
      geometry.rowHeight,
      geometry.maxRows - 1,
    )
    next = { ...next, fromRow: from.index, fromRowOffset: pxToEmu(from.offset) }
  }
  // Keep a minimum frame: a degenerate resize must not collapse the chart.
  const frame = anchorFrame(geometry, next)
  if (frame.width < 24 || frame.height < 24) return { ...anchor }
  return next
}

// ── Interactive floating chart frame ─────────────────────────────────

const MIN_FRAME_PIXELS = 24
const RESIZE_CORNERS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
type ResizeCorner = (typeof RESIZE_CORNERS)[number]

/**
 * The live chart spec a frame renders: the file chart's canonical state
 * with the pending semantic edit overlaid (desktop withChartEdit
 * parity — the on-screen chart previews the pending save), or the
 * session creation's ChartAdd rendered through the same SVG family.
 */
export function liveChartSpec(
  store: ChartEditingStore,
  key: string,
): {
  anchor: DrawingAnchor
  anchorType: 'two-cell' | 'one-cell'
  widthPx?: number | undefined
  heightPx?: number | undefined
  chart: ChartVisualState | ChartAdd
  isSession: boolean
} | null {
  const file = store.fileCharts.get(key)
  if (file !== undefined) {
    if (store.removals.has(key)) return null
    return {
      anchor: file.info.anchor,
      anchorType: file.info.anchorType,
      widthPx: file.info.widthPx,
      heightPx: file.info.heightPx,
      chart: applyChartStateEdit(file.info.chart, store.edits.get(key)),
      isSession: false,
    }
  }
  const session = store.sessionAdds.get(key)
  if (session !== undefined) {
    return {
      anchor: session.anchor,
      anchorType: 'two-cell',
      chart: session.chart,
      isSession: true,
    }
  }
  return null
}

/**
 * The interactive chart frame: renders the SVG chart (live state from the
 * shared store, so pending semantic edits preview immediately), selects
 * on click, moves on drag, resizes through 8 handles (two-cell anchors
 * only — a one-cell chart keeps its fixed size, so resize fails closed),
 * and deletes via its button. The drag commits on window mouseup, so a
 * release outside the frame still lands.
 */
function ChartFrame({
  chartKey,
  store,
  geometry,
  onGeometryCommit,
  onStructureChange,
}: {
  readonly chartKey: string
  readonly store: ChartEditingStore
  readonly geometry: WorksheetGeometry
  readonly onGeometryCommit: (key: string, anchor: DrawingAnchor) => void
  readonly onStructureChange: () => void
}): React.JSX.Element | null {
  useChartStoreVersion(store)
  const spec = liveChartSpec(store, chartKey)
  const dragRef = useRef<{
    mode: 'move' | 'resize'
    corner: ResizeCorner
    startX: number
    startY: number
    dx: number
    dy: number
  } | null>(null)
  const [dragTick, setDragTick] = useState(0)
  const drag = dragRef.current
  const isSelected = store.selection === chartKey

  useEffect(() => {
    const onMove = (event: MouseEvent): void => {
      const current = dragRef.current
      if (current === null) return
      const dx = event.clientX - current.startX
      const dy = event.clientY - current.startY
      if (dx === current.dx && dy === current.dy) return
      dragRef.current = { ...current, dx, dy }
      setDragTick((tick) => tick + 1)
    }
    const onUp = (): void => {
      const current = dragRef.current
      dragRef.current = null
      if (current !== null && spec !== null && (current.dx !== 0 || current.dy !== 0)) {
        if (current.mode === 'move') {
          const fixedSize =
            spec.anchorType === 'one-cell'
              ? { widthPx: spec.widthPx, heightPx: spec.heightPx }
              : undefined
          onGeometryCommit(
            chartKey,
            anchorMoved(geometry, spec.anchor, current.dx, current.dy, fixedSize),
          )
        } else if (spec.anchorType !== 'one-cell') {
          onGeometryCommit(
            chartKey,
            anchorResized(geometry, spec.anchor, current.corner, current.dx, current.dy),
          )
        }
      }
      setDragTick((tick) => tick + 1)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [chartKey, geometry, onGeometryCommit, spec, dragTick])

  if (spec === null) return null
  const chart = spec.chart
  const isAdd = 'chartType' in chart
  const title = chart.title
  const chartState = isAdd ? undefined : chart
  const chartAdd = isAdd ? chart : undefined
  const oneCell = spec.anchorType === 'one-cell'

  return (
    <div
      className={`chart-frame${isSelected ? ' chart-selected' : ''}`}
      data-chart-key={chartKey}
      style={{
        width: '100%',
        height: '100%',
        ...(drag !== null ? { transform: `translate(${drag.dx}px, ${drag.dy}px)` } : {}),
      }}
      onMouseDown={(event) => {
        if (event.button !== 0) return
        store.selection = chartKey
        store.bump()
        dragRef.current = {
          mode: 'move',
          corner: 'se',
          startX: event.clientX,
          startY: event.clientY,
          dx: 0,
          dy: 0,
        }
        setDragTick((tick) => tick + 1)
      }}
    >
      {title !== '' && (
        <div
          className="chart-title"
          style={{
            ...(chartState?.titleStyle?.size !== undefined
              ? { fontSize: `${chartState.titleStyle.size}pt` }
              : {}),
            ...(chartState?.titleStyle?.bold === false ? { fontWeight: 400 } : {}),
            ...(chartState?.titleStyle?.color !== undefined
              ? { color: chartState.titleStyle.color }
              : {}),
          }}
        >
          {title}
        </div>
      )}
      {chartState !== undefined && <ChartSvg chart={chartState} />}
      {chartAdd !== undefined && <ChartSvg chart={chartAddToVisualState(chartAdd)} />}
      {isSelected && (
        <>
          <button
            className="chart-delete-button"
            title="Delete chart"
            data-testid="chart-delete"
            onMouseDown={(event) => {
              event.stopPropagation()
            }}
            onClick={(event) => {
              event.stopPropagation()
              if (store.fileCharts.has(chartKey)) {
                store.removals.add(chartKey)
                store.dirty.delete(chartKey)
                store.edits.delete(chartKey)
              } else {
                store.sessionAdds.delete(chartKey)
                store.edits.delete(chartKey)
              }
              store.selection = null
              store.bump()
              onStructureChange()
            }}
          >
            ✕
          </button>
          {!oneCell &&
            RESIZE_CORNERS.map((corner) => (
              <span
                key={corner}
                className={`chart-handle handle-${corner}`}
                data-testid={`chart-handle-${corner}`}
                onMouseDown={(event) => {
                  event.stopPropagation()
                  dragRef.current = {
                    mode: 'resize',
                    corner,
                    startX: event.clientX,
                    startY: event.clientY,
                    dx: 0,
                    dy: 0,
                  }
                  setDragTick((tick) => tick + 1)
                }}
              />
            ))}
        </>
      )}
    </div>
  )
}

// ── Install (registerComponent + addFloatDomToRange — public facades) ─

/**
 * The worksheet surface the chart layer needs, as a structural interface
 * satisfied by the drawing-preset-augmented FWorksheet (the same
 * sheet-images.ts ImageWorksheetFacade pattern — no casts, no private
 * internals). addFloatDomToRange/registerComponent are Univer's PUBLIC
 * floating-DOM facade, documented with examples in
 * @univerjs/sheets-drawing-ui's f-worksheet declarations; the desktop
 * installs its charts through the exact same pair.
 */
export interface ChartWorksheetFacade {
  getRange(row: number, column: number, numRows: number, numColumns: number): unknown
  getColumnWidth(column: number): number
  getRowHeight(row: number): number
  getMaxRows(): number
  getMaxColumns(): number
  setRowCount(rows: number): void
  setColumnCount(columns: number): void
  addFloatDomToRange(
    range: unknown,
    layer: { componentKey: string },
    domLayout: {
      width?: number | undefined
      height?: number | undefined
      marginX?: number | undefined
      marginY?: number | undefined
    },
    id?: string,
    // Univer's Nullable<T> includes void (a missing float returns
    // undefined-ish from overloads) — the union must accept it.
  ): { dispose(): void } | undefined | null | void
}

/** The univerAPI surface the chart layer needs (public facade subset). */
export interface ChartUniverFacade {
  registerComponent(name: string, component: unknown): { dispose(): void }
}

export interface ChartHost {
  readonly univerAPI: ChartUniverFacade
  readonly store: ChartEditingStore
  onGeometryCommit(key: string, anchor: DrawingAnchor): void
  /** Structural change (insert/delete) — the host reinstalls the floats. */
  onStructureChange(): void
}

/**
 * Installs one chart as a floating DOM visual over the sheet's grid:
 * registerComponent + addFloatDomToRange (the desktop's exact install
 * pair), with the pixel-exact twoCellAnchor frame derived from the live
 * grid. Anchors beyond the grid grow it first (desktop parity — the
 * float keeps its frame instead of being clamped into a sliver). The
 * registered component reads LIVE state from the shared store, so
 * pending semantic edits preview without reinstalling.
 */
export function installSheetChart(
  host: ChartHost,
  ws: ChartWorksheetFacade,
  id: string,
): { dispose(): void } | null {
  const store = host.store
  const spec = liveChartSpec(store, id)
  if (spec === null) return null
  const geometry = worksheetGeometryOf(ws)
  if (spec.anchor.toRow >= ws.getMaxRows()) ws.setRowCount(spec.anchor.toRow + 1)
  if (spec.anchor.toColumn >= ws.getMaxColumns()) ws.setColumnCount(spec.anchor.toColumn + 1)
  const fromRow = Math.min(spec.anchor.fromRow, ws.getMaxRows() - 1)
  const fromColumn = Math.min(spec.anchor.fromColumn, ws.getMaxColumns() - 1)
  const toRow = Math.min(spec.anchor.toRow, ws.getMaxRows() - 1)
  const toColumn = Math.min(spec.anchor.toColumn, ws.getMaxColumns() - 1)
  const range = ws.getRange(
    fromRow,
    fromColumn,
    Math.max(1, toRow + 1 - fromRow),
    Math.max(1, toColumn + 1 - fromColumn),
  )
  // Two-cell anchors size from the live marker span; the explicit a:ext
  // pixel size applies to ONE-CELL charts only (their to markers carry
  // no size — desktop anchor semantics).
  const frame = anchorFrame(geometry, spec.anchor, {
    widthPx: spec.anchorType === 'one-cell' ? spec.widthPx : undefined,
    heightPx: spec.anchorType === 'one-cell' ? spec.heightPx : undefined,
  })
  const layout =
    frame.width >= MIN_FRAME_PIXELS / 2 && frame.height >= MIN_FRAME_PIXELS / 2
      ? frame
      : {
          width: Math.max(frame.width, 240),
          height: Math.max(frame.height, 160),
          marginX: frame.marginX,
          marginY: frame.marginY,
        }
  const componentKey = `web-chart-${sanitizeFloatId(id)}`
  const registration = host.univerAPI.registerComponent(componentKey, () => (
    <ChartFrame
      chartKey={id}
      store={store}
      geometry={geometry}
      onGeometryCommit={host.onGeometryCommit}
      onStructureChange={host.onStructureChange}
    />
  ))
  // The float DOM id must be DOM-safe: the canonical locator key carries
  // `:`, `/`, and `#` which break Univer's internal drawing/DOM lookups.
  const floating = ws.addFloatDomToRange(range, { componentKey }, layout, sanitizeFloatId(id))
  if (floating === undefined || floating === null) {
    registration.dispose()
    return null
  }
  return {
    dispose: () => {
      floating.dispose()
      registration.dispose()
    },
  }
}

function worksheetGeometryOf(ws: ChartWorksheetFacade): WorksheetGeometry {
  return {
    columnWidth: (index) => Math.max(ws.getColumnWidth(index), 1),
    rowHeight: (index) => Math.max(ws.getRowHeight(index), 1),
    maxColumns: ws.getMaxColumns(),
    maxRows: ws.getMaxRows(),
  }
}

/// DOM-safe float/component id for a canonical locator key (the key's
/// `:`, `/`, `#` break Univer's internal drawing and DOM id lookups —
/// the desktop likewise uses plain `${sessionId}-${visual.id}` ids).
function sanitizeFloatId(key: string): string {
  return key.replace(/[^A-Za-z0-9_-]/g, '_')
}

/// A session creation's ChartAdd rendered as the canonical domain state
/// (the wire keeps ChartAdd — the visual layer projects it once for the
/// SVG family, mirroring buildChartVisual's type mapping).
export function chartAddToVisualState(add: ChartAdd): ChartVisualState {
  return {
    chartTypes:
      add.chartType === 'combo'
        ? ['barChart', 'lineChart']
        : [
            add.chartType === 'column' || add.chartType === 'bar'
              ? 'barChart'
              : `${add.chartType}Chart`,
          ],
    ...(add.chartType === 'bar' ? { barDirection: 'bar' } : {}),
    ...(add.chartType === 'column' ? { barDirection: 'col' } : {}),
    title: add.title,
    series: add.series.map((series) => ({
      name: series.name,
      categories: [...series.categories],
      values: [...series.values],
      ...(series.color !== undefined ? { color: series.color } : {}),
    })),
    ...(add.legend !== undefined ? { legend: add.legend } : {}),
    ...(add.dataLabels !== undefined ? { dataLabels: add.dataLabels } : {}),
    ...(add.grouping !== undefined ? { grouping: add.grouping } : {}),
    ...(add.gridlines !== undefined ? { gridlines: add.gridlines } : {}),
    ...(add.gapWidthPct !== undefined ? { gapWidthPct: add.gapWidthPct } : {}),
    ...(add.holeSizePct !== undefined ? { holeSizePct: add.holeSizePct } : {}),
  }
}
