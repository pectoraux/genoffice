/// EXCEL-023 canonical chart reader: resolves a worksheet's drawing
/// relationship chain (worksheet → drawing part → chart relationships →
/// xl/charts/*.xml) into typed SheetChartInfo entries carrying the canonical
/// ChartVisualState (the shared domain model the desktop sidecar also
/// emits), so the browser renders and edits charts without any OOXML
/// knowledge.
///
/// Fail-closed semantics mirror the other per-sheet read families
/// (filters / validations / notes / tables / images): unreadable drawing
/// wiring surfaces NO charts for that sheet — the workbook still opens and
/// a no-op save preserves the file's parts byte-for-byte. Individual
/// charts outside the supported envelope (3-D plots, bubble/stock/surface,
/// chartEx extension parts, multi-plot combinations other than the
/// canonical bar+line combo, missing parts, oversized series sets) are
/// skipped; their anchors still count toward drawingIndex parity with
/// xlsx-drawing-edit.ts and xlsx-image-read.ts (the index counts EVERY
/// anchor element in document order).
///
/// absoluteAnchor charts are OMITTED from the browser model (EXCEL-022
/// precedent, architect review PR #20 blocker 2): their fixed-sheet
/// geometry has no two-cell representation and must never be relocated
/// through a zero-marker approximation. The anchor still counts toward
/// drawingIndex parity.

import type { ChartAxisInfoState, ChartSeriesVisualState, ChartVisualState } from '../domain/chart-visual.js'
import type { DrawingAnchor } from './xlsx-drawing-add.js'

export class ChartReadError extends Error {}

/// One chart on one sheet — the typed wire model the browser consumes.
/// `drawingPath` + `drawingIndex` locate the ANCHOR (move/resize/delete
/// ride the EXCEL-022 visualEdits family); `chartPath` locates the CHART
/// PART (semantic edits ride the chartEdits family). `chart` is the
/// canonical domain state the desktop renders from — structurally the
/// same model the desktop sidecar emits for its file visuals.
export interface SheetChartInfo {
  readonly drawingPath: string
  readonly drawingIndex: number
  readonly chartPath: string
  readonly anchorType: 'two-cell' | 'one-cell'
  readonly anchor: DrawingAnchor
  /// One-cell charts: explicit a:ext size in px (rounded) — the anchor's
  /// to markers carry no size, so the browser renders at this fixed size.
  readonly widthPx?: number | undefined
  readonly heightPx?: number | undefined
  /// xdr:cNvPr/@name — display label / alt text echo.
  readonly name?: string | undefined
  readonly chart: ChartVisualState
}

/// Entry source (text-only — charts need no binary reads).
export interface ChartEntrySource {
  readText(path: string): Promise<string>
  has(path: string): Promise<boolean>
}

const DRAWING_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'

/// Anchors count in document order — identical pattern to
/// xlsx-drawing-edit.ts and xlsx-image-read.ts, so a visualEdit's
/// drawingIndex from this reader addresses the exact same anchor the
/// editor patches.
const ANCHOR_PATTERN = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g

/// The single-plot families the canonical model represents (mirrors the
/// plot elements buildChartXml writes and applyChartEdit converts).
const SUPPORTED_PLOTS = [
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
  'scatterChart',
  'radarChart',
] as const
type SupportedPlot = (typeof SUPPORTED_PLOTS)[number]

/// Any other plot element means the chart is outside the envelope.
const KNOWN_UNSUPPORTED_PLOTS =
  /<c:(bar3DChart|line3DChart|area3DChart|pie3DChart|doughnut3DChart|ofPieChart|bubbleChart|stockChart|surfaceChart|surface3DChart|radar3DChart|scatter3DChart)\b/

/// Fail-closed caps (per chart / per sheet) — charts beyond them surface
/// nothing for that chart, matching the image reader's size discipline.
const MAX_CHARTS_PER_SHEET = 64
const MAX_SERIES_PER_CHART = 64
const MAX_POINTS_PER_SERIES = 8192

const EMU_PER_PX = 9525

export async function parseSheetCharts(
  reader: ChartEntrySource,
  worksheetPath: string,
  worksheetXml: string,
): Promise<readonly SheetChartInfo[]> {
  const drawingRelId = /<drawing\b[^>]*\br:id="([^"]+)"/.exec(worksheetXml)?.[1]
  if (drawingRelId === undefined) return []
  const drawingPath = await resolveDrawingPath(reader, worksheetPath, drawingRelId)
  if (drawingPath === null) {
    throw new ChartReadError(
      `The worksheet's drawing relationship "${drawingRelId}" could not be resolved.`,
    )
  }
  if (!(await reader.has(drawingPath))) {
    throw new ChartReadError(`Workbook is missing ${drawingPath}.`)
  }
  const drawingXml = await reader.readText(drawingPath)
  const chartRels = await parseChartRelationships(reader, drawingPath)
  const charts: SheetChartInfo[] = []
  let anchorIndex = 0
  for (const match of drawingXml.matchAll(ANCHOR_PATTERN)) {
    const index = anchorIndex
    anchorIndex += 1
    const anchorXml = match[0]
    const kind = match[1]
    if (kind === 'absoluteAnchor') {
      // Fail closed (EXCEL-022 precedent): fixed-sheet geometry has no
      // two-cell representation — omit, never relocate. The index already
      // counted above, so later anchors keep their locators.
      continue
    }
    // Only chart graphic frames belong here — pictures (image reader) and
    // plain shapes stay out of this family.
    if (!anchorXml.includes('<xdr:graphicFrame')) continue
    const chartRelId = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(anchorXml)?.[1]
    if (chartRelId === undefined) continue
    const chartPath = chartRels.get(chartRelId)
    if (chartPath === undefined) continue
    if (!(await reader.has(chartPath))) continue
    if (charts.length >= MAX_CHARTS_PER_SHEET) {
      throw new ChartReadError(
        `The sheet carries more than ${MAX_CHARTS_PER_SHEET} charts — charts are not available.`,
      )
    }
    const anchor = parseChartAnchor(anchorXml, kind as 'twoCellAnchor' | 'oneCellAnchor')
    if (anchor === null) continue
    let chartXml: string
    try {
      chartXml = await reader.readText(chartPath)
    } catch {
      continue
    }
    let state: ChartVisualState | null = null
    try {
      state = parseChartXml(chartXml)
    } catch (error) {
      if (!(error instanceof ChartReadError)) throw error
      // Per-chart fail closed: an unrepresentable chart never surfaces.
      state = null
    }
    if (state === null) continue
    charts.push({
      drawingPath,
      drawingIndex: index,
      chartPath,
      anchorType: anchor.anchorType,
      anchor: anchor.anchor,
      ...(anchor.widthPx !== undefined ? { widthPx: anchor.widthPx } : {}),
      ...(anchor.heightPx !== undefined ? { heightPx: anchor.heightPx } : {}),
      ...(anchor.name !== undefined && anchor.name !== '' ? { name: anchor.name } : {}),
      chart: state,
    })
  }
  return charts
}

/// worksheet rels → drawing part path (same two-step lookup the image
/// reader uses).
async function resolveDrawingPath(
  reader: ChartEntrySource,
  worksheetPath: string,
  relId: string,
): Promise<string | null> {
  const relsPath = worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
  if (!(await reader.has(relsPath))) return null
  const relsXml = await reader.readText(relsPath)
  const relationship = new RegExp(
    `<Relationship\\b[^>]*\\bId="${escapeRegExp(relId)}"[^>]*/?>`,
  ).exec(relsXml)?.[0]
  if (relationship === undefined) return null
  if (!relationship.includes(`Type="${DRAWING_REL_TYPE}"`)) return null
  const target = /\bTarget="([^"]+)"/.exec(relationship)?.[1]
  if (target === undefined) return null
  if (target.startsWith('/')) return target.slice(1)
  return resolveRelativePart(worksheetPath.split('/').slice(0, -1), target)
}

/// drawing rels → map of chart relationship id → resolved chart part path.
async function parseChartRelationships(
  reader: ChartEntrySource,
  drawingPath: string,
): Promise<Map<string, string>> {
  const relsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
  const map = new Map<string, string>()
  if (!(await reader.has(relsPath))) return map
  const relsXml = await reader.readText(relsPath)
  const base = drawingPath.split('/').slice(0, -1)
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const element = match[0]
    if (!element.includes(`Type="${CHART_REL_TYPE}"`)) continue
    const id = /\bId="([^"]+)"/.exec(element)?.[1]
    const target = /\bTarget="([^"]+)"/.exec(element)?.[1]
    if (id === undefined || target === undefined) continue
    if (/\bTargetMode="External"/.test(element)) continue
    map.set(id, resolveRelativePart(base, target))
  }
  return map
}

interface ParsedChartAnchor {
  readonly anchorType: 'two-cell' | 'one-cell'
  readonly anchor: DrawingAnchor
  readonly widthPx?: number | undefined
  readonly heightPx?: number | undefined
  readonly name?: string | undefined
}

function parseChartAnchor(
  anchorXml: string,
  kind: 'twoCellAnchor' | 'oneCellAnchor',
): ParsedChartAnchor | null {
  const name = /<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(anchorXml)?.[1]
  // graphicFrame carries its size in nvGraphicFramePr-sibling xfrm/ext
  // (a:ext cx/cy, EMU) — the one-cell chart renders at this fixed size.
  const extMatch = /<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(anchorXml)
  const widthPx = extMatch !== null ? Math.round(Number(extMatch[1]) / EMU_PER_PX) : undefined
  const heightPx = extMatch !== null ? Math.round(Number(extMatch[2]) / EMU_PER_PX) : undefined
  const from = parseMarker(anchorXml, 'from')
  if (from === null) return null
  if (kind === 'oneCellAnchor') {
    // One-cell charts move with their from marker but keep the a:ext size;
    // the canonical edit family rewrites only the from marker, so the
    // browser treats them as move-only (resize fails closed).
    return {
      anchorType: 'one-cell',
      anchor: {
        ...from,
        toRow: from.fromRow,
        toColumn: from.fromColumn,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      ...(widthPx !== undefined ? { widthPx } : {}),
      ...(heightPx !== undefined ? { heightPx } : {}),
      ...(name !== undefined ? { name } : {}),
    }
  }
  const to = parseMarker(anchorXml, 'to')
  if (to === null) return null
  return {
    anchorType: 'two-cell',
    anchor: {
      fromRow: from.fromRow,
      fromColumn: from.fromColumn,
      fromRowOffset: from.fromRowOffset,
      fromColumnOffset: from.fromColumnOffset,
      toRow: to.fromRow,
      toColumn: to.fromColumn,
      toRowOffset: to.fromRowOffset,
      toColumnOffset: to.fromColumnOffset,
    },
    ...(widthPx !== undefined ? { widthPx } : {}),
    ...(heightPx !== undefined ? { heightPx } : {}),
    ...(name !== undefined ? { name } : {}),
  }
}

/// Parses an <xdr:from> or <xdr:to> marker (col, colOff, row, rowOff).
function parseMarker(
  anchorXml: string,
  tag: 'from' | 'to',
): Omit<DrawingAnchor, 'toRow' | 'toColumn' | 'toRowOffset' | 'toColumnOffset'> | null {
  const section = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(anchorXml)
  if (section === null) return null
  const body = section[1] ?? ''
  const col = Number(/<xdr:col>(\d+)<\/xdr:col>/.exec(body)?.[1])
  const colOff = Number(/<xdr:colOff>(\d+)<\/xdr:colOff>/.exec(body)?.[1])
  const row = Number(/<xdr:row>(\d+)<\/xdr:row>/.exec(body)?.[1])
  const rowOff = Number(/<xdr:rowOff>(\d+)<\/xdr:rowOff>/.exec(body)?.[1])
  if (![col, colOff, row, rowOff].every((value) => Number.isInteger(value) && value >= 0)) {
    return null
  }
  return { fromRow: row, fromColumn: col, fromRowOffset: rowOff, fromColumnOffset: colOff }
}

// ── Chart part parsing (xl/charts/chartN.xml → ChartVisualState) ──────

const PLOT_PATTERN =
  /<c:(barChart|lineChart|areaChart|pieChart|doughnutChart|scatterChart|radarChart)>([\s\S]*?)<\/c:\1>/g

export function parseChartXml(chartXml: string): ChartVisualState | null {
  if (KNOWN_UNSUPPORTED_PLOTS.test(chartXml)) return null
  const plots = [...chartXml.matchAll(PLOT_PATTERN)]
  if (plots.length === 0) return null
  const plotNames = plots.map((plot) => plot[1] as SupportedPlot)
  const uniquePlots = new Set(plotNames)
  // Multi-plot support is the canonical combo ONLY: one barChart plot plus
  // one lineChart plot (series ordered bars-then-line, matching what
  // buildChartXml itself writes for 'combo').
  if (plots.length > 1) {
    const isCombo =
      plots.length === 2 && uniquePlots.has('barChart') && uniquePlots.has('lineChart')
    if (!isCombo) return null
  }
  if (uniquePlots.size !== plots.length && plots.length > 1) return null

  // Series: combo orders the bar plot's series first, then the line plot's.
  const orderedPlots =
    plots.length === 2 && plotNames[0] === 'lineChart' && plotNames[1] === 'barChart'
      ? [plots[1], plots[0]]
      : plots
  const series: ChartSeriesVisualState[] = []
  for (const plot of orderedPlots) {
    for (const serMatch of (plot[2] ?? '').matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)) {
      const parsed = parseSeries(serMatch[0], plot[1] as SupportedPlot)
      if (parsed !== null) series.push(parsed)
      if (series.length > MAX_SERIES_PER_CHART) return null
    }
  }
  if (series.length === 0) return null

  const barPlot = plots.find((plot) => plot[1] === 'barChart')
  const barDirection = /<c:barDir val="([^"]+)"/.exec(barPlot?.[2] ?? '')?.[1]
  const anyPlot = plots[0]?.[2] ?? ''
  const grouping = /<c:grouping val="([^"]+)"/.exec(anyPlot)?.[1]
  const gapWidthRaw = /<c:gapWidth val="(\d+)"/.exec(anyPlot)?.[1]
  const holeSizeRaw = /<c:holeSize val="(\d+)"/.exec(anyPlot)?.[1]
  const scatterStyle = /<c:scatterStyle val="([^"]+)"/.exec(anyPlot)?.[1]

  // Chart-level title: the first <c:title> before <c:plotArea>.
  const plotAreaAt = chartXml.search(/<c:plotArea[\s>]/)
  const chartScope = plotAreaAt === -1 ? chartXml : chartXml.slice(0, plotAreaAt)
  const titleElement = /<c:title>[\s\S]*?<\/c:title>/.exec(chartScope)
  const titleRuns = titleElement ? [...titleElement[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)] : []
  const title =
    titleRuns.length > 0
      ? decodeXmlText(titleRuns.map((run) => run[1] ?? '').join('')).slice(0, 255)
      : ''

  // Title style shorthand: c:title/c:txPr//a:defRPr.
  const titleStyle = parseTitleStyle(titleElement?.[0])

  // Plot-level data labels (the element the edit path writes — after the
  // last series, directly inside the plot).
  const plotLabels = /<\/c:ser>(<c:dLbls>[\s\S]*?<\/c:dLbls>)/.exec(anyPlot)?.[1]
  const dataLabels = parseDataLabels(plotLabels)
  const dataLabelPosition = parseDataLabelPosition(plotLabels)
  const dataLabelFormat =
    plotLabels !== undefined
      ? /<c:numFmt\b[^>]*\bformatCode="([^"]*)"/.exec(plotLabels)?.[1]
      : undefined

  // Legend.
  const legendElement = /<c:legend>[\s\S]*?<\/c:legend>/.exec(chartXml)
  let legend: ChartVisualState['legend']
  if (legendElement === null) {
    legend = undefined
  } else {
    const pos = /<c:legendPos val="([^"]+)"/.exec(legendElement[0])?.[1]
    legend =
      pos === 'b'
        ? 'bottom'
        : pos === 't'
          ? 'top'
          : pos === 'l'
            ? 'left'
            : pos === 'tr'
              ? 'right'
              : 'right'
  }

  // Axes: catAx/dateAx → x, valAx → y (first) and secondary (second).
  const catAx = /<c:(catAx|dateAx)>([\s\S]*?)<\/c:\1>/.exec(chartXml)
  const valAxes = [...chartXml.matchAll(/<c:valAx>([\s\S]*?)<\/c:valAx>/g)]
  const xAxis = catAx !== null ? parseAxisInfo(catAx[2] ?? '') : undefined
  const yAxis = valAxes[0] !== undefined ? parseAxisInfo(valAxes[0][1] ?? '') : undefined
  const secondaryYAxis =
    valAxes[1] !== undefined ? parseAxisInfo(valAxes[1][1] ?? '') : undefined
  const axisTitles =
    xAxis !== undefined || yAxis !== undefined
      ? {
          category: xAxis?.title ?? null,
          value: yAxis?.title ?? null,
        }
      : undefined

  const state: ChartVisualState = {
    chartTypes: orderedPlots.map((plot) => plot[1] as string),
    ...(barDirection !== undefined ? { barDirection } : {}),
    title,
    series,
    ...(legend !== undefined ? { legend } : {}),
    ...(axisTitles !== undefined ? { axisTitles } : {}),
    ...(dataLabels !== undefined ? { dataLabels } : {}),
    ...(dataLabelPosition !== undefined ? { dataLabelPosition } : {}),
    ...(dataLabelFormat !== undefined ? { dataLabelFormat } : {}),
    ...(grouping !== undefined ? { grouping: grouping as ChartVisualState['grouping'] } : {}),
    ...(yAxis !== undefined ? { gridlines: yAxis.majorGridlines } : {}),
    ...(yAxis !== undefined &&
    (yAxis.min !== undefined || yAxis.max !== undefined)
      ? {
          valueAxis: {
            ...(yAxis.min !== undefined ? { min: yAxis.min } : {}),
            ...(yAxis.max !== undefined ? { max: yAxis.max } : {}),
          },
        }
      : {}),
    ...(xAxis?.numFmt !== undefined ? { categoryAxisFormat: xAxis.numFmt } : {}),
    ...(gapWidthRaw !== undefined ? { gapWidthPct: Number(gapWidthRaw) } : {}),
    ...(holeSizeRaw !== undefined ? { holeSizePct: Number(holeSizeRaw) } : {}),
    ...(xAxis !== undefined ? { xAxis } : {}),
    ...(yAxis !== undefined ? { yAxis } : {}),
    ...(secondaryYAxis !== undefined ? { secondaryYAxis } : {}),
    ...(scatterStyle !== undefined ? { scatterStyle } : {}),
    ...(titleStyle !== undefined ? { titleStyle } : {}),
  }
  return state
}

/// One <c:ser> → ChartSeriesVisualState. Returns null for an unparseable
/// series (missing values block).
function parseSeries(serXml: string, plot: SupportedPlot): ChartSeriesVisualState | null {
  const name = parseSeriesName(serXml)
  const isScatter = plot === 'scatterChart'
  const valuesBlock = isScatter
    ? /<c:yVal>([\s\S]*?)<\/c:yVal>/.exec(serXml)?.[1]
    : /<c:val>([\s\S]*?)<\/c:val>/.exec(serXml)?.[1]
  if (valuesBlock === undefined) return null
  const valuesRef = /<c:f>([\s\S]*?)<\/c:f>/.exec(valuesBlock)?.[1]
  const values = parseNumericCache(valuesBlock)

  const categoriesBlock = isScatter
    ? /<c:xVal>([\s\S]*?)<\/c:xVal>/.exec(serXml)?.[1]
    : /<c:cat>([\s\S]*?)<\/c:cat>/.exec(serXml)?.[1]
  let categories: string[] = []
  let categoriesRef: string | undefined
  let categoryFormat: string | undefined
  if (categoriesBlock !== undefined) {
    categoriesRef = /<c:f>([\s\S]*?)<\/c:f>/.exec(categoriesBlock)?.[1]
    const format = /<c:formatCode>([\s\S]*?)<\/c:formatCode>/.exec(categoriesBlock)?.[1]
    if (format !== undefined && format !== 'General') categoryFormat = format
    const cached = parsePointCache(categoriesBlock)
    if (cached !== null) categories = cached.slice(0, MAX_POINTS_PER_SERIES)
  }

  const color = parseSolidFillColor(serXml)
  const lineInfo = parseLineColor(serXml)
  const pointColors = parsePointColors(serXml)
  const explosionRaw = /<c:explosion val="(\d+)"/.exec(serXml)?.[1]
  const pointExplosions = parsePointExplosions(serXml)
  const smoothRaw = /<c:smooth val="([^"]+)"/.exec(serXml)?.[1]
  const markerSymbol = /<c:marker><c:symbol val="([^"]+)"/.exec(serXml)?.[1]
  const numberFormatRaw = /<c:formatCode>([\s\S]*?)<\/c:formatCode>/.exec(valuesBlock)?.[1]

  return {
    name,
    categories,
    values: values.slice(0, MAX_POINTS_PER_SERIES),
    ...(numberFormatRaw !== undefined && numberFormatRaw !== 'General'
      ? { numberFormat: numberFormatRaw }
      : {}),
    ...(categoryFormat !== undefined ? { categoryFormat } : {}),
    ...(color !== undefined ? { color } : {}),
    ...(pointColors !== undefined ? { pointColors } : {}),
    ...(explosionRaw !== undefined ? { explosionPct: Number(explosionRaw) } : {}),
    ...(pointExplosions !== undefined ? { pointExplosions } : {}),
    ...(lineInfo !== undefined ? { lineColor: lineInfo } : {}),
    ...(smoothRaw !== undefined ? { smooth: smoothRaw === '1' } : {}),
    ...(markerSymbol !== undefined ? { marker: markerSymbol } : {}),
    ...(valuesRef !== undefined ? { valuesRef: decodeXmlText(valuesRef) } : {}),
    ...(categoriesRef !== undefined ? { categoriesRef: decodeXmlText(categoriesRef) } : {}),
  }
}

/// c:tx → series name (strRef cache, strLit, or a bare c:v literal).
function parseSeriesName(serXml: string): string {
  const tx = /<c:tx>([\s\S]*?)<\/c:tx>/.exec(serXml)?.[1]
  if (tx === undefined) return ''
  const strCache = /<c:strCache>([\s\S]*?)<\/c:strCache>/.exec(tx)?.[1]
  const scope = strCache ?? tx
  const values = [...scope.matchAll(/<c:pt\b[^>]*>(?:<c:v>([\s\S]*?)<\/c:v>)?<\/c:pt>/g)]
    .map((point) => point[1] ?? '')
    .join('')
  if (values !== '') return decodeXmlText(values).slice(0, 255)
  const literal = /<c:v>([\s\S]*?)<\/c:v>/.exec(tx)?.[1]
  return literal !== undefined ? decodeXmlText(literal).slice(0, 255) : ''
}

/// c:val/c:yVal block → numbers (numCache pts, or numLit). Files written
/// without a cache (openpyxl, pandas) carry only the c:f reference — the
/// series surfaces with EMPTY values plus its valuesRef, and the browser
/// hydrates from its live cell state exactly like the desktop
/// (WorkbookVisuals numCache hydration).
function parseNumericCache(block: string): number[] {
  const cache = /<c:numCache>([\s\S]*?)<\/c:numCache>/.exec(block)?.[1]
  const scope = cache ?? /<c:numLit>([\s\S]*?)<\/c:numLit>/.exec(block)?.[1]
  if (scope === undefined) return []
  const values: number[] = []
  for (const point of scope.matchAll(/<c:pt\b[^>]*>(?:<c:v>([\s\S]*?)<\/c:v>)?<\/c:pt>/g)) {
    const raw = (point[1] ?? '').trim()
    if (raw === '') continue
    const numeric = Number(raw)
    if (!Number.isFinite(numeric)) continue
    values.push(numeric)
  }
  return values
}

/// c:cat/c:xVal block → string labels (strCache/numCache/strLit/numLit).
function parsePointCache(block: string): string[] | null {
  const cache =
    /<c:strCache>([\s\S]*?)<\/c:strCache>/.exec(block)?.[1] ??
    /<c:numCache>([\s\S]*?)<\/c:numCache>/.exec(block)?.[1]
  const literal =
    /<c:strLit>([\s\S]*?)<\/c:strLit>/.exec(block)?.[1] ??
    /<c:numLit>([\s\S]*?)<\/c:numLit>/.exec(block)?.[1]
  const scope = cache ?? literal
  if (scope === undefined) return null
  const values: string[] = []
  for (const point of scope.matchAll(/<c:pt\b[^>]*>(?:<c:v>([\s\S]*?)<\/c:v>)?<\/c:pt>/g)) {
    values.push(decodeXmlText(point[1] ?? '').slice(0, 255))
  }
  return values
}

/// First spPr solidFill srgbClr → '#RRGGBB'. Theme scheme colors are left
/// unresolved (undefined) — the renderer's palette fallback applies.
function parseSolidFillColor(serXml: string): string | undefined {
  const spPr = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(serXml)?.[1]
  if (spPr === undefined) return undefined
  const fill = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(spPr)
  return fill !== null && fill[1] !== undefined ? `#${fill[1].toUpperCase()}` : undefined
}

/// spPr/a:ln color — srgbClr → '#RRGGBB', explicit noFill → 'none'.
function parseLineColor(serXml: string): string | undefined {
  const spPr = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(serXml)?.[1]
  if (spPr === undefined) return undefined
  const ln = /<a:ln\b[\s\S]*?<\/a:ln>/.exec(spPr)?.[0]
  if (ln === undefined) return undefined
  if (/<a:noFill\/>/.test(ln)) return 'none'
  const fill = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(ln)
  return fill !== null && fill[1] !== undefined ? `#${fill[1].toUpperCase()}` : undefined
}

/// c:dPt entries → per-point fills (idx + spPr fill), ascending by idx.
function parsePointColors(
  serXml: string,
): { index: number; color: string }[] | undefined {
  const entries: { index: number; color: string }[] = []
  for (const dpt of serXml.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)) {
    const idx = Number(/<c:idx val="(\d+)"/.exec(dpt[0])?.[1])
    if (!Number.isInteger(idx) || idx < 0) continue
    const color = parseSolidFillColor(dpt[0])
    if (color === undefined) continue
    entries.push({ index: idx, color })
  }
  if (entries.length === 0) return undefined
  return entries.sort((left, right) => left.index - right.index)
}

/// c:dPt c:explosion → per-slice explosion overrides, ascending by idx.
function parsePointExplosions(
  serXml: string,
): { index: number; pct: number }[] | undefined {
  const entries: { index: number; pct: number }[] = []
  for (const dpt of serXml.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)) {
    const idx = Number(/<c:idx val="(\d+)"/.exec(dpt[0])?.[1])
    const pct = Number(/<c:explosion val="(\d+)"/.exec(dpt[0])?.[1])
    if (!Number.isInteger(idx) || idx < 0 || !Number.isFinite(pct)) continue
    entries.push({ index: idx, pct })
  }
  if (entries.length === 0) return undefined
  return entries.sort((left, right) => left.index - right.index)
}

/// Plot-level dLbls → the canonical label mode.
function parseDataLabels(plotLabels: string | undefined): ChartVisualState['dataLabels'] {
  if (plotLabels === undefined) return undefined
  const showVal = /<c:showVal val="1"\/>/.test(plotLabels)
  const showPercent = /<c:showPercent val="1"\/>/.test(plotLabels)
  const showCatName = /<c:showCatName val="1"\/>/.test(plotLabels)
  if (showVal) return 'value'
  if (showPercent && showCatName) return 'category-percent'
  if (showPercent) return 'percent'
  return undefined
}

function parseDataLabelPosition(
  plotLabels: string | undefined,
): ChartVisualState['dataLabelPosition'] {
  if (plotLabels === undefined) return undefined
  const pos = /<c:dLblPos val="([^"]+)"/.exec(plotLabels)?.[1]
  if (pos === 'ctr') return 'center'
  if (pos === 'inEnd') return 'inside-end'
  if (pos === 'outEnd') return 'outside-end'
  return undefined
}

/// One axis element body → ChartAxisInfoState.
function parseAxisInfo(axisXml: string): ChartAxisInfoState {
  const titleRuns = [...axisXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
  const title =
    titleRuns.length > 0
      ? decodeXmlText(titleRuns.map((run) => run[1] ?? '').join('')).slice(0, 255)
      : undefined
  const numFmtElement = /<c:numFmt\b([^>]*)\/?>/.exec(axisXml)?.[1] ?? ''
  const sourceLinked = /sourceLinked="1"/.test(numFmtElement)
  const formatCode = /formatCode="([^"]*)"/.exec(numFmtElement)?.[1]
  const min = Number(/<c:min val="(-?[\d.eE+]+)"/.exec(axisXml)?.[1])
  const max = Number(/<c:max val="(-?[\d.eE+]+)"/.exec(axisXml)?.[1])
  const majorUnit = Number(/<c:majorUnit val="(-?[\d.eE+]+)"/.exec(axisXml)?.[1])
  return {
    ...(title !== undefined && title !== '' ? { title } : {}),
    ...(Number.isFinite(min) ? { min } : {}),
    ...(Number.isFinite(max) ? { max } : {}),
    ...(Number.isFinite(majorUnit) ? { majorUnit } : {}),
    ...(formatCode !== undefined && !sourceLinked && formatCode !== 'General'
      ? { numFmt: formatCode }
      : {}),
    majorGridlines: /<c:majorGridlines\b/.test(axisXml),
    hidden: /<c:delete val="1"\/>/.test(axisXml),
  }
}

function parseTitleStyle(titleElement: string | undefined):
  | { size?: number | undefined; bold?: boolean | undefined; color?: string | undefined }
  | undefined {
  if (titleElement === undefined) return undefined
  const defRPr = /<c:txPr>[\s\S]*?<a:defRPr\b([^>]*)\/?>/.exec(titleElement)?.[1]
  if (defRPr === undefined) return undefined
  const sizeAttr = /\bsz="(\d+)"/.exec(defRPr)?.[1]
  const bold = /\bb="1"/.test(defRPr)
  const colorMatch = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(
    /<c:txPr>[\s\S]*?<a:defRPr\b[^>]*>([\s\S]*?)<\/a:defRPr>/.exec(titleElement)?.[1] ?? '',
  )
  const style: { size?: number; bold?: boolean; color?: string } = {}
  if (sizeAttr !== undefined) style.size = Number(sizeAttr) / 100
  if (bold) style.bold = true
  if (colorMatch !== null && colorMatch[1] !== undefined) {
    style.color = `#${colorMatch[1].toUpperCase()}`
  }
  return Object.keys(style).length > 0 ? style : undefined
}

function decodeXmlText(input: string): string {
  return input
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_full, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&amp;/g, '&')
}

function resolveRelativePart(base: readonly string[], target: string): string {
  const segments = [...base]
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.') segments.push(part)
  }
  return segments.join('/')
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
