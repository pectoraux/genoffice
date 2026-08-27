/// EXCEL-023: the web's chart domain surface. The canonical chart domain
/// (ChartVisualState, ChartStateEdit overlay merge, Excel-fidelity axis
/// scales, selection→series parsing, insert-visual assembly) lives in
/// @genoffice/xlsx-gateway's PURE domain module — no archive code, no XML,
/// no relationship work. The desktop consumes the exact same module
/// (apps/sheets/src/domain/chart-visual.ts re-exports it verbatim); the
/// web follows that precedent so the two renderers can never drift.
///
/// ARCHITECTURE BOUNDARY: value imports from the gateway package ROOT
/// remain forbidden (the engine must never bundle into the browser) — the
/// architecture tests permit exactly this one pure-domain subpath.
export {
  applyChartStateEdit,
  buildChartVisual,
  chartCategoryFormat,
  chartDataFromValues,
  chartSupportsDataLabels,
  chartSupportsSeriesReplace,
  formatCategoryLabel,
  scatterAxisBounds,
  valueAxisScale,
  withDefaultBarLabels,
  CHART_EDIT_TYPES,
} from '@genoffice/xlsx-gateway/src/domain/chart-visual.js'
export type {
  BuildChartVisualInput,
  ChartAxisInfoState,
  ChartGridValue,
  ChartSeriesVisualState,
  ChartStateEdit,
  ChartVisualState,
  ParsedChartData,
  SheetVisual,
  SheetVisualAnchor,
} from '@genoffice/xlsx-gateway/src/domain/chart-visual.js'
