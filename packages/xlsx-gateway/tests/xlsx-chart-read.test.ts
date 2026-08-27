import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'
import type { CellEdit } from '../src/index.js'
import type { WorkbookVisualEdit } from '../src/types.js'

/// ── Fixture kit: a workbook whose sheet carries a drawing part with
///    configurable chart graphic-frame anchors and chart parts.

interface ChartAnchorSpec {
  kind: 'twoCellAnchor' | 'oneCellAnchor' | 'absoluteAnchor'
  chartRelId?: string
  name?: string
  from?: { col: number; colOff: number; row: number; rowOff: number }
  to?: { col: number; colOff: number; row: number; rowOff: number }
  ext?: { cx: number; cy: number }
  picture?: boolean
  plainShape?: boolean
}

function marker(m: { col: number; colOff: number; row: number; rowOff: number }): string {
  return (
    `<xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff>` +
    `<xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff>`
  )
}

function chartAnchorXml(spec: ChartAnchorSpec): string {
  const relId = spec.chartRelId ?? 'rIdChart1'
  const name = spec.name ?? 'Chart 1'
  const ext = spec.ext ?? { cx: 914400, cy: 609600 }
  const body = spec.picture
    ? '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Pic"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdPic"/></xdr:blipFill>' +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></a:xfrm></xdr:spPr></xdr:pic>`
    : spec.plainShape
      ? '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="9" name="Shape"/></xdr:nvSpPr></xdr:sp>'
      : '<xdr:graphicFrame macro="">' +
        `<xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="${name}"/></xdr:nvGraphicFramePr></xdr:nvGraphicFramePr>` +
        `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></xdr:xfrm>` +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/>` +
        '</a:graphicData></a:graphic></xdr:graphicFrame>'
  return (
    `<xdr:${spec.kind}>` +
    (spec.kind === 'absoluteAnchor'
      ? `<xdr:pos x="47625" y="9525"/><xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>`
      : `<xdr:from>${marker(spec.from ?? { col: 1, colOff: 0, row: 2, rowOff: 0 })}</xdr:from>` +
        (spec.kind === 'twoCellAnchor'
          ? `<xdr:to>${marker(spec.to ?? { col: 8, colOff: 0, row: 18, rowOff: 0 })}</xdr:to>`
          : `<xdr:ext cx="${ext.cx}" cy="${ext.cy}"/>`)) +
    body +
    '<xdr:clientData/></xdr:' +
    spec.kind +
    '>'
  )
}

const CAT_CACHE =
  '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>' +
  '<c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>' +
  '<c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>'

function seriesXml(
  name: string,
  values: readonly number[],
  options: {
    color?: string
    valuesRef?: string
    numberFormat?: string
    dPt?: { idx: number; color: string; explosion?: number }
  } = {},
): string {
  const nameRef =
    '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/>' +
    `<c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
  const spPr =
    options.color !== undefined
      ? `<c:spPr><a:solidFill><a:srgbClr val="${options.color}"/></a:solidFill></c:spPr>`
      : ''
  const dPt =
    options.dPt !== undefined
      ? `<c:dPt><c:idx val="${options.dPt.idx}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
        `<c:spPr><a:solidFill><a:srgbClr val="${options.dPt.color}"/></a:solidFill></c:spPr>` +
        (options.dPt.explosion !== undefined
          ? `<c:explosion val="${options.dPt.explosion}"/>`
          : '') +
        '</c:dPt>'
      : ''
  const formatCode =
    options.numberFormat !== undefined
      ? `<c:formatCode>${options.numberFormat}</c:formatCode>`
      : '<c:formatCode>General</c:formatCode>'
  const valuesRef = options.valuesRef ?? 'Sheet1!$B$2:$B$4'
  const valuePoints = values
    .map((value, idx) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`)
    .join('')
  return (
    `<c:ser><c:idx val="0"/><c:order val="0"/>${nameRef}${spPr}${dPt}${CAT_CACHE}` +
    `<c:val><c:numRef><c:f>${valuesRef}</c:f><c:numCache>${formatCode}` +
    `<c:ptCount val="${values.length}"/>${valuePoints}</c:numCache></c:numRef></c:val></c:ser>`
  )
}

function barChartXml(
  series: readonly string[],
  options: {
    title?: string
    legendPos?: string
    dLbls?: string
    grouping?: string
    gapWidth?: number
    catTitle?: string
    valTitle?: string
    valAxisScaling?: string
    valMajorUnit?: number
    valNumFmt?: { formatCode: string; sourceLinked: boolean }
    catNumFmt?: { formatCode: string; sourceLinked: boolean }
  } = {},
): string {
  const title =
    options.title !== undefined
      ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
        `<a:t>${options.title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
      : ''
  const legend =
    options.legendPos !== undefined
      ? `<c:legend><c:legendPos val="${options.legendPos}"/></c:legend>`
      : ''
  const catAxisNumFmt = options.catNumFmt
    ? `<c:numFmt formatCode="${options.catNumFmt.formatCode}" sourceLinked="${options.catNumFmt.sourceLinked ? 1 : 0}"/>`
    : ''
  const valAxisNumFmt = options.valNumFmt
    ? `<c:numFmt formatCode="${options.valNumFmt.formatCode}" sourceLinked="${options.valNumFmt.sourceLinked ? 1 : 0}"/>`
    : ''
  const catTitle =
    options.catTitle !== undefined
      ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${options.catTitle}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
      : ''
  const valTitle =
    options.valTitle !== undefined
      ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${options.valTitle}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
      : ''
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:chart>' +
    title +
    '<c:autoTitleDeleted val="0"/>' +
    '<c:plotArea><c:layout/>' +
    '<c:barChart><c:barDir val="col"/>' +
    `<c:grouping val="${options.grouping ?? 'clustered'}"/>` +
    '<c:varyColors val="0"/>' +
    series.join('') +
    (options.dLbls ?? '') +
    `<c:gapWidth val="${options.gapWidth ?? 150}"/>` +
    '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
    '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    `<c:delete val="0"/><c:axPos val="b"/>${catTitle}${catAxisNumFmt}<c:crossAx val="222"/></c:catAx>` +
    '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/>' +
    (options.valAxisScaling ?? '') +
    '</c:scaling>' +
    `<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>${valTitle}${valAxisNumFmt}` +
    (options.valMajorUnit !== undefined
      ? `<c:majorUnit val="${options.valMajorUnit}"/>`
      : '') +
    '<c:crossAx val="111"/></c:valAx>' +
    '</c:plotArea>' +
    legend +
    '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
  )
}

function pieChartXml(
  series: readonly string[],
  options: { title?: string; hole?: number; explosion?: number } = {},
): string {
  const title =
    options.title !== undefined
      ? '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r>' +
        `<a:t>${options.title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
      : ''
  const plot = options.hole !== undefined ? 'doughnutChart' : 'pieChart'
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:chart>' +
    title +
    '<c:autoTitleDeleted val="0"/>' +
    '<c:plotArea><c:layout/>' +
    `<c:${plot}><c:varyColors val="1"/>` +
    series
      .map(
        (ser) =>
          ser +
          (options.explosion !== undefined
            ? `<c:explosion val="${options.explosion}"/>`
            : ''),
      )
      .join('') +
    '<c:firstSliceAng val="0"/>' +
    (options.hole !== undefined ? `<c:holeSize val="${options.hole}"/>` : '') +
    `</c:${plot}></c:plotArea>` +
    '<c:legend><c:legendPos val="r"/></c:legend>' +
    '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
  )
}

interface ChartFixtureOptions {
  anchors?: ChartAnchorSpec[]
  chartParts?: Record<string, string>
  chartRels?: string[]
  drawingRels?: string[]
}

async function buildChartFixture(options: ChartFixtureOptions = {}): Promise<Buffer> {
  const anchors = options.anchors ?? [{ kind: 'twoCellAnchor' as const }]
  const chartParts = options.chartParts ?? {
    'xl/charts/chart1.xml': barChartXml(
      [
        seriesXml('Revenue', [10, 20, 30], { color: '4472C4' }),
        seriesXml('Cost', [5, 8, 11], { color: 'ED7D31', valuesRef: 'Sheet1!$C$2:$C$4' }),
      ],
      { title: 'Sales &amp; Cost', legendPos: 'r' },
    ),
  }
  const chartRels = options.chartRels ?? []
  const drawingRels =
    options.drawingRels ??
    Array.from(new Set(anchors.map((anchor) => anchor.chartRelId ?? 'rIdChart1'))).map(
      (relId, index) =>
        `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${index + 1}.xml"/>`,
    )
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData>
  <drawing r:id="rIdDrawing1"/>
</worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/drawings/drawing1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors
      .map((anchor) => chartAnchorXml(anchor))
      .join('')}</xdr:wsDr>`,
  )
  zip.file(
    'xl/drawings/_rels/drawing1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRels.join('')}</Relationships>`,
  )
  for (const [path, xml] of Object.entries(chartParts)) {
    zip.file(path, xml)
  }
  if (chartRels.length > 0) {
    zip.file(
      'xl/charts/_rels/chart1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${chartRels.join('')}</Relationships>`,
    )
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function readEntry(buffer: Buffer, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  return entry === null ? null : await entry.async('string')
}

async function listEntries(buffer: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer)
  return Object.entries(zip.files)
    .filter(([, file]) => !file.dir)
    .map(([path]) => path)
}

async function entryBytes(buffer: Buffer, path: string): Promise<Buffer | null> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  return entry === null ? null : await entry.async('nodebuffer')
}

const NO_CELL_EDITS: readonly CellEdit[] = []

/// ── Read path ────────────────────────────────────────────────────────

describe('chart read: canonical parseSheetCharts', () => {
  it('surfaces a two-cell chart with anchor, both locators, and full state', async () => {
    const buffer = await buildChartFixture()
    const imported = await readBasicWorkbook(buffer)
    const charts = imported.snapshot.sheets[0]?.charts ?? []
    expect(charts).toHaveLength(1)
    const chart = charts[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.drawingPath).toBe('xl/drawings/drawing1.xml')
    expect(chart.drawingIndex).toBe(0)
    expect(chart.chartPath).toBe('xl/charts/chart1.xml')
    expect(chart.anchorType).toBe('two-cell')
    expect(chart.anchor.fromColumn).toBe(1)
    expect(chart.anchor.fromRow).toBe(2)
    expect(chart.anchor.toColumn).toBe(8)
    expect(chart.anchor.toRow).toBe(18)
    expect(chart.chart.chartTypes).toEqual(['barChart'])
    expect(chart.chart.barDirection).toBe('col')
    expect(chart.chart.title).toBe('Sales & Cost')
    expect(chart.chart.series).toHaveLength(2)
    expect(chart.chart.series[0]?.name).toBe('Revenue')
    expect(chart.chart.series[0]?.values).toEqual([10, 20, 30])
    expect(chart.chart.series[0]?.valuesRef).toBe('Sheet1!$B$2:$B$4')
    expect(chart.chart.series[0]?.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(chart.chart.series[0]?.categoriesRef).toBe('Sheet1!$A$2:$A$4')
    expect(chart.chart.series[0]?.color).toBe('#4472C4')
    expect(chart.chart.series[1]?.color).toBe('#ED7D31')
    expect(chart.chart.grouping).toBe('clustered')
    expect(chart.chart.gapWidthPct).toBe(150)
    expect(chart.chart.gridlines).toBe(true)
  })

  it('reads title, legend, axis titles, explicit axis bounds, and label flags', async () => {
    const buffer = await buildChartFixture({
      chartParts: {
        'xl/charts/chart1.xml': barChartXml(
          [seriesXml('Revenue', [10, 20, 30])],
          {
            title: 'Quarterly Sales',
            legendPos: 'b',
            dLbls:
              '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showPercent val="0"/></c:dLbls>',
            catTitle: 'Quarter',
            valTitle: 'USD',
            valAxisScaling: '<c:min val="0"/><c:max val="100"/>',
            valMajorUnit: 20,
            valNumFmt: { formatCode: '#,##0', sourceLinked: false },
          },
        ),
      },
    })
    const imported = await readBasicWorkbook(buffer)
    const chart = imported.snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.chart.title).toBe('Quarterly Sales')
    expect(chart.chart.legend).toBe('bottom')
    expect(chart.chart.dataLabels).toBe('value')
    expect(chart.chart.axisTitles?.category).toBe('Quarter')
    expect(chart.chart.axisTitles?.value).toBe('USD')
    expect(chart.chart.valueAxis).toEqual({ min: 0, max: 100 })
    expect(chart.chart.yAxis?.majorUnit).toBe(20)
    expect(chart.chart.yAxis?.numFmt).toBe('#,##0')
    expect(chart.chart.yAxis?.majorGridlines).toBe(true)
    expect(chart.chart.xAxis?.title).toBe('Quarter')
  })

  it('reads pie slice colors, explosion, and doughnut hole size', async () => {
    const buffer = await buildChartFixture({
      chartParts: {
        'xl/charts/chart1.xml': pieChartXml(
          [
            seriesXml('Share', [40, 35, 25], {
              dPt: { idx: 1, color: 'FF0000', explosion: 25 },
            }),
          ],
          { title: 'Share', hole: 55 },
        ),
      },
    })
    const imported = await readBasicWorkbook(buffer)
    const chart = imported.snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.chart.chartTypes).toEqual(['doughnutChart'])
    expect(chart.chart.holeSizePct).toBe(55)
    expect(chart.chart.series[0]?.pointColors).toEqual([{ index: 1, color: '#FF0000' }])
    expect(chart.chart.series[0]?.pointExplosions).toEqual([{ index: 1, pct: 25 }])
    expect(chart.chart.legend).toBe('right')
  })

  it('reads a scatter chart through xVal/yVal with scatterStyle', async () => {
    const scatterSeries =
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Points</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
      '<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>' +
      '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:xVal>' +
      '<c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>' +
      '<c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:yVal>' +
      '</c:ser>'
    const scatterXml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<c:chart><c:plotArea><c:layout/>' +
      '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' +
      scatterSeries +
      '<c:axId val="111"/><c:axId val="222"/></c:scatterChart>' +
      '<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' +
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>' +
      '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
    const buffer = await buildChartFixture({
      chartParts: { 'xl/charts/chart1.xml': scatterXml },
    })
    const imported = await readBasicWorkbook(buffer)
    const chart = imported.snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.chart.chartTypes).toEqual(['scatterChart'])
    expect(chart.chart.scatterStyle).toBe('lineMarker')
    expect(chart.chart.series[0]?.values).toEqual([2, 4, 9])
    expect(chart.chart.series[0]?.categories).toEqual(['1', '2', '3'])
    expect(chart.chart.series[0]?.valuesRef).toBe('Sheet1!$B$2:$B$4')
    expect(chart.chart.series[0]?.categoriesRef).toBe('Sheet1!$A$2:$A$4')
  })

  it('reads the canonical bar+line combo with bars-then-line series order', async () => {
    const comboXml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<c:chart><c:plotArea><c:layout/>' +
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
      seriesXml('Revenue', [10, 20, 30]) +
      seriesXml('Cost', [5, 8, 11]) +
      '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
      '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
      seriesXml('Margin', [5, 12, 19]).replace(
        '<c:idx val="0"/>',
        '<c:idx val="2"/>',
      ) +
      '<c:marker val="1"/><c:axId val="111"/><c:axId val="333"/></c:lineChart>' +
      '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>' +
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>' +
      '<c:valAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="111"/></c:valAx>' +
      '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
    const buffer = await buildChartFixture({
      chartParts: { 'xl/charts/chart1.xml': comboXml },
    })
    const imported = await readBasicWorkbook(buffer)
    const chart = imported.snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.chart.chartTypes).toEqual(['barChart', 'lineChart'])
    expect(chart.chart.series.map((series) => series.name)).toEqual([
      'Revenue',
      'Cost',
      'Margin',
    ])
    expect(chart.chart.secondaryYAxis).toBeDefined()
  })

  it('surfaces numCache-less series (openpyxl) with refs for browser hydration', async () => {
    const refOnlySeries =
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>' +
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>'
    const buffer = await buildChartFixture({
      chartParts: {
        'xl/charts/chart1.xml': barChartXml([refOnlySeries]),
      },
    })
    const imported = await readBasicWorkbook(buffer)
    const chart = imported.snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.chart.series[0]?.values).toEqual([])
    expect(chart.chart.series[0]?.valuesRef).toBe('Sheet1!$B$2:$B$4')
    expect(chart.chart.series[0]?.categories).toEqual([])
    expect(chart.chart.series[0]?.categoriesRef).toBe('Sheet1!$A$2:$A$4')
  })

  it('keeps drawingIndex parity across pictures, shapes, and skipped charts', async () => {
    const buffer = await buildChartFixture({
      anchors: [
        { kind: 'twoCellAnchor', picture: true },
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart1' },
        { kind: 'twoCellAnchor', plainShape: true },
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart2' },
      ],
      chartParts: {
        'xl/charts/chart1.xml': barChartXml([seriesXml('Revenue', [10, 20, 30])]),
        'xl/charts/chart2.xml': pieChartXml([seriesXml('Share', [1, 2, 3])]),
      },
      drawingRels: [
        '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
        '<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>',
      ],
    })
    const imported = await readBasicWorkbook(buffer)
    const charts = imported.snapshot.sheets[0]?.charts ?? []
    expect(charts.map((chart) => chart.drawingIndex)).toEqual([1, 3])
    expect(charts[0]?.chartPath).toBe('xl/charts/chart1.xml')
    expect(charts[1]?.chartPath).toBe('xl/charts/chart2.xml')
  })

  it('omits absolute-anchored charts but keeps later locators stable', async () => {
    const buffer = await buildChartFixture({
      anchors: [
        { kind: 'absoluteAnchor', chartRelId: 'rIdChart1' },
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart2' },
      ],
      chartParts: {
        'xl/charts/chart1.xml': barChartXml([seriesXml('A', [1, 2, 3])]),
        'xl/charts/chart2.xml': barChartXml([seriesXml('B', [4, 5, 6])]),
      },
      drawingRels: [
        '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
        '<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>',
      ],
    })
    const imported = await readBasicWorkbook(buffer)
    const charts = imported.snapshot.sheets[0]?.charts ?? []
    expect(charts).toHaveLength(1)
    expect(charts[0]?.drawingIndex).toBe(1)
    expect(charts[0]?.chartPath).toBe('xl/charts/chart2.xml')
  })

  it('omits unsupported chart families (3-D, bubble) per chart', async () => {
    const threeDee =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<c:chart><c:plotArea><c:layout/>' +
      '<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
      seriesXml('Revenue', [1, 2, 3]) +
      '<c:axId val="111"/><c:axId val="222"/></c:bar3DChart>' +
      '</c:plotArea></c:chart></c:chartSpace>'
    const buffer = await buildChartFixture({
      anchors: [
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart1' },
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart2' },
      ],
      chartParts: {
        'xl/charts/chart1.xml': threeDee,
        'xl/charts/chart2.xml': barChartXml([seriesXml('B', [4, 5, 6])]),
      },
      drawingRels: [
        '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
        '<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>',
      ],
    })
    const imported = await readBasicWorkbook(buffer)
    const charts = imported.snapshot.sheets[0]?.charts ?? []
    expect(charts).toHaveLength(1)
    expect(charts[0]?.chartPath).toBe('xl/charts/chart2.xml')
    expect(charts[0]?.drawingIndex).toBe(1)
  })

  it('carries one-cell charts as move-only with explicit pixel size', async () => {
    const buffer = await buildChartFixture({
      anchors: [{ kind: 'oneCellAnchor', ext: { cx: 476250, cy: 285750 } }],
    })
    const imported = await readBasicWorkbook(buffer)
    const chart = imported.snapshot.sheets[0]?.charts?.[0]
    expect(chart).toBeDefined()
    if (!chart) return
    expect(chart.anchorType).toBe('one-cell')
    expect(chart.widthPx).toBe(50)
    expect(chart.heightPx).toBe(30)
    expect(chart.anchor.toColumn).toBe(chart.anchor.fromColumn)
    expect(chart.anchor.toRow).toBe(chart.anchor.fromRow)
  })

  it('omits non-combo multi-plot combinations', async () => {
    const dualBar =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<c:chart><c:plotArea><c:layout/>' +
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
      seriesXml('A', [1, 2, 3]) +
      '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
      seriesXml('B', [4, 5, 6]) +
      '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
      '</c:plotArea></c:chart></c:chartSpace>'
    const buffer = await buildChartFixture({
      chartParts: { 'xl/charts/chart1.xml': dualBar },
    })
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]?.charts ?? []).toHaveLength(0)
  })

  it('surfaces no charts when the drawing relationship is unresolvable (fail closed)', async () => {
    const buffer = await buildChartFixture({
      drawingRels: [
        '<Relationship Id="rIdOther" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>',
      ],
    })
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]?.charts ?? []).toHaveLength(0)
    expect(imported.snapshot.sheets).toHaveLength(1)
  })
})

/// ── Save integration: the read locators drive the canonical families ──

interface MutateOptions {
  chartEdits?: readonly import('../src/types.js').WorkbookChartEdit[]
  visualAdditions?: readonly import('../src/gateway/xlsx-gateway.js').SheetVisualAddition[]
  visualEdits?: readonly WorkbookVisualEdit[]
}

async function mutate(buffer: Buffer, options: MutateOptions = {}) {
  return applyCellEditsToXlsx(
    buffer,
    NO_CELL_EDITS,
    [],
    options.chartEdits ?? [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    null,
    [],
    options.visualAdditions ?? [],
    options.visualEdits ?? [],
  )
}

describe('chart read: locator integration with the save families', () => {
  it('a chartEdit against the read chartPath patches the chart part', async () => {
    const buffer = await buildChartFixture()
    const mutated = await mutate(buffer, {
      chartEdits: [{ chartPath: 'xl/charts/chart1.xml', title: 'Renamed' }],
    })
    const chartXml = await readEntry(mutated.buffer, 'xl/charts/chart1.xml')
    expect(chartXml).toContain('Renamed')
    expect(chartXml).toContain('Q1')
  })

  it('a visualEdit remove against the read locator cascades the chart part', async () => {
    const buffer = await buildChartFixture({
      chartRels: [
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>',
      ],
    })
    const mutated = await mutate(buffer, {
      visualEdits: [
        { drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, remove: true },
      ],
    })
    const entries = await listEntries(mutated.buffer)
    expect(entries).not.toContain('xl/charts/chart1.xml')
    expect(entries).not.toContain('xl/charts/_rels/chart1.xml.rels')
    const rels = await readEntry(mutated.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(rels).not.toContain('rIdChart1')
    const contentTypes = await readEntry(mutated.buffer, '[Content_Types].xml')
    expect(contentTypes).not.toContain('chart1.xml')
  })

  it('a no-op save preserves the chart part, drawing, rels, and content types byte-for-byte', async () => {
    const buffer = await buildChartFixture({
      chartRels: [
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>',
      ],
    })
    const mutated = await mutate(buffer)
    for (const path of [
      'xl/charts/chart1.xml',
      'xl/charts/_rels/chart1.xml.rels',
      'xl/drawings/drawing1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
      '[Content_Types].xml',
    ]) {
      expect(mutated.touchedEntries).not.toContain(path)
      const before = await entryBytes(buffer, path)
      const after = await entryBytes(mutated.buffer, path)
      expect(after).not.toBeNull()
      expect(before?.equals(after ?? Buffer.alloc(0))).toBe(true)
    }
  })

  it('editing one chart leaves an unrelated chart byte-identical', async () => {
    const buffer = await buildChartFixture({
      anchors: [
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart1' },
        { kind: 'twoCellAnchor', chartRelId: 'rIdChart2' },
      ],
      chartParts: {
        'xl/charts/chart1.xml': barChartXml([seriesXml('Revenue', [10, 20, 30])]),
        'xl/charts/chart2.xml': pieChartXml([seriesXml('Share', [1, 2, 3])]),
      },
      drawingRels: [
        '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
        '<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>',
      ],
    })
    const mutated = await mutate(buffer, {
      chartEdits: [{ chartPath: 'xl/charts/chart1.xml', title: 'Edited' }],
    })
    const untouched = await entryBytes(mutated.buffer, 'xl/charts/chart2.xml')
    const original = await entryBytes(buffer, 'xl/charts/chart2.xml')
    expect(untouched?.equals(original ?? Buffer.alloc(0))).toBe(true)
    const edited = await readEntry(mutated.buffer, 'xl/charts/chart1.xml')
    expect(edited).toContain('Edited')
  })

  it('a chart creation persists and returns its locator with the chart part path', async () => {
    const buffer = await buildChartFixture()
    const mutated = await mutate(buffer, {
      visualAdditions: [
        {
          sheetName: 'Sheet1',
          anchor: {
            fromRow: 20,
            fromColumn: 1,
            fromRowOffset: 0,
            fromColumnOffset: 0,
            toRow: 30,
            toColumn: 10,
            toRowOffset: 0,
            toColumnOffset: 0,
          },
          chart: {
            chartType: 'column',
            title: 'Created',
            series: [{ name: 'S1', categories: ['a', 'b'], values: [1, 2] }],
          },
        },
      ],
    })
    expect(mutated.addedVisuals).toHaveLength(1)
    expect(mutated.addedVisuals?.[0]?.chartPath).toMatch(/^xl\/charts\/chart\d+\.xml$/)
    const drawing = await readEntry(mutated.buffer, 'xl/drawings/drawing1.xml')
    expect(drawing).toContain('graphicFrame')
    const contentTypes = await readEntry(mutated.buffer, '[Content_Types].xml')
    expect(contentTypes).toContain(mutated.addedVisuals?.[0]?.chartPath ?? '')
  })
})
