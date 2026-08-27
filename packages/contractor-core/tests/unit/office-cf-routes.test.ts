/**
 * Unit tests — /office/workbooks/save request validation for the cfStates
 * family (EXCEL-024 — Home → Conditional Formatting).
 *
 * Proves routeOffice accepts ONLY strictly-typed canonical SheetCfState
 * payloads (the Univer wire shape the gateway's writer consumes) and
 * rejects unknown rule families, unsupported operators, time-period rules,
 * x14-only icon sets, mixed icon sets, malformed ranges, invalid colors,
 * out-of-bounds ranks/percentages, missing dataBar fields, unknown rule
 * fields, and guard-rail overruns with 400s — nothing unvalidated reaches
 * the engine.
 *
 * The XLSX application itself is covered by @genoffice/xlsx-gateway tests;
 * these tests pin the WIRE contract.
 */
import { describe, it, expect } from 'vitest'
import { routeOffice } from '../../src/api/office-routes.js'

/** Placeholder bytes: validation-only tests never reach the engine stage. */
const FILE_BYTES = Buffer.from('placeholder-xlsx-bytes').toString('base64')

async function save(plan: {
  edits?: unknown[]
  cfStates?: unknown[]
}): Promise<{ status: number; body: Record<string, unknown> }> {
  const res = await routeOffice({
    path: '/office/workbooks/save',
    method: 'POST',
    body: {
      fileName: 'validation.xlsx',
      fileBytes: FILE_BYTES,
      savePlan: { edits: [], ...plan },
    },
  })
  if (res === null) throw new Error('routeOffice returned null for the save route')
  return { status: res.status, body: res.body as Record<string, unknown> }
}

const AREA = { startRow: 1, endRow: 5, startColumn: 0, endColumn: 0 }

describe('workbooks/save cfStates validation', () => {
  it('accepts a canonical cellIs rule with resolved style', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'number',
                operator: 'greaterThan',
                value: 10,
                style: { bl: 1, cl: { rgb: '#9C0006' }, bg: { rgb: 'rgb(255,199,206)' } },
              },
            },
          ],
        },
      ],
    })
    // Validation passes; the engine stage fails on the placeholder bytes
    // with a MALFORMED (not validation) error — proving the CF state
    // itself cleared validation.
    expect(res.body.error).toBe('malformed')
  })

  it('accepts every saveable family: text, rank, average, expression, color scale, data bar, icon set', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'text',
                operator: 'containsText',
                value: 'urgent',
                style: {},
              },
            },
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'rank',
                value: 10,
                isPercent: true,
                isBottom: false,
                style: {},
              },
            },
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'average',
                operator: 'greaterThanOrEqual',
                style: {},
              },
            },
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'formula', value: '=A2>5', style: {} },
            },
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'duplicateValues', style: {} },
            },
            {
              ranges: [AREA],
              stopIfTrue: true,
              rule: {
                type: 'colorScale',
                config: [
                  { index: 0, color: '#F8696B', value: { type: 'min' } },
                  { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } },
                  { index: 2, color: '#63BE7B', value: { type: 'max' } },
                ],
              },
            },
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'dataBar',
                isShowValue: true,
                config: {
                  min: { type: 'min' },
                  max: { type: 'max' },
                  isGradient: true,
                  positiveColor: '#638EC6',
                  nativeColor: '#FF0000',
                },
              },
            },
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'iconSet',
                isShowValue: false,
                config: [
                  {
                    iconType: '3TrafficLights1',
                    iconId: '0',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'num', value: 67 },
                  },
                  {
                    iconType: '3TrafficLights1',
                    iconId: '1',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'num', value: 33 },
                  },
                  {
                    iconType: '3TrafficLights1',
                    iconId: '2',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'min' },
                  },
                ],
              },
            },
          ],
        },
      ],
    })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects malformed ranges (inverted area)', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [{ startRow: 9, endRow: 5, startColumn: 0, endColumn: 0 }],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'duplicateValues', style: {} },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
  })

  it('rejects rules without ranges or with a non-boolean stopIfTrue', async () => {
    const empty = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'uniqueValues', style: {} },
            },
          ],
        },
      ],
    })
    expect(empty.status).toBe(400)
    const badFlag = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: 'yes',
              rule: { type: 'highlightCell', subType: 'uniqueValues', style: {} },
            },
          ],
        },
      ],
    })
    expect(badFlag.status).toBe(400)
    expect(badFlag.body.error).toBe('validation')
  })

  it('rejects invalid colors', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'colorScale',
                config: [
                  { index: 0, color: 'not-a-color', value: { type: 'min' } },
                  { index: 1, color: '#63BE7B', value: { type: 'max' } },
                ],
              },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
  })

  it('rejects unsupported cell-value and text operators', async () => {
    const numberOperator = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'number',
                operator: 'containsText',
                value: 1,
                style: {},
              },
            },
          ],
        },
      ],
    })
    expect(numberOperator.status).toBe(400)
    expect(numberOperator.body.error).toBe('validation')
    const textOperator = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'text',
                operator: 'greaterThan',
                value: 'x',
                style: {},
              },
            },
          ],
        },
      ],
    })
    expect(textOperator.status).toBe(400)
  })

  it('rejects time-period (date-occurring) rules fail-closed', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'timePeriod',
                operator: 'yesterday',
                style: {},
              },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
    expect(String(res.body.message)).toContain('date-occurring')
  })

  it('rejects x14-only icon sets and mixed icon sets', async () => {
    const x14Only = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'iconSet',
                config: [
                  {
                    iconType: '3Triangles',
                    iconId: '0',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'num', value: 67 },
                  },
                  {
                    iconType: '3Triangles',
                    iconId: '1',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'num', value: 33 },
                  },
                  {
                    iconType: '3Triangles',
                    iconId: '2',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'min' },
                  },
                ],
              },
            },
          ],
        },
      ],
    })
    expect(x14Only.status).toBe(400)
    expect(x14Only.body.error).toBe('validation')
    const mixed = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'iconSet',
                config: [
                  {
                    iconType: '3Arrows',
                    iconId: '0',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'num', value: 67 },
                  },
                  {
                    iconType: '3Flags',
                    iconId: '1',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'num', value: 33 },
                  },
                  {
                    iconType: '3Arrows',
                    iconId: '2',
                    operator: 'greaterThanOrEqual',
                    value: { type: 'min' },
                  },
                ],
              },
            },
          ],
        },
      ],
    })
    expect(mixed.status).toBe(400)
    expect(String(mixed.body.message)).toContain('x14-only')
  })

  it('rejects out-of-bounds ranks, non-numeric values, and between pairs', async () => {
    const rank = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'rank', value: 5000, style: {} },
            },
          ],
        },
      ],
    })
    expect(rank.status).toBe(400)
    const value = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'number',
                operator: 'greaterThan',
                value: 'big',
                style: {},
              },
            },
          ],
        },
      ],
    })
    expect(value.status).toBe(400)
    const between = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'number',
                operator: 'between',
                value: [1],
                style: {},
              },
            },
          ],
        },
      ],
    })
    expect(between.status).toBe(400)
  })

  it('rejects data bars missing required fields', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: { type: 'dataBar', config: { min: { type: 'min' }, positiveColor: '#638EC6' } },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
  })

  it('rejects unknown fields at every level', async () => {
    const ruleLevel = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              cfId: 'rule-1',
              rule: { type: 'highlightCell', subType: 'uniqueValues', style: {} },
            },
          ],
        },
      ],
    })
    expect(ruleLevel.status).toBe(400)
    const familyLevel = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'uniqueValues',
                style: {},
                extLst: '<x14:id>1</x14:id>',
              },
            },
          ],
        },
      ],
    })
    expect(familyLevel.status).toBe(400)
    const stateLevel = await save({
      cfStates: [{ sheetName: 'Data', sheetId: 'sheet-1', rules: [] }],
    })
    expect(stateLevel.status).toBe(400)
  })

  it('rejects unknown rule families and sub-types', async () => {
    const family = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [{ ranges: [AREA], stopIfTrue: false, rule: { type: 'sparkline', style: {} } }],
        },
      ],
    })
    expect(family.status).toBe(400)
    const subType = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'mystery', style: {} },
            },
          ],
        },
      ],
    })
    expect(subType.status).toBe(400)
  })

  it('rejects color scales with too few or too many stops', async () => {
    const one = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'colorScale',
                config: [{ index: 0, color: '#FF0000', value: { type: 'min' } }],
              },
            },
          ],
        },
      ],
    })
    expect(one.status).toBe(400)
    const six: Array<Record<string, unknown>> = []
    for (let index = 0; index < 6; index += 1) {
      six.push({ index, color: '#FF0000', value: { type: 'num', value: index } })
    }
    const many = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [{ ranges: [AREA], stopIfTrue: false, rule: { type: 'colorScale', config: six } }],
        },
      ],
    })
    expect(many.status).toBe(400)
  })

  it('rejects unsupported threshold types', async () => {
    const res = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'dataBar',
                config: {
                  min: { type: 'autoMin' },
                  max: { type: 'max' },
                  positiveColor: '#638EC6',
                  nativeColor: '#FF0000',
                },
              },
            },
          ],
        },
      ],
    })
    expect(res.status).toBe(400)
    expect(res.body.error).toBe('validation')
  })

  it('rejects oversize cfStates payloads', async () => {
    const states = Array.from({ length: 1_001 }, (_, index) => ({
      sheetName: `S${index}`,
      rules: [],
    }))
    const res = await save({ cfStates: states })
    expect(res.status).toBe(400)
    expect(String(res.body.message)).toContain('exceeds')
    const manyRules = Array.from({ length: 501 }, () => ({
      ranges: [AREA],
      stopIfTrue: false,
      rule: { type: 'highlightCell', subType: 'uniqueValues', style: {} },
    }))
    const rulesRes = await save({ cfStates: [{ sheetName: 'Data', rules: manyRules }] })
    expect(rulesRes.status).toBe(400)
  })

  it('accepts an empty rules list (clearing a sheet conditional formatting)', async () => {
    const res = await save({ cfStates: [{ sheetName: 'Data', rules: [] }] })
    expect(res.body.error).toBe('malformed')
  })

  it('rejects styles with unknown fields or malformed decoration objects', async () => {
    const unknown = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'uniqueValues', style: { fontSize: 14 } },
            },
          ],
        },
      ],
    })
    expect(unknown.status).toBe(400)
    const deco = await save({
      cfStates: [
        {
          sheetName: 'Data',
          rules: [
            {
              ranges: [AREA],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'uniqueValues',
                style: { ul: { s: 1, val: 'double' } },
              },
            },
          ],
        },
      ],
    })
    expect(deco.status).toBe(400)
  })
})
