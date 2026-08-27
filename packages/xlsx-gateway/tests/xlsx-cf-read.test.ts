import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { parseConditionalFormatting, CfReadError } from '../src/gateway/xlsx-cf'
import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'

/// dxf styles used across the fixtures (styles.xml dxfs order).
const DXFS = [
  '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
  '<dxf><font><i/><strike/><u/><color rgb="FF006100"/></font></dxf>',
  '<dxf><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
  '<dxf><numFmt numFmtId="164" formatCode="0.000"/></dxf>',
]

const dxfAt = (dxfId: number): string | undefined => DXFS[dxfId]

const sheet = (inner: string): string =>
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${inner}<pageMargins left="0.7"/></worksheet>`

const parse = (xml: string) => parseConditionalFormatting(xml, dxfAt)

describe('parseConditionalFormatting — cell value operators', () => {
  it('parses a numeric cellIs rule with its dxf-resolved style', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="B2:B10"><cfRule type="cellIs" dxfId="0" priority="3" operator="equal"><formula>42</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules).toEqual([
      {
        ranges: [{ startRow: 1, endRow: 9, startColumn: 1, endColumn: 1 }],
        stopIfTrue: false,
        rule: {
          type: 'highlightCell',
          subType: 'number',
          operator: 'equal',
          value: 42,
          style: { bl: 1, cl: { rgb: '#9C0006' }, bg: { rgb: '#FFC7CE' } },
        },
      },
    ])
  })

  it('parses every unary numeric operator', () => {
    // With operand 10, lessThan / lessThanOrEqual paint Excel blanks
    // (0 < 10) while Univer's native conditions skip them — small ranges
    // swap to the equivalent formula rule (desktop parity). The other
    // operators keep the native number rule.
    const native: Array<[string, string]> = [
      ['greaterThan', '=A1>10'],
      ['greaterThanOrEqual', '=A1>=10'],
      ['notEqual', '=A1<>10'],
      ['equal', '=A1=10'],
    ]
    for (const [operator, formula] of native) {
      const rules = parse(
        sheet(
          `<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" priority="1" operator="${operator}"><formula>10</formula></cfRule></conditionalFormatting>`,
        ),
      )
      expect(rules[0]!.rule).toEqual({
        type: 'highlightCell',
        subType: 'number',
        operator,
        value: 10,
        style: {},
      })
      void formula
    }
    for (const operator of ['lessThan', 'lessThanOrEqual']) {
      const rules = parse(
        sheet(
          `<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" priority="1" operator="${operator}"><formula>10</formula></cfRule></conditionalFormatting>`,
        ),
      )
      expect(rules[0]!.rule).toMatchObject({
        type: 'highlightCell',
        subType: 'formula',
        value: operator === 'lessThan' ? '=A1<10' : '=A1<=10',
      })
    }
  })

  it('parses between and notBetween into two-value number rules', () => {
    for (const operator of ['between', 'notBetween']) {
      const rules = parse(
        sheet(
          `<conditionalFormatting sqref="C1:C4"><cfRule type="cellIs" priority="1" operator="${operator}"><formula>1</formula><formula>9</formula></cfRule></conditionalFormatting>`,
        ),
      )
      const rule = rules[0]!.rule as Record<string, unknown>
      expect(rule.subType).toBe('number')
      expect(rule.value).toEqual([1, 9])
    }
  })

  it('swaps blank-diverging operands to the equivalent formula rule (desktop parity)', () => {
    // lessThan 5: Excel paints blanks (0 < 5), Univer skips them → the
    // small range swaps to the formula form anchored at the range top-left.
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" priority="1" operator="lessThan"><formula>5</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'formula',
      value: '=A1<5',
      style: {},
    })
  })

  it('keeps the native number rule for blank-diverging operands on huge ranges', () => {
    // A1:A1048576 = 1,048,576 covered cells > the 20,000-cell formula
    // budget → the cheaper native condition survives (desktop parity).
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A1048576"><cfRule type="cellIs" priority="1" operator="lessThan"><formula>5</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'number',
      operator: 'lessThan',
      value: 5,
      style: {},
    })
  })

  it('parses quoted string equality into text rules (exact writer inverse)', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="2" priority="1" operator="equal"><formula>"done"</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'text',
      operator: 'equal',
      value: 'done',
      style: { bg: { rgb: '#FFEB9C' } },
    })
    const notEqual = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" priority="1" operator="notEqual"><formula>"a""b"</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(notEqual[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'text',
      operator: 'notEqual',
      value: 'a"b',
      style: {},
    })
  })

  it('parses reference/expression operands into anchored formula rules', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="B2:B9"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>$B$1</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'formula',
      value: '=B2>($B$1)',
      style: {},
    })
  })

  it('anchors relative formulas at the top-left-sorted first range', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="C3:C9 A1:A5"><cfRule type="cellIs" priority="1" operator="lessThanOrEqual"><formula>3</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.ranges).toEqual([
      { startRow: 2, endRow: 8, startColumn: 2, endColumn: 2 },
      { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
    ])
    expect((rules[0]!.rule as Record<string, unknown>).value).toBe('=A1<=3')
  })

  it('rejects unknown or missing cellIs operators', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="containsText"><formula>1</formula></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1"><formula>1</formula></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('rejects between rules with a missing second operand', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="between"><formula>1</formula></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })
})

describe('parseConditionalFormatting — text and presence families', () => {
  it('parses the four text operators with their text attribute', () => {
    for (const type of ['containsText', 'notContainsText', 'beginsWith', 'endsWith']) {
      const rules = parse(
        sheet(
          `<conditionalFormatting sqref="A1:A9"><cfRule type="${type}" dxfId="1" priority="1" text="urgent" operator="${type}"><formula>placeholder</formula></cfRule></conditionalFormatting>`,
        ),
      )
      expect(rules[0]!.rule).toEqual({
        type: 'highlightCell',
        subType: 'text',
        operator: type,
        value: 'urgent',
        style: { it: 1, st: { s: 1 }, ul: { s: 1 }, cl: { rgb: '#006100' } },
      })
    }
  })

  it('parses blanks, noBlanks, errors and noErrors', () => {
    const cases: Array<[string, string]> = [
      ['containsBlanks', 'containsBlanks'],
      ['notContainsBlanks', 'notContainsBlanks'],
      ['containsErrors', 'containsErrors'],
      ['notContainsErrors', 'notContainsErrors'],
    ]
    for (const [type, operator] of cases) {
      const rules = parse(
        sheet(
          `<conditionalFormatting sqref="D1:D9"><cfRule type="${type}" priority="1"><formula>unused</formula></cfRule></conditionalFormatting>`,
        ),
      )
      expect(rules[0]!.rule).toEqual({
        type: 'highlightCell',
        subType: 'text',
        operator,
        value: '',
        style: {},
      })
    }
  })
})

describe('parseConditionalFormatting — duplicate, rank and average families', () => {
  it('parses duplicateValues and uniqueValues', () => {
    const duplicates = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A9"><cfRule type="duplicateValues" dxfId="2" priority="1"/></conditionalFormatting>',
      ),
    )
    expect(duplicates[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'duplicateValues',
      style: { bg: { rgb: '#FFEB9C' } },
    })
    const unique = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A9"><cfRule type="uniqueValues" priority="2"/></conditionalFormatting>',
      ),
    )
    expect(unique[0]!.rule).toEqual({ type: 'highlightCell', subType: 'uniqueValues', style: {} })
  })

  it('parses top10 with rank, percent and bottom flags', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="B1:B20"><cfRule type="top10" dxfId="0" priority="1" rank="5" percent="1" bottom="1"/></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'rank',
      value: 5,
      isPercent: true,
      isBottom: true,
      style: { bl: 1, cl: { rgb: '#9C0006' }, bg: { rgb: '#FFC7CE' } },
    })
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="top10" priority="1"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('parses all four aboveAverage operator forms', () => {
    const cases: Array<[string, string]> = [
      ['', 'greaterThan'],
      [' equalAverage="1"', 'greaterThanOrEqual'],
      [' aboveAverage="0"', 'lessThan'],
      [' aboveAverage="0" equalAverage="1"', 'lessThanOrEqual'],
    ]
    for (const [extra, operator] of cases) {
      const rules = parse(
        sheet(
          `<conditionalFormatting sqref="C1:C9"><cfRule type="aboveAverage" dxfId="2" priority="1"${extra}/></conditionalFormatting>`,
        ),
      )
      expect(rules[0]!.rule).toEqual({
        type: 'highlightCell',
        subType: 'average',
        operator,
        style: { bg: { rgb: '#FFEB9C' } },
      })
    }
  })

  it('parses expression rules with the = prefix added', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="E1:E9"><cfRule type="expression" dxfId="1" priority="1"><formula>AND(A1&gt;1,B1&lt;2)</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'highlightCell',
      subType: 'formula',
      value: '=AND(A1>1,B1<2)',
      style: { it: 1, st: { s: 1 }, ul: { s: 1 }, cl: { rgb: '#006100' } },
    })
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })
})

describe('parseConditionalFormatting — visual families', () => {
  it('parses a three-color scale with mixed threshold types', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="F1:F9"><cfRule type="colorScale" priority="1"><colorScale>' +
          '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
          '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>' +
          '</colorScale></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'colorScale',
      config: [
        { index: 0, color: '#F8696B', value: { type: 'min' } },
        { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } },
        { index: 2, color: '#63BE7B', value: { type: 'max' } },
      ],
    })
  })

  it('parses a two-color scale with num and formula thresholds', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="F1:F9"><cfRule type="colorScale" priority="1"><colorScale>' +
          '<cfvo type="num" val="0"/><cfvo type="formula">$G$1</cfvo>' +
          '<color rgb="FF0000FF"/><color rgb="FF00FF00"/>' +
          '</colorScale></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'colorScale',
      config: [
        { index: 0, color: '#0000FF', value: { type: 'num', value: 0 } },
        { index: 1, color: '#00FF00', value: { type: 'formula', value: '=$G$1' } },
      ],
    })
  })

  it('rejects color scales with mismatched stops or unrepresentable colors', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color theme="1"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('parses a base data bar with showValue and threshold configs', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="G1:G9"><cfRule type="dataBar" priority="1"><dataBar showValue="0">' +
          '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>' +
          '</dataBar></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'dataBar',
      isShowValue: false,
      config: {
        min: { type: 'min' },
        max: { type: 'max' },
        isGradient: true,
        positiveColor: '#638EC6',
        nativeColor: '#FF0000',
      },
    })
  })

  it('rejects data bars without exactly two thresholds or a color', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/></dataBar></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('parses a natural 3TrafficLights1 icon set (best-first conversion)', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="H1:H9"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1">' +
          '<cfvo type="percent" val="0"/><cfvo type="num" val="33"/><cfvo type="num" val="67"/>' +
          '</iconSet></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'iconSet',
      isShowValue: true,
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
    })
  })

  it('parses a reversed 3Arrows icon set and strict thresholds', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="H1:H9"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows" reverse="1">' +
          '<cfvo type="percent" val="0"/><cfvo type="num" val="33" gte="0"/><cfvo type="num" val="67"/>' +
          '</iconSet></cfRule></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.rule).toEqual({
      type: 'iconSet',
      isShowValue: true,
      config: [
        {
          iconType: '3Arrows',
          iconId: '2',
          operator: 'greaterThanOrEqual',
          value: { type: 'num', value: 67 },
        },
        {
          iconType: '3Arrows',
          iconId: '1',
          operator: 'greaterThan',
          value: { type: 'num', value: 33 },
        },
        {
          iconType: '3Arrows',
          iconId: '0',
          operator: 'greaterThanOrEqual',
          value: { type: 'min' },
        },
      ],
    })
  })

  it('parses a worst-first 5Rating icon set', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="H1:H9"><cfRule type="iconSet" priority="1"><iconSet iconSet="5Rating">' +
          '<cfvo type="percent" val="0"/><cfvo type="percent" val="20"/><cfvo type="percent" val="40"/><cfvo type="percent" val="60"/><cfvo type="percent" val="80"/>' +
          '</iconSet></cfRule></conditionalFormatting>',
      ),
    )
    const config = (rules[0]!.rule as Record<string, unknown>).config as Array<
      Record<string, unknown>
    >
    expect(config.map((entry) => entry.iconId)).toEqual(['4', '3', '2', '1', '0'])
  })

  it('rejects icon sets outside the OOXML whitelist or with wrong counts', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Triangles"><cfvo type="percent" val="0"/><cfvo type="num" val="33"/><cfvo type="num" val="67"/></iconSet></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="num" val="33"/></iconSet></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })
})

describe('parseConditionalFormatting — structure, priority and styles', () => {
  it('orders rules by ascending priority across blocks and within blocks', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A9"><cfRule type="uniqueValues" priority="7"/><cfRule type="duplicateValues" priority="2" stopIfTrue="1"/></conditionalFormatting>' +
          '<conditionalFormatting sqref="B1:B9"><cfRule type="containsBlanks" priority="4"/></conditionalFormatting>',
      ),
    )
    expect(rules.map((rule) => (rule.rule as Record<string, unknown>).subType)).toEqual([
      'duplicateValues',
      'text',
      'uniqueValues',
    ])
    expect(rules[0]!.stopIfTrue).toBe(true)
    expect(rules[1]!.ranges).toEqual([{ startRow: 0, endRow: 8, startColumn: 1, endColumn: 1 }])
  })

  it('parses multi-range sqref into the ranges array (file order)', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A5 C3:D4"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      ),
    )
    expect(rules[0]!.ranges).toEqual([
      { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
      { startRow: 2, endRow: 3, startColumn: 2, endColumn: 3 },
    ])
  })

  it('resolves numFmt-only dxfs into the style pattern', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="3" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
      ),
    )
    expect((rules[0]!.rule as Record<string, unknown>).style).toEqual({ n: { pattern: '0.000' } })
  })

  it('treats rules without dxfId as unstyled', () => {
    const rules = parse(
      sheet(
        '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      ),
    )
    expect((rules[0]!.rule as Record<string, unknown>).style).toEqual({})
  })

  it('returns an empty list when the sheet has no conditional formatting', () => {
    expect(parse(sheet('<pageMargins left="0.7"/>'))).toEqual([])
  })
})

describe('parseConditionalFormatting — fail-closed constructs', () => {
  it('rejects x14 extension rules in the worksheet extLst', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1:A5"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>' +
            '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" x14:id="x"></x14:cfRule></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('rejects x14-linked base halves (cfRule extLst)', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1:A5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-BA67-4E7D1D9ED3A3}"><x14:id>1</x14:id></ext></extLst></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('rejects timePeriod rules', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1:A5"><cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="yesterday"><formula>FLOOR(A1,1)=TODAY()-1</formula></cfRule></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('rejects unknown rule types', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="mysteryRule" priority="1"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('rejects malformed sqref and unreadable priorities', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="not-a-range"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
  })

  it('rejects unresolvable or unrepresentable dxf styling', () => {
    expect(() =>
      parse(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" dxfId="99" priority="1"/></conditionalFormatting>',
        ),
      ),
    ).toThrow(CfReadError)
    const invalidDxf = ['not-a-dxf-entry']
    expect(() =>
      parseConditionalFormatting(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" dxfId="0" priority="1"/></conditionalFormatting>',
        ),
        (dxfId) => invalidDxf[dxfId],
      ),
    ).toThrow(CfReadError)
  })

  it('rejects dxfs carrying constructs the writer would drop', () => {
    const borderDxf = ['<dxf><border><left style="thin"/></border></dxf>']
    expect(() =>
      parseConditionalFormatting(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" dxfId="0" priority="1"/></conditionalFormatting>',
        ),
        (dxfId) => borderDxf[dxfId],
      ),
    ).toThrow(CfReadError)
    const themeDxf = ['<dxf><font><color theme="4"/></font></dxf>']
    expect(() =>
      parseConditionalFormatting(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" dxfId="0" priority="1"/></conditionalFormatting>',
        ),
        (dxfId) => themeDxf[dxfId],
      ),
    ).toThrow(CfReadError)
    const sizeDxf = ['<dxf><font><sz val="14"/></font></dxf>']
    expect(() =>
      parseConditionalFormatting(
        sheet(
          '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" dxfId="0" priority="1"/></conditionalFormatting>',
        ),
        (dxfId) => sizeDxf[dxfId],
      ),
    ).toThrow(CfReadError)
  })
})

/// Two-sheet fixture with configurable CF XML per sheet + the DXFS list in
/// styles.xml — the readBasicWorkbook / applyCellEditsToXlsx integration
/// surface for EXCEL-024.
async function buildCfFixture(sheet1Cf: string, sheet2Cf: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Numbers" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>' +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  const worksheet = (cf: string) =>
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>7</v></c></row></sheetData>${cf}` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
  zip.file('xl/worksheets/sheet1.xml', worksheet(sheet1Cf))
  zip.file('xl/worksheets/sheet2.xml', worksheet(sheet2Cf))
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>' +
      `<dxfs count="${DXFS.length}">${DXFS.join('')}</dxfs>` +
      '</styleSheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const readEntry = async (buffer: Buffer, entry: string): Promise<string> => {
  const zip = new JSZip()
  const loaded = await zip.loadAsync(buffer)
  return loaded.file(entry)!.async('string')
}

describe('EXCEL-024 readBasicWorkbook integration', () => {
  it('exposes cfRules on the sheet snapshot (priority-ascending, styles resolved)', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A1:B4"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>3</formula></cfRule><cfRule type="duplicateValues" priority="5" stopIfTrue="1"/></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1:B9"><cfRule type="expression" dxfId="1" priority="1"><formula>B1&gt;A1</formula></cfRule></conditionalFormatting>',
      '',
    )
    const imported = await readBasicWorkbook(buffer)
    const first = imported.snapshot.sheets[0]!
    expect(first.cfRules).toBeDefined()
    expect(first.cfLocked).toBeUndefined()
    const rules = first.cfRules!
    expect(rules).toHaveLength(3)
    expect((rules[0]!.rule as Record<string, unknown>).subType).toBe('formula')
    expect(rules[0]!.ranges).toEqual([{ startRow: 0, endRow: 8, startColumn: 1, endColumn: 1 }])
    expect((rules[1]!.rule as Record<string, unknown>).subType).toBe('number')
    expect(rules[1]!.ranges).toEqual([{ startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }])
    expect(rules[2]!.stopIfTrue).toBe(true)
    expect(imported.snapshot.sheets[1]!.cfRules).toBeUndefined()
  })

  it('fails closed per sheet: x14 locks sheet 1 while sheet 2 still parses', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1"><extLst><ext uri="{B025F937-C7B1-47D3-BA67-4E7D1D9ED3A3}"><x14:id>1</x14:id></ext></extLst></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="A1:A3"><cfRule type="duplicateValues" dxfId="2" priority="1"/></conditionalFormatting>',
    )
    const imported = await readBasicWorkbook(buffer)
    expect(imported.snapshot.sheets[0]!.cfRules).toBeUndefined()
    expect(imported.snapshot.sheets[0]!.cfLocked).toBe(true)
    expect(imported.snapshot.sheets[1]!.cfRules).toBeDefined()
    expect(imported.snapshot.sheets[1]!.cfLocked).toBeUndefined()
  })
})

describe('EXCEL-024 save integration (applyCellEditsToXlsx cfStates)', () => {
  it('round-trips rules through the writer and back (write → reopen)', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A1:A4"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="A1:A3"><cfRule type="duplicateValues" dxfId="2" priority="1"/></conditionalFormatting>',
    )
    const imported = await readBasicWorkbook(buffer)
    const rules = imported.snapshot.sheets[0]!.cfRules!
    // Edit the rule (change the operand) — the browser would send the full
    // live rule set of the now-CF-dirty sheet.
    const edited = rules.map((rule) =>
      rule === rules[0]
        ? { ...rule, rule: { ...(rule.rule as Record<string, unknown>), value: 99 } }
        : rule,
    )
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [{ sheetName: 'Numbers', rules: edited }],
      [],
      [],
      null,
      [],
      [],
      [],
      null,
      [],
      [],
      [],
    )
    const savedXml = await readEntry(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(savedXml).toContain('operator="greaterThan"')
    expect(savedXml).toContain('<formula>99</formula>')
    // Reopen: the saved bytes parse back into the same wire shape.
    const reopened = await readBasicWorkbook(mutation.buffer)
    const reopenedRules = reopened.snapshot.sheets[0]!.cfRules!
    expect(reopenedRules).toEqual(edited)
    // Sheet 2 was never CF-dirty — its section is untouched.
    const sheet2Xml = await readEntry(mutation.buffer, 'xl/worksheets/sheet2.xml')
    expect(sheet2Xml).toContain('type="duplicateValues"')
  })

  it('clears all rules of a CF-dirty sheet with an empty snapshot', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A1:A4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>',
      '',
    )
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [{ sheetName: 'Numbers', rules: [] }],
      [],
      [],
      null,
      [],
      [],
      [],
      null,
      [],
      [],
      [],
    )
    const savedXml = await readEntry(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(savedXml).not.toContain('conditionalFormatting')
  })

  it('preserves CF bytes untouched on a no-CF-dirty save', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A1:A4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1:B4"><cfRule type="timePeriod" dxfId="1" priority="2" timePeriod="today"><formula>placeholder</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="A1:A3"><cfRule type="duplicateValues" dxfId="2" priority="1"/></conditionalFormatting>',
    )
    const originalSheet1 = await readEntry(buffer, 'xl/worksheets/sheet1.xml')
    const originalSheet2 = await readEntry(buffer, 'xl/worksheets/sheet2.xml')
    const originalStyles = await readEntry(buffer, 'xl/styles.xml')
    // A cell edit only — no sheet is CF-dirty, so no cfStates are sent and
    // applyCfRules never runs (the locked sheet's timePeriod rule and the
    // unlocked sheets' rules all stay byte-identical).
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [{ sheetName: 'Numbers', row: 0, column: 0, writeValue: true, cell: { value: 2 } }],
      [],
      [],
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
      [],
      [],
    )
    // The cell edit rewrites the value and injects a dimension element —
    // everything from the CF sections onward stays byte-identical.
    const fromConditional = (xml: string): string =>
      xml.slice(xml.indexOf('<conditionalFormatting'))
    expect(fromConditional(await readEntry(mutation.buffer, 'xl/worksheets/sheet1.xml'))).toBe(
      fromConditional(originalSheet1),
    )
    expect(await readEntry(mutation.buffer, 'xl/worksheets/sheet2.xml')).toBe(originalSheet2)
    expect(await readEntry(mutation.buffer, 'xl/styles.xml')).toBe(originalStyles)
  })

  it('shifts CF sqref and formulas through structural ops and reopens shifted', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A3:A9"><cfRule type="expression" dxfId="1" priority="1"><formula>A3&gt;5</formula></cfRule></conditionalFormatting>',
      '',
    )
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [{ sheetName: 'Numbers', ops: [{ kind: 'insert-rows', index: 1, count: 2 }] }],
      [],
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
      [],
      [],
    )
    const savedXml = await readEntry(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(savedXml).toContain('sqref="A5:A11"')
    expect(savedXml).toContain('<formula>A5&gt;5</formula>')
    const reopened = await readBasicWorkbook(mutation.buffer)
    const rules = reopened.snapshot.sheets[0]!.cfRules!
    expect(rules[0]!.ranges).toEqual([{ startRow: 4, endRow: 10, startColumn: 0, endColumn: 0 }])
    expect((rules[0]!.rule as Record<string, unknown>).value).toBe('=A5>5')
  })

  it('edits siblings while preserving unrelated-package parts byte-identically', async () => {
    const buffer = await buildCfFixture(
      '<conditionalFormatting sqref="A1:A4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1:B4"><cfRule type="containsBlanks" dxfId="2" priority="2"/></conditionalFormatting>',
      '',
    )
    const originalWorkbook = await readEntry(buffer, 'xl/workbook.xml')
    const originalRels = await readEntry(buffer, 'xl/_rels/workbook.xml.rels')
    const originalSheet2 = await readEntry(buffer, 'xl/worksheets/sheet2.xml')
    const originalContentTypes = await readEntry(buffer, '[Content_Types].xml')
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [
        {
          sheetName: 'Numbers',
          rules: [
            {
              ranges: [{ startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 }],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'number',
                operator: 'lessThan',
                value: 8,
                style: { bg: { rgb: '#FFC7CE' } },
              },
            },
            {
              ranges: [{ startRow: 0, endRow: 3, startColumn: 1, endColumn: 1 }],
              stopIfTrue: false,
              rule: {
                type: 'highlightCell',
                subType: 'text',
                operator: 'containsBlanks',
                value: '',
                style: {},
              },
            },
          ],
        },
      ],
      [],
      [],
      null,
      [],
      [],
      [],
      null,
      [],
      [],
      [],
    )
    const savedSheet1 = await readEntry(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(savedSheet1).toContain('operator="lessThan"')
    expect(savedSheet1).toContain('type="containsBlanks"')
    expect(savedSheet1).not.toContain('type="cellIs" dxfId="0" priority="1" operator="greaterThan"')
    // workbook.xml only gains the repository's documented global
    // fullCalcOnLoad marker; everything else stays byte-identical.
    expect(await readEntry(mutation.buffer, 'xl/workbook.xml')).toBe(
      originalWorkbook.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>'),
    )
    expect(await readEntry(mutation.buffer, 'xl/_rels/workbook.xml.rels')).toBe(originalRels)
    expect(await readEntry(mutation.buffer, 'xl/worksheets/sheet2.xml')).toBe(originalSheet2)
    expect(await readEntry(mutation.buffer, '[Content_Types].xml')).toBe(originalContentTypes)
  })
})
