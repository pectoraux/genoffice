import { describe, expect, it } from 'vitest'
import { buildRowWindow, rowWindowIsEmpty } from '../../src/index.js'

describe('PROJECT-022 virtualization — the shared row window', () => {
  it('builds the clamped inclusive window for plain inputs', () => {
    expect(buildRowWindow(100, { firstRow: 10, visibleRows: 20 })).toEqual({
      firstIndex: 10,
      lastIndex: 29,
    })
  })

  it('clamps firstRow into the row range and visibleRows into [0, rowCount]', () => {
    expect(buildRowWindow(10, { firstRow: -5, visibleRows: 4 })).toEqual({
      firstIndex: 0,
      lastIndex: 3,
    })
    expect(buildRowWindow(10, { firstRow: 99, visibleRows: 4 })).toEqual({
      firstIndex: 9,
      lastIndex: 9,
    })
    expect(buildRowWindow(10, { firstRow: 2, visibleRows: 999 })).toEqual({
      firstIndex: 2,
      lastIndex: 9,
    })
  })

  it('applies overscan symmetrically, clamped to the row bounds', () => {
    expect(buildRowWindow(100, { firstRow: 10, visibleRows: 5, overscan: 3 })).toEqual({
      firstIndex: 7,
      lastIndex: 17,
    })
    expect(buildRowWindow(100, { firstRow: 0, visibleRows: 5, overscan: 3 })).toEqual({
      firstIndex: 0,
      lastIndex: 7,
    })
    expect(buildRowWindow(4, { firstRow: 3, visibleRows: 1, overscan: 2 })).toEqual({
      firstIndex: 1,
      lastIndex: 3,
    })
  })

  it('degenerates to the canonical empty window (lastIndex < firstIndex)', () => {
    expect(buildRowWindow(0, { firstRow: 0, visibleRows: 10 })).toEqual({
      firstIndex: 0,
      lastIndex: -1,
    })
    expect(buildRowWindow(10, { firstRow: 0, visibleRows: 0 })).toEqual({
      firstIndex: 0,
      lastIndex: -1,
    })
    expect(buildRowWindow(-1, { firstRow: 0, visibleRows: 3 })).toEqual({
      firstIndex: 0,
      lastIndex: -1,
    })
    expect(buildRowWindow(Number.NaN, { firstRow: 0, visibleRows: 3 })).toEqual({
      firstIndex: 0,
      lastIndex: -1,
    })
    expect(rowWindowIsEmpty({ firstIndex: 0, lastIndex: -1 })).toBe(true)
    expect(rowWindowIsEmpty({ firstIndex: 4, lastIndex: 4 })).toBe(false)
  })

  it('is deterministic (3× identical)', () => {
    const run = () => buildRowWindow(1000, { firstRow: 401, visibleRows: 33, overscan: 7 })
    expect(run()).toEqual(run())
    expect(run()).toEqual(run())
  })
})
