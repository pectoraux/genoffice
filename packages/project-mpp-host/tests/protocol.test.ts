/**
 * PROJECT-018 — sidecar protocol frame validation tests (pure).
 */
import { describe, expect, it } from 'vitest'
import { parseSidecarFrame } from '../src/protocol.js'

const OK_FRAME =
  '{"version":1,"requestId":"r1","ok":true,"counts":{"tasks":8,"resources":0,"calendars":1,"predecessorLinks":0,"assignments":0}}'

describe('parseSidecarFrame', () => {
  it('parses a valid success frame with counts', () => {
    const frame = parseSidecarFrame(`${OK_FRAME}\n`)
    expect(frame).not.toBeNull()
    expect(frame?.version).toBe(1)
    expect(frame?.requestId).toBe('r1')
    expect(frame?.ok).toBe(true)
    expect(frame?.counts).toEqual({
      tasks: 8,
      resources: 0,
      calendars: 1,
      predecessorLinks: 0,
      assignments: 0,
    })
  })

  it('parses a failure frame with a protocol error code', () => {
    const frame = parseSidecarFrame(
      '{"version":1,"requestId":"r2","ok":false,"error":{"code":"UNSUPPORTED_MPP_FORMAT","message":"nope"}}',
    )
    expect(frame?.ok).toBe(false)
    expect(frame?.error?.code).toBe('UNSUPPORTED_MPP_FORMAT')
    expect(frame?.error?.message).toBe('nope')
  })

  it('rejects a wrong protocol version', () => {
    expect(parseSidecarFrame('{"version":2,"requestId":"r","ok":true}')).toBeNull()
  })

  it('rejects non-JSON, empty output, and non-object JSON', () => {
    expect(parseSidecarFrame('THIS IS NOT JSON')).toBeNull()
    expect(parseSidecarFrame('')).toBeNull()
    expect(parseSidecarFrame('   \n  \n')).toBeNull()
    expect(parseSidecarFrame('[1,2,3]')).toBeNull()
    expect(parseSidecarFrame('"a string"')).toBeNull()
  })

  it('rejects structurally malformed envelopes (missing fields, wrong types)', () => {
    expect(parseSidecarFrame('{"version":1,"ok":true}')).toBeNull() // no requestId
    expect(parseSidecarFrame('{"version":1,"requestId":"r"}')).toBeNull() // no ok
    expect(parseSidecarFrame('{"version":1,"requestId":42,"ok":true}')).toBeNull()
    expect(
      parseSidecarFrame('{"version":1,"requestId":"r","ok":true,"counts":{"tasks":"many"}}'),
    ).toBeNull()
    expect(
      parseSidecarFrame('{"version":1,"requestId":"r","ok":true,"error":"notAnObject"}'),
    ).toBeNull()
  })

  it('treats the LAST non-empty line as authoritative (JVM banner tolerance)', () => {
    const frame = parseSidecarFrame(`WARNING: banner noise\nLog4j notice\n${OK_FRAME}`)
    expect(frame?.ok).toBe(true)
  })

  it('rejects a frame where the last line is noise after a valid frame', () => {
    // Protocol violation: trailing output after the frame is NOT tolerated —
    // the last line must be the frame.
    expect(parseSidecarFrame(`${OK_FRAME}\ntrailing junk`)).toBeNull()
  })
})
