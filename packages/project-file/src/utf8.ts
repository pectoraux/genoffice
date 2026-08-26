/**
 * PROJECT-014 — Pure-TypeScript UTF-8 codec.
 *
 * The `.gproj` adapter is a foundation package and MUST remain host-neutral
 * (`spec/project/architecture-lock.md` §13 forbids Node `fs`/`process`,
 * browser globals, and host APIs). To avoid relying on `Buffer` (Node) or
 * `TextEncoder` (a runtime global that, while standard, is a host-provided
 * API), the adapter ships a tiny deterministic UTF-8 encoder/decoder written in
 * pure TypeScript. It is sufficient for JSON-text round-tripping (the BMP is
 * handled directly; supplementary-plane code points are handled via surrogate
 * pairs, which `JSON.stringify` already emits).
 */

/** Encode a JS string as UTF-8 bytes. */
export function encodeUtf8(input: string): Uint8Array {
  // Fast path: ASCII-only strings are extremely common for .gproj fixtures
  // (ISO-8601 timestamps, ASCII identifiers, ASCII enum literals). For those,
  // a direct charCode map avoids per-code-point branching.
  let asciiOnly = true
  for (let i = 0; i < input.length; i++) {
    if (input.charCodeAt(i) > 0x7f) {
      asciiOnly = false
      break
    }
  }
  if (asciiOnly) {
    const out = new Uint8Array(input.length)
    for (let i = 0; i < input.length; i++) out[i] = input.charCodeAt(i)
    return out
  }
  const bytes: number[] = []
  for (let i = 0; i < input.length; i++) {
    let code = input.charCodeAt(i)
    // Collapse surrogate pairs into a single code point so supplementary-plane
    // characters encode as a 4-byte UTF-8 sequence (matches JSON.stringify /
    // TextEncoder byte output).
    if (code >= 0xd800 && code <= 0xdbff && i + 1 < input.length) {
      const low = input.charCodeAt(i + 1)
      if (low >= 0xdc00 && low <= 0xdfff) {
        code = 0x10000 + ((code - 0xd800) << 10) + (low - 0xdc00)
        i += 1
      }
    }
    if (code < 0x80) {
      bytes.push(code)
    } else if (code < 0x800) {
      bytes.push(0xc0 | (code >> 6), 0x80 | (code & 0x3f))
    } else if (code < 0x10000) {
      bytes.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 0x3f), 0x80 | (code & 0x3f))
    } else {
      bytes.push(
        0xf0 | (code >> 18),
        0x80 | ((code >> 12) & 0x3f),
        0x80 | ((code >> 6) & 0x3f),
        0x80 | (code & 0x3f),
      )
    }
  }
  return Uint8Array.from(bytes)
}

/** Decode UTF-8 bytes into a JS string. Throws on invalid sequences. */
export function decodeUtf8(input: Uint8Array): string {
  let result = ''
  let i = 0
  while (i < input.length) {
    const b0 = input[i]
    if (b0 < 0x80) {
      result += String.fromCharCode(b0)
      i += 1
    } else if (b0 < 0xc0) {
      throw new Error('Invalid UTF-8: unexpected continuation byte at index ' + i)
    } else if (b0 < 0xe0) {
      const b1 = input[i + 1]
      if (b1 === undefined || (b1 & 0xc0) !== 0x80)
        throw new Error('Invalid UTF-8: truncated 2-byte sequence at index ' + i)
      result += String.fromCharCode(((b0 & 0x1f) << 6) | (b1 & 0x3f))
      i += 2
    } else if (b0 < 0xf0) {
      const b1 = input[i + 1]
      const b2 = input[i + 2]
      if (b1 === undefined || b2 === undefined || (b1 & 0xc0) !== 0x80 || (b2 & 0xc0) !== 0x80)
        throw new Error('Invalid UTF-8: truncated 3-byte sequence at index ' + i)
      const code = ((b0 & 0x0f) << 12) | ((b1 & 0x3f) << 6) | (b2 & 0x3f)
      result += String.fromCharCode(code)
      i += 3
    } else {
      const b1 = input[i + 1]
      const b2 = input[i + 2]
      const b3 = input[i + 3]
      if (
        b1 === undefined ||
        b2 === undefined ||
        b3 === undefined ||
        (b1 & 0xc0) !== 0x80 ||
        (b2 & 0xc0) !== 0x80 ||
        (b3 & 0xc0) !== 0x80
      )
        throw new Error('Invalid UTF-8: truncated 4-byte sequence at index ' + i)
      const code = ((b0 & 0x07) << 18) | ((b1 & 0x3f) << 12) | ((b2 & 0x3f) << 6) | (b3 & 0x3f)
      const adjusted = code - 0x10000
      result += String.fromCharCode(0xd800 | (adjusted >> 10), 0xdc00 | (adjusted & 0x3ff))
      i += 4
    }
  }
  return result
}
