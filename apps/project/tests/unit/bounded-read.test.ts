// @vitest-environment node
/**
 * PROJECT-027 correction — the canonical bounded native-read helper battery.
 *
 * Direct unit tests of the ONE transport read (`src/main/bounded-read.ts`):
 * the cap is enforced on the STAT (an oversized file is rejected without a
 * byte being read), the boundary size reads exactly, missing/unreadable/
 * directory paths are error VALUES (never throws), and the error variant
 * carries no bytes — the structural "the renderer never receives uncapped
 * file contents" proof. The default cap (MAX_FILE_BYTES) is pinned; the
 * boundary algebra is exercised through the injectable `maxBytes` seam with
 * small caps, and the real-cap boundary + both routed IPC surfaces are
 * proven end-to-end by the desktop E2E (E13).
 */
import { chmod, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { MAX_FILE_BYTES, boundedReadFile } from '../../src/main/bounded-read.js'

async function scratchFile(content: string): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-bounded-read-'))
  const path = join(dir, 'fixture.gproj')
  await writeFile(path, content)
  return path
}

describe('the canonical bounded native read', () => {
  it('reads a small file exactly (default cap, full bytes, ok value)', async () => {
    const path = await scratchFile('{"format":"gproj"}')
    const result = await boundedReadFile(path)
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.bytes.byteLength).toBe(18)
      expect(new TextDecoder().decode(result.bytes)).toBe('{"format":"gproj"}')
    }
  })

  it('reads an EMPTY file as ok with zero bytes', async () => {
    const path = await scratchFile('')
    const result = await boundedReadFile(path)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.bytes.byteLength).toBe(0)
  })

  it('reads a file at EXACTLY the boundary size (size === cap → ok)', async () => {
    const path = await scratchFile('0123456789')
    const result = await boundedReadFile(path, 10)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.bytes.byteLength).toBe(10)
  })

  it('rejects size cap + 1 on the STAT — no bytes in the result', async () => {
    const path = await scratchFile('0123456789A')
    const result = await boundedReadFile(path, 10)
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.error).toBe('File exceeds the 10 byte limit')
      // The structural proof: the error variant has NO bytes field at all —
      // uncapped content cannot cross, by shape.
      expect('bytes' in result).toBe(false)
    }
  })

  it('the oversized rejection is decided BEFORE any read (the stat branch)', async () => {
    // A directory's stat size (4096) exceeds the small cap, so the STAT
    // branch must fire: if the helper had opened/read first, a directory
    // read would fail with EISDIR instead — a different error than the
    // size rejection asserted here.
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-bounded-read-'))
    const result = await boundedReadFile(dir, 16)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toBe('File exceeds the 16 byte limit')
  })

  it('a missing path is an ENOENT error value (never a throw)', async () => {
    const result = await boundedReadFile(join(tmpdir(), 'genoffice-bounded-read-missing.gproj'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('ENOENT')
  })

  it('a directory path under the cap fails on the read (EISDIR)', async () => {
    // The directory's stat size (4096) is under the default cap, so the
    // stat passes, the open succeeds, and the READ on a directory fd fails
    // with EISDIR — an error value, never a throw.
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-bounded-read-'))
    const result = await boundedReadFile(dir)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('EISDIR')
  })

  it('an unreadable (mode 000) path is an EACCES error value', async () => {
    const path = await scratchFile('secret')
    await chmod(path, 0o000)
    const result = await boundedReadFile(path)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('EACCES')
  })

  it('the default transport cap is MAX_FILE_BYTES (100 MiB), defined once', () => {
    expect(MAX_FILE_BYTES).toBe(100 * 1024 * 1024)
  })
})
