/**
 * PROJECT-027 correction — the ONE canonical bounded native-read helper.
 *
 * Every native file read in the host crosses this function: the open-picker
 * path AND the argv / second-instance path share ONE transport policy (the
 * review finding was exactly that the latter had escaped the cap). Stat
 * first — an oversized file is rejected without reading a byte — then a
 * single bounded read window of stat-size + 1: a file that GREW after the
 * stat can never smuggle uncapped bytes across (it either still fits the
 * cap or trips the size rejection on the final check).
 *
 * Pure Node (no Electron import) so the transport cap is unit-testable
 * directly. Errors are VALUES, never throws: the result's error variant
 * carries no bytes, which is what makes "the renderer never receives
 * uncapped file contents" structural rather than conventional.
 */
import { open, stat } from 'node:fs/promises'
import type { FileHandle } from 'node:fs/promises'
import type { NativeReadResult } from '../shared/ipc.js'

/** The maximum size the host will read or write (defense in depth; the
 * canonical adapters carry their own caps — this is transport hygiene).
 * Defined exactly once, here: every read surface and the write cap consume
 * this constant, so no handler carries its own cap logic. */
export const MAX_FILE_BYTES = 100 * 1024 * 1024

/**
 * Reads at most `maxBytes` from `path` (default: the transport cap).
 * Never throws — a missing, unreadable, oversized, or mid-read-failing
 * path is an `{ ok: false, error }` value.
 */
export async function boundedReadFile(
  path: string,
  maxBytes: number = MAX_FILE_BYTES,
): Promise<NativeReadResult> {
  let size: number
  try {
    size = (await stat(path)).size
  } catch (error) {
    return { ok: false, error: describeError(error) }
  }
  // Reject oversized files on the STAT, before a single byte is read.
  if (size > maxBytes) {
    return { ok: false, error: `File exceeds the ${maxBytes} byte limit` }
  }
  let handle: FileHandle
  try {
    handle = await open(path, 'r')
  } catch (error) {
    return { ok: false, error: describeError(error) }
  }
  try {
    // One bounded window: stat size + 1 byte. The +1 detects growth: a file
    // that grew past the cap after the stat reads at most size + 1 bytes
    // here, and the final check rejects it — uncapped content can never
    // cross the boundary even under a stat/read race.
    const window = Buffer.alloc(size + 1)
    let total = 0
    while (total < window.length) {
      const { bytesRead } = await handle.read(window, total, window.length - total, total)
      if (bytesRead === 0) break
      total += bytesRead
    }
    if (total > maxBytes) {
      return { ok: false, error: `File exceeds the ${maxBytes} byte limit` }
    }
    return { ok: true, bytes: window.subarray(0, total) }
  } catch (error) {
    return { ok: false, error: describeError(error) }
  } finally {
    await handle.close().catch(() => undefined)
  }
}

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
