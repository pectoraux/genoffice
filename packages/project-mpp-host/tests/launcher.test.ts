/**
 * PROJECT-018 — launcher failure-injection suite.
 *
 * Every failure class required by the brief is exercised against the REAL
 * process-management logic using a deterministic stand-in executable
 * (node + the fake-sidecar fixtures): unavailable, timeout, nonzero exit,
 * invalid frame, request-id mismatch, unsupported-format frame, missing
 * output, oversized output, banner-noise tolerance, and the success path.
 * The real java/MPXJ path is covered end-to-end by e2e-real-corpus.test.ts.
 *
 * Also covers the pipeline-level failure classes: input-size limit,
 * malformed sidecar MSPDI (mspdi stage), and canonical rejection
 * (atomicity), plus the shell-injection counter-proof.
 */
import {
  chmodSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  truncateSync,
  writeFileSync,
  existsSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  MppSidecarLauncher,
  buildSidecarCommand,
  importMppFromBytes,
  importMppFromFile,
  probeNetworkIsolation,
  stageSchedulingDiagnostics,
  type MppSidecarLauncherConfig,
} from '../src/index.js'
import {
  MPP_INPUT_TOO_LARGE,
  MPP_INPUT_UNREADABLE,
  MPP_MAX_INPUT_BYTES,
  MPP_OUTPUT_TOO_LARGE,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_UNSUPPORTED_FORMAT,
  emptyProjectDocument,
} from '@genoffice/project-file'

const FIXTURES = join(import.meta.dirname, 'fixtures')
let workspace: string

beforeEach(() => {
  workspace = mkdtempSync(join(tmpdir(), 'launcher-test-'))
})
afterEach(() => {
  rmSync(workspace, { recursive: true, force: true })
})

/** A launcher driving a fake sidecar via the commandBuilder test seam.
 * The fake-sidecar suite exercises the PROCESS-MANAGEMENT logic
 * (timeouts, exit codes, frame validation) — orthogonal to network
 * isolation, which has its own dedicated suite (network-isolation.test.ts,
 * incl. the real java path in e2e) — so it explicitly opts out of the
 * wrapper to stay deterministic on hosts without the mechanism. */
function fakeLauncher(
  fixture: string,
  overrides: Partial<MppSidecarLauncherConfig> = {},
): MppSidecarLauncher {
  return new MppSidecarLauncher({
    mpxjHome: '/nonexistent-mpxj',
    networkIsolation: 'off',
    commandBuilder: (inputPath, outputPath, requestId) => ({
      command: process.execPath,
      args: [join(FIXTURES, fixture), inputPath, outputPath, requestId],
    }),
    ...overrides,
  })
}

describe('launcher failure classes (fake sidecar, real process management)', () => {
  it('MPP_SIDECAR_UNAVAILABLE: the executable cannot be started (wrapped: exit 127 + "failed to execute" maps precisely; probe-aware)', async () => {
    // On a host WITH the mechanism (default 'required' policy): the missing
    // java executable surfaces as the wrapper's exit 127 + "failed to
    // execute" stderr — mapped back to the semantically-correct
    // MPP_SIDECAR_UNAVAILABLE, never a misleading "conversion exited". On a
    // host WITHOUT the mechanism (e.g. a CI runner that restricts
    // unprivileged user namespaces): the default policy FAILS CLOSED with
    // MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE instead — equally precise,
    // and the wrapper-branch mapping itself is then covered by the
    // deterministic fake-unshare behavioral test. Both outcomes asserted:
    const capability = await probeNetworkIsolation()
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      javaExecutable: '/definitely/not/a/real/binary',
      timeoutMs: 5_000,
    })
    const result = await launcher.convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics).toHaveLength(1)
      if (capability.supported) {
        expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_UNAVAILABLE)
        expect(result.diagnostics[0].message).toContain('could not be started')
      } else {
        expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE)
      }
      expect(result.diagnostics[0].severity).toBe('error')
      expect(result.diagnostics[0].stage).toBe('sidecar')
    }
  })

  it('MPP_SIDECAR_UNAVAILABLE: the direct (unwrapped) spawn-error branch stays covered under the explicit off policy', async () => {
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      javaExecutable: '/definitely/not/a/real/binary',
      networkIsolation: 'off',
      timeoutMs: 5_000,
    })
    const result = await launcher.convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics).toHaveLength(1)
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_UNAVAILABLE)
      expect(result.diagnostics[0].stage).toBe('sidecar')
    }
  })

  it('the timeout kill path works THROUGH the isolation wrapper (default policy, probe-aware)', async () => {
    // On a host with the mechanism: the hanging fake sidecar runs WRAPPED
    // (unshare execs it in place, so child.kill still reaches it) and the
    // timeout still terminates it. On a host without the mechanism: the
    // default policy fails closed instead — both outcomes asserted.
    const capability = await probeNetworkIsolation()
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      timeoutMs: 700,
      commandBuilder: (inputPath, outputPath, requestId) => ({
        command: process.execPath,
        args: [join(FIXTURES, 'fake-timeout.mjs'), inputPath, outputPath, requestId],
      }),
    })
    const start = Date.now()
    const result = await launcher.convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
    const elapsed = Date.now() - start
    expect(result.ok).toBe(false)
    if (!result.ok) {
      if (capability.supported) {
        expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_TIMEOUT)
        expect(result.diagnostics[0].message).toContain('700 ms')
        expect(elapsed).toBeGreaterThanOrEqual(600)
        expect(elapsed).toBeLessThan(10_000)
      } else {
        expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE)
      }
    }
  })

  it('MPP_SIDECAR_TIMEOUT: a hanging sidecar is terminated and diagnosed', async () => {
    const launcher = fakeLauncher('fake-timeout.mjs', { timeoutMs: 700 })
    const start = Date.now()
    const result = await launcher.convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
    const elapsed = Date.now() - start
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_TIMEOUT)
      expect(result.diagnostics[0].message).toContain('700 ms')
    }
    // The timeout actually fired (not an accidental fast failure):
    expect(elapsed).toBeGreaterThanOrEqual(600)
    expect(elapsed).toBeLessThan(10_000)
  })

  it('MPP_SIDECAR_EXIT: a crashing sidecar (exit 1 + stderr) is diagnosed with the stderr tail', async () => {
    const result = await fakeLauncher('fake-crash.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_EXIT)
      expect(result.diagnostics[0].message).toContain('boom (fake sidecar crash)')
    }
  })

  it('MPP_SIDECAR_RESPONSE_INVALID: non-JSON stdout is rejected', async () => {
    const result = await fakeLauncher('fake-garbage.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_RESPONSE_INVALID)
    }
  })

  it('MPP_SIDECAR_RESPONSE_INVALID: a request-id mismatch is rejected (correlation safety)', async () => {
    const result = await fakeLauncher('fake-wrong-request-id.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
      'expected-id',
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_RESPONSE_INVALID)
    }
  })

  it('MPP_UNSUPPORTED_FORMAT: the protocol failure frame maps to the foundation code', async () => {
    const result = await fakeLauncher('fake-unsupported.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_UNSUPPORTED_FORMAT)
      expect(result.diagnostics[0].message).toContain('unrecognized project format')
    }
  })

  it('MPP_SIDECAR_RESPONSE_INVALID: ok=true without an output file is rejected', async () => {
    const result = await fakeLauncher('fake-no-output.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_RESPONSE_INVALID)
      expect(result.diagnostics[0].message).toContain('no MSPDI output file')
    }
  })

  it('MPP_OUTPUT_TOO_LARGE: an oversized output is rejected and the file removed', async () => {
    const result = await fakeLauncher('fake-oversize.mjs', { maxOutputBytes: 1_024 }).convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_OUTPUT_TOO_LARGE)
      expect(result.diagnostics[0].message).toContain('65536')
    }
    expect(existsSync(join(workspace, 'out.xml'))).toBe(false)
  })

  it('banner noise before the frame is tolerated (last-line authority)', async () => {
    const result = await fakeLauncher('fake-noise.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
    )
    expect(result.ok).toBe(true)
  })

  it('the success path returns the MSPDI bytes and the validated frame', async () => {
    const result = await fakeLauncher('fake-ok.mjs').convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
      'req-1',
    )
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.frame.requestId).toBe('req-1')
      expect(result.frame.counts?.tasks).toBe(1)
      const text = new TextDecoder().decode(result.mspdiBytes)
      expect(text).toContain('<Project')
      expect(text).toContain('Fake Project')
    }
  })
})

describe('pipeline failure classes (importMpp* over the fake sidecar)', () => {
  it('MPP_INPUT_TOO_LARGE: oversized input is refused before any process spawns', async () => {
    let spawned = false
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      commandBuilder: (inputPath, outputPath, requestId) => {
        spawned = true
        return {
          command: process.execPath,
          args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
        }
      },
    })
    const huge = new Uint8Array(MPP_MAX_INPUT_BYTES + 1)
    const result = await importMppFromBytes(huge, { launcher })
    expect(spawned).toBe(false)
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].code).toBe(MPP_INPUT_TOO_LARGE)
    expect(result.diagnostics[0].stage).toBe('sidecar')
  })

  it('malformed MSPDI from the sidecar is an mspdi-stage INVALID_MSPDI error (atomic)', async () => {
    const result = await importMppFromBytes(new Uint8Array([1, 2, 3]), {
      launcher: fakeLauncher('fake-bad-xml.mjs'),
    })
    expect(result.document).toEqual(emptyProjectDocument())
    const invalid = result.diagnostics.filter(
      (d) => d.code === 'INVALID_MSPDI' && d.stage === 'mspdi',
    )
    expect(invalid).toHaveLength(1)
    expect(invalid[0].severity).toBe('error')
  })

  it('a canonical-validation rejection is atomic: empty document + canonical-stage error', async () => {
    const result = await importMppFromBytes(new Uint8Array([1, 2, 3]), {
      launcher: fakeLauncher('fake-cycle-xml.mjs'),
    })
    expect(result.document).toEqual(emptyProjectDocument())
    const canonical = result.diagnostics.filter((d) => d.stage === 'canonical')
    expect(canonical.length).toBeGreaterThan(0)
    expect(canonical.every((d) => d.severity === 'error')).toBe(true)
    expect(result.diagnostics.some((d) => d.stage === 'mspdi')).toBe(true)
  })

  it('a sidecar failure propagates only sidecar-stage diagnostics (atomic)', async () => {
    const result = await importMppFromBytes(new Uint8Array([1, 2, 3]), {
      launcher: fakeLauncher('fake-crash.mjs'),
    })
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.schedule).toBeUndefined()
    expect(result.diagnostics.every((d) => d.stage === 'sidecar')).toBe(true)
  })

  it('the success path returns a scheduled, validated document with staged diagnostics', async () => {
    const result = await importMppFromBytes(new Uint8Array([1, 2, 3]), {
      launcher: fakeLauncher('fake-ok.mjs'),
    })
    expect(result.document.tasks).toHaveLength(1)
    expect(result.document.tasks[0]?.name).toBe('Only Task')
    expect(result.schedule).toBeDefined()
    expect(result.schedule?.diagnostics).toEqual([])
    expect(result.schedule?.projectFinish).toBeDefined()
    const stages = new Set(result.diagnostics.map((d) => d.stage))
    expect(stages.has('mspdi')).toBe(true)
    expect(result.diagnostics.every((d) => d.severity !== 'error')).toBe(true)
  })

  it('importMppFromFile works from a path and never mutates the input', async () => {
    const inputPath = join(workspace, 'input.mpp')
    const original = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3])
    writeFileSync(inputPath, original)
    const result = await importMppFromFile(inputPath, { launcher: fakeLauncher('fake-ok.mjs') })
    expect(result.document.tasks).toHaveLength(1)
    expect([...readFileSync(inputPath)]).toEqual([...original])
    // The temp workspace is cleaned up (no stray converted.mspdi anywhere):
    expect(existsSync(join(workspace, 'converted.mspdi'))).toBe(false)
  })

  it('scheduling-failure diagnostics are staged verbatim (stage mapping)', () => {
    const staged = stageSchedulingDiagnostics([
      { code: 'DEPENDENCY_CYCLE', severity: 'error', message: 'cycle detected' },
      { code: 'CALC_ERROR', severity: 'error', message: 'boom' },
    ])
    expect(staged).toEqual([
      {
        code: 'DEPENDENCY_CYCLE',
        severity: 'error',
        message: 'cycle detected',
        stage: 'scheduling',
      },
      { code: 'CALC_ERROR', severity: 'error', message: 'boom', stage: 'scheduling' },
    ])
  })
})

describe('input-side diagnostic provenance (unreadable ≠ oversized)', () => {
  it('MPP_INPUT_UNREADABLE: a MISSING input file is an input-side failure, never MPP_INPUT_TOO_LARGE', async () => {
    let built = false
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      networkIsolation: 'off',
      commandBuilder: (inputPath, outputPath, requestId) => {
        built = true
        return {
          command: process.execPath,
          args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
        }
      },
    })
    const result = await importMppFromFile(join(workspace, 'missing.mpp'), { launcher })
    // The sidecar command is never even built — let alone spawned:
    expect(built).toBe(false)
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.schedule).toBeUndefined()
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].code).toBe(MPP_INPUT_UNREADABLE)
    expect(result.diagnostics[0].code).not.toBe(MPP_INPUT_TOO_LARGE)
    expect(result.diagnostics[0].severity).toBe('error')
    expect(result.diagnostics[0].stage).toBe('sidecar')
    // The OS reason is preserved for provenance (missing file → ENOENT):
    expect(result.diagnostics[0].message).toContain('missing.mpp')
    expect(result.diagnostics[0].message).toContain('ENOENT')
  })

  it('MPP_INPUT_UNREADABLE: a PERMISSION-DENIED input is distinguished from both missing and oversized', async () => {
    const inputPath = join(workspace, 'no-read-permission.mpp')
    writeFileSync(inputPath, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]))
    chmodSync(inputPath, 0o000)
    let built = false
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      networkIsolation: 'off',
      commandBuilder: (inputPath, outputPath, requestId) => {
        built = true
        return {
          command: process.execPath,
          args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
        }
      },
    })
    const result = await importMppFromFile(inputPath, { launcher })
    // The input gate opens the file for reading, so the permission failure
    // is caught input-side (stat alone would pass — it needs no read
    // permission on the file itself) and nothing is ever built or spawned:
    expect(built).toBe(false)
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.schedule).toBeUndefined()
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].code).toBe(MPP_INPUT_UNREADABLE)
    expect(result.diagnostics[0].stage).toBe('sidecar')
    // The OS reason is preserved (permission → EACCES), and the failure is
    // not mislabeled as a size problem:
    expect(result.diagnostics[0].message).toContain('EACCES')
    expect(result.diagnostics[0].code).not.toBe(MPP_INPUT_TOO_LARGE)
  })

  it('MPP_INPUT_TOO_LARGE: an oversized input FILE (not bytes) is refused before any spawn', async () => {
    const inputPath = join(workspace, 'oversized.mpp')
    writeFileSync(inputPath, '')
    // Sparse file: the input gate reads size metadata, so this is a real
    // >100 MiB input without materializing 100 MiB of data:
    truncateSync(inputPath, MPP_MAX_INPUT_BYTES + 1)
    let built = false
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      networkIsolation: 'off',
      commandBuilder: (inputPath, outputPath, requestId) => {
        built = true
        return {
          command: process.execPath,
          args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
        }
      },
    })
    const result = await importMppFromFile(inputPath, { launcher })
    expect(built).toBe(false)
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].code).toBe(MPP_INPUT_TOO_LARGE)
    expect(result.diagnostics[0].code).not.toBe(MPP_INPUT_UNREADABLE)
    expect(result.diagnostics[0].message).toContain(String(MPP_MAX_INPUT_BYTES))
  })
})

describe('no-shell-injection guarantees', () => {
  it('buildSidecarCommand produces a direct argument array with the paths as argv entries', () => {
    const config: MppSidecarLauncherConfig = {
      mpxjHome: '/opt/mpxj-16.7.0',
      sidecarSource: '/app/java/MppSidecar.java',
      javaArgs: ['-Djava.awt.headless=true', '-Xmx512m'],
    }
    const { command, args } = buildSidecarCommand(
      config,
      '/tmp/in; rm -rf ~.mpp',
      '/tmp/out $(whoami).xml',
      'req-"; echo pwned',
    )
    expect(command).toBe('java')
    // The hostile paths survive as SINGLE argv entries — no interpolation
    // into a command string ever happens:
    expect(args[args.length - 3]).toBe('/tmp/in; rm -rf ~.mpp')
    expect(args[args.length - 2]).toBe('/tmp/out $(whoami).xml')
    expect(args[args.length - 1]).toBe('req-"; echo pwned')
    expect(args).toContain('-Djava.awt.headless=true')
    expect(args).toContain('-Xmx512m')
    expect(args.join(' ')).toContain('/opt/mpxj-16.7.0/mpxj.jar:/opt/mpxj-16.7.0/lib/*')
  })

  it('a hostile FILENAME cannot break out of argv (behavioral counter-proof)', async () => {
    // A file whose name is a shell-injection payload: with any shell
    // interpolation, this would create an extra artifact; with argv it is
    // just a filename.
    const hostile = join(workspace, 'in; touch PWNED.mpp')
    writeFileSync(hostile, new Uint8Array([1, 2, 3]))
    const result = await importMppFromFile(hostile, { launcher: fakeLauncher('fake-ok.mjs') })
    expect(result.document.tasks).toHaveLength(1)
    expect(existsSync(join(workspace, 'PWNED.mpp'))).toBe(false)
    expect(existsSync(join(process.cwd(), 'PWNED.mpp'))).toBe(false)
  })
})
