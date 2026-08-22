/**
 * Increment 5B — Real Electron CDP smoke test (vitest wrapper).
 *
 * This test launches the REAL Sheets Electron app under Xvfb, drives the
 * renderer via CDP, and verifies the full migration path:
 *
 *   renderer → preload → ipcRenderer.invoke('workbook:read-range')
 *   → migrated handler → SheetsShellCoordinator → SpreadsheetService
 *   → ElectronXlsxSidecarEngine → shared sidecar process → Rust binary
 *   → response → renderer
 *
 * SKIPPED if the Rust sidecar binary or the XLSX fixture is unavailable,
 * OR if Xvfb is not installed. Run with: npx vitest run tests/sheets-cdp-real.test.ts
 *
 * The actual driver logic lives in scripts/sheets-cdp-smoke.mjs. This test
 * file spawns the script and asserts ALL CHECKS PASSED appears in stdout.
 */
import { describe, test, expect } from 'vitest'
import { existsSync } from 'node:fs'
import { join, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

const here = dirname(fileURLToPath(import.meta.url))
// here = .../apps/sheets/tests — go up 3 levels to reach the repo root
const repoRoot = resolve(here, '..', '..', '..')
const sidecarBin = join(repoRoot, 'apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar')
const fixture = join(repoRoot, 'apps/sheets/fixtures/generated/compatibility-basic.xlsx')
const smokeScript = join(repoRoot, 'scripts/sheets-cdp-smoke.mjs')
const electronBin = join(repoRoot, 'node_modules/electron/dist/electron')

function have(cmd: string): boolean {
  try {
    const r = spawnSync('which', [cmd], { encoding: 'utf8' })
    return r.status === 0 && r.stdout.trim().length > 0
  } catch { return false }
}

const sidecarAvailable = existsSync(sidecarBin)
const fixtureAvailable = existsSync(fixture)
const scriptAvailable = existsSync(smokeScript)
const electronAvailable = existsSync(electronBin)
const xvfbAvailable = have('Xvfb')

const canRun = sidecarAvailable && fixtureAvailable && scriptAvailable && electronAvailable && xvfbAvailable

describe.skipIf(!canRun)('Increment 5B — Real Electron CDP smoke test', () => {
  test('real renderer-driven read-range crosses the full migration path', () => {
    const result = spawnSync('node', [smokeScript], {
      encoding: 'utf8',
      cwd: repoRoot,
      timeout: 120_000,
      env: {
        ...process.env,
        // Use non-default ports to avoid conflicts with any leftover processes
        CDP_PORT: '9888',
      },
    })

    const stdout = result.stdout ?? ''
    const stderr = result.stderr ?? ''
    const combined = stdout + '\n' + stderr

    if (result.status !== 0) {
      console.error('CDP smoke script failed:')
      console.error('--- STDOUT ---')
      console.error(stdout)
      console.error('--- STDERR ---')
      console.error(stderr)
    }

    expect(result.status).toBe(0)
    expect(combined).toContain('[LEGACY-SELECT] SUCCESS')
    expect(combined).toContain('[SIDECAR-SHARING] SUCCESS — exactly ONE sidecar process')
    expect(combined).toContain('[MIGRATED-READ-RANGE] SUCCESS')
    expect(combined).toContain('[MIGRATED-READ-FORMULAS] SUCCESS')
    expect(combined).toContain('[MIGRATED-SAVE] SUCCESS')
    expect(combined).toContain('[SESSION-CONTINUITY] SUCCESS')
    expect(combined).toContain('[SAVE-RESPONSE-FIDELITY] SUCCESS')
    expect(combined).toContain('[SAVE-CONTENT-FIDELITY] SUCCESS')
    expect(combined).toContain('[MIGRATED-PDF-EXPORT] SUCCESS')
    expect(combined).toContain('[MIGRATED-SCREEN-CAPTURE] SUCCESS')
    expect(combined).toContain('[MIGRATED-FILES]')
    expect(combined).toContain('[PIVOT-READ] SUCCESS')
    expect(combined).toContain('[AUTO-RENAME]')
    expect(combined).toContain('[INVALID-SESSION] SUCCESS')
    expect(combined).toContain('ALL CHECKS PASSED')
  }, 180_000)
})

describe('Increment 5B — Real Electron CDP smoke test (availability report)', () => {
  test('reports availability of prerequisites', () => {
    console.log('=== REAL ELECTRON CDP SMOKE TEST PREREQUISITES ===')
    console.log(`Sidecar binary: ${sidecarBin}`)
    console.log(`  available: ${sidecarAvailable}`)
    console.log(`XLSX fixture: ${fixture}`)
    console.log(`  available: ${fixtureAvailable}`)
    console.log(`Smoke script: ${smokeScript}`)
    console.log(`  available: ${scriptAvailable}`)
    console.log(`Electron binary: ${electronBin}`)
    console.log(`  available: ${electronAvailable}`)
    console.log(`Xvfb: available: ${xvfbAvailable}`)

    if (!canRun) {
      console.log('REAL SHEETS E2E IPC: SKIPPED (prerequisites not met)')
    } else {
      console.log('REAL SHEETS E2E IPC: READY (test will run)')
    }
    expect(true).toBe(true)
  })
})
