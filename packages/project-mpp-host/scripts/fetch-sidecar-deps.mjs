#!/usr/bin/env node
/**
 * PROJECT-018 — fetch the pinned MPXJ distribution and the real-MPP corpus.
 *
 * Downloads (idempotent, SHA-256-verified, never committed — .gitignore'd
 * under packages/project-mpp-host/.sidecar-deps/):
 *
 *   .sidecar-deps/mpxj-16.7.0/   the MPXJ distribution (mpxj.jar + lib/)
 *   .sidecar-deps/corpus/        the 8 real corpus .mpp files
 *
 * Provenance and checksums live in corpus/corpus-manifest.json (the corpus
 * is pinned to MPXJ commit abdbf6ef85654e3eff35c11c5e76cf08da842dce =
 * tag v16.7.0). The script uses only Node built-ins + the system `unzip`
 * (present on GitHub Actions runners and standard dev environments).
 *
 * LGPL posture: nothing downloaded here enters the repository or any
 * distribution artifact — the repository ships only this script, the
 * manifest, and the in-repo Apache-2.0 sidecar source. Obligations for a
 * future BUNDLED distribution are documented in LICENSE-THIRD-PARTY.md.
 */
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const pkgRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const depsRoot = join(pkgRoot, '.sidecar-deps')
const manifest = JSON.parse(readFileSync(join(pkgRoot, 'corpus', 'corpus-manifest.json'), 'utf8'))

async function fetchBytes(url) {
  const response = await fetch(url, { redirect: 'follow' })
  if (!response.ok) throw new Error(`download failed (${response.status}): ${url}`)
  return new Uint8Array(await response.arrayBuffer())
}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}

async function ensureFile(url, expectedSha256, destination) {
  if (existsSync(destination)) {
    const current = readFileSync(destination)
    if (sha256(current) === expectedSha256) {
      console.log(`  ✓ ${destination} (cached)`)
      return
    }
    console.log(`  ! ${destination} exists with a wrong checksum — re-downloading`)
  }
  mkdirSync(dirname(destination), { recursive: true })
  process.stdout.write(`  ↓ ${url} … `)
  const bytes = await fetchBytes(url)
  const digest = sha256(bytes)
  if (digest !== expectedSha256) {
    throw new Error(
      `checksum mismatch for ${url}\n  expected ${expectedSha256}\n  actual   ${digest}`,
    )
  }
  writeFileSync(destination, bytes)
  console.log(`${(bytes.length / 1024 / 1024).toFixed(1)} MiB ✓`)
}

async function main() {
  // ── MPXJ distribution ─────────────────────────────────────────────────
  const mpxjDir = join(depsRoot, `mpxj-${manifest.mpxj.version}`)
  const mpxjMarker = join(mpxjDir, 'mpxj.jar')
  if (existsSync(mpxjMarker) && statSync(mpxjMarker).size > 0) {
    console.log(`✓ MPXJ ${manifest.mpxj.version} already extracted at ${mpxjDir}`)
  } else {
    const zipPath = join(depsRoot, `mpxj-${manifest.mpxj.version}.zip`)
    console.log(`MPXJ ${manifest.mpxj.version}:`)
    await ensureFile(manifest.mpxj.url, manifest.mpxj.sha256, zipPath)
    const staging = join(depsRoot, '.unzip-staging')
    execFileSync('unzip', ['-oq', zipPath, '-d', staging])
    // The zip extracts a top-level directory; locate the one containing mpxj.jar.
    const { readdirSync } = await import('node:fs')
    let extracted = null
    for (const entry of readdirSync(staging)) {
      const candidate = join(staging, entry)
      if (existsSync(join(candidate, 'mpxj.jar'))) extracted = candidate
    }
    if (extracted === null) throw new Error('mpxj.jar not found inside the downloaded zip')
    mkdirSync(depsRoot, { recursive: true })
    const { renameSync, rmSync } = await import('node:fs')
    rmSync(mpxjDir, { recursive: true, force: true })
    renameSync(extracted, mpxjDir)
    rmSync(staging, { recursive: true, force: true })
    console.log(`  ✓ extracted to ${mpxjDir}`)
  }

  // ── Corpus ─────────────────────────────────────────────────────────────
  const corpusDir = join(depsRoot, 'corpus')
  console.log('Real-MPP corpus (pinned to MPXJ commit ' + manifest.source.pinnedCommit + '):')
  mkdirSync(corpusDir, { recursive: true })
  for (const file of manifest.corpus) {
    const url = `https://raw.githubusercontent.com/joniles/mpxj/${manifest.source.pinnedCommit}/${file.sourcePath}`
    await ensureFile(url, file.sha256, join(corpusDir, file.filename))
  }
  console.log('All sidecar dependencies are in place and checksum-verified.')
}

main().catch((error) => {
  console.error(String(error))
  process.exit(1)
})
