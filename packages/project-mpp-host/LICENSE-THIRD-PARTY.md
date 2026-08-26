# Third-Party Notices — MPP Import Sidecar (PROJECT-018)

This package (`@genoffice/project-mpp-host`) implements GenOffice's MPP import
via an externalized MPXJ conversion sidecar, exactly as approved by the
PROJECT-017 feasibility investigation (`spec/project/mpp-feasibility.md`).
The redistribution and licensing posture is documented here in full.

## Shipped artifacts vs. downloaded artifacts

**Shipped in this repository (Apache-2.0, GenOffice-authored):**

- `java/MppSidecar.java` — the sidecar converter program (our own source;
  single-file source-launcher mode, no build step, runs on a JRE).
- `src/**` — the host-side launcher, wire protocol, and pipeline code.
- `corpus/corpus-manifest.json` + `scripts/fetch-sidecar-deps.mjs` — pinned
  download manifest and checksum-verifying fetch script.
- `tests/**` — the test suite (unit, failure-injection, and real-corpus
  end-to-end).

**NOT shipped — downloaded at development/CI time, never committed**
(`.sidecar-deps/` is gitignored; verified by SHA-256 against
`corpus/corpus-manifest.json`):

- **MPXJ 16.7.0** (`net.sf.mpxj:mpxj`) — https://github.com/joniles/mpxj
  - License: **GNU Lesser General Public License (LGPL)** —
    https://www.mpxj.org/ , https://github.com/joniles/mpxj/blob/master/LICENSE
  - Pinned distribution:
    `mpxj-16.7.0.zip` (release `v16.7.0`), SHA-256
    `2a149f3ae9f6ac1034d0b0fadeca8f98e3f1f900764640c6bd16188ec2f078bd`.
  - Transitive runtime jars ship inside the distribution (Apache POI 5.5.1
    and others; Apache-2.0 / BSD-style / EPL-class licenses per MPXJ's own
    distribution `legal/` notices).
- **The MPXJ LGPL test corpus** (8 real `.mpp` files used only by tests) —
  part of the MPXJ source distribution at pinned commit
  `abdbf6ef85654e3eff35c11c5e76cf08da842dce` (tag `v16.7.0`); consumed
  externally by pinned download with checksum verification, never copied
  into this repository.

## LGPL distribution obligations (current posture)

Today, GenOffice **does not distribute MPXJ**: the artifacts are fetched at
build/test time into a non-committed workspace. Under this model no LGPL
section-4 redistribution duty is triggered by this repository.

**If a future release BUNDLES the sidecar** (desktop installer, server
image), the following obligations apply and must be discharged before
shipping:

1. Include the MPXJ LGPL license text alongside the distributed binaries.
2. Provide the corresponding MPXJ source (or a written offer valid for at
   least three years) — the pinned-release source archive satisfies this.
3. Convey the MPXJ license text and the LGPL notice for the aggregate.
4. Preserve the user's ability to relink/replace the library — the
   sidecar-process topology (an independent JVM process communicating via
   argv + files, mere aggregation rather than a derivative work) keeps
   these obligations minimal and the proprietary GenOffice code
   uncontaminated. Do NOT convert to an in-process jar binding without
   re-evaluating LGPL section 4 duties.

## Version pinning / updates

- The MPXJ version is pinned in TWO places that a test keeps in sync:
  `MPXJ_PINNED_VERSION` (exported by `@genoffice/project-file`, the
  foundation contract) and `corpus/corpus-manifest.json` → `mpxj.version`.
- Upgrading MPXJ is a deliberate act: update both, re-verify the
  distribution SHA-256, re-run the full real-corpus suite, and re-record
  the determinism evidence (the PROJECT-017/018 pattern).

## Rejected dependency routes (documented in PROJECT-017 §18)

- `@byteink/mppjs` (MIT wrapper around a GraalVM AOT native MPXJ binary):
  failed on headless Linux in our own spike (`UnsatisfiedLinkError: No
  awt`), and LGPL §4 relinking duties are practically impossible with an
  AOT static binary. Not used.
- Aspose.Tasks (the only programmatic MPP writer): commercial SDK —
  rejected (MPP export is out of scope by direction).
- CheerpJ / browser-WASM MPXJ: commercial runtime — rejected.
