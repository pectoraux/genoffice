// Fake sidecar: UNSUPPORTED FORMAT — protocol ok=false frame, exit 0.
const [, , , , requestId] = process.argv
console.log(
  JSON.stringify({
    version: 1,
    requestId,
    ok: false,
    error: { code: 'UNSUPPORTED_MPP_FORMAT', message: 'fake: unrecognized project format' },
  }),
)
