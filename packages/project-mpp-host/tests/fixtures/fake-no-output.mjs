// Fake sidecar: MISSING OUTPUT — reports ok=true but writes no file.
const [, , , , requestId] = process.argv
console.log(
  JSON.stringify({
    version: 1,
    requestId,
    ok: true,
    counts: { tasks: 0, resources: 0, calendars: 0, predecessorLinks: 0, assignments: 0 },
  }),
)
