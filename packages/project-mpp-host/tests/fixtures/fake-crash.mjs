// Fake sidecar: CRASH — prints a stack trace to stderr and exits 1.
console.error('java.lang.RuntimeException: boom (fake sidecar crash)')
process.exit(1)
