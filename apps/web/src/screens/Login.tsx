/**
 * LoginScreen — password login + signup (waitlist) + demo quick-login.
 *
 * The primary auth is email + password (admin, approved users).
 * Sign-up goes to the waitlist (admin approves later).
 * Demo buttons provide instant login for each role (owner/member/viewer).
 */
import { useState } from 'react'
import { authApi } from '../api/client'
import { styles } from '../styles'

export function LoginScreen({ onLoggedIn }: { onLoggedIn: () => Promise<void> }) {
  const [mode, setMode] = useState<'login' | 'signup'>('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [displayName, setDisplayName] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [success, setSuccess] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError(null)
    setSuccess(null)
    setLoading(true)
    try {
      if (mode === 'login') {
        await authApi.passwordLogin(email, password)
        await onLoggedIn()
      } else {
        const result = await authApi.signup(email, displayName || undefined)
        setSuccess(result.message)
        setMode('login')
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed')
    } finally {
      setLoading(false)
    }
  }

  const demoLogin = async (role: 'owner' | 'member' | 'viewer') => {
    setError(null)
    setLoading(true)
    try {
      await authApi.demoLogin(role)
      await onLoggedIn()
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Demo login failed')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div style={{ ...styles.app, alignItems: 'center', justifyContent: 'center', padding: 24 }}>
      <form
        onSubmit={submit}
        style={{ ...styles.card, width: 380, display: 'flex', flexDirection: 'column', gap: 16 }}
      >
        <div>
          <h1 style={styles.title}>Contractor GenOffice</h1>
          <p style={styles.subtitle}>Construction business operating system</p>
        </div>
        <div style={styles.tabRow}>
          <button
            type="button"
            style={mode === 'login' ? styles.tabActive : styles.tab}
            onClick={() => {
              setMode('login')
              setError(null)
              setSuccess(null)
            }}
          >
            Sign in
          </button>
          <button
            type="button"
            style={mode === 'signup' ? styles.tabActive : styles.tab}
            onClick={() => {
              setMode('signup')
              setError(null)
              setSuccess(null)
            }}
          >
            Request access
          </button>
        </div>
        {mode === 'login' ? (
          <>
            <label style={styles.label}>
              Email
              <input
                style={styles.input}
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                autoFocus
                autoComplete="email"
              />
            </label>
            <label style={styles.label}>
              Password
              <input
                style={styles.input}
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                autoComplete="current-password"
              />
            </label>
            <button
              type="submit"
              style={styles.buttonPrimary}
              disabled={loading || !email || !password}
            >
              {loading ? 'Signing in…' : 'Sign in'}
            </button>
          </>
        ) : (
          <>
            <label style={styles.label}>
              Email
              <input
                style={styles.input}
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                autoFocus
              />
            </label>
            <label style={styles.label}>
              Name (optional)
              <input
                style={styles.input}
                value={displayName}
                onChange={(e) => setDisplayName(e.target.value)}
              />
            </label>
            <button type="submit" style={styles.buttonPrimary} disabled={loading || !email}>
              {loading ? 'Submitting…' : 'Join waitlist'}
            </button>
          </>
        )}
        {error && <div style={styles.error}>{error}</div>}
        {success && <div style={styles.success}>{success}</div>}
        <div style={{ ...styles.card, background: 'var(--surface-subtle, #f6f7f9)' }}>
          <div style={{ ...styles.label, marginBottom: 8 }}>Demo accounts (quick login)</div>
          <div style={styles.row}>
            <button
              type="button"
              style={styles.button}
              onClick={() => demoLogin('owner')}
              disabled={loading}
            >
              Owner
            </button>
            <button
              type="button"
              style={styles.button}
              onClick={() => demoLogin('member')}
              disabled={loading}
            >
              Member
            </button>
            <button
              type="button"
              style={styles.button}
              onClick={() => demoLogin('viewer')}
              disabled={loading}
            >
              Viewer
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
