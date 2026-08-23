import { useState, useEffect } from 'react'
import { authApi, type SessionInfo } from '../api/client'
import { styles } from '../styles'
import { ProjectsScreen } from './Projects'
import { ProjectWorkspace } from './ProjectWorkspace'
import { OfficeHome } from './OfficeHome'
import { WordEditor } from './WordEditor'
import { ExcelEditor } from './ExcelEditor'

export function AppShell({
  route, onRoute, onLogout,
}: { route: string; onRoute: (r: string) => void; onLogout: () => Promise<void> }) {
  const [session, setSession] = useState<SessionInfo | null>(null)

  useEffect(() => {
    authApi.session().then(setSession).catch(() => setSession(null))
  }, [route])

  const logout = async () => {
    await authApi.logout()
    await onLogout()
  }

  if (route === '/office' || route === '/office/') {
    return <OfficeFrame session={session} onRoute={onRoute} onLogout={logout}><OfficeHome onRoute={onRoute} /></OfficeFrame>
  }
  if (route === '/office/word') return <OfficeFrame session={session} onRoute={onRoute} onLogout={logout}><WordEditor onRoute={onRoute} /></OfficeFrame>
  if (route === '/office/excel') return <OfficeFrame session={session} onRoute={onRoute} onLogout={logout}><ExcelEditor onRoute={onRoute} /></OfficeFrame>

  const m = route.match(/^\/projects\/([^/]+)$/)
  const projectId = m ? m[1] : null
  return (
    <div style={styles.app}>
      <header style={styles.header}>
        <h1 style={styles.headerTitle}>Contractor GenOffice{session?.displayName ? ` — ${session.displayName}` : ''}</h1>
        <div style={styles.headerRight}>
          <button style={styles.button} onClick={() => onRoute('/office')}>Office</button>
          <button style={styles.button} onClick={() => onRoute('/projects')}>Projects</button>
          <button style={styles.button} onClick={logout}>Sign out</button>
        </div>
      </header>
      <main style={styles.main}>{projectId ? <ProjectWorkspace projectId={projectId} onRoute={onRoute} /> : <ProjectsScreen onRoute={onRoute} />}</main>
    </div>
  )
}

function OfficeFrame({ children, session, onRoute, onLogout }: { children: React.ReactNode; session: SessionInfo | null; onRoute: (r: string) => void; onLogout: () => Promise<void> }) {
  return <div style={{ minHeight: '100vh' }}>
    <header style={styles.header}>
      <h1 style={styles.headerTitle}>GenOffice Office{session?.displayName ? ` — ${session.displayName}` : ''}</h1>
      <div style={styles.headerRight}>
        <button style={styles.button} onClick={() => onRoute('/office')}>Office</button>
        <button style={styles.button} onClick={onLogout}>Sign out</button>
      </div>
    </header>
    {children}
  </div>
}
