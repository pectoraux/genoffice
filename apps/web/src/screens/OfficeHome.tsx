import { useMemo } from 'react'

export function OfficeHome({ onRoute }: { onRoute: (route: string) => void }) {
  const cards = useMemo(() => [
    { key: 'word', title: 'Word', description: 'Create and edit rich documents in the browser.', route: '/office/word' },
    { key: 'excel', title: 'Excel', description: 'Create and edit spreadsheets in the browser.', route: '/office/excel' },
  ], [])

  return (
    <section style={{ maxWidth: 1180, margin: '0 auto', padding: 32 }}>
      <div style={{ marginBottom: 28 }}>
        <div style={{ fontSize: 13, fontWeight: 700, letterSpacing: 1.2, textTransform: 'uppercase', opacity: 0.62 }}>GenOffice Web</div>
        <h2 style={{ margin: '8px 0', fontSize: 34 }}>Office</h2>
        <p style={{ margin: 0, opacity: 0.72, maxWidth: 720, lineHeight: 1.6 }}>
          Browser-first Word and Excel workspaces built on the same product surface as the desktop editors.
        </p>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 18 }}>
        {cards.map((card) => (
          <button key={card.key} onClick={() => onRoute(card.route)} style={{ textAlign: 'left', border: '1px solid #d7dce4', borderRadius: 18, padding: 24, background: '#fff', cursor: 'pointer', boxShadow: '0 8px 24px rgba(20,30,50,.06)' }}>
            <div style={{ fontSize: 28, fontWeight: 800, marginBottom: 10 }}>{card.title}</div>
            <div style={{ color: '#536071', lineHeight: 1.55 }}>{card.description}</div>
            <div style={{ marginTop: 22, fontWeight: 700 }}>Open editor →</div>
          </button>
        ))}
      </div>
    </section>
  )
}
