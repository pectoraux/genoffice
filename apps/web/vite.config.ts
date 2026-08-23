import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Browser dev server. /api/* is proxied to the Node HTTP host (default :5179).
// The browser NEVER imports Electron, pg, pglite, repositories, or services.
export default defineConfig({
  plugins: [react()],
  // Pre-bundle the prosemirror-tables re-export so the first E2E import of the table-actions test host does not trigger a mid-test dep re-optimization (full page reload).
  optimizeDeps: { include: ['@tiptap/pm/tables'] },
  server: {
    port: 5178,
    proxy: {
      '/api': {
        target: 'http://localhost:5179',
        changeOrigin: false,
      },
    },
  },
})
