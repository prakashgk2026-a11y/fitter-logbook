import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'

// VITE_USE_SHIM=1 swaps Supabase for an in-browser Postgres (used only by the automated UI tests).
const shim = process.env.VITE_USE_SHIM === '1'

export default defineConfig(({ mode }) => {
  const env = { ...loadEnv(mode, process.cwd(), 'VITE_'), ...process.env }
  const appName = (env.VITE_APP_NAME || '').trim() || 'ITI Logbook'     // shown in the browser-tab title
  return {
    base: './',                       // works from any GitHub Pages sub-path
    plugins: [
      react(),
      { name: 'app-title', transformIndexHtml: (html) => html.replace(/<title>.*?<\/title>/, `<title>${appName.replace(/[<>&]/g, '')}</title>`) },
    ],
    resolve: shim ? { alias: { '@supabase/supabase-js': fileURLToPath(new URL('./tests/shim/supabase-shim.js', import.meta.url)) } } : {},
    optimizeDeps: { exclude: ['@electric-sql/pglite'] },
    build: { chunkSizeWarningLimit: 1500 },
  }
})
