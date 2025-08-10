import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: './',          // <= important: use relative asset paths
  build: { outDir: 'dist' }
})