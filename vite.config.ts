import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'node:fs'

// Word 加载项强制 https:本地用 mkcert 证书,前后端共用 5173 单端口
const https = fs.existsSync('certs/cert.pem')
  ? { cert: fs.readFileSync('certs/cert.pem'), key: fs.readFileSync('certs/key.pem') }
  : undefined

export default defineConfig({
  plugins: [react()],
  server: {
    host: 'localhost',
    port: 5173,
    strictPort: true,
    https,
    proxy: {
      '/api': { target: 'http://localhost:8787', changeOrigin: true },
    },
  },
})
