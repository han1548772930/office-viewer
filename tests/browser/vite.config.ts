import { defineConfig } from 'vite'

export default defineConfig({
  root: process.cwd(),
  publicDir: 'public',
  server: { fs: { allow: [process.cwd()] } },
})
