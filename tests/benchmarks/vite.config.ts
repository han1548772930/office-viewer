import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { defineConfig } from 'vite'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))

export default defineConfig({
  root: workspaceRoot,
  publicDir: resolve(workspaceRoot, 'public'),
  server: { fs: { allow: [workspaceRoot] } },
  logLevel: 'error',
})
