import { readdir } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { spawn } from 'node:child_process'

const mode = process.argv[2] ?? 'unit'
if (!['unit', 'components'].includes(mode)) throw new Error(`Unknown test suite: ${mode}`)
const workspace = fileURLToPath(new URL('../', import.meta.url))
const root = join(workspace, 'packages/office-studio/tests')
const paths = await readdir(root, { recursive: true })
const files = paths.filter((path) => path.endsWith('.test.mjs') && path.endsWith('.browser.test.mjs') === (mode === 'components'))
  .sort().map((path) => join(root, path))
if (!files.length) throw new Error(`No tests found for ${mode}`)
const child = spawn(process.execPath, ['--experimental-strip-types', '--test', `--test-concurrency=${mode === 'components' ? 1 : 4}`, ...files], { cwd: workspace, stdio: 'inherit' })
child.on('error', (error) => { console.error(error); process.exitCode = 1 })
child.on('exit', (code, signal) => { process.exitCode = signal ? 1 : code ?? 1 })
