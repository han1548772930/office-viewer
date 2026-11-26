import { chromium } from '@playwright/test'
import { createServer } from 'vite'
import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

// Execution entry draft. It has not been invoked and produces no claimed
// result until the real local WASM artifact and browser workload are run.
const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))
const options = parseArgs(process.argv.slice(2))
let server, browser
try {
  server = await createServer({ configFile: resolve(workspaceRoot, 'tests/benchmarks/vite.config.ts'),
    server: { host: '127.0.0.1', port: options.port, strictPort: true } })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No slicer benchmark server port')
  browser = await chromium.launch({ channel: options.channel, headless: true })
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1 }), errors = []
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(`http://127.0.0.1:${address.port}/tests/benchmarks/slicer-member-benchmark.html`, { waitUntil: 'networkidle' })
  await page.waitForFunction(() => document.querySelector('#status')?.textContent === 'ready')
  const result = await page.evaluate(({ sizes, iterations, locale }) => window.slicerMemberBenchmark.run({ sizes, iterations, locale }), options)
  if (errors.length) throw new Error(`Slicer benchmark browser errors: ${errors.join('\n')}`)
  result.environment.browserVersion = browser.version()
  result.environment.browserChannel = options.channel
  const stamp = result.generatedAt.replaceAll(/[:.]/g, '-'), destination = options.output ?? resolve(workspaceRoot, 'docs/feature-evidence/benchmarks')
  const jsonPath = resolve(destination, `slicer-members-${stamp}.json`)
  await mkdir(destination, { recursive: true }); await writeFile(jsonPath, `${JSON.stringify(result, null, 2)}\n`, 'utf8')
  process.stdout.write(`${JSON.stringify({ jsonPath, summaries: result.workloads.map(({ rowCount, summaries }) => ({ rowCount, summaries })) }, null, 2)}\n`)
} finally { await browser?.close(); await server?.close() }

function parseArgs(values) {
  const args = { sizes: [4096, 16384, 65536], iterations: 3, locale: 'en-US', channel: 'msedge', port: 0, output: undefined }
  for (let index = 0; index < values.length; index++) {
    const option = values[index], value = values[++index]
    if (value === undefined) throw new Error(`Missing value for ${option}`)
    if (option === '--sizes') args.sizes = value.split(',').map(Number)
    else if (option === '--iterations') args.iterations = Number(value)
    else if (option === '--locale') args.locale = value
    else if (option === '--channel') args.channel = value
    else if (option === '--port') args.port = Number(value)
    else if (option === '--output') args.output = resolve(workspaceRoot, value)
    else throw new Error(`Unknown argument: ${option}`)
  }
  if (!args.sizes.length || args.sizes.some(size => !Number.isSafeInteger(size) || size < 1 || size > 1_000_000)
    || !Number.isSafeInteger(args.iterations) || args.iterations < 1 || args.iterations > 20
    || !Number.isSafeInteger(args.port) || args.port < 0 || args.port > 65535 || !args.locale.trim() || !args.channel.trim()) throw new Error('Invalid slicer benchmark arguments')
  return args
}
