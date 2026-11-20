import { chromium } from '@playwright/test'
import { createServer } from 'vite'
import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const workspaceRoot = fileURLToPath(new URL('../..', import.meta.url))
const configFile = resolve(workspaceRoot, 'tests/benchmarks/vite.config.ts')
const outputDirectory = resolve(workspaceRoot, 'docs/feature-evidence/benchmarks')
const args = parseArgs(process.argv.slice(2))

const server = await createServer({ configFile, server: { host: '127.0.0.1', port: args.port, strictPort: true } })
await server.listen()
const address = server.httpServer.address()
if (!address || typeof address === 'string') throw new Error('Vite server did not bind an IPv4/IPv6 TCP port')
const baseUrl = `http://127.0.0.1:${address.port}`
const browser = await chromium.launch({ channel: args.channel, headless: true })
try {
  const page = await browser.newPage({ viewport: { width: args.width, height: args.height }, deviceScaleFactor: args.dpr })
  await page.goto(`${baseUrl}/tests/benchmarks/compute-benchmark.html`, { waitUntil: 'networkidle' })
  await page.waitForFunction(() => document.querySelector('#status')?.textContent === 'ready')
  const result = await page.evaluate(({ sizes, iterations }) => window.computeBenchmark.run({ sizes, iterations }), {
    sizes: args.sizes,
    iterations: args.iterations,
  })
  const stamp = result.generatedAt.replaceAll(/[:.]/g, '-')
  const jsonPath = resolve(args.output ?? outputDirectory, `compute-${stamp}.json`)
  const markdownPath = resolve(args.output ?? outputDirectory, `compute-${stamp}.md`)
  await mkdir(args.output ?? outputDirectory, { recursive: true })
  await writeFile(jsonPath, `${JSON.stringify(result, null, 2)}\n`, 'utf8')
  await writeFile(markdownPath, renderMarkdown(result), 'utf8')
  process.stdout.write(`${JSON.stringify({ jsonPath, markdownPath, result }, null, 2)}\n`)
} finally {
  await browser.close()
  await server.close()
}

function parseArgs(values) {
  const options = { sizes: [4096, 16384], iterations: 3, port: 0, channel: 'chrome', width: 1440, height: 900, dpr: 1, output: undefined }
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index]
    if (value === '--sizes') options.sizes = String(values[++index] ?? '').split(',').map((entry) => Number(entry)).filter((entry) => Number.isSafeInteger(entry) && entry > 0)
    else if (value === '--iterations') options.iterations = Number(values[++index])
    else if (value === '--port') options.port = Number(values[++index])
    else if (value === '--channel') options.channel = String(values[++index])
    else if (value === '--width') options.width = Number(values[++index])
    else if (value === '--height') options.height = Number(values[++index])
    else if (value === '--dpr') options.dpr = Number(values[++index])
    else if (value === '--output') options.output = resolve(workspaceRoot, String(values[++index]))
    else throw new Error(`Unknown argument: ${value}`)
  }
  if (!options.sizes.length || !Number.isSafeInteger(options.iterations) || options.iterations < 1) throw new Error('Use positive --sizes and --iterations values')
  if (!Number.isSafeInteger(options.port) || options.port < 0 || options.port > 65535) throw new Error('Invalid --port')
  if (!Number.isFinite(options.dpr) || options.dpr <= 0) throw new Error('Invalid --dpr')
  return options
}

function renderMarkdown(result) {
  const lines = [
    '# Office compute benchmark',
    '',
    `Generated: ${result.generatedAt}`,
    '',
    'This is a measurement record. It does not claim a performance target was met.',
    '',
    '## Environment',
    '',
    `- User agent: ${result.environment.userAgent}`,
    `- Hardware concurrency: ${result.environment.hardwareConcurrency}`,
    `- Device memory (GiB): ${result.environment.deviceMemoryGiB ?? 'unavailable'}`,
    `- Viewport: ${result.environment.viewport.width}×${result.environment.viewport.height} @ DPR ${result.environment.viewport.devicePixelRatio}`,
    `- Long task observer: ${result.environment.longTaskObserverAvailable ? 'available' : 'unavailable'}`,
    '',
    '## Routing',
    '',
    `- Configured NumericIndexService minRows: ${result.routing.configuredMinRows}`,
    `- Threshold status: ${result.routing.thresholdStatus}`,
    '',
    '## Measurements',
    '',
    '| Rows | Operation | Regime | P50 wall (ms) | P95 wall (ms) | Worker compute (ms) | Queue/startup (ms) | Boundary bytes | Peak estimate (bytes) | Long-task max (ms) |',
    '| ---: | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |',
  ]
  for (const workload of result.workloads) {
    const all = [...workload.typescript, ...workload.cold, ...workload.warm]
    for (const summary of workload.summary) {
      const matching = all.filter((sample) => sample.operation === summary.operation && sample.regime === summary.regime)
      const workerCompute = averageOrDash(matching.map((sample) => sample.workerComputeMs))
      const queue = averageOrDash(matching.map((sample) => sample.workerQueueAndStartupMs))
      const boundary = averageOrDash(matching.map((sample) => sample.boundaryBytes))
      const peak = averageOrDash(matching.map((sample) => sample.estimatedPeakBytes))
      const longTask = averageOrDash(matching.map((sample) => sample.longTasks.maxMs))
      lines.push(`| ${workload.rowCount} | ${summary.operation} | ${summary.regime} | ${format(summary.p50WallMs)} | ${format(summary.p95WallMs)} | ${workerCompute} | ${queue} | ${boundary} | ${peak} | ${longTask} |`)
    }
  }
  lines.push('', '## Limitations', '', ...result.methodology.memoryLimitations.map((item) => `- ${item}`), '')
  return `${lines.join('\n')}\n`
}

function averageOrDash(values) {
  const finite = values.filter((value) => typeof value === 'number' && Number.isFinite(value))
  return finite.length ? format(finite.reduce((sum, value) => sum + value, 0) / finite.length) : '—'
}
function format(value) { return typeof value === 'number' && Number.isFinite(value) ? value.toFixed(3) : '—' }
