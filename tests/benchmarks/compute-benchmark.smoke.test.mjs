import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { mkdtemp, readFile, rm, stat } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { tmpdir } from 'node:os'
import { test } from 'vitest'

const workspaceRoot = resolve(import.meta.dirname, '../..')
const benchmarkScript = resolve(workspaceRoot, 'scripts/benchmarks/compute-benchmark.mjs')

test('compute benchmark smoke keeps JS and Worker/WASM indices equivalent at small sizes', async () => {
  const outputDirectory = await mkdtemp(join(tmpdir(), 'office-compute-benchmark-'))
  try {
    const run = spawnSync(process.execPath, [benchmarkScript, '--sizes', '128,512', '--iterations', '1', '--output', outputDirectory], {
      cwd: workspaceRoot,
      encoding: 'utf8',
      timeout: 120_000,
    })
    assert.equal(run.status, 0, `benchmark failed\nstdout:\n${run.stdout}\nstderr:\n${run.stderr}`)
    const output = JSON.parse(run.stdout)
    const result = output.result
    assert.deepEqual(result.workloads.map((workload) => workload.rowCount), [128, 512])
    for (const workload of result.workloads) {
      assert.equal(workload.defaultRoute.sortUsesWorker, false)
      assert.equal(workload.defaultRoute.filterUsesWorker, false)
      const samples = [...workload.typescript, ...workload.cold, ...workload.warm]
      assert.ok(samples.length > 0)
      assert.ok(samples.every((sample) => sample.parity))
      assert.ok(samples.every((sample) => sample.rowCount === workload.rowCount))
      assert.ok(workload.cold.every((sample) => sample.workerComputeMs !== null && sample.workerQueueAndStartupMs !== null))
    }
    assert.match(result.routing.thresholdStatus, /provisional/)
    assert.match(result.methodology.coldDefinition, /not a guaranteed disk-cold load/)
    await stat(output.jsonPath)
    await stat(output.markdownPath)
    const markdown = await readFile(output.markdownPath, 'utf8')
    assert.match(markdown, /does not claim a performance target was met/)
    assert.match(markdown, /Limitations/)
  } finally {
    await rm(outputDirectory, { recursive: true, force: true })
  }
})
