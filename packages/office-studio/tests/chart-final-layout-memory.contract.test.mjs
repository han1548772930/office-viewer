import assert from 'node:assert/strict'
import { test } from 'vitest'
import { estimateChartFinalLayoutMemory } from '../src/features/charts/chart-compute-memory.ts'
import { assertChartPreparationMemory } from '../src/features/charts/chart-compute-lifecycle.ts'
import { estimateChartTaskMemory } from '../src/compute/chart-tasks.ts'

// Deferred admission contracts; these are not performance or heap measurements.
const source = { categories: ['A', 'B'], series: [{ field: 1, name: 'Value', values: [0, 1] }] }
const chart = type => ({ type, options: { mapping: { useDateAxis: true } } })

test('a small histogram source cannot admit ten thousand final bins using its original row count', () => {
  const small = { histogram: { validSampleCount: 2, bins: Array(2).fill({ kind: 'regular', count: 1 }) } }
  const expanded = { histogram: { validSampleCount: 2, bins: Array(10_000).fill({ kind: 'regular', count: 0 }) } }
  const budget = estimateChartFinalLayoutMemory(source, chart('histogram'), small)
  assertChartPreparationMemory(budget, { memoryBudgetBytes: budget })
  const actualAdmission = estimateChartFinalLayoutMemory(source, chart('histogram'), expanded)
  assert.ok(actualAdmission > budget + 4 * 1024 * 1024)
  assert.throws(() => assertChartPreparationMemory(actualAdmission, { memoryBudgetBytes: budget }),
    error => error.code === 'COMPUTE_MEMORY_LIMIT')
})

test('final boxplot outliers and trend/error extents are included above retained statistics', () => {
  const empty = estimateChartFinalLayoutMemory(source, chart('boxplot'), {})
  const full = estimateChartFinalLayoutMemory(source, chart('boxplot'), {
    boxplots: [{ field: 1, seriesPosition: 0, statistics: { min: 0, q1: 0, median: 0, q3: 0, max: 0,
      outliers: Array(20_000).fill(1), outlierSourceIndices: Array.from({ length: 20_000 }, (_value, index) => index) } }],
    trendlines: [{ configIndex: 0, seriesPosition: 0, statistics: { points: Array(1000).fill([0, 1]) } }],
    errorBars: [{ configIndex: 0, seriesPosition: 0, intervals: Array(2000).fill({ x: 0, y: 1, minus: 1, plus: 1, categoryIndex: 0 }) }],
  })
  assert.ok(full > empty + 2 * 1024 * 1024)
})

test('boxplot Worker admission reserves an owned identity buffer even when every sample can be an outlier', () => {
  const params = { rowCount: 12, columnCount: 1, categoryCount: 12,
    series: [{ field: 1, seriesPosition: 0, column: 0 }],
    jobs: [{ kind: 'boxplot', seriesIndex: 0, slot: 0, quartileMethod: 'inclusive', whiskerMultiplier: 0 }] }
  const estimate = estimateChartTaskMemory(params)
  assert.equal(estimate.indexCapacity, params.rowCount)
  assert.ok(estimate.outputBytes >= (params.rowCount + 5) * 8 + params.rowCount * 4)
  assert.ok(estimate.estimatedPeakBytes > estimate.inputBytes * 2 + estimate.outputBytes * 2 + params.rowCount * 16)
})
