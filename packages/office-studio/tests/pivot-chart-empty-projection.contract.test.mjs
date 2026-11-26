import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Projection/ownership drafts only; not executed during implementation.
let server, projection, charts, fixture
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[projection, charts, fixture] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-record-projection.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs'),
  ])
})
afterAll(async () => { await server?.close() })
const clone = value => JSON.parse(JSON.stringify(value))

test.each(['column', 'line', 'chord', 'treemap'])('%s accepts an explicit unconfigured Pivot report without inventing records', async type => {
  const source = { fields: [], records: [], sourceCategoryFields: [], sourceValueFields: [], visibleFields: [] }
  const binding = { type, categoryField: 0, valueFields: [1] }, before = clone(binding)
  const result = projection.projectChartRecords(source, binding)
  assert.deepEqual(result.categories, []); assert.deepEqual(result.series, []); assert.deepEqual(result.records, [])
  assert.deepEqual(result.visibleFields, []); assert.deepEqual(binding, before)
  assert.deepEqual(await projection.projectChartRecordsAsync(source, binding), result)
})

test.each(['sourceCategoryFields', 'sourceValueFields'])('missing %s keeps existing author roles and actual retained field metadata', async area => {
  const source = { fields: ['Region', 'Amount'], records: [], sourceCategoryFields: [0], sourceValueFields: [1], visibleFields: [0, 1], date1904: true }
  source[area] = []
  const binding = { type: 'line', categoryField: 0, valueFields: [1], options: { mapping: { categoryIndexes: [0] } } }, before = clone(binding)
  const data = await projection.projectChartRecordsAsync(source, binding)
  assert.deepEqual(data.fields, source.fields); assert.equal(data.date1904, true)
  assert.deepEqual(data.categories, []); assert.deepEqual(data.series, []); assert.deepEqual(binding, before)
})

test('a fully configured but filtered empty report still validates invalid authored field references', () => {
  const source = { fields: ['Region', 'Amount'], records: [], sourceCategoryFields: [0], sourceValueFields: [1], visibleFields: [0, 1] }
  assert.throws(() => projection.projectChartRecords(source, { type: 'line', categoryField: 0, valueFields: [999] }), /字段|field/i)
  const valid = projection.projectChartRecords(source, { type: 'line', categoryField: 0, valueFields: [1] })
  assert.deepEqual(valid.categories, []); assert.deepEqual(valid.records, [])
})

test('ordinary empty and nonempty record sources do not use the unconfigured Pivot escape path', () => {
  const binding = { type: 'line', categoryField: 0, valueFields: [999] }
  assert.throws(() => projection.projectChartRecords({ fields: ['Region', 'Amount'], records: [] }, binding), /字段|field/i)
  assert.throws(() => projection.projectChartRecords({ fields: ['Region', 'Amount'], records: [['North', 10]], sourceCategoryFields: [], sourceValueFields: [1] }, binding), /字段|field/i)
})

test('reading an empty owned report uses the real report path and leaves every worksheet cell unchanged', () => {
  const workbook = fixture.pivotChartEngineWorkbook({ ownership: 'owned', configured: false })
  const before = clone(workbook), chart = workbook.charts[0], data = charts.readChartData(workbook, chart)
  assert.deepEqual(data.series, []); assert.deepEqual(data.categories, [])
  assert.deepEqual(workbook, before); assert.equal(chart.ownedPivot.outputRange, undefined)
  assert.equal(workbook.pivotTables.length, 0)
})
