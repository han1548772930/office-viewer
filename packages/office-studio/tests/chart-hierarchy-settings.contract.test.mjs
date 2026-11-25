import assert from 'node:assert/strict'
import { test } from 'vitest'
import { addChartHierarchyField, chartHierarchySettingsMapping, moveChartHierarchyField, removeChartHierarchyField,
  replaceChartHierarchyField, setChartHierarchySettingsMapping, setChartHierarchyWeightField } from '../src/features/charts/chart-hierarchy-settings.ts'
import { createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { updateChart } from '../src/features/charts/chart-engine.ts'
import { treemapFixture } from './fixtures/chart-treemap-fixtures.mjs'
import { sunburstFixture } from './fixtures/chart-sunburst-fixtures.mjs'

// Ordered hierarchy UI/model drafts. They have not been executed.
const json = value => JSON.parse(JSON.stringify(value))

test('add, replace, move, remove and weight changes preserve root-to-leaf order and detached inputs', () => {
  const fields = [0, 1], config = { type: 'sunburst', categoryField: 0, valueFields: [4] }
  setChartHierarchySettingsMapping(config, fields, 4, 5); fields.reverse()
  assert.deepEqual(chartHierarchySettingsMapping(config, 5), { hierarchyIndexes: [0, 1], valueIndex: 4 })
  addChartHierarchyField(config, 2, 5); moveChartHierarchyField(config, 2, -1, 5)
  assert.deepEqual(config.options.mapping.hierarchyIndexes, [0, 2, 1])
  replaceChartHierarchyField(config, 0, 3, 5); removeChartHierarchyField(config, 1, 5)
  assert.deepEqual(config.options.mapping.hierarchyIndexes, [3, 1])
  setChartHierarchyWeightField(config, 0, 5); assert.deepEqual(config.valueFields, [0])
})

test('invalid add/replace/move/remove/weight roles reject atomically, including removal of the last hierarchy', () => {
  const config = { type: 'sunburst', categoryField: 0, valueFields: [3], options: { mapping: { hierarchyIndexes: [0, 1] } } }, before = json(config)
  for (const action of [() => addChartHierarchyField(config, 1, 4), () => addChartHierarchyField(config, 3, 4),
    () => addChartHierarchyField(config, 4, 4), () => replaceChartHierarchyField(config, 0, 1, 4),
    () => replaceChartHierarchyField(config, 4, 2, 4), () => moveChartHierarchyField(config, 0, -1, 4),
    () => moveChartHierarchyField(config, 1, 1, 4), () => moveChartHierarchyField(config, 0, 0, 4),
    () => removeChartHierarchyField(config, -1, 4), () => setChartHierarchyWeightField(config, 0, 4),
    () => setChartHierarchySettingsMapping(config, [], 3, 4), () => setChartHierarchySettingsMapping(config, [0, 1], 3, 2)]) {
    assert.throws(action); assert.deepEqual(config, before)
  }
  removeChartHierarchyField(config, 1, 4); const one = json(config)
  assert.throws(() => removeChartHierarchyField(config, 0, 4), /至少保留/); assert.deepEqual(config, one)
  const empty = { type: 'treemap' }
  assert.throws(() => setChartHierarchySettingsMapping(empty, [0], 0, 2)); assert.deepEqual(empty, { type: 'treemap' })
})

test('both production hierarchy types hold draft changes until one update and retain independent type settings', () => {
  for (const fixture of [treemapFixture, sunburstFixture]) {
    const { workbook, chart } = fixture(), before = json(chart), draft = createChartSettingsDraft(chart)
    moveChartHierarchyField(draft.config, 0, 1, 3)
    assert.deepEqual(draft.config.options.mapping.hierarchyIndexes, [1, 0]); assert.deepEqual(json(chart), before)
    updateChart(workbook, chart, chartSettingsPatch(draft))
    assert.equal(chart.revision, before.revision + 1); assert.deepEqual(chart.options.mapping.hierarchyIndexes, [1, 0])
    assert.deepEqual(chart.valueFields, [2]); assert.equal(chart.type, before.type)
  }
})
