import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { applyPivotChartEdit, getPivotChartReport } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { getPivotShowDataAsBaseItems } from '../src/features/pivot/services/pivot-engine.ts'
import { movePivotChartDraftField, removePivotChartDraftFields, setPivotChartDraftField,
  pivotChartDraftShowDataAsReport } from '../src/features/pivot/services/pivot-chart-field-draft.ts'
import { configurePivotFromPanel, createPivotPanelDraft, pivotPanelDraftReport, preparePivotPanelInput } from '../src/features/pivot/services/pivot-panel-state.ts'
import { changePivotShowDataAsBaseField, changePivotShowDataAsBaseItemType, changePivotShowDataAsType, createPivotShowDataAsDraft,
  PIVOT_SHOW_DATA_AS_OPTIONS, pivotShowDataAsBaseFieldOptions, pivotShowDataAsDraftError, pivotShowDataAsMemberRevision, preparePivotShowDataAsDraft } from '../src/features/pivot/services/pivot-show-data-as-ui.ts'
import { showAsUICases, showAsUIIds as ids, showAsUIRule as rule, showAsUIValue, showAsUIWorkbook } from './fixtures/pivot-show-data-as-ui-fixture.mjs'

// Deferred model/UI contracts. This file and every parameterized case are unexecuted.
const clone = value => structuredClone(value)
const report = workbook => getPivotChartReport(workbook, workbook.charts[0])
const config = workbook => ({ fields: clone(report(workbook).fields), options: clone(report(workbook).options) })
const panel = workbook => createPivotPanelDraft(workbook, ids.host, 'C2:G10', ids.pivot)
const valueField = workbook => report(workbook).fields.find(field => field.id === ids.amount)
const close = (actual, expected) => assert.ok(typeof actual === 'number' && Math.abs(actual - expected) < 1e-10, `${actual} ~= ${expected}`)

test('default picker has the fixed fifteen entries in order and excludes the historical rule', () => {
  assert.deepEqual(PIVOT_SHOW_DATA_AS_OPTIONS.map(option => [option.value, option.label]), showAsUICases.map(([value, label]) => [value.type, label]))
  assert.equal(PIVOT_SHOW_DATA_AS_OPTIONS.length, 15)
  assert.equal(PIVOT_SHOW_DATA_AS_OPTIONS.some(option => option.value === 'percentOfParentGrandTotal'), false)
})

test('initial/current rules are defensive drafts, and null means an explicit normal reset', () => {
  const original = rule('differenceFrom', ids.product, 'B'), draft = createPivotShowDataAsDraft(original)
  draft.baseItem = 'C'; assert.equal(original.baseItem, 'B')
  assert.deepEqual(createPivotShowDataAsDraft(null), rule('normal'))
})

test.each(showAsUICases)('UI rule %j reaches real chart values through one full field configure', (value, _label, expected) => {
  const workbook = showAsUIWorkbook({ ownership: 'owned' }), before = clone(workbook.sheets), owner = report(workbook)
  const normalized = preparePivotShowDataAsDraft(owner, value, value.baseItem ? getPivotShowDataAsBaseItems(workbook, owner, value.baseFieldId) : undefined)
  const draft = setPivotChartDraftField(config(workbook), ids.amount, { showDataAs: normalized })
  assert.deepEqual(workbook.sheets, before); assert.equal(valueField(workbook).showDataAs, undefined)
  applyPivotChartEdit(workbook, workbook.charts[0], { type: 'setFieldsConfig', config: draft })
  close(showAsUIValue(workbook), expected); assert.deepEqual(workbook.sheets, before)
})

test('base fields use report-local IDs even when a canonical source ID collides with another local ID', () => {
  const workbook = showAsUIWorkbook(), pivot = report(workbook)
  pivot.fields.find(field => field.id === ids.region).id = 'product-data'
  const options = pivotShowDataAsBaseFieldOptions(pivot, 'differenceFrom')
  assert.deepEqual(options.map(option => option.value), ['product-data', ids.product, ids.year, ids.quarter])
  assert.equal(options.some(option => option.value === 'region-data'), false)
  const result = preparePivotShowDataAsDraft(pivot, rule('runningTotal', 'product-data'))
  assert.equal(result.baseFieldId, 'product-data')
})

test('specified-parent picker exposes only fields with actual children, on both axes', () => {
  const pivot = report(showAsUIWorkbook())
  assert.deepEqual(pivotShowDataAsBaseFieldOptions(pivot, 'percentOfParent').map(option => option.value), [ids.region, ids.year])
  assert.match(pivotShowDataAsDraftError(pivot, rule('percentOfParent', ids.product)), /子层级/)
})

test('repeated labels remain distinct picker options with local identities and readable disambiguation', () => {
  const pivot = report(showAsUIWorkbook())
  pivot.fields.find(field => field.id === ids.year).displayName = 'Region'
  const options = pivotShowDataAsBaseFieldOptions(pivot, 'runningTotal').filter(option => option.value === ids.region || option.value === ids.year)
  assert.equal(options.length, 2); assert.notEqual(options[0].label, options[1].label)
  assert.ok(options.every(option => option.label.includes(option.value)))
})

test.each(['normal', 'percentOfGrandTotal', 'percentOfRowTotal', 'percentOfColumnTotal', 'percentOfParentRowTotal', 'percentOfParentColumnTotal', 'index'])('changing to %s clears every stale base parameter', type => {
  const pivot = report(showAsUIWorkbook()), previous = rule('differenceFrom', ids.product, '', 'previous')
  assert.deepEqual(changePivotShowDataAsType(pivot, previous, type), rule(type))
})

test('changing parameterized modes retains only an eligible local field and clears the previous member', () => {
  const pivot = report(showAsUIWorkbook())
  assert.deepEqual(changePivotShowDataAsType(pivot, rule('differenceFrom', ids.product, 'B'), 'percentOf'), rule('percentOf', ids.product))
  assert.deepEqual(changePivotShowDataAsType(pivot, rule('differenceFrom', ids.product, 'B'), 'percentOfParent'), rule('percentOfParent'))
})

test('changing the base field clears concrete and relative parameters', () => {
  const pivot = report(showAsUIWorkbook())
  assert.deepEqual(changePivotShowDataAsBaseField(pivot, rule('differenceFrom', ids.product, '', 'next'), ids.quarter), rule('differenceFrom', ids.quarter))
  assert.throws(() => changePivotShowDataAsBaseField(pivot, rule('percentOfParent', ids.region), ids.product), /行或列/)
})

test.each(['item', 'previous', 'next'])('relative picker %s clears the old member and produces canonical parameters', type => {
  const previous = rule('differenceFrom', ids.product, 'B')
  assert.deepEqual(changePivotShowDataAsBaseItemType(previous, type), rule('differenceFrom', ids.product, '', type === 'item' ? undefined : type))
})

test('relative controls cannot be used for a fixed member or running/ranking mode', () => {
  for (const type of ['percentOf', 'runningTotal', 'rankAscending']) assert.throws(() => changePivotShowDataAsBaseItemType(rule(type, ids.product), 'next'), /相对/)
})

test('missing, stale and empty members remain invalid drafts rather than silently selecting a member', () => {
  const pivot = report(showAsUIWorkbook())
  assert.match(pivotShowDataAsDraftError(pivot, rule('percentOf', ids.product)), /基准项/)
  assert.match(pivotShowDataAsDraftError(pivot, rule('percentOf', ids.product, 'missing'), ['A', 'B']), /真实来源/)
  assert.match(pivotShowDataAsDraftError(undefined, null), /读取/)
})

test('grouped base members come from the existing engine and raw members are rejected', () => {
  const workbook = showAsUIWorkbook({ grouping: true }), pivot = report(workbook), draft = panel(workbook), projected = pivotPanelDraftReport(workbook, draft)
  const before = clone(workbook), items = getPivotShowDataAsBaseItems(workbook, projected, ids.product)
  assert.deepEqual(items, ['Mix', 'C', 'D'])
  assert.deepEqual(preparePivotShowDataAsDraft(projected, rule('percentOf', ids.product, 'Mix'), items), rule('percentOf', ids.product, 'Mix'))
  assert.throws(() => preparePivotShowDataAsDraft(projected, rule('percentOf', ids.product, 'A'), items), /真实来源/)
  assert.deepEqual(projected.fields.find(field => field.id === ids.product).manualGroups, pivot.fields.find(field => field.id === ids.product).manualGroups)
  assert.deepEqual(workbook, before)
})

test('compatibility data stays readable but the default picker cannot create that mode', () => {
  const pivot = report(showAsUIWorkbook({ compatibility: true })), imported = pivot.fields.find(field => field.id === ids.amount).showDataAs
  assert.deepEqual(preparePivotShowDataAsDraft(pivot, createPivotShowDataAsDraft(imported)), imported)
  assert.throws(() => changePivotShowDataAsType(pivot, rule('normal'), 'percentOfParentGrandTotal'), /固定版本/)
  assert.equal(preparePivotShowDataAsDraft(pivot, null), null)
})

test('chart mode and member edits change only a defensive local field configuration', () => {
  const workbook = showAsUIWorkbook(), before = clone(workbook), initial = config(workbook)
  const next = setPivotChartDraftField(initial, ids.amount, { showDataAs: rule('percentOf', ids.product, 'B') })
  next.fields.find(field => field.id === ids.amount).showDataAs.baseItem = 'C'
  assert.equal(initial.fields.find(field => field.id === ids.amount).showDataAs, undefined)
  assert.deepEqual(workbook, before)
  const projected = pivotChartDraftShowDataAsReport(report(workbook), next)
  projected.fields[0].name = 'local'; assert.equal(report(workbook).fields[0].name, 'Region')
})

test.each(['inherit', 'override', 'reset'])('Show As chart projection preserves defensive collapse state for %s', mode => {
  const workbook = showAsUIWorkbook(), pivot = report(workbook), draft = config(workbook)
  pivot.collapseInfo = { [ids.region]: { North: true } }
  if (mode === 'override') draft.collapseInfo = { [ids.product]: { A: true } }
  if (mode === 'reset') draft.collapseInfo = null
  const projected = pivotChartDraftShowDataAsReport(pivot, draft)
  if (mode === 'reset') assert.equal(projected.collapseInfo, undefined)
  else {
    assert.deepEqual(projected.collapseInfo, mode === 'inherit' ? pivot.collapseInfo : draft.collapseInfo)
    projected.collapseInfo[mode === 'inherit' ? ids.region : ids.product].local = true
    assert.equal(pivot.collapseInfo[ids.region].local, undefined)
    if (mode === 'override') assert.equal(draft.collapseInfo[ids.product].local, undefined)
  }
})

test('moving a value away and back clears the Show As rule without resurrecting old authority', () => {
  const workbook = showAsUIWorkbook(), initial = setPivotChartDraftField(config(workbook), ids.amount, { showDataAs: rule('runningTotal', ids.product) })
  const moved = movePivotChartDraftField(initial, ids.amount, 'filter', 0)
  const back = movePivotChartDraftField(moved, ids.amount, 'value', 0)
  assert.equal(moved.fields.find(field => field.id === ids.amount).showDataAs, undefined)
  assert.equal(back.fields.find(field => field.id === ids.amount).showDataAs, undefined)
})

test('removing the base field retains an invalid value draft for explicit correction', () => {
  const workbook = showAsUIWorkbook(), draft = setPivotChartDraftField(config(workbook), ids.amount, { showDataAs: rule('runningTotal', ids.product) })
  const removed = removePivotChartDraftFields(draft, [ids.product]), projected = pivotChartDraftShowDataAsReport(report(workbook), removed)
  assert.match(pivotShowDataAsDraftError(projected, projected.fields.find(field => field.id === ids.amount).showDataAs), /行或列/)
  assert.equal(valueField(workbook).showDataAs, undefined)
})

test('null and normal reset the chart field, while non-value fields refuse Show As', () => {
  const workbook = showAsUIWorkbook(), initial = setPivotChartDraftField(config(workbook), ids.amount, { showDataAs: rule('percentOfGrandTotal') })
  for (const reset of [null, rule('normal')]) assert.equal(setPivotChartDraftField(initial, ids.amount, { showDataAs: reset }).fields.find(field => field.id === ids.amount).showDataAs, undefined)
  assert.throws(() => setPivotChartDraftField(initial, ids.region, { showDataAs: rule('index') }), /值字段/)
  assert.throws(() => setPivotChartDraftField(initial, ids.amount, { showDataAs: undefined }), /显式/)
})

test('ordinary panel reload retains custom local IDs, repeated values, grouping and source-cache identities', () => {
  const workbook = showAsUIWorkbook({ repeated: true, grouping: true }), before = clone(workbook), draft = panel(workbook)
  assert.deepEqual(new Set(draft.fields.map(field => field.fieldId)), new Set(report(workbook).fields.map(field => field.id)))
  assert.equal(draft.fields.filter(field => field.sourceFieldId === 'amount-data').length, 2)
  const input = preparePivotPanelInput(workbook, draft, ids.pivot)
  assert.ok(input.fields.every(field => Number.isSafeInteger(field.sourceColumn) && field.sourceFieldId))
  assert.equal(input.source.pivotCacheId, 'show-cache'); assert.deepEqual(workbook, before)
})

test('ordinary configuration writes a real rule and numeric output, retaining source cells and grouped field settings', () => {
  const workbook = showAsUIWorkbook(), draft = panel(workbook), sourceCells = clone(workbook.sheets[0].cells)
  draft.fields.find(field => field.fieldId === ids.amount).showDataAs = rule('percentOfGrandTotal')
  configurePivotFromPanel(workbook, preparePivotPanelInput(workbook, draft, ids.pivot))
  assert.deepEqual(valueField(workbook).showDataAs, rule('percentOfGrandTotal')); close(showAsUIValue(workbook), .05)
  assert.deepEqual(workbook.sheets[0].cells, sourceCells); assert.equal(report(workbook).pivotCacheId, 'show-cache')
})

test('normalized ordinary configure no-op preserves report identity, revision and output cells', () => {
  const workbook = showAsUIWorkbook(), pivot = report(workbook), before = clone(workbook), draft = panel(workbook)
  const returned = configurePivotFromPanel(workbook, preparePivotPanelInput(workbook, draft, ids.pivot))
  assert.equal(returned, pivot); assert.deepEqual(workbook, before)
})

test('ordinary reset removes the rule and a second identical configure stays a no-op', () => {
  const workbook = showAsUIWorkbook(), draft = panel(workbook)
  draft.fields.find(field => field.fieldId === ids.amount).showDataAs = rule('percentOfGrandTotal')
  configurePivotFromPanel(workbook, preparePivotPanelInput(workbook, draft, ids.pivot))
  const reset = panel(workbook); reset.fields.find(field => field.fieldId === ids.amount).showDataAs = null
  configurePivotFromPanel(workbook, preparePivotPanelInput(workbook, reset, ids.pivot)); close(showAsUIValue(workbook), 10)
  assert.equal(valueField(workbook).showDataAs, undefined)
  const before = clone(workbook); configurePivotFromPanel(workbook, preparePivotPanelInput(workbook, panel(workbook), ids.pivot)); assert.deepEqual(workbook, before)
})

test('invalid ordinary member configuration is rejected before clearing prior output or changing fields', () => {
  const workbook = showAsUIWorkbook({ grouping: true }), draft = panel(workbook), before = clone(workbook)
  draft.fields.find(field => field.fieldId === ids.amount).showDataAs = rule('percentOf', ids.product, 'A')
  const input = preparePivotPanelInput(workbook, draft, ids.pivot)
  assert.throws(() => configurePivotFromPanel(workbook, input), /真实来源/); assert.deepEqual(workbook, before)
})

test('source-column mismatches and duplicate local instances are refused before mutation', () => {
  const workbook = showAsUIWorkbook(), before = clone(workbook), input = preparePivotPanelInput(workbook, panel(workbook), ids.pivot)
  input.fields[0].sourceFieldId = 'amount-data'; assert.throws(() => configurePivotFromPanel(workbook, input), /不匹配/)
  const duplicate = preparePivotPanelInput(workbook, panel(workbook), ids.pivot); duplicate.fields.push(clone(duplicate.fields[0]))
  assert.throws(() => configurePivotFromPanel(workbook, duplicate), /实例/); assert.deepEqual(workbook, before)
})

test('JSON reorder and snapshot reopen preserve normalized local base identity and a stable configure', () => {
  const workbook = showAsUIWorkbook(), draft = panel(workbook)
  draft.fields.find(field => field.fieldId === ids.amount).showDataAs = { baseItem: 'B', baseFieldId: ids.product, type: 'percentOf' }
  configurePivotFromPanel(workbook, preparePivotPanelInput(workbook, draft, ids.pivot))
  const reopened = workbookFromInput(JSON.parse(JSON.stringify(workbook))), before = clone(reopened), loaded = panel(reopened)
  assert.equal(loaded.fields.find(field => field.fieldId === ids.amount).showDataAs.baseFieldId, ids.product)
  configurePivotFromPanel(reopened, preparePivotPanelInput(reopened, loaded, ids.pivot)); assert.deepEqual(reopened, before)
})

test('member revision changes with grouping/field structure and stays stable for value-mode and display-name edits', () => {
  const workbook = showAsUIWorkbook(), pivot = report(workbook), before = pivotShowDataAsMemberRevision(pivot)
  valueField(workbook).showDataAs = rule('percentOfGrandTotal'); valueField(workbook).displayName = 'Revenue'
  assert.equal(pivotShowDataAsMemberRevision(pivot), before)
  pivot.fields.find(field => field.id === ids.product).manualGroups = [{ name: 'Mix', members: ['A', 'B'] }]
  assert.notEqual(pivotShowDataAsMemberRevision(pivot), before)
})
