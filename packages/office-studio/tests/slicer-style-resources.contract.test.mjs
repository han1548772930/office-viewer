import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { getSlicerStyle, getSlicerStyleType, getDefaultSlicerStyleName, hasSlicerStyle, listSlicerStyleNames,
  projectSlicerItemStyle, projectSlicerStyleFormatting, projectTimelineStyleFormatting } from '../src/features/slicer/slicer-style-catalog.ts'
import { importSlicerStyleResources, normalizeSlicerStyleFormatting, normalizeSlicerStylesDefinition,
  normalizeTimelineStylesDefinition, projectSlicerStyles, projectTimelineStyles, readSlicerStyleResources,
  removeSlicerStyleDefinition, setDefaultSlicerStyle, setSlicerStyleResources, upsertSlicerStyleDefinition } from '../src/features/slicer/slicer-style-resources.ts'
import { customSlicerStyle, customTimelineStyle, slicerStyleStates, slicerStyleWorkbook, styleFormatting, timelineStyleElements } from './fixtures/slicer-style-fixture.mjs'

// Real source/catalog/import and author-mutation drafts. Every declaration and planned case remains unexecuted.
const clone = value => structuredClone(value), types = ['slicer', 'timeline']
const key = type => type === 'slicer' ? 'slicerStyles' : 'timelineStyles'
const defaultKey = type => type === 'slicer' ? 'defaultSlicerStyle' : 'defaultTimelineStyle'
const elementKey = type => type === 'slicer' ? 'slicerStyleElements' : 'timelineStyleElements'
const definition = type => type === 'slicer' ? customSlicerStyle() : customTimelineStyle()
const customName = type => type === 'slicer' ? 'Regional members' : 'Calendar bands'
const unusedName = type => type === 'slicer' ? 'Unused slicer' : 'Unused timeline'
const normalize = type => type === 'slicer' ? normalizeSlicerStylesDefinition : normalizeTimelineStylesDefinition
const project = type => type === 'slicer' ? projectSlicerStyles : projectTimelineStyles
const authority = workbook => clone({ pivotTables: workbook.pivotTables, pivotCaches: workbook.pivotCaches,
  charts: workbook.charts, tables: workbook.tables, slicerCaches: workbook.slicerCaches,
  selections: workbook.slicers.map(control => ({ id: control.id, selectedItems: control.selectedItems, timelineRange: control.timelineRange })) })

for (const type of types) test(`${type} lists exact built-ins then document styles with stable references and defensive results`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook)
  const prefix = type === 'slicer' ? 'Slicer' : 'Timeline'
  const expected = [...Array.from({ length: 6 }, (_, index) => `${prefix}StyleLight${index + 1}`),
    ...(type === 'slicer' ? ['SlicerStyleOther1', 'SlicerStyleOther2'] : []),
    ...Array.from({ length: 6 }, (_, index) => `${prefix}StyleDark${index + 1}`), customName(type), unusedName(type)]
  assert.deepEqual(listSlicerStyleNames(workbook, type), expected)
  for (const name of expected) assert.equal(hasSlicerStyle(workbook, type, name), true)
  assert.equal(hasSlicerStyle(workbook, type, type === 'slicer' ? 'TimelineStyleLight1' : 'SlicerStyleLight1'), false)
  const names = listSlicerStyleNames(workbook, type); names.pop(); assert.deepEqual(listSlicerStyleNames(workbook, type), expected)
  const resolved = getSlicerStyle(workbook, type, customName(type)); resolved.definition.wholeStyle.bg.rgb = '#ffffff'
  assert.deepEqual(workbook, before)
})

for (const [selected, hasData, hovered, expected] of [
  [false, true, false, 'unselectedItemWithData'], [true, true, false, 'selectedItemWithData'],
  [false, false, false, 'unselectedItemWithNoData'], [true, false, false, 'selectedItemWithNoData'],
  [false, true, true, 'hoveredUnselectedItemWithData'], [true, true, true, 'hoveredSelectedItemWithData'],
  [false, false, true, 'hoveredUnselectedItemWithNoData'], [true, false, true, 'hoveredSelectedItemWithNoData'],
]) test(`typed ${expected} projects the authored state without changing member selection or availability`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook), state = { selected, hasData, hovered }
  assert.equal(getSlicerStyleType(state), expected)
  const item = projectSlicerItemStyle(workbook, undefined, state)
  assert.equal(item.type, expected)
  assert.deepEqual(item.formatting, workbook.slicerStyles.styles[0].slicerStyleElements.find(element => element.type === expected).formatting)
  item.formatting.bg.th = 11; assert.deepEqual(workbook, before); assert.deepEqual(state, { selected, hasData, hovered })
})

for (const type of types) test(`${type} normalizes all published color, overline, border and alignment fields including null resets`, () => {
  const workbook = slicerStyleWorkbook(), resource = normalize(type)(workbook[key(type)])
  assert.deepEqual(resource.styles[0].wholeStyle, styleFormatting)
  assert.deepEqual(normalizeSlicerStyleFormatting({ bg: null, cl: null, bd: null, ht: null, vt: null, bl: 0 }),
    { bg: null, cl: null, bd: null, ht: null, vt: null, bl: 0 })
  const projected = type === 'slicer' ? projectSlicerStyleFormatting(workbook) : projectTimelineStyleFormatting(workbook)
  const values = type === 'slicer' ? projected.items : projected.elements
  assert.deepEqual(Object.keys(values), type === 'slicer' ? slicerStyleStates : timelineStyleElements)
  assert.deepEqual(values, Object.fromEntries(resource.styles[0][elementKey(type)].map(element => [element.type, element.formatting])))
  projected.wholeStyle.ol.cl.rgb = '#000000'; assert.deepEqual(workbook[key(type)].styles[0].wholeStyle, styleFormatting)
})

for (const type of types) test(`${type} resolves omitted/default/explicit/unknown references without materializing author style`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook)
  assert.equal(getDefaultSlicerStyleName(workbook, type), customName(type))
  const inherited = getSlicerStyle(workbook, type); assert.equal(inherited.name, customName(type)); assert.equal(inherited.usesDefault, true)
  const explicit = getSlicerStyle(workbook, type, `${type === 'slicer' ? 'Slicer' : 'Timeline'}StyleDark3`)
  assert.equal(explicit.builtIn, true); assert.equal(explicit.usesDefault, false); assert.equal(explicit.definition, undefined)
  const unresolved = getSlicerStyle(workbook, type, 'Missing author reference')
  assert.equal(unresolved.name, customName(type)); assert.equal(unresolved.unresolvedName, 'Missing author reference')
  delete workbook[key(type)]
  assert.equal(getSlicerStyle(workbook, type).name, `${type === 'slicer' ? 'Slicer' : 'Timeline'}StyleLight1`)
  assert.deepEqual(workbook.slicers, before.slicers)
})

const invalidFormatting = [null, [], { bg: '#fff' }, { cl: { rgb: 1 } }, { bg: { th: 12 } }, { bg: { th: -1 } }, { bg: { th: 0.5 } },
  { ol: null }, { ol: { s: true } }, { ol: { s: 1, c: 2 } }, { ol: { s: 1, t: 20 } }, { ol: { s: 1, cl: null } },
  { bd: { t: { s: 14, cl: {} } } }, { bd: { r: { s: 1, cl: null } } }, { bd: [] }, { bl: true }, { bl: null }, { ht: 7 }, { vt: 4 }, { ht: 0.5 }]
for (const formatting of invalidFormatting) test(`invalid formatting ${JSON.stringify(formatting)} rejects before resource edits`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook)
  assert.throws(() => normalizeSlicerStyleFormatting(formatting), error => error.code === 'SLICER_INVALID_STYLE_RESOURCE')
  const resource = clone(workbook.slicerStyles); resource.styles[0].wholeStyle = formatting
  assert.equal(projectSlicerStyles(resource), undefined)
  assert.throws(() => setSlicerStyleResources(workbook, 'slicer', resource), /样式资源无效/); assert.deepEqual(workbook, before)
})

const invalidResourceEdits = [
  resource => { resource[defaultKey(resource.kind)] = 'Unknown default' },
  resource => { resource.styles = null },
  resource => { resource.styles[0].name = '' },
  resource => { resource.styles[0].name = 'x'.repeat(256) },
  resource => { resource.styles[1].name = resource.styles[0].name.toUpperCase() },
  resource => { resource.styles[0].name = `${resource.kind === 'slicer' ? 'Slicer' : 'Timeline'}StyleLight1` },
  resource => { resource.styles[0][elementKey(resource.kind)] = null },
  resource => { resource.styles[0][elementKey(resource.kind)][0].type = 'Unknown state' },
  resource => { resource.styles[0][elementKey(resource.kind)].push(clone(resource.styles[0][elementKey(resource.kind)][0])) },
  resource => { resource.styles[0].headerRowStyle = null },
  resource => { resource.styles[0][elementKey(resource.kind)][0].formatting = null },
]
for (const type of types) for (const [index, mutate] of invalidResourceEdits.entries()) test(`${type} invalid definition ${index} rejects duplicate names/states/defaults atomically`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook), resource = { ...clone(workbook[key(type)]), kind: type }
  mutate(resource); delete resource.kind
  assert.equal(project(type)(resource), undefined)
  assert.throws(() => setSlicerStyleResources(workbook, type, resource), error => error.code === 'SLICER_INVALID_STYLE_RESOURCE')
  assert.deepEqual(workbook, before)
})

for (const type of types) test(`${type} import retains original extensions with defensive known-field projection and stable provenance`, () => {
  const workbook = slicerStyleWorkbook(), input = clone(workbook)
  input[key(type)].styles[0].wholeStyle.futureFormatting = { rawOfficeAttribute: 'keep me' }
  input[key(type)].futureResource = ['unresolved source field']
  const before = clone(input), imported = importSlicerStyleResources(input)
  assert.deepEqual(imported.metadata.slicerStyleResourcesSource[key(type)], before[key(type)])
  assert.equal(imported[key(type)].styles[0].wholeStyle.futureFormatting, undefined)
  assert.ok(imported.metadata.slicerStyleResourceDiagnostics.some(diagnostic => diagnostic.code === 'SLICER_STYLE_SOURCE_EXTENSIONS' && diagnostic.kind === type))
  assert.deepEqual(imported.metadata.preservedField, before.metadata.preservedField)
  const reopened = importSlicerStyleResources({ ...before, ...clone(imported) }); assert.deepEqual(reopened, imported)
  imported.metadata.slicerStyleResourcesSource[key(type)].styles[0].name = 'mutated source'; assert.deepEqual(input, before)
})

for (const type of types) test(`${type} invalid imported family preserves raw source while the other family remains usable`, () => {
  const input = slicerStyleWorkbook(), before = clone(input), other = type === 'slicer' ? 'timeline' : 'slicer'
  input[key(type)].styles[0][elementKey(type)][0].type = 'Unknown preserved state'
  const rejected = clone(input[key(type)]), imported = importSlicerStyleResources(input)
  assert.equal(imported[key(type)], undefined); assert.deepEqual(imported[key(other)], before[key(other)])
  assert.deepEqual(imported.metadata.slicerStyleResourcesSource[key(type)], rejected)
  assert.ok(imported.metadata.slicerStyleResourceDiagnostics.some(diagnostic => diagnostic.code === 'SLICER_INVALID_STYLE_RESOURCE' && diagnostic.kind === type))
  assert.equal(getSlicerStyle(imported, type).builtIn, true)
  assert.deepEqual(input.slicers, before.slicers)
})

for (const type of types) test(`${type} unresolved imported control references are diagnosed and retained without changing filters`, () => {
  const input = slicerStyleWorkbook(), control = input.slicers.find(control => control.type === type)
  control.style = 'Office custom style not projected'; const before = clone(input), imported = importSlicerStyleResources(input)
  assert.ok(imported.metadata.slicerStyleResourceDiagnostics.some(diagnostic => diagnostic.code === 'SLICER_UNRESOLVED_STYLE_REFERENCE' && diagnostic.kind === type))
  assert.equal(getSlicerStyle(imported, type, control.style).unresolvedName, control.style)
  assert.deepEqual(input, before)
})

test('published resource control arrays diagnose separate native Slicer and Timeline references', () => {
  const imported = importSlicerStyleResources({ slicers: [{ style: 'Native missing slicer' }], timelines: [{ style: 'Native missing timeline' }] })
  assert.deepEqual(imported.metadata.slicerStyleResourceDiagnostics.map(diagnostic => [diagnostic.kind, diagnostic.path]),
    [['slicer', 'slicers[0].style'], ['timeline', 'timelines[0].style']])
})

for (const shape of ['domain model', 'raw snapshot']) test(`${shape} import and JSON reopening retain document styles, provenance and sparse author references`, () => {
  const input = slicerStyleWorkbook(); input.slicerStyles.styles[0].wholeStyle.officeAttribute = { unrendered: 'retained' }
  if (shape === 'raw snapshot') delete input.kind
  const workbook = workbookFromInput(clone(input)), reopened = workbookFromInput(JSON.parse(JSON.stringify(workbook)))
  assert.deepEqual(workbook.slicerStyles, normalizeSlicerStylesDefinition(input.slicerStyles))
  assert.deepEqual(workbook.timelineStyles, normalizeTimelineStylesDefinition(input.timelineStyles))
  assert.deepEqual(reopened.slicerStyles, workbook.slicerStyles); assert.deepEqual(reopened.timelineStyles, workbook.timelineStyles)
  assert.deepEqual(reopened.metadata.slicerStyleResourcesSource, workbook.metadata.slicerStyleResourcesSource)
  assert.deepEqual(reopened.metadata.slicerStyleResourceDiagnostics, workbook.metadata.slicerStyleResourceDiagnostics)
  assert.equal(workbook.slicers.find(control => control.type === 'slicer').style, undefined)
  assert.equal(reopened.slicers.find(control => control.type === 'timeline').style, undefined)
})

for (const kind of ['pivot', 'table']) test(`${kind} default-style author changes invalidate only inherited controls and preserve every filter/cache authority`, () => {
  const workbook = slicerStyleWorkbook(kind), before = clone(workbook), filtered = authority(workbook)
  assert.equal(setDefaultSlicerStyle(workbook, 'slicer', 'SlicerStyleDark4'), true)
  assert.equal(workbook.slicers[0].revision, before.slicers[0].revision + 1)
  assert.equal(workbook.slicers[1].revision, before.slicers[1].revision)
  assert.equal(workbook.slicers[2].revision, before.slicers[2].revision)
  assert.deepEqual(authority(workbook), filtered)
  assert.equal(workbook.slicers[0].style, undefined); assert.equal(getSlicerStyle(workbook, 'slicer').name, 'SlicerStyleDark4')
  const after = clone(workbook); assert.equal(setDefaultSlicerStyle(workbook, 'slicer', 'SlicerStyleDark4'), false); assert.deepEqual(workbook, after)
})

for (const type of types) test(`${type} custom style upsert updates effective views once and preserves defensive resource state`, () => {
  const workbook = slicerStyleWorkbook(), filtered = authority(workbook), before = clone(workbook.slicers)
  const style = definition(type); style.wholeStyle.bg.rgb = '#807060'
  assert.equal(upsertSlicerStyleDefinition(workbook, type, style), true)
  for (const control of workbook.slicers) assert.equal(control.revision, before.find(prior => prior.id === control.id).revision
    + (control.type === type && control.style === undefined ? 1 : 0))
  assert.deepEqual(authority(workbook), filtered)
  const after = clone(workbook); assert.equal(upsertSlicerStyleDefinition(workbook, type, clone(style)), false); assert.deepEqual(workbook, after)
  const snapshot = readSlicerStyleResources(workbook, type); snapshot.styles[0].wholeStyle.bg.rgb = '#ffffff'
  style.wholeStyle.bg.rgb = '#000000'; assert.equal(workbook[key(type)].styles[0].wholeStyle.bg.rgb, '#807060')
})

for (const type of types) test(`${type} removals refuse default/in-use/built-in references and remove unused resources atomically`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook)
  assert.throws(() => removeSlicerStyleDefinition(workbook, type, customName(type)), /默认.*样式不存在/); assert.deepEqual(workbook, before)
  assert.throws(() => removeSlicerStyleDefinition(workbook, type, `${type === 'slicer' ? 'Slicer' : 'Timeline'}StyleLight1`), /内置样式不能删除/)
  assert.equal(removeSlicerStyleDefinition(workbook, type, unusedName(type)), true)
  const control = workbook.slicers.find(control => control.type === type); control.style = customName(type)
  setDefaultSlicerStyle(workbook, type, `${type === 'slicer' ? 'Slicer' : 'Timeline'}StyleLight1`)
  const referenced = clone(workbook)
  assert.throws(() => removeSlicerStyleDefinition(workbook, type, customName(type)), /仍被控件引用/); assert.deepEqual(workbook, referenced)
  assert.equal(removeSlicerStyleDefinition(workbook, type, 'Never authored'), false); assert.deepEqual(workbook, referenced)
})

for (const type of types) test(`${type} clearing unreferenced resource defaults leaves optional control state and no-op revisions intact`, () => {
  const workbook = slicerStyleWorkbook(), filtered = authority(workbook)
  assert.equal(setSlicerStyleResources(workbook, type, undefined), true)
  assert.equal(workbook[key(type)], undefined); assert.equal(getDefaultSlicerStyleName(workbook, type), `${type === 'slicer' ? 'Slicer' : 'Timeline'}StyleLight1`)
  assert.deepEqual(authority(workbook), filtered)
  const after = clone(workbook); assert.equal(setSlicerStyleResources(workbook, type, undefined), false); assert.deepEqual(workbook, after)
})

for (const type of types) test(`${type} author edits reject unsupported extensions while import retains their original data`, () => {
  const workbook = slicerStyleWorkbook(), before = clone(workbook), style = definition(type); style.wholeStyle.officeExtension = 'preserve by import'
  assert.throws(() => upsertSlicerStyleDefinition(workbook, type, style), /尚未支持的扩展字段/); assert.deepEqual(workbook, before)
  const resource = clone(workbook[key(type)]); resource.styles[0] = style
  assert.throws(() => setSlicerStyleResources(workbook, type, resource), /尚未支持的扩展字段/); assert.deepEqual(workbook, before)
  const imported = importSlicerStyleResources({ [key(type)]: resource })
  assert.equal(imported.metadata.slicerStyleResourcesSource[key(type)].styles[0].wholeStyle.officeExtension, 'preserve by import')
})

test('style resources remain isolated between workbooks and preserve metadata when their defaults change', () => {
  const first = slicerStyleWorkbook(), second = slicerStyleWorkbook(), before = clone(second), metadata = clone(first.metadata)
  setDefaultSlicerStyle(first, 'slicer', 'SlicerStyleOther2'); upsertSlicerStyleDefinition(first, 'timeline', customTimelineStyle('Additional calendar'))
  assert.deepEqual(second, before); assert.deepEqual(first.metadata, metadata)
  assert.equal(hasSlicerStyle(first, 'timeline', 'Additional calendar'), true); assert.equal(hasSlicerStyle(second, 'timeline', 'Additional calendar'), false)
})
