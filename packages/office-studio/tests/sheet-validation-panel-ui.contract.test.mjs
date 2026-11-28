import assert from 'node:assert/strict'
import { test } from 'vitest'
import { cloneModel } from '../src/domain/serialization.ts'
import { changeValidationType, moveValidationListItem, validationListItemsFromCsv, validationPanelSettings,
  validationRuleFromPanel } from '../src/features/sheet-validation/validation-panel-state.ts'

// Implementation-phase behavioral drafts replace the previous source-string UI checks. They have not been run.
const ranges = [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }]

test('CSV bulk edits keep existing item IDs/colors by case-folded label and never split separators or embedded newlines', () => {
  const previous = [{ id: 'red', label: 'Red', color: '#ff0000' }, { id: 'green', label: 'Green, dark', color: '#00ff00' }], before = cloneModel(previous)
  const items = validationListItemsFromCsv('"green, DARK",RED,red,"Line\nbreak"', previous, () => 'new')
  assert.deepEqual(items, [{ id: 'green', label: 'green, DARK', color: '#00ff00' }, { id: 'red', label: 'RED', color: '#ff0000' }, { id: 'new', label: 'Line\nbreak', color: '' }])
  assert.deepEqual(previous, before)
})

test('item ordering moves colors together with labels and saves one stable CSV/color mapping', () => {
  const previous = { id: 'colored', type: 'listMultiple', ranges, formula1: '"Red,"Green, dark",It""s blue"', formula2: '#ff0000,#00ff00,#0000ff' }
  const settings = validationPanelSettings(previous), original = cloneModel(settings.listItems), green = settings.listItems[1]
  settings.listItems = moveValidationListItem(settings.listItems, green.id, -1)
  const next = validationRuleFromPanel(settings, ranges, {}, previous)
  assert.equal(next.id, 'colored'); assert.equal(next.formula1, '""Green, dark",Red,It""s blue"'); assert.equal(next.formula2, '#00ff00,#ff0000,#0000ff')
  assert.deepEqual(original.map(item => item.label), ['Red', 'Green, dark', 'It"s blue'])
  assert.deepEqual(moveValidationListItem(settings.listItems, green.id, -1), settings.listItems)
})

test('an unedited list preserves source CSV bytes, empty color slots and trailing source positions across scope and list-type edits', () => {
  const previous = { id: 'raw', type: 'list', ranges, formula1: '" Red, Green "', formula2: '#ff0000,,,#123456', metadata: { retained: true } }
  const settings = changeValidationType(validationPanelSettings(previous), 'listMultiple'), next = validationRuleFromPanel(settings,
    [{ startRow: 4, endRow: 6, startCol: 1, endCol: 1 }], {}, previous)
  assert.equal(next.formula1, previous.formula1); assert.equal(next.formula2, previous.formula2); assert.equal(next.id, previous.id); assert.deepEqual(next.metadata, previous.metadata)
})

test('reference color edits retain the reference formula and source anchor without introducing a default numeric operand', () => {
  const previous = { id: 'reference', type: 'list', ranges, formula1: '=$D$1:$D$2', formula2: '#ff0000,', formulaAnchor: { row: 0, col: 0 } }
  const settings = validationPanelSettings(previous); settings.formula2 = '#ff0000,#00ff00'
  const next = validationRuleFromPanel(settings, ranges, {}, previous)
  assert.equal(next.formula1, previous.formula1); assert.equal(next.formula2, '#ff0000,#00ff00'); assert.deepEqual(next.formulaAnchor, previous.formulaAnchor)
  const number = { id: 'number', type: 'whole', operator: 'between', formula1: '1', formula2: '10', ranges }, converted = changeValidationType(validationPanelSettings(number), 'list')
  converted.listItems = [{ id: 'new', label: '1', color: '' }, { id: 'other', label: '=1+1', color: '' }]
  const list = validationRuleFromPanel(converted, ranges, {}, number); assert.equal(list.formula1, '"1,=1+1"'); assert.equal(list.formula2, undefined)
})

test('existing unsupported color strings survive unchanged, edited invalid colors reject, and an empty list cannot be saved', () => {
  const previous = { id: 'legacy', type: 'list', ranges, formula1: '"Red"', formula2: 'theme.legacy' }, settings = validationPanelSettings(previous)
  assert.equal(validationRuleFromPanel(settings, ranges, {}, previous).formula2, 'theme.legacy')
  settings.listItems[0].color = 'edited.invalid'; assert.throws(() => validationRuleFromPanel(settings, ranges, {}, previous), /颜色/)
  settings.listItems = []; assert.throws(() => validationRuleFromPanel(settings, ranges, {}, previous), /选项/)
})
