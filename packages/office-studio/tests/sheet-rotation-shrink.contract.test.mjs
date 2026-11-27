import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { createCanvas2DDriver } from '../src/rendering/canvas2d.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { ribbonInvocation } from '../src/platform/ribbon/ribbon-schema.ts'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { setSheetTextFit } from '../src/features/sheet-workbench/sheet-toolbar-format.ts'

const state = {
  canUndo: false, canRedo: false, hasSheet: true, canMerge: false, canUnmerge: false,
  canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true,
  canSort: true, fontFamily: 'Aptos', fontSize: 11, numberFormat: 'general',
  gridlines: true, frozen: false, zoom: 100,
}

class MockOffscreenCanvas {
  getContext() {
    return {
      font: '',
      measureText(text) {
        const size = Number(/([\d.]+)px/.exec(this.font)?.[1] ?? 12)
        return { width: text.length * size * 0.75 }
      },
    }
  }
}
globalThis.OffscreenCanvas = MockOffscreenCanvas

function itemsFor(format = {}) {
  return createSheetRibbon({ ...state, ...format })
    .flatMap(tab => tab.groups.flatMap(group => group.items.flatMap(item => [item, ...(item.children ?? [])])))
}

test('text direction entries dispatch supported rotation degrees and reflect mixed selection state', () => {
  const items = itemsFor({ rotation: null })
  const rotations = items.filter(item => item.id.startsWith('sheet.menu.text-rotation.'))
  assert.deepEqual(rotations.map(item => item.params), [0, -45, 45, 'v', -90, 90])
  assert.deepEqual(ribbonInvocation(rotations.find(item => item.params === 'v')), {
    commandId: 'sheet.command.set-text-rotation', params: 'v',
  })
  assert.equal(rotations.every(item => item.activated === false), true)
  assert.equal(items.find(item => item.id === 'sheet.menu.text-rotation').disabled, false)
})

test('shrink-to-fit toggles the selected value through the workbook formatting command', () => {
  const disabled = itemsFor({ shrinkToFit: false }).find(item => item.id === 'sheet.command.set-shrink-to-fit')
  const enabled = itemsFor({ shrinkToFit: true }).find(item => item.id === 'sheet.command.set-shrink-to-fit')
  assert.deepEqual(ribbonInvocation(disabled), { commandId: 'sheet.command.set-shrink-to-fit', params: undefined })
  assert.equal(disabled.activated, false)
  assert.deepEqual(ribbonInvocation(enabled), { commandId: 'sheet.command.set-shrink-to-fit', params: undefined })
  assert.equal(enabled.activated, true)
  const wrapped = itemsFor({ shrinkToFit: true, wrap: true }).find(item => item.id === 'sheet.command.set-shrink-to-fit')
  assert.equal(wrapped.disabled, true)
  assert.equal(wrapped.activated, false)
})

test('layout input capability disables source roots and choices while preserving primary state and font editing', () => {
  for (const context of [
    { layoutEditable: undefined, disabled: false },
    { layoutEditable: true, disabled: false },
    { layoutEditable: false, disabled: true },
    { layoutEditable: true, fontEditable: false, disabled: true },
    { layoutEditable: true, hasSheet: false, disabled: true },
  ]) {
    const tabs = createSheetRibbon({ ...state, ...context, canMerge: true, canUnmerge: true, mergeEditable: true,
      align: 'center', vertical: 'top', wrap: false, wrapStrategy: 'clip', rotation: 45, shrinkToFit: true })
    const layout = tabs.find(tab => tab.id === 'home').groups.find(group => group.id === 'layout').items
    assert.deepEqual(layout.map(item => item.id), ['sheet.menu.horizontal-align', 'sheet.menu.vertical-align', 'sheet.menu.text-wrap',
      'sheet.command.set-shrink-to-fit', 'sheet.menu.text-rotation', 'sheet.menu.merge'])
    assert.deepEqual(layout.map(item => item.gridLayout), [{ row: 1, column: 1 }, { row: 2, column: 1 }, { row: 1, column: 2 },
      { row: 2, column: 2 }, { row: 1, column: 3 }, { row: 2, column: 3 }])
    const choices = layout.flatMap(item => [item, ...(item.children ?? [])])
    assert.equal(choices.length, 25); assert.ok(choices.every(item => item.disabled === context.disabled))
    assert.ok(choices.filter(item => item.type !== 'menu').every(item => Boolean(ribbonInvocation(item)) === !context.disabled))
    assert.equal(layout[0].children.find(item => item.params === 'center').activated, true)
    assert.equal(layout[1].children.find(item => item.params === 'top').activated, true)
    assert.equal(layout[2].children.find(item => item.params === 'clip').activated, true)
    assert.equal(layout[3].activated, true)
    assert.equal(layout[4].children.find(item => item.params === 45).activated, true)
    assert.equal(tabs.find(tab => tab.id === 'home').groups.find(group => group.id === 'font').items[4].disabled,
      context.hasSheet === false || context.fontEditable === false)
  }
  const drawing = itemsFor({ drawingFocused: true, layoutEditable: false })
  assert.ok(['sheet.menu.horizontal-align', 'sheet.menu.vertical-align', 'sheet.menu.text-wrap', 'sheet.menu.text-rotation', 'sheet.menu.merge']
    .every(id => drawing.find(item => item.id === id).hidden === true))
  assert.equal(drawing.find(item => item.id === 'sheet.command.set-shrink-to-fit').hidden, undefined)
})

test('wrap and shrink-to-fit clear each other when enabled', () => {
  const wrapped = setSheetTextFit({ value: 'Text', shrinkToFit: true }, 'wrap', true)
  assert.equal(wrapped.wrap, true)
  assert.equal(wrapped.shrinkToFit, false)
  const fitted = setSheetTextFit(wrapped, 'shrinkToFit', true)
  assert.equal(fitted.wrap, false)
  assert.equal(fitted.shrinkToFit, true)
})

test('sheet scene carries rotation and uniformly scaled fit while preserving selected-cell paint', () => {
  const geometry = new SpreadsheetGeometry({
    id: 's', name: 'Fit', cells: {
      A1: { value: 'Long worksheet heading', rotation: 45, shrinkToFit: true, underline: true },
      B1: { value: 'Short', shrinkToFit: true },
      C1: { value: '中🙂A', rotationVertical: true },
    },
    rowCount: 3, columnCount: 3, columnWidths: { 0: 64, 1: 128 },
  })
  const scene = createSpreadsheetScene(geometry, {
    width: 360, height: 180, scrollLeft: 0, scrollTop: 0, zoom: 1,
  }, { activeCell: 'C3', selectedCells: ['A1'] })
  const longText = scene.commands.find(command => command.kind === 'text' && command.text === 'Long worksheet heading')
  const shortText = scene.commands.find(command => command.kind === 'text' && command.text === 'Short')

  assert.equal(longText.rotation, 45)
  assert.equal(longText.maxWidth, undefined)
  const fittedSize = Number(/([\d.]+)px/.exec(longText.font)?.[1])
  assert.ok(fittedSize > 0 && fittedSize < 11 * 4 / 3)
  const radians = Math.PI / 4
  const width = 'Long worksheet heading'.length * fittedSize * .75, height = fittedSize * 1.2
  assert.ok(width * Math.cos(radians) + height * Math.sin(radians) <= 52 + 1e-8)
  assert.ok(width * Math.sin(radians) + height * Math.cos(radians) <= 20 + 1e-8)
  assert.equal(shortText.maxWidth, undefined)
  assert.deepEqual(scene.commands.filter(command => command.kind === 'text' && command.clip?.y > 26 && ['中', '🙂', 'A'].includes(command.text)).map(command => command.text), ['中', '🙂', 'A'])
  assert.ok(scene.commands.some(command => command.kind === 'rect' && command.fill === 'rgba(36,107,253,0.09)'))
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.y1 !== command.y2))
})

test('parser normalizes OOXML vertical sentinel and model/view mapping preserves it', () => {
  const workbook = normalizeWorkbook({
    schemaVersion: 1, parserVersion: 'test', format: 'xlsx', fileName: 'vertical.xlsx', size: 1, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { sheets: [{ id: 's', name: 'S', cells: [{
      row: 0, col: 0, value: 'Vertical', style: { rotation: 255, rotationVertical: true },
    }] }] } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  })
  const imported = workbook.sheets[0].cells['0:0']
  assert.equal(imported.style.rotationVertical, true)
  assert.equal(imported.style.rotation, 0)

  const model = createBlankWorkbook()
  model.sheets[0].cells['0:0'] = createCell(0, 0, 'Vertical')
  model.sheets[0].cells['0:0'].style.rotationVertical = true
  const view = toExcelWorkbookView(model)
  assert.equal(view.sheets[0].cells.A1.rotationVertical, true)
  const restored = applyExcelWorkbookView(model, view)
  assert.equal(restored.sheets[0].cells['0:0'].style.rotationVertical, true)
})

test('Canvas2D applies the same text rotation and max-width constraint', () => {
  const calls = []
  const context = {
    setTransform() {}, fillRect() {}, save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
    translate(...args) { calls.push(['translate', ...args]) },
    rotate(...args) { calls.push(['rotate', ...args]) },
    fillText(...args) { calls.push(['fillText', ...args]) },
  }
  const driver = createCanvas2DDriver({ getContext: () => context })
  driver.resize(100, 100, 1)
  driver.render({ width: 100, height: 100, commands: [{
    kind: 'text', x: 20, y: 30, text: 'abc', font: '12px sans-serif', color: '#000000',
    rotation: 45, maxWidth: 10,
  }] })

  assert.deepEqual(calls.find(call => call[0] === 'translate'), ['translate', 20, 30])
  assert.deepEqual(calls.find(call => call[0] === 'rotate'), ['rotate', Math.PI / 4])
  assert.deepEqual(calls.find(call => call[0] === 'fillText'), ['fillText', 'abc', 0, 0, 10])
})
