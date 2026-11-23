import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { createSheetContextMenuItems } from '../src/components/spreadsheet/sheet-context-menu.ts'
import { ribbonInvocation, visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'
import { createBlankWorkbook, createBlankSheet, createCell } from '../src/domain/model.ts'
import { captureSheetClipboard } from '../src/features/sheet-workbench/sheet-clipboard-data.ts'
import { sheetClipboardCapabilities } from '../src/features/sheet-workbench/sheet-clipboard-capabilities.ts'

// Draft only: no Ribbon/component verification was run in this batch.
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
const state = { canUndo: false, canRedo: false, hasSheet: true, canMerge: false, canUnmerge: false, canDeleteRow: true,
  canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true, canSort: true,
  fontFamily: 'Arial', fontSize: 11, numberFormat: 'general', gridlines: true, frozen: false, zoom: 100, fullscreen: false }
const context = { sheetId: 'target', canDeleteRow: true, canDeleteColumn: true, frozen: false, hasHiddenRows: false, hasHiddenColumns: false }
const hit = { kind: 'cell', row: 4, column: 4, key: 'E5' }
const flatten = items => items.flatMap(item => [item, ...flatten(item.children ?? [])])
function capabilities(extra = {}, cache = true) {
  const book = createBlankWorkbook(); book.id = 'clipboard-menu-book'
  book.sheets = ['source', 'target'].map(id => ({ ...createBlankSheet(id), id, rowCount: 40, columnCount: 12 }))
  book.sheets[0].cells['0:0'] = createCell(0, 0, 'text')
  const clipboard = cache ? captureSheetClipboard(book, { unitId: book.id, sheetId: 'source', range: range(0, 0) }) : undefined
  return sheetClipboardCapabilities({ workbook: book, sheetId: 'target', clipboard, selection: { ranges: [range(4, 4)] },
    userId: 'local-user', systemClipboardAvailable: false, ...extra })
}
function pasteRoot(caps) { return createSheetRibbon({ ...state, clipboardCapabilities: caps })[0].groups[0].items.find(item => item.id === 'sheet.command.paste') }

test('Ribbon and context clipboard leaves share each mode capability and block disabled invocations', () => {
  const caps = capabilities({ permission: request => request.point !== 'setCellStyle' }), ribbon = pasteRoot(caps)
  const menu = flatten(createSheetContextMenuItems(hit, { ...context, clipboardCapabilities: caps }))
  for (const [command, mode] of [['sheet.command.paste', 'all'], ['sheet.command.paste-values', 'values'],
    ['sheet.command.paste-formulas', 'formulas'], ['sheet.command.paste-formats', 'formats'],
    ['sheet.command.paste-column-widths', 'columnWidths'], ['sheet.command.paste-except-borders', 'exceptBorders']]) {
    const ribbonItem = ribbon.children.find(item => item.commandId === command), menuItem = menu.find(item => item.commandId === command)
    assert.equal(ribbonItem.disabled, caps.paste[mode].disabled); assert.equal(menuItem.disabled, ribbonItem.disabled)
    assert.equal(ribbonItem.disabledReason, caps.paste[mode].reason); assert.equal(menuItem.disabledReason, ribbonItem.disabledReason)
    if (ribbonItem.disabled) assert.equal(ribbonInvocation(ribbonItem), undefined)
    else assert.equal(ribbonInvocation(ribbonItem).commandId, command)
  }
})

test('an unavailable default paste does not prevent opening permitted special modes', () => {
  const selective = pasteRoot(capabilities({ permission: request => request.point !== 'setCellValue' }))
  assert.equal(selective.type, 'menu'); assert.equal(selective.disabled, false); assert.equal(ribbonInvocation(selective), undefined)
  assert.equal(selective.children.find(item => item.commandId === 'sheet.command.paste').disabled, true)
  assert.equal(selective.children.find(item => item.commandId === 'sheet.command.paste-formats').disabled, false)
  assert.equal(selective.children.find(item => item.commandId === 'sheet.command.paste-column-widths').disabled, false)
  const blocked = pasteRoot(capabilities({ permission: request => request.point !== 'edit' }))
  assert.equal(blocked.disabled, true); assert.ok(blocked.children.every(item => item.disabled)); assert.ok(blocked.disabledReason)
})

test('exclusive plugin ranges hide special containers and the Ribbon root while ordinary quick actions remain distinct', () => {
  const caps = capabilities({ exclusiveRange: true }), ribbon = visibleRibbonTabs(createSheetRibbon({ ...state, clipboardCapabilities: caps }))
  assert.ok(!ribbon.flatMap(tab => tab.groups.flatMap(group => group.items)).some(item => item.id === 'sheet.command.paste'))
  const menu = createSheetContextMenuItems(hit, { ...context, clipboardCapabilities: caps })
  assert.ok(!menu.some(item => item.id === 'copy-special' || item.id === 'paste-special'))
  assert.deepEqual(menu.filter(item => item.quick).map(item => item.id), ['copy', 'cut', 'paste'])
  const multi = capabilities({ selection: { ranges: [range(4, 4), range(10, 8)], activeRangeIndex: 0 } })
  assert.equal(pasteRoot(multi).hidden, false)
  assert.ok(createSheetContextMenuItems(hit, { ...context, clipboardCapabilities: multi }).some(item => item.id === 'paste-special'))
})

test('absent clipboard content disables paste everywhere without disabling copy and cut or fabricating source content', () => {
  const caps = capabilities({}, false), menu = flatten(createSheetContextMenuItems(hit, { ...context, clipboardCapabilities: caps }))
  assert.equal(pasteRoot(caps).disabled, true)
  assert.equal(menu.find(item => item.id === 'copy').disabled, false); assert.equal(menu.find(item => item.id === 'cut').disabled, false)
  assert.ok(menu.filter(item => item.commandId.startsWith('sheet.command.paste')).every(item => item.disabled && item.disabledReason))
  const reader = capabilities({ systemClipboardAvailable: true }, false)
  assert.equal(pasteRoot(reader).disabled, false); assert.equal(reader.source.kind, 'none')
})
