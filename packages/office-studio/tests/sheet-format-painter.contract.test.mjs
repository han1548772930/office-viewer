import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createPinia } from 'pinia'
import { reactive } from 'vue'
import { createBlankWorkbook, createCell, defaultCellStyle } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import {
  applySheetFormatPainterCommand, applySheetViewFormatPainterCommand, captureSheetFormatPainterSource,
  captureSheetViewFormatPainterSource, cloneSheetFormatPainterSource, createSheetFormatPainterPlan,
  MAXIMUM_FORMAT_PAINTER_CELLS, previewSheetFormatPainterCommand,
} from '../src/features/sheet-workbench/sheet-format-painter.ts'
import {
  formatPainterActionForCommand, SHEET_FORMAT_PAINTER_ALIASES, SHEET_FORMAT_PAINTER_COMMAND_IDS,
  SHEET_FORMAT_PAINTER_COMMANDS as commands, SheetFormatPainterSession, sheetFormatPainterCommandInput,
} from '../src/features/sheet-workbench/sheet-format-painter-session.ts'
import { useSheetFormatPainterStore } from '../src/features/sheet-workbench/sheet-format-painter-store.ts'
import { conditionalFormatFormulaAt } from '../src/features/conditional-formatting/conditional-format-formulas.ts'

const range = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const coordinateRange = value => ({ startRow: value.startRow, endRow: value.endRow, startColumn: value.startCol, endColumn: value.endCol })
const contents = sheet => Object.fromEntries(Object.entries(sheet.cells).map(([key, { value, formula, dataType, cachedDataType, richText, metadata, hyperlink, sparkline }]) => [key, { value, formula, dataType, cachedDataType, richText, metadata, hyperlink, sparkline }]))
const protection = ranges => [{ id: 'painter-locked', permissionId: 'painter-locked-p', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges }]
function sample() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'painter-contract'; workbook.activeSheetId = sheet.id = 'painter-main'; workbook.calculationMode = 'manual'
  sheet.rowCount = 200; sheet.columnCount = 26
  sheet.cells['0:0'] = { ...createCell(0, 0, 123.456), metadata: { sourceValue: 'excluded' }, richText: [{ text: 'source rich text', style: { italic: true } }] }
  sheet.cells['0:0'].style = { ...defaultCellStyle(), font: { family: 'Courier New', size: 18, bold: true, italic: true, underline: true, strike: true, color: '#123456' }, fill: '#abcdef', horizontal: 'right', vertical: 'bottom', wrap: true, wrapStrategy: 'wrap', numberFormat: '0.00', border: { bottom: { style: 'double', color: '#336699' } }, rotation: 45, rotationVertical: false, indent: 2, shrinkToFit: true, metadata: { vendor: { id: 'source-style' } } }
  sheet.cells['0:1'] = createCell(0, 1, 'source B1'); sheet.cells['0:1'].style.font.italic = true
  sheet.cells['1:0'] = createCell(1, 0, 'source A2'); sheet.cells['1:0'].style.fill = '#00ff00'
  sheet.cells['1:1'] = createCell(1, 1, 'default source')
  sheet.cells['3:3'] = { ...createCell(3, 3, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number', display: '0.25', metadata: { destination: { keep: true } }, hyperlink: 'https://destination.test' }
  sheet.cells['3:4'] = { ...createCell(3, 4, '00123'), dataType: 'text', richText: [{ text: '00123', style: { italic: true } }], metadata: { target: 'text' } }
  sheet.cells['4:3'] = createCell(4, 3, 7); sheet.cells['4:3'].style.fill = '#ff0000'
  sheet.cells['4:4'] = createCell(4, 4, false); sheet.cells['4:4'].style.font.bold = true
  sheet.cells['10:13'] = createCell(10, 13, 'outside')
  return workbook
}
function mergedSample() {
  const workbook = sample()
  workbook.sheets[0].mergedRanges = [range(0, 0, 0, 1)]
  return workbook
}
const sourceFor = (sheet, selection = range(0, 0, 1, 1)) => captureSheetFormatPainterSource(sheet, selection)

test('sampling stores sparse, deeply immutable formatting and excludes source content', () => {
  const sheet = sample().sheets[0], source = sourceFor(sheet)
  assert.deepEqual(Object.keys(source.styles).sort(), ['0:0', '0:1', '1:0'])
  assert.deepEqual(source.range, range(0, 0, 1, 1)); assert.deepEqual(source.merges, [])
  assert.equal(Object.isFrozen(source), true); assert.equal(Object.isFrozen(source.styles['0:0'].font), true)
  assert.equal(Object.isFrozen(source.styles['0:0'].metadata.vendor), true)
  assert.equal(JSON.stringify(source).includes('sourceValue'), false); assert.equal(JSON.stringify(source).includes('source rich text'), false)
  const before = cloneModel(source)
  sheet.cells['0:0'].style.font.size = 72; sheet.cells['0:0'].style.metadata.vendor.id = 'edited later'
  assert.deepEqual(cloneModel(source), before)
  assert.throws(() => { source.styles['0:0'].font.size = 99 }, TypeError)
})

test('view and domain capture share formatting, relative coordinates and merge snapshots', () => {
  const workbook = mergedSample(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0]
  const modelSource = sourceFor(sheet, range(0, 1)), viewSource = captureSheetViewFormatPainterSource(view, range(0, 1))
  assert.deepEqual(cloneModel(viewSource), cloneModel(modelSource))
  assert.deepEqual(modelSource.range, range(0, 0, 0, 1)); assert.deepEqual(modelSource.merges, [range(0, 0, 0, 1)])
})

test('full worksheet sampling stays sparse and clones reactive source data safely', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.rowCount = 1_048_576; sheet.columnCount = 16_384
  sheet.cells['1048575:16383'] = createCell(1_048_575, 16_383, 'kept'); sheet.cells['1048575:16383'].style.font.bold = true
  const source = captureSheetFormatPainterSource(reactive(sheet), range(0, 0, 1_048_575, 16_383))
  assert.deepEqual(Object.keys(source.styles), ['1048575:16383'])
  const copy = cloneSheetFormatPainterSource(reactive(cloneModel(source)))
  assert.deepEqual(cloneModel(copy), cloneModel(source)); assert.equal(Object.isFrozen(copy.styles['1048575:16383']), true)
})

test('single-cell targets expand to the complete sampled block and retain destination content', () => {
  const sheet = sample().sheets[0], before = contents(sheet), source = sourceFor(sheet)
  const result = applySheetFormatPainterCommand(sheet, source, [range(3, 3)])
  assert.deepEqual(result.ranges, [range(3, 3, 4, 4)]); assert.equal(result.applied, true); assert.equal(result.changed, true)
  assert.equal(sheet.cells['3:3'].style.font.family, 'Courier New'); assert.equal(sheet.cells['3:4'].style.font.italic, true)
  assert.equal(sheet.cells['4:3'].style.fill, '#00ff00'); assert.equal(sheet.cells['4:4'].style.font.bold, false)
  assert.deepEqual(contents(sheet), before)
})

test('rectangular targets tile the sampled rows and columns without expanding the selected size', () => {
  const sheet = sample().sheets[0], source = sourceFor(sheet), target = range(6, 6, 8, 10)
  const result = applySheetFormatPainterCommand(sheet, source, [target])
  assert.deepEqual(result.ranges, [target])
  for (let row = 6; row <= 8; row++) for (let col = 6; col <= 10; col++) {
    const style = sheet.cells[row + ':' + col]?.style ?? defaultCellStyle()
    assert.equal(style.font.bold, (row - 6) % 2 === 0 && (col - 6) % 2 === 0)
    assert.equal(style.fill, (col - 6) % 2 === 1 ? null : (row - 6) % 2 ? '#00ff00' : '#abcdef')
  }
  assert.equal(sheet.cells['9:10'], undefined); assert.equal(sheet.cells['8:11'], undefined)
})

test('later disjoint-selection entries control the tiling phase at overlapping cells', () => {
  const sheet = sample().sheets[0], source = sourceFor(sheet)
  applySheetFormatPainterCommand(sheet, source, [range(5, 5, 7, 7), range(6, 6, 8, 8)])
  assert.equal(sheet.cells['6:6'].style.font.bold, true)
  assert.equal(sheet.cells['7:7']?.style.font.bold ?? false, false)
  assert.equal(sheet.cells['5:5'].style.font.bold, true)
})

test('default source cells reset all supported styling while retaining target values and extensions', () => {
  const sheet = sample().sheets[0], target = sheet.cells['3:3'], before = contents(sheet)
  target.style = cloneModel(sheet.cells['0:0'].style)
  applySheetFormatPainterCommand(sheet, sourceFor(sheet, range(1, 1)), [range(3, 3)])
  const style = target.style
  assert.deepEqual(style.font, defaultCellStyle().font); assert.equal(style.fill, null); assert.equal(style.numberFormat, 'General')
  assert.equal(style.rotation, 0); assert.equal(style.rotationVertical, false); assert.equal(style.indent, 0)
  assert.equal(style.shrinkToFit, false); assert.equal(style.wrapStrategy, 'overflow'); assert.deepEqual(style.border, {})
  assert.equal(style.metadata, undefined); assert.deepEqual(contents(sheet), before)
})

test('formula cache and numeric type remain destination-owned when the number format changes', () => {
  const sheet = sample().sheets[0]; sheet.cells['0:0'].style.numberFormat = '0%'
  const source = sourceFor(sheet, range(0, 0))
  applySheetFormatPainterCommand(sheet, source, [range(3, 3)])
  const target = sheet.cells['3:3']
  assert.equal(target.value, .25); assert.equal(target.formula, '=1/4'); assert.equal(target.cachedDataType, 'number')
  assert.equal(target.display, '25%'); assert.notEqual(target.value, sheet.cells['0:0'].value)
  assert.deepEqual(target.metadata, { destination: { keep: true } }); assert.equal(target.hyperlink, 'https://destination.test')
})

test('style-only application retains rich text, hyperlinks and unrelated cells', () => {
  const sheet = sample().sheets[0], original = cloneModel(sheet.cells['3:4']), outside = cloneModel(sheet.cells['10:13'])
  applySheetFormatPainterCommand(sheet, sourceFor(sheet, range(0, 0)), [range(3, 4)])
  const target = sheet.cells['3:4']
  assert.equal(target.value, '00123'); assert.deepEqual(target.richText, original.richText); assert.deepEqual(target.metadata, original.metadata)
  assert.notDeepEqual(target.style, original.style); assert.deepEqual(sheet.cells['10:13'], outside)
})

test('view formatting retains raw formula cache, rich text and style metadata through model adaptation', () => {
  const workbook = sample(), before = contents(workbook.sheets[0]), view = toExcelWorkbookView(workbook)
  const source = captureSheetViewFormatPainterSource(view.sheets[0], range(0, 0, 1, 1))
  const result = applySheetViewFormatPainterCommand(view.sheets[0], source, [range(3, 3)])
  assert.equal(result.applied, true)
  const reopened = applyExcelWorkbookView(workbook, view), sheet = reopened.sheets[0]
  assert.deepEqual(contents(sheet), before); assert.equal(sheet.cells['3:3'].style.font.size, 18)
  assert.deepEqual(sheet.cells['3:3'].style.metadata, { vendor: { id: 'source-style' } })
  assert.equal(sheet.cells['3:3'].style.rotation, 45); assert.equal(sheet.cells['3:3'].style.shrinkToFit, true)
})

test('successful default paints do not create blank cells and repeated paints produce no extra changes', () => {
  const sheet = sample().sheets[0], source = sourceFor(sheet, range(1, 1)), before = cloneModel(sheet)
  const result = applySheetFormatPainterCommand(sheet, source, [range(8, 8, 9, 9)])
  assert.equal(result.applied, true); assert.equal(result.changed, false); assert.deepEqual(sheet, before)
  const nondefault = sourceFor(sheet, range(0, 0))
  assert.equal(applySheetFormatPainterCommand(sheet, nondefault, [range(3, 3)]).changed, true)
  assert.equal(applySheetFormatPainterCommand(sheet, nondefault, [range(3, 3)]).changed, false)
})

test('an unmerged source preserves pre-existing target merge topology', () => {
  const sheet = sample().sheets[0]; sheet.mergedRanges = [range(3, 3, 4, 4)]
  const before = cloneModel(sheet.mergedRanges)
  const result = applySheetFormatPainterCommand(sheet, sourceFor(sheet, range(0, 0)), [range(3, 3, 4, 4)])
  assert.equal(result.requiresConfirmation, false); assert.deepEqual(sheet.mergedRanges, before)
  assert.deepEqual(result.added, []); assert.deepEqual(result.removed, [])
})

test('source merges repeat only in complete source blocks and may extend outside a smaller multi-cell target', () => {
  const sheet = sample().sheets[0]; sheet.mergedRanges = [range(0, 0, 0, 2)]
  const source = sourceFor(sheet, range(0, 0, 1, 2))
  const tiled = createSheetFormatPainterPlan([], source, [range(6, 6, 10, 13)])
  assert.deepEqual(tiled.mergeRanges, [range(6, 6, 6, 8), range(6, 9, 6, 11), range(8, 6, 8, 8), range(8, 9, 8, 11)])
  const small = createSheetFormatPainterPlan([], source, [range(3, 3, 3, 4)])
  assert.deepEqual(small.ranges, [range(3, 3, 3, 4)]); assert.deepEqual(small.mergeRanges, [range(3, 3, 3, 5)])
})

test('merge confirmation counts discarded content and leaves formatting and topology unchanged until accepted', () => {
  const sheet = mergedSample().sheets[0], source = sourceFor(sheet, range(0, 1)), before = cloneModel(sheet)
  const preview = previewSheetFormatPainterCommand(sheet, source, [range(3, 3)])
  assert.equal(preview.requiresConfirmation, true); assert.equal(preview.discardedContentCount, 1)
  const result = applySheetFormatPainterCommand(sheet, source, [range(3, 3)])
  assert.equal(result.applied, false); assert.deepEqual(sheet, before)
  const accepted = applySheetFormatPainterCommand(sheet, source, [range(3, 3)], { mergeConfirmed: true })
  assert.equal(accepted.applied, true); assert.equal(accepted.contentChanged, true)
  assert.ok(sheet.mergedRanges.some(value => JSON.stringify(value) === JSON.stringify(range(3, 3, 3, 4))))
  assert.equal(sheet.cells['3:3'].formula, '=1/4'); assert.equal(sheet.cells['3:3'].value, .25)
  assert.equal(sheet.cells['3:4'].value, null); assert.equal(sheet.cells['3:4'].style.font.italic, true)
})

test('merge painting shares first-nonempty donor semantics in domain and standalone view commands', () => {
  const workbook = mergedSample(), sheet = workbook.sheets[0]
  sheet.cells['3:3'] = createCell(3, 3, null)
  const donor = cloneModel(sheet.cells['3:4']), view = toExcelWorkbookView(workbook)
  const result = applySheetViewFormatPainterCommand(view.sheets[0], sourceFor(sheet, range(0, 0)), [range(3, 3)], { mergeConfirmed: true })
  assert.equal(result.applied, true)
  const reopened = applyExcelWorkbookView(workbook, view).sheets[0]
  assert.equal(reopened.cells['3:3'].value, donor.value); assert.deepEqual(reopened.cells['3:3'].richText, donor.richText)
  assert.deepEqual(reopened.cells['3:3'].metadata, donor.metadata); assert.equal(reopened.cells['3:4'].value, null)
  assert.equal(reopened.cells['3:3'].style.font.bold, true)
})

test('protection covers merge-expanded sampling, expanded targets and replaced merge boundaries', () => {
  const workbook = mergedSample(), sheet = workbook.sheets[0]
  sheet.rangeProtectionRules = protection([range(0, 0)])
  assert.throws(() => sourceFor(sheet, range(0, 1)), /保护/)
  delete sheet.rangeProtectionRules
  const source = sourceFor(sheet, range(0, 0)); sheet.mergedRanges.push(range(3, 4, 3, 5))
  sheet.rangeProtectionRules = protection([range(3, 5)])
  const before = cloneModel(sheet), view = toExcelWorkbookView(workbook).sheets[0], beforeView = cloneModel(view)
  assert.throws(() => applySheetFormatPainterCommand(sheet, source, [range(3, 3)], { mergeConfirmed: true }), /保护/)
  assert.throws(() => applySheetViewFormatPainterCommand(view, source, [range(3, 3)], { mergeConfirmed: true }), /保护/)
  assert.deepEqual(sheet, before); assert.deepEqual(view, beforeView)
})

test('hidden and filtered target rows retain the physical formatting scope', () => {
  const sheet = sample().sheets[0], source = sourceFor(sheet, range(0, 0)), before = contents(sheet)
  sheet.hiddenRows = [4]; sheet.autoFilter = { range: range(3, 3, 5, 3), columns: {}, filteredRows: [3] }
  applySheetFormatPainterCommand(sheet, source, [range(3, 3, 4, 3)])
  assert.equal(sheet.cells['3:3'].style.font.size, 18); assert.equal(sheet.cells['4:3'].style.font.size, 18)
  assert.deepEqual(sheet.hiddenRows, [4]); assert.deepEqual(sheet.autoFilter.filteredRows, [3]); assert.deepEqual(contents(sheet), before)
})

test('invalid style coordinates, malformed styles and uncloneable snapshots reject before merging content', () => {
  const sheet = mergedSample().sheets[0], source = sourceFor(sheet, range(0, 0)), before = cloneModel(sheet)
  const invalid = [
    value => { value.styles['-1:0'] = defaultCellStyle() },
    value => { value.styles['00:0'] = defaultCellStyle() },
    value => { value.styles['0:2'] = defaultCellStyle() },
    value => { value.styles['0:0'].font = null },
    value => { value.styles['0:0'].font.size = NaN },
    value => { value.styles['0:0'].wrapStrategy = 'invalid' },
    value => { value.styles['0:0'].metadata = { bad: () => 1 } },
    value => { value.styles['0:0'].metadata = {}; value.styles['0:0'].metadata.self = value.styles['0:0'].metadata },
    value => { value.merges = [range(0, 0, 0, 3)] },
  ]
  for (const mutate of invalid) {
    const value = cloneModel(source); mutate(value)
    assert.throws(() => applySheetFormatPainterCommand(sheet, value, [range(3, 3)], { mergeConfirmed: true }))
    assert.deepEqual(sheet, before)
  }
})

test('invalid and overflowing target coordinates cannot leave partial domain or view changes', () => {
  const workbook = sample(), sheet = workbook.sheets[0], source = sourceFor(sheet), view = toExcelWorkbookView(workbook).sheets[0]
  const before = cloneModel(sheet), beforeView = cloneModel(view)
  for (const targets of [[range(-1, 0)], [range(1_048_575, 16_383)], [range(3, 3), range(3.5, 4)], [range(2, 2, 1, 1)]]) {
    assert.throws(() => applySheetFormatPainterCommand(sheet, source, targets), RangeError)
    assert.throws(() => applySheetViewFormatPainterCommand(view, source, targets), RangeError)
    assert.deepEqual(sheet, before); assert.deepEqual(view, beforeView)
  }
  assert.equal(applySheetFormatPainterCommand(sheet, source, []).applied, false); assert.deepEqual(sheet, before)
})

test('synchronous cost checks include single-cell expansion and merge content outside the nominal target', () => {
  const sheet = sample().sheets[0], before = cloneModel(sheet)
  const large = sourceFor(sheet, range(0, 0, 1000, 100))
  assert.ok((large.range.endRow + 1) * (large.range.endCol + 1) > MAXIMUM_FORMAT_PAINTER_CELLS)
  assert.throws(() => applySheetFormatPainterCommand(sheet, large, [range(2000, 0)]), /最多同步处理/)
  assert.deepEqual(sheet, before)
  const wide = { range: range(0, 0, 0, 1000), styles: {}, merges: [range(0, 0, 0, 1000)] }
  const targets = Array.from({ length: 101 }, (_, row) => range(row * 2, 0, row * 2, 1))
  assert.throws(() => applySheetFormatPainterCommand(sheet, wide, targets, { mergeConfirmed: true }), /最多同步处理/)
  assert.deepEqual(sheet, before)
})

test('replacing a large target merge is bounded before restoring content outside the brush rectangle', () => {
  const workbook = mergedSample(), sheet = workbook.sheets[0], source = sourceFor(sheet, range(0, 0))
  sheet.mergedRanges.push(range(3, 3, 100_003, 3))
  const before = cloneModel(sheet), view = toExcelWorkbookView(workbook).sheets[0], beforeView = cloneModel(view)
  assert.throws(() => applySheetFormatPainterCommand(sheet, source, [range(3, 3)], { mergeConfirmed: true }), /最多同步处理/)
  assert.throws(() => applySheetViewFormatPainterCommand(view, source, [range(3, 3)], { mergeConfirmed: true }), /最多同步处理/)
  assert.deepEqual(sheet, before); assert.deepEqual(view, beforeView)
})

test('all command aliases resolve to their observable modes and validate input contracts', () => {
  assert.equal(new Set(SHEET_FORMAT_PAINTER_COMMAND_IDS).size, SHEET_FORMAT_PAINTER_COMMAND_IDS.length)
  for (const id of [commands.activate, commands.once, SHEET_FORMAT_PAINTER_ALIASES.activate]) assert.deepEqual(formatPainterActionForCommand(id), { kind: 'activate', continuous: false })
  for (const id of [commands.continuous, SHEET_FORMAT_PAINTER_ALIASES.continuous]) assert.deepEqual(formatPainterActionForCommand(id), { kind: 'activate', continuous: true })
  assert.deepEqual(formatPainterActionForCommand(commands.infinite), { kind: 'activate', continuous: true, legacyToggle: true })
  for (const [status, mode] of [[0, 'off'], [1, 'once'], [2, 'continuous'], ['off', 'off'], ['once', 'once'], ['continuous', 'continuous']]) assert.deepEqual(formatPainterActionForCommand(commands.setStatus, { status }), { kind: 'set-mode', mode })
  for (const params of [null, [], { status: 'infinite' }, { expectedRevision: -1 }, { requestId: '' }, { mergeConfirmed: 1 }, { ranges: null }]) assert.throws(() => sheetFormatPainterCommandInput(params))
})

test('ONCE is retained after rejected or failed applications and consumed only by success', () => {
  const session = new SheetFormatPainterSession(), source = sourceFor(sample().sheets[0])
  assert.equal(session.activate('book', 'sheet', () => source), true)
  assert.equal(session.apply('book', () => false), false); assert.equal(session.getState().mode, 'once')
  assert.throws(() => session.apply('book', () => { throw new Error('retryable') }), /retryable/)
  assert.equal(session.getState().mode, 'once'); assert.equal(session.getState().applying, false)
  assert.equal(session.apply('book', () => true), true); assert.equal(session.getState().mode, 'off'); assert.equal(session.getSource(), null)
})

test('shared continuous upgrades preserve the immutable sample and never recapture the live selection', () => {
  const session = new SheetFormatPainterSession(), source = sourceFor(sample().sheets[0]); let captures = 0
  const capture = () => { captures++; return source }
  session.activate('book', 'sheet', capture)
  const initial = session.getState().sessionId
  session.activate('book', 'other-sheet', capture, true); session.activate('book', 'other-sheet', capture, true)
  assert.equal(captures, 1); assert.equal(session.getState().sessionId, initial); assert.equal(session.getState().sourceSheetId, 'sheet')
  const copy = session.getSource(); copy.styles['0:0'].font.size = 99
  assert.equal(session.apply('book', captured => { assert.equal(captured.styles['0:0'].font.size, 18); return true }), true)
  assert.equal(session.getState().mode, 'continuous')
})

test('legacy infinite toggles off while shared continuous activation keeps an existing continuous session', () => {
  const session = new SheetFormatPainterSession(), source = sourceFor(sample().sheets[0])
  session.activate('book', 'sheet', () => source, true)
  session.activate('book', 'sheet', () => null, true); assert.equal(session.getState().mode, 'continuous')
  session.activate('book', 'sheet', () => null, true, true); assert.equal(session.getState().mode, 'off')
})

test('sessions reject another workbook and stale session tokens without changing their source', () => {
  const session = new SheetFormatPainterSession(), source = sourceFor(sample().sheets[0])
  session.activate('book', 'sheet', () => source); const before = session.getState()
  assert.equal(session.activate('other', 'sheet', () => source, true), false)
  assert.equal(session.setMode('other', 'sheet', () => source, 'off'), false)
  assert.equal(session.apply('other', () => true), false); assert.equal(session.apply('book', () => true, 'stale'), false)
  assert.deepEqual(session.getState(), before)
})

test('a recursive application is rejected while the outer application succeeds once', () => {
  const session = new SheetFormatPainterSession(); session.activate('book', 'sheet', () => sourceFor(sample().sheets[0]))
  let writes = 0
  assert.equal(session.apply('book', () => {
    assert.equal(session.apply('book', () => { writes++; return true }), false); writes++; return true
  }), true)
  assert.equal(writes, 1); assert.equal(session.getState().applying, false); assert.equal(session.getState().mode, 'off')
})

test('an observer replacing the session before execution prevents the old source from applying', () => {
  const source = sourceFor(sample().sheets[0]); let replace = false, session
  session = new SheetFormatPainterSession(state => {
    if (replace && state.applying) { replace = false; session.cancel(); session.activate('book', 'fresh-sheet', () => source, true) }
  })
  session.activate('book', 'old-sheet', () => source); const oldId = session.getState().sessionId; replace = true
  let writes = 0
  assert.equal(session.apply('book', () => { writes++; return true }), false)
  assert.equal(writes, 0); assert.notEqual(session.getState().sessionId, oldId)
  assert.equal(session.getState().sourceSheetId, 'fresh-sheet'); assert.equal(session.getState().mode, 'continuous'); assert.equal(session.getState().applying, false)
})

test('a newly activated session during a successful application is not consumed with the old ONCE session', () => {
  const session = new SheetFormatPainterSession(), source = sourceFor(sample().sheets[0])
  session.activate('book', 'old-sheet', () => source)
  assert.equal(session.apply('book', () => { session.cancel(); session.activate('book', 'fresh-sheet', () => source); return true }), true)
  assert.equal(session.getState().mode, 'once'); assert.equal(session.getState().sourceSheetId, 'fresh-sheet')
})

test('notification errors release the applying guard and leave the source available for retry', () => {
  let reject = false
  const session = new SheetFormatPainterSession(state => { if (reject && state.applying) { reject = false; throw new Error('observer failed') } })
  session.activate('book', 'sheet', () => sourceFor(sample().sheets[0])); reject = true
  assert.throws(() => session.apply('book', () => true), /observer failed/)
  assert.equal(session.getState().applying, false); assert.equal(session.getState().mode, 'once')
  assert.equal(session.apply('book', () => true), true)
})

test('confirmation requests are isolated snapshots, scoped to the session and canceled without consuming it', () => {
  const session = new SheetFormatPainterSession(), source = sourceFor(sample().sheets[0]); session.activate('book', 'sheet', () => source)
  const request = { requestId: 'request', sessionId: session.getState().sessionId, workbookId: 'book', sheetId: 'target', sourceSheetId: 'sheet', source: cloneModel(source), baseRevision: 0, discardedContentCount: 1, input: { ranges: [range(3, 3)] } }
  session.requestConfirmation(request); request.source.styles['0:0'].font.size = 99
  assert.equal(session.getConfirmation().source.styles['0:0'].font.size, 18)
  assert.equal(session.cancelConfirmation('wrong'), false); assert.equal(session.cancelConfirmation('request'), true)
  assert.equal(session.getState().mode, 'once'); assert.equal(session.getState().confirmationId, undefined)
  assert.throws(() => session.requestConfirmation({ ...request, sessionId: 'stale' }), /失效/)
  session.requestConfirmation(request); session.cancel(); assert.equal(session.getConfirmation(), null)
})

test('separate Pinia instances keep interaction sessions isolated from document state', () => {
  const first = useSheetFormatPainterStore(createPinia()), second = useSheetFormatPainterStore(createPinia())
  first.activate('book', 'sheet', () => sourceFor(sample().sheets[0]))
  assert.equal(first.state.mode, 'once'); assert.equal(second.state.mode, 'off')
  second.cancel(); assert.equal(first.state.mode, 'once')
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run, workbook = sample()) {
  const studio = new OfficeStudio({ workbook, currentUserId: 'local-user' })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}
const undoCount = studio => studio.editor.history.units[studio.activeWorkbookId]?.undo.length ?? 0
const paintInput = (studio, ranges, extra = {}) => ({ unitId: studio.activeWorkbookId, sheetId: studio.editor.workbook.model.sheets[0].id, ranges, ...extra })
const activate = (studio, ranges = [range(0, 0, 1, 1)], id = commands.activate) => studio.executeCommandSync(id, paintInput(studio, ranges))

test('registered activation samples the last selection and explicit apply creates a single history entry', () => withStudio((studio, api) => {
  const executed = [], before = contents(studio.editor.workbook.model.sheets[0])
  studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
  assert.equal(activate(studio, [range(1, 1), range(0, 0)]), true)
  assert.deepEqual(studio.sheetFormatPainter.state.sourceRange, range(0, 0))
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [coordinateRange(range(3, 3))])), true)
  assert.equal(studio.editor.workbook.model.sheets[0].cells['3:3'].style.font.size, 18)
  assert.deepEqual(contents(studio.editor.workbook.model.sheets[0]), before); assert.equal(undoCount(studio), 1)
  assert.deepEqual(executed, [commands.activate, commands.apply]); assert.equal(studio.sheetFormatPainter.state.mode, 'off')
  api.getActiveWorkbook().undo(); assert.equal(studio.editor.workbook.model.sheets[0].cells['3:3'].style.font.size, 11)
  api.getActiveWorkbook().redo(); assert.equal(studio.editor.workbook.model.sheets[0].cells['3:3'].style.font.size, 18)
  const saved = api.getActiveWorkbook().save(), reopened = api.createWorkbook({ ...saved, id: 'painter-reopened' })
  assert.equal(studio.getWorkbookModel(reopened.getId()).sheets[0].cells['3:3'].style.font.size, 18)
  assert.deepEqual(contents(studio.getWorkbookModel(reopened.getId()).sheets[0]), before)
}))

test('single-cell expansion updates a matching API selection without resampling the source', () => withStudio((studio, api) => {
  const sheet = api.getActiveWorkbook().getActiveSheet()
  assert.equal(activate(studio), true); sheet.getRange('D4').activate()
  assert.equal(studio.executeCommandSync(commands.apply), true)
  const selection = studio.getSelection()
  assert.deepEqual({ startRow: selection.startRow, endRow: selection.endRow, startCol: selection.startColumn, endCol: selection.endColumn }, range(3, 3, 4, 4))
  assert.equal(undoCount(studio), 1)
}))

test('API continuous mode paints another sheet in the same workbook using the original source', () => {
  const workbook = sample(), second = cloneModel(createBlankWorkbook().sheets[0]); second.id = 'painter-second'; second.cells['3:3'] = createCell(3, 3, 'other target'); workbook.sheets.push(second)
  withStudio((studio, api) => {
    activate(studio, [range(0, 0)], commands.continuous)
    const token = studio.sheetFormatPainter.state.sessionId
    api.getActiveWorkbook().setActiveSheet(second.id)
    assert.equal(studio.executeCommandSync(commands.apply, { unitId: workbook.id, sheetId: second.id, ranges: [range(3, 3)] }), true)
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[1].cells['3:3'].value, 'other target')
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[1].cells['3:3'].style.font.size, 18)
    assert.equal(studio.sheetFormatPainter.state.sessionId, token); assert.equal(studio.sheetFormatPainter.state.sourceSheetId, 'painter-main')
    assert.equal(studio.sheetFormatPainter.state.mode, 'continuous')
  }, workbook)
})

test('shared UI aliases cannot activate from a Docs context with stale cell coordinates', () => withStudio(studio => {
  studio.editor.state.activeKind = 'document'
  assert.equal(studio.executeCommandSync(SHEET_FORMAT_PAINTER_ALIASES.activate, paintInput(studio, [range(0, 0)])), false)
  assert.equal(studio.sheetFormatPainter.state.mode, 'off'); assert.equal(undoCount(studio), 0)
}))

test('BeforeCommandExecute cancellation preserves captured source and publishes no apply success', () => withStudio(studio => {
  activate(studio, [range(0, 0)]); const before = cloneModel(studio.editor.workbook.model), events = []
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  const canceled = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commands.apply) event.cancel = true })
  assert.throws(() => studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), /canceled/)
  assert.deepEqual(studio.editor.workbook.model, before); assert.equal(studio.sheetFormatPainter.state.mode, 'once')
  assert.deepEqual(events, []); assert.equal(undoCount(studio), 0)
  canceled.dispose(); assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), true)
  assert.equal(undoCount(studio), 1)
}))

test('a successful redundant API paint consumes ONCE and emits the command without adding history', () => withStudio(studio => {
  activate(studio, [range(1, 1)]); const before = cloneModel(studio.editor.workbook.model), events = []
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(8, 8)])), true)
  assert.equal(studio.sheetFormatPainter.state.mode, 'off'); assert.equal(undoCount(studio), 0)
  assert.deepEqual(studio.editor.workbook.model, before); assert.deepEqual(events, [commands.apply])
}))

test('target permission refusal leaves the sample available and allows a later authorized retry', () => withStudio(studio => {
  activate(studio, [range(0, 0)])
  const locked = cloneModel(studio.editor.workbook.model); locked.sheets[0].rangeProtectionRules = protection([range(3, 3)])
  studio.editor.workbook.replace(locked)
  assert.throws(() => studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), /保护/)
  assert.deepEqual(studio.editor.workbook.model, locked); assert.equal(undoCount(studio), 0); assert.equal(studio.sheetFormatPainter.state.mode, 'once')
  const unlocked = cloneModel(locked); delete unlocked.sheets[0].rangeProtectionRules; studio.editor.workbook.replace(unlocked)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), true); assert.equal(undoCount(studio), 1)
}))

test('merge confirmation is a separate command with exact source, target and revision contracts', () => withStudio(studio => {
  const events = [], values = []; studio.events.addEvent('SheetFormatPainterConfirmationRequested', event => events.push(cloneModel(event)))
  studio.events.addEvent('SheetValueChanged', event => values.push(event))
  activate(studio, [range(0, 1)]); const before = cloneModel(studio.editor.workbook.model)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), false)
  assert.equal(events.length, 1); assert.equal(events[0].discardedContentCount, 1); assert.equal(undoCount(studio), 0)
  assert.deepEqual(events[0].input.ranges, [range(3, 3)]); assert.deepEqual(events[0].source.range, range(0, 0, 0, 1))
  assert.equal(events[0].baseRevision, studio.getWorkbookRevision(before.id)); assert.deepEqual(studio.editor.workbook.model, before)
  assert.equal(studio.sheetFormatPainter.state.mode, 'once')
  assert.equal(studio.executeCommandSync(commands.confirm, { requestId: events[0].requestId, expectedRevision: events[0].baseRevision }), true)
  assert.equal(undoCount(studio), 1); assert.equal(studio.sheetFormatPainter.state.mode, 'off')
  assert.equal(studio.editor.workbook.model.sheets[0].cells['3:4'].value, null); assert.equal(values.length, 1)
}, mergedSample()))

test('canceling confirmation preserves ONCE and a new target produces an independent request', () => withStudio(studio => {
  activate(studio, [range(0, 0)])
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), false)
  const first = studio.sheetFormatPainter.getConfirmation()
  assert.equal(studio.executeCommandSync(commands.cancelConfirmation, { requestId: first.requestId }), true)
  assert.equal(studio.sheetFormatPainter.state.mode, 'once'); assert.equal(studio.sheetFormatPainter.getConfirmation(), null)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), false)
  assert.notEqual(studio.sheetFormatPainter.getConfirmation().requestId, first.requestId); assert.equal(undoCount(studio), 0)
}, mergedSample()))

test('content revisions invalidate merge confirmation without changing target formatting or topology', () => withStudio((studio, api) => {
  activate(studio, [range(0, 0)])
  studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])); const request = studio.sheetFormatPainter.getConfirmation()
  api.getActiveWorkbook().getActiveSheet().getRange('N11').setValue('new revision')
  const before = cloneModel(studio.editor.workbook.model)
  assert.throws(() => studio.executeCommandSync(commands.confirm, { requestId: request.requestId, expectedRevision: request.baseRevision }), /已更改/)
  assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 1)
  assert.equal(studio.sheetFormatPainter.state.mode, 'once'); assert.equal(studio.sheetFormatPainter.state.confirmationId, request.requestId)
}, mergedSample()))

test('confirmed merge painting validates all changed content in the candidate transaction', () => {
  const workbook = mergedSample(), sheet = workbook.sheets[0]
  sheet.cells['3:3'] = createCell(3, 3, 1); sheet.cells['3:4'] = createCell(3, 4, 2)
  sheet.dataValidations = [{ id: 'positive', type: 'custom', formula1: '=E4>0', allowBlank: false, showErrorMessage: true, errorStyle: 'stop', error: '数值必须大于零。', ranges: [range(3, 4)] }]
  withStudio(studio => {
    activate(studio, [range(0, 0)]); const before = cloneModel(studio.editor.workbook.model)
    assert.throws(() => studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)], { mergeConfirmed: true })), /大于零/)
    assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 0); assert.equal(studio.sheetFormatPainter.state.mode, 'once')
  }, workbook)
})

test('format clear ends the session while content clear and rejected format clear retain it', () => withStudio((studio, api) => {
  const target = api.getActiveWorkbook().getActiveSheet().getRange('N11')
  activate(studio, [range(0, 0)], commands.continuous); target.clearContent()
  assert.equal(studio.sheetFormatPainter.state.mode, 'continuous')
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'sheet.command.clear-format') event.cancel = true })
  assert.throws(() => target.clearFormat(), /canceled/); assert.equal(studio.sheetFormatPainter.state.mode, 'continuous')
  stop.dispose(); target.clearFormat(); assert.equal(studio.sheetFormatPainter.state.mode, 'off')
}))

test('API rejects cross-workbook activation and apply requests while retaining the owning session', () => withStudio(studio => {
  activate(studio, [range(0, 0)]); const before = cloneModel(studio.editor.workbook.model), state = cloneModel(studio.sheetFormatPainter.state)
  assert.equal(studio.executeCommandSync(commands.continuous, { unitId: 'other', sheetId: 'painter-main', ranges: [range(0, 0)] }), false)
  assert.equal(studio.executeCommandSync(commands.setStatus, { unitId: 'other', status: 0 }), false)
  assert.equal(studio.executeCommandSync(commands.apply, { unitId: 'other', sheetId: 'painter-main', ranges: [range(3, 3)] }), false)
  assert.deepEqual(studio.sheetFormatPainter.state, state); assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 0)
}))

test('source deletion, loss of view permission, editor switching and workbook switching clear ephemeral sessions', () => withStudio((studio, api) => {
  activate(studio, [range(0, 0)])
  const removed = cloneModel(studio.editor.workbook.model)
  removed.sheets = [cloneModel(createBlankWorkbook().sheets[0])]; removed.sheets[0].id = removed.activeSheetId = 'remaining-sheet'
  studio.editor.workbook.replace(removed); assert.equal(studio.sheetFormatPainter.state.mode, 'off')
  studio.editor.workbook.replace(sample()); activate(studio, [range(0, 0)])
  const hidden = cloneModel(studio.editor.workbook.model); hidden.sheets[0].rangeProtectionRules = [{ ...protection([range(0, 0)])[0], viewState: 'noOneElseCanView' }]
  hidden.sheets[0].worksheetProtection = { permissionId: 'source-view', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'noOneElseCanView' }
  studio.editor.workbook.replace(hidden); assert.equal(studio.sheetFormatPainter.state.mode, 'off')
  studio.editor.workbook.replace(sample()); activate(studio, [range(0, 0)])
  studio.editor.state.activeKind = 'document'; assert.equal(studio.sheetFormatPainter.state.mode, 'off')
  studio.editor.state.activeKind = 'workbook'; activate(studio, [range(0, 0)])
  const second = api.createWorkbook({ ...sample(), id: 'painter-second-book' }); assert.equal(studio.sheetFormatPainter.state.mode, 'off')
  studio.activateWorkbook('painter-contract'); activate(studio, [range(0, 0)])
  api.disposeUnit('painter-contract'); assert.equal(studio.sheetFormatPainter.state.mode, 'off'); assert.equal(studio.activeWorkbookId, second.getId())
}))

test('state notifications expose only interaction state and dispose releases the captured source', () => {
  const studio = new OfficeStudio({ workbook: sample() }), states = []
  studio.events.addEvent('SheetFormatPainterStateChanged', event => states.push(cloneModel(event.state)))
  activate(studio, [range(0, 0)], commands.continuous)
  assert.equal(states.some(state => state.mode === 'continuous'), true)
  assert.equal(states.some(state => 'styles' in state || 'source' in state), false)
  studio.dispose(); assert.equal(studio.sheetFormatPainter.state.mode, 'off'); assert.equal(studio.sheetFormatPainter.getSource(), null)
})

function conditionalSample(merged = false, formatting = false) {
  const workbook = merged ? mergedSample() : sample(), sheet = workbook.sheets[0]
  if (!formatting) for (const cell of Object.values(sheet.cells)) cell.style = defaultCellStyle()
  sheet.conditionalFormats = [
    { id: 'painter-cf-source', type: 'expression', ranges: [range(0, 0)], operator: 'greaterThan', formula1: '=A1>0', style: { fill: '#35a65f' }, priority: 1, stopIfTrue: true, metadata: { vendor: { keep: true } } },
    { id: 'painter-cf-target', type: 'cellIs', ranges: [range(3, 3, 4, 4)], operator: 'greaterThan', formula1: '0', style: { fill: '#dc3344' }, priority: 2, stopIfTrue: false },
  ]
  return workbook
}

test('conditional-only API painting shares one history transaction and skips value events and value validation', () => {
  const workbook = conditionalSample()
  workbook.sheets[0].dataValidations = [{ id: 'reject-values', type: 'custom', ranges: [range(3, 3)], formula1: '=FALSE()', allowBlank: false, showErrorMessage: true, errorStyle: 'stop', error: 'Value edits are blocked' }]
  withStudio((studio, api) => {
    const values = [], before = cloneModel(studio.editor.workbook.model), beforeContents = contents(before.sheets[0])
    studio.events.addEvent('SheetValueChanged', event => values.push(event))
    activate(studio, [range(0, 0)]); assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), true)
    const painted = cloneModel(studio.editor.workbook.model)
    assert.deepEqual(contents(painted.sheets[0]), beforeContents); assert.equal(undoCount(studio), 1); assert.deepEqual(values, [])
    assert.equal(painted.sheets[0].cells['3:3'].style.fill, null); assert.equal(painted.sheets[0].conditionalFormats[0].ranges.length, 2)
    api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
    api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, painted.sheets[0].conditionalFormats)
    const saved = api.getActiveWorkbook().save(), reopened = api.createWorkbook({ ...saved, id: 'painter-cf-reopened' })
    assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].conditionalFormats, painted.sheets[0].conditionalFormats)
  }, workbook)
})

test('formatting and conditional-rule replacement enter the same API edit and redundant continuous painting adds no history', () => withStudio((studio, api) => {
  const before = cloneModel(studio.editor.workbook.model)
  activate(studio, [range(0, 0)], commands.continuous)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), true)
  const painted = cloneModel(studio.editor.workbook.model)
  assert.equal(painted.sheets[0].cells['3:3'].style.font.size, 18); assert.equal(undoCount(studio), 1)
  assert.notDeepEqual(painted.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), true); assert.equal(undoCount(studio), 1)
  assert.equal(studio.sheetFormatPainter.state.mode, 'continuous')
  api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
  assert.equal(studio.editor.workbook.model.sheets[0].cells['3:3'].style.font.size, 11)
}, conditionalSample(false, true)))

test('cross-sheet API rule copies retain independent IDs and their relative formula through history and snapshot save', () => {
  const workbook = conditionalSample(), second = cloneModel(createBlankWorkbook().sheets[0]); second.id = 'painter-cf-second'; second.cells['3:3'] = createCell(3, 3, 10); workbook.sheets.push(second)
  withStudio((studio, api) => {
    activate(studio, [range(0, 0)], commands.continuous); api.getActiveWorkbook().setActiveSheet(second.id)
    const input = { unitId: workbook.id, sheetId: second.id, ranges: [range(3, 3)] }
    assert.equal(studio.executeCommandSync(commands.apply, input), true)
    const painted = cloneModel(studio.getWorkbookModel(workbook.id).sheets[1].conditionalFormats)
    assert.notEqual(painted[0].id, 'painter-cf-source'); assert.equal(painted[0].formula1, '=D4>0')
    assert.equal(studio.executeCommandSync(commands.apply, input), true); assert.equal(undoCount(studio), 1)
    assert.deepEqual(studio.getWorkbookModel(workbook.id).sheets[1].conditionalFormats, painted)
    api.getActiveWorkbook().undo(); assert.equal(studio.getWorkbookModel(workbook.id).sheets[1].conditionalFormats, undefined)
    api.getActiveWorkbook().redo(); assert.deepEqual(api.getActiveWorkbook().save().sheets[1].conditionalFormats, painted)
  }, workbook)
})

test('conditional-rule sampling remains immutable when source rules are edited during a continuous API session', () => withStudio(studio => {
  activate(studio, [range(0, 0)], commands.continuous)
  studio.commitWorkbook(studio.activeWorkbookId, '修改源条件格式', workbook => { workbook.sheets[0].conditionalFormats[0].style.fill = '#3344dd' }, { recalculate: false })
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), true)
  const rules = studio.editor.workbook.model.sheets[0].conditionalFormats
  assert.equal(rules.find(rule => rule.id === 'painter-cf-source').style.fill, '#3344dd')
  const copied = rules.find(rule => rule.ranges.some(area => area.startRow === 3 && area.startCol === 3))
  assert.equal(copied.style.fill, '#35a65f'); assert.equal(copied.formula1, '=D4>0'); assert.equal(studio.sheetFormatPainter.state.mode, 'continuous')
  assert.equal(undoCount(studio), 2)
}, conditionalSample()))

test('merge confirmation defers conditional replacement and acceptance commits merge, rules and formatting once', () => withStudio((studio, api) => {
  const before = cloneModel(studio.editor.workbook.model)
  activate(studio, [range(0, 0)]); assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)])), false)
  const request = studio.sheetFormatPainter.getConfirmation()
  assert.deepEqual(studio.editor.workbook.model, before); assert.equal(request.source.conditional.rules[0].id, 'painter-cf-source')
  assert.equal(studio.executeCommandSync(commands.confirm, { requestId: request.requestId, expectedRevision: request.baseRevision }), true)
  const painted = cloneModel(studio.editor.workbook.model)
  assert.equal(undoCount(studio), 1); assert.equal(painted.sheets[0].cells['3:4'].value, null)
  assert.notDeepEqual(painted.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
  api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
  assert.equal(studio.editor.workbook.model.sheets[0].cells['3:4'].value, '00123')
  api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, painted.sheets[0].conditionalFormats)
}, conditionalSample(true, true)))

test('a rejected merge value mutation rolls back conditional replacement with the rest of the API transaction', () => {
  const workbook = conditionalSample(true, true), sheet = workbook.sheets[0]
  sheet.cells['3:3'] = createCell(3, 3, 1); sheet.cells['3:4'] = createCell(3, 4, 2)
  sheet.dataValidations = [{ id: 'positive-target', type: 'custom', ranges: [range(3, 4)], formula1: '=E4>0', allowBlank: false, showErrorMessage: true, errorStyle: 'stop', error: '必须大于零。' }]
  withStudio(studio => {
    const before = cloneModel(studio.editor.workbook.model); activate(studio, [range(0, 0)])
    assert.throws(() => studio.executeCommandSync(commands.apply, paintInput(studio, [range(3, 3)], { mergeConfirmed: true })), /大于零/)
    assert.deepEqual(studio.editor.workbook.model, before); assert.equal(undoCount(studio), 0); assert.equal(studio.sheetFormatPainter.state.mode, 'once')
  }, workbook)
})

function conditionalBoundarySample() {
  const workbook = conditionalSample()
  workbook.sheets[0].conditionalFormats = [{ ...workbook.sheets[0].conditionalFormats[0], ranges: [range(1, 0, 4, 0)], formula1: '=A1>5' }]
  return workbook
}

test('partial-overflow API painting retains one conditional origin through undo, redo and internal reopen', () => withStudio((studio, api) => {
  const before = cloneModel(studio.editor.workbook.model), values = []
  studio.events.addEvent('SheetValueChanged', value => values.push(value))
  assert.equal(activate(studio, [range(1, 0, 4, 0)], commands.continuous), true)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(0, 0, 3, 0)])), true)
  const painted = cloneModel(studio.editor.workbook.model), entry = painted.sheets[0].conditionalFormats[0]
  assert.equal(painted.sheets[0].conditionalFormats.length, 1); assert.deepEqual(entry.formulaAnchor, { row: 1, col: 0 })
  assert.equal(conditionalFormatFormulaAt(entry.formula1, entry, 0, 0), '=#REF!>5'); assert.equal(conditionalFormatFormulaAt(entry.formula1, entry, 1, 0), '=A1>5')
  assert.deepEqual(contents(painted.sheets[0]), contents(before.sheets[0])); assert.deepEqual(values, []); assert.equal(undoCount(studio), 1)
  assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(0, 0, 3, 0)])), true); assert.equal(undoCount(studio), 1)
  api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
  api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, painted.sheets[0].conditionalFormats)
  const saved = api.getActiveWorkbook().save(), reopened = api.createWorkbook({ ...saved, id: 'painter-partial-ref-reopened' })
  assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].conditionalFormats, painted.sheets[0].conditionalFormats)
}, conditionalBoundarySample()))

test('cross-sheet partial-overflow API copying keeps an independent rule and the immutable original origin', () => {
  const workbook = conditionalBoundarySample(), second = cloneModel(createBlankWorkbook().sheets[0]); second.id = 'painter-ref-second'; workbook.sheets.push(second)
  withStudio((studio, api) => {
    const original = cloneModel(studio.editor.workbook.model.sheets[0].conditionalFormats)
    activate(studio, [range(1, 0, 4, 0)], commands.continuous); api.getActiveWorkbook().setActiveSheet(second.id)
    const input = { unitId: workbook.id, sheetId: second.id, ranges: [range(0, 0, 3, 0)] }
    assert.equal(studio.executeCommandSync(commands.apply, input), true)
    const copied = cloneModel(studio.editor.workbook.model.sheets[1].conditionalFormats)
    assert.notEqual(copied[0].id, 'painter-cf-source'); assert.deepEqual(copied[0].formulaAnchor, { row: 1, col: 0 })
    assert.equal(conditionalFormatFormulaAt(copied[0].formula1, copied[0], 2, 0), '=A2>5')
    assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, original)
    assert.equal(studio.executeCommandSync(commands.apply, input), true); assert.equal(undoCount(studio), 1)
    api.getActiveWorkbook().undo(); assert.equal(studio.editor.workbook.model.sheets[1].conditionalFormats, undefined)
    api.getActiveWorkbook().redo(); assert.deepEqual(api.getActiveWorkbook().save().sheets[1].conditionalFormats, copied)
  }, workbook)
})

test('merge confirmation defers a partially invalid conditional origin and commits it with content loss once', () => {
  const workbook = conditionalBoundarySample(); workbook.sheets[0].mergedRanges = [range(1, 0, 1, 1)]
  withStudio((studio, api) => {
    const before = cloneModel(studio.editor.workbook.model)
    activate(studio, [range(1, 0, 1, 1)]); assert.equal(studio.executeCommandSync(commands.apply, paintInput(studio, [range(0, 0)])), false)
    const request = studio.sheetFormatPainter.getConfirmation(); assert.deepEqual(studio.editor.workbook.model, before)
    assert.equal(studio.executeCommandSync(commands.confirm, { requestId: request.requestId, expectedRevision: request.baseRevision }), true)
    const painted = cloneModel(studio.editor.workbook.model), entry = painted.sheets[0].conditionalFormats[0]
    assert.deepEqual(entry.formulaAnchor, { row: 1, col: 0 }); assert.equal(conditionalFormatFormulaAt(entry.formula1, entry, 0, 0), '=#REF!>5')
    assert.equal(painted.sheets[0].cells['0:1'].value, null); assert.equal(undoCount(studio), 1)
    api.getActiveWorkbook().undo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, before.sheets[0].conditionalFormats)
    assert.deepEqual(contents(studio.editor.workbook.model.sheets[0]), contents(before.sheets[0]))
    api.getActiveWorkbook().redo(); assert.deepEqual(studio.editor.workbook.model.sheets[0].conditionalFormats, painted.sheets[0].conditionalFormats)
  }, workbook)
})
