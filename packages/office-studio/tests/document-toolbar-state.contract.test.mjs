import assert from 'node:assert/strict'
import { afterAll, test } from 'vitest'
import { createServer } from 'vite'
const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
afterAll(() => vite.close())
const { createBlankDocument } = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
const { createParagraph, createRun, collapsedSelection } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { clampDocumentZoom, documentCharacterCount, documentSelectionCharacterCount, documentToolbarState } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-toolbar-state.ts')
const { createDocumentRibbon, documentCommandFromRibbon } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-ribbon.ts')
const { normalizeDocument: importDocument } = await vite.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts')
const { executeDocumentCommand } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
const { createRibbonMenuModel } = await vite.ssrLoadModule('/packages/office-studio/src/platform/ribbon/ribbon-command-menu.ts')
const { layoutDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
const { documentOrientedPage, documentPageSettingsError, documentPaperSize } = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-page-settings.ts')

test('a DOCX title retains its font slots from normalization through caret state, glyphs and range formatting', () => {
  const model = importDocument({ fileName: 'fonts.docx', format: 'docx', payload: { data: { blocks: [{ id: 'p', type: 'section_heading', details: { styleId: 'Title', inlines: [{ type: 'text', text: 'A中', style: { fontFamilyAscii: 'Arial', fontFamilyEastAsia: 'SimSun' } }] } }] } } })
  const caret = offset => collapsedSelection({ blockId: 'p', offset })
  assert.equal(documentToolbarState(model, caret(1)).blockStyle, 'title')
  assert.equal(documentToolbarState(model, caret(1)).marks.fontFamily, 'Arial')
  assert.equal(documentToolbarState(model, caret(2)).marks.fontFamily, 'SimSun')
  const glyphs = layoutDocument(model, { measure: () => 8 }).lines.flatMap(line => line.glyphs)
  assert.match(glyphs.find(g => g.text === 'A').font, /px Arial,/)
  assert.match(glyphs.find(g => g.text === '中').font, /px SimSun,/)
  executeDocumentCommand(model, { anchor: caret(1).anchor, focus: caret(2).focus }, { id: 'doc.command.format-text', patch: { fontFamily: 'Kai' } })
  assert.deepEqual(model.blocks[0].runs.map(r => [r.text, r.fontFamily, r.eastAsiaFontFamily]), [['A', 'Arial', 'SimSun'], ['中', 'Kai', 'Kai']])
  executeDocumentCommand(model, caret(2), { id: 'doc.command.set-block-type', type: 'quote' })
  assert.equal(documentToolbarState(model, caret(2)).blockStyle, 'quote')
})

test('toolbar reports uniform and mixed formats from the actual selected run slices', () => {
  const model = createBlankDocument()
  const first = createParagraph()
  first.runs = [createRun('ab', { bold: true, fontFamily: 'Arial' }), createRun('cd', { fontFamily: 'Georgia' })]
  model.blocks = [first]
  const mixed = documentToolbarState(model, { anchor: { blockId: first.id, offset: 1 }, focus: { blockId: first.id, offset: 3 } })
  assert.equal(mixed.marks.bold, undefined)
  assert.equal(mixed.marks.fontFamily, undefined)
  assert.ok(mixed.mixedMarks.includes('bold'))
  assert.ok(mixed.mixedMarks.includes('fontFamily'))
  const uniform = documentToolbarState(model, { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: first.id, offset: 2 } })
  assert.equal(uniform.marks.bold, true)
  assert.equal(uniform.marks.fontFamily, 'Arial')
})

test('collapsed pending marks affect only the caret and backwards ranges resolve the same state', () => {
  const model = createBlankDocument()
  const paragraph = createParagraph('abcd', { fontSize: 14 })
  model.blocks = [paragraph]
  const caret = documentToolbarState(model, collapsedSelection({ blockId: paragraph.id, offset: 2 }), { bold: true })
  assert.equal(caret.marks.bold, true)
  assert.equal(caret.marks.fontSize, 14)
  const range = documentToolbarState(model, { anchor: { blockId: paragraph.id, offset: 3 }, focus: { blockId: paragraph.id, offset: 1 } }, { fontSize: 28 })
  assert.equal(range.marks.fontSize, 14)
  assert.equal(range.hasSelection, true)
})

test('subscript and superscript expose uniform and mixed selection state', () => {
  const model = createBlankDocument()
  const paragraph = createParagraph()
  paragraph.runs = [createRun('H', { verticalAlign: 'baseline' }), createRun('2', { verticalAlign: 'subscript' }), createRun('O', { verticalAlign: 'superscript' })]
  model.blocks = [paragraph]

  const subscript = documentToolbarState(model, { anchor: { blockId: paragraph.id, offset: 1 }, focus: { blockId: paragraph.id, offset: 2 } })
  assert.equal(subscript.marks.verticalAlign, 'subscript')
  assert.deepEqual(subscript.mixedMarks, [])
  const mixed = documentToolbarState(model, { anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 2 } })
  assert.equal(mixed.marks.verticalAlign, undefined)
  assert.ok(mixed.mixedMarks.includes('verticalAlign'))
})

test('a range ending at the next paragraph offset zero excludes its toolbar styles', () => {
  const model = createBlankDocument()
  const first = createParagraph('first', { bold: true })
  first.type = 'heading'; first.level = 2
  const second = createParagraph('second')
  second.style.align = 'right'
  model.blocks = [first, second]
  const state = documentToolbarState(model, { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: second.id, offset: 0 } })
  assert.equal(state.marks.bold, true)
  assert.equal(state.paragraph.align, 'left')
  assert.equal(state.blockStyle, 'H2')
})

test('the named-style selector follows the ordered starting paragraph while mixed block and character state stay independent', () => {
  const model = createBlankDocument(), first = createParagraph('Title'), last = createParagraph('Heading')
  executeDocumentCommand({ ...model, blocks: [first] }, collapsedSelection({ blockId: first.id, offset: 0 }), { id: 'doc.command.set-paragraph-named-style', namedStyle: 'title' })
  executeDocumentCommand({ ...model, blocks: [last] }, collapsedSelection({ blockId: last.id, offset: 0 }), { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading2' })
  model.blocks = [first, last]
  const selection = { anchor: { blockId: first.id, offset: 1 }, focus: { blockId: last.id, offset: 2 } }
  const state = documentToolbarState(model, selection)
  assert.equal(state.namedStyle, 'title')
  assert.equal(state.blockStyle, undefined)
  assert.ok(state.mixedMarks.includes('fontSize'))
  assert.deepEqual(documentToolbarState(model, { anchor: selection.focus, focus: selection.anchor }), state)
  const item = createDocumentRibbon({ state, canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false,
    paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true }).flatMap(tab => tab.groups.flatMap(group => group.items)).find(item => item.id === 'doc-block-style')
  assert.equal(item.value, 'title')
  assert.deepEqual(item.options.map(option => option.params.namedStyle), ['normal', 'title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5'])
})

test('checklist and numbered-heading controls retain their own state after applying a named style', () => {
  const model = createBlankDocument(), checklist = { ...createParagraph('Task'), type: 'checklist', level: 1, checklist: { listId: 'tasks', checked: true } }
  const numbered = { ...createParagraph('Item'), type: 'numbered', level: 2, list: { id: 'numbers', preset: 'ORDER_LIST' } }
  model.blocks = [checklist, numbered]
  for (const block of model.blocks) executeDocumentCommand(model, collapsedSelection({ blockId: block.id, offset: 0 }),
    { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading3' })
  const checklistState = documentToolbarState(model, collapsedSelection({ blockId: checklist.id, offset: 0 }))
  assert.equal(checklistState.namedStyle, 'heading3')
  assert.equal(checklistState.blockStyle, 'checklist')
  const numberedState = documentToolbarState(model, collapsedSelection({ blockId: numbered.id, offset: 0 }))
  assert.equal(numberedState.namedStyle, 'heading3')
  assert.equal(numberedState.blockStyle, 'H3')
  assert.equal(numberedState.listPreset, 'ORDER_LIST')
  const items = createDocumentRibbon({ state: numberedState, canUndo: false, canRedo: false, outline: false, focus: false,
    fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true }).flatMap(tab => tab.groups.flatMap(group => group.items))
  assert.equal(items.find(item => item.id === 'doc-numbered').activated, true)
  assert.equal(items.find(item => item.id === 'doc-block-style').value, 'H3')
})

test('paragraph state includes mixed indentation, spacing units and wrap flags across the real selection', () => {
  const model = createBlankDocument()
  const first = createParagraph('first'), second = createParagraph('second')
  first.style = { ...first.style, indentStart: 24, indentEnd: 8, indentFirstLine: 20, hanging: 4, spaceBefore: 6, lineHeight: 2, spacingRule: 'auto', wordWrap: false }
  second.style = { ...second.style, indentStart: 48, indentEnd: 8, indentFirstLine: 40, hanging: 4, spaceBefore: 12, lineHeight: 2, spacingRule: 'exact', wordWrap: true }
  model.blocks = [first, second]
  const selection = { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: second.id, offset: 6 } }
  const state = documentToolbarState(model, selection)
  assert.equal(state.paragraph.indentEnd, 8)
  assert.equal(state.paragraph.hanging, 4)
  assert.equal(state.paragraph.indentStart, undefined)
  assert.equal(state.paragraph.indentFirstLine, undefined)
  assert.equal(state.paragraph.lineHeight, undefined)
  assert.deepEqual([...state.mixedParagraph].sort(), ['indentStart', 'indentFirstLine', 'spaceBefore', 'lineHeight', 'spacingRule', 'wordWrap'].sort())
  assert.deepEqual(documentToolbarState(model, { anchor: selection.focus, focus: selection.anchor }), state)
})

test('missing optional paragraph values and explicit display defaults are uniform, including legacy indentation', () => {
  const model = createBlankDocument()
  const first = createParagraph('first'), second = createParagraph('second')
  first.style.indent = 1
  second.style = { ...second.style, indent: 1, indentStart: 24, indentEnd: 0, indentFirstLine: 0, hanging: 0, spacingRule: 'auto', wordWrap: false, snapToGrid: false }
  model.blocks = [first, second]
  const state = documentToolbarState(model, { anchor: { blockId: first.id, offset: 0 }, focus: { blockId: second.id, offset: 6 } })
  assert.deepEqual(state.mixedParagraph, [])
  assert.equal(state.paragraph.indentStart, 24)
  assert.equal(state.paragraph.spacingRule, 'auto')
  assert.equal(state.paragraph.wordWrap, false)
})

test('character statistics count selected graphemes and ignore whitespace', () => {
  assert.equal(documentCharacterCount('中文 e\u0301 👨‍👩‍👧‍👦\n'), 4)
  const model = createBlankDocument()
  const block = createParagraph('甲 乙丙')
  model.blocks = [block]
  assert.equal(documentSelectionCharacterCount(model, { anchor: { blockId: block.id, offset: 0 }, focus: { blockId: block.id, offset: 3 } }), 2)
  assert.equal(documentSelectionCharacterCount(model, collapsedSelection({ blockId: block.id, offset: 0 })), 0)
})

test('zoom clamps invalid, fractional and out of range values at the shared boundary', () => {
  assert.equal(clampDocumentZoom(-10), 10)
  assert.equal(clampDocumentZoom(1000), 400)
  assert.equal(clampDocumentZoom(125.4), 125)
  assert.equal(clampDocumentZoom(Number.NaN), 100)
})

test('symbol and emoji roots share the fixed value command and preserve whole UTF-16 strings with the legacy string adapter', () => {
  const model = createBlankDocument(), state = documentToolbarState(model, collapsedSelection({ blockId: model.blocks[0].id, offset: 0 }))
  const context = { state, canUndo: false, canRedo: false, outline: false, focus: false, paper: 'A4', orientation: 'portrait', zoom: 100 }
  const items = createDocumentRibbon(context).flatMap(tab => tab.groups.flatMap(group => group.items))
  for (const [id, mode, row] of [['doc-emoji', 'emoji', 1], ['doc-symbol', 'symbol', 2]]) {
    const item = items.find(item => item.id === id), value = mode === 'emoji' ? '👨‍👩‍👧‍👦' : '|a|e\u0301'
    assert.equal(item.popupCommandId, 'doc.command.insert-special-character')
    assert.deepEqual(item.popupProps, { mode })
    assert.deepEqual(item.gridLayout, { row, column: 3 })
    assert.deepEqual(documentCommandFromRibbon(item.popupCommandId, { value }), { id: 'doc.command.insert-text', text: value })
  }
  for (const params of ['', { value: '' }, { value: 1 }, { text: '©' }, null, ['©']]) assert.equal(documentCommandFromRibbon('doc.command.insert-special-character', params), undefined)
  assert.deepEqual(documentCommandFromRibbon('doc.command.insert-special-character', { value: ' ' }), { id: 'doc.command.insert-text', text: ' ' })
  assert.deepEqual(documentCommandFromRibbon('doc.command.insert-special-character', { value: '\u00a0' }), { id: 'doc.command.insert-text', text: '\u00a0' })
  assert.deepEqual(documentCommandFromRibbon('doc.command.insert-character', '©'), { id: 'doc.command.insert-text', text: '©' })
  assert.equal(documentCommandFromRibbon('doc.command.insert-character', ''), undefined)
  const disabled = createDocumentRibbon({ ...context, editingDisabled: true }).flatMap(tab => tab.groups.flatMap(group => group.items))
  assert.ok(disabled.filter(item => item.id === 'doc-emoji' || item.id === 'doc-symbol').every(item => item.disabled))
})

test('ribbon schemas derive current values, mixed toggle behavior and dispatch validated Canvas commands', () => {
  const state = { marks: { fontSize: 18 }, mixedMarks: ['bold', 'fontFamily'], paragraph: { align: 'right', indent: 0 }, blockStyle: 'H2', hasSelection: true }
  const tabs = createDocumentRibbon({ state, canUndo: true, canRedo: false, outline: true, focus: false, paper: 'A4', orientation: 'portrait', zoom: 125, contentControlsAvailable: true })
  const items = tabs.flatMap((tab) => tab.groups.flatMap((group) => group.items))
  assert.deepEqual(tabs.map((tab) => tab.id), ['home', 'insert', 'view'])
  const divider = items.find(item => item.id === 'doc-divider')
  assert.equal(divider.commandId, 'doc.command.horizontal-line')
  const dividerCommand = documentCommandFromRibbon(divider.commandId, divider.params)
  assert.deepEqual(dividerCommand, { id: 'doc.command.horizontal-line' })
  const dividerModel = createBlankDocument()
  executeDocumentCommand(dividerModel, collapsedSelection({ blockId: dividerModel.blocks[0].id, offset: 0 }), dividerCommand)
  assert.equal(dividerModel.blocks.some(block => block.type === 'divider'), true)
  for (const id of ['doc.command.insert-divider', 'doc.command.insert-horizontal-line-bellow']) assert.deepEqual(documentCommandFromRibbon(id, undefined), { id })
  const disabledDivider = createDocumentRibbon({ state, editingDisabled: true, canUndo: false, canRedo: false, outline: false, focus: false, paper: 'A4', orientation: 'portrait', zoom: 100 })
    .flatMap(tab => tab.groups.flatMap(group => group.items)).find(item => item.id === 'doc-divider')
  assert.equal(disabledDivider.disabled, true)
  const breakMenu = items.find((item) => item.id === 'doc-breaks')
  assert.ok(breakMenu)
  const columnBreak = breakMenu.children.find((item) => item.id === 'doc-column-break')
  assert.deepEqual(documentCommandFromRibbon(columnBreak.commandId, columnBreak.params), { id: 'doc.command.insert-column-break' })
  const sectionMenu = breakMenu.children.find((item) => item.id === 'doc-section-breaks')
  assert.deepEqual(sectionMenu.children.map((item) => item.params.type), ['nextPage', 'continuous', 'nextColumn', 'evenPage', 'oddPage'])
  for (const item of sectionMenu.children) {
    assert.deepEqual(documentCommandFromRibbon(item.commandId, item.params), { id: 'doc.command.insert-section-break', type: item.params.type })
  }
  assert.equal(documentCommandFromRibbon('doc.command.insert-section-break', { type: 'unexpected' }), undefined)
  const font = items.find((item) => item.id === 'doc-font-family')
  assert.equal(font.value, null)
  assert.equal(font.placeholder, '混合字体')
  const bold = items.find((item) => item.id === 'doc-bold')
  const alignMenu = createRibbonMenuModel(items.find(item => item.id === 'doc-align').children)
  assert.deepEqual(alignMenu.selectedKeys, ['doc-align-right'])
  assert.deepEqual(alignMenu.actions.get('doc-align-center').item.params, { patch: { align: 'center' } })
  assert.equal(bold.activated, false)
  assert.deepEqual(documentCommandFromRibbon(bold.commandId, bold.params), { id: 'doc.command.format-text', patch: { bold: true } })
  const superscript = items.find((item) => item.id === 'doc-superscript')
  assert.equal(superscript.activated, false)
  assert.deepEqual(documentCommandFromRibbon(superscript.commandId, superscript.params), { id: 'doc.command.format-text', patch: { verticalAlign: 'superscript' } })
  const activeSubscript = createDocumentRibbon({ state: { ...state, marks: { verticalAlign: 'subscript' } }, canUndo: true, canRedo: false, outline: true, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 125, contentControlsAvailable: true })
    .flatMap((tab) => tab.groups.flatMap((group) => group.items)).find((item) => item.id === 'doc-subscript')
  assert.equal(activeSubscript.activated, true)
  assert.deepEqual(documentCommandFromRibbon(activeSubscript.commandId, activeSubscript.params), { id: 'doc.command.format-text', patch: { verticalAlign: 'baseline' } })
  assert.deepEqual(documentCommandFromRibbon(font.commandId, font.valueParams('Georgia')), { id: 'doc.command.format-text', patch: { fontFamily: 'Georgia' } })
  assert.equal(items.find((item) => item.id === 'doc-redo').disabled, true)
  const viewItems = tabs.find((tab) => tab.id === 'view').groups.flatMap((group) => group.items)
  assert.ok(viewItems.find((item) => item.id === 'doc-page-settings'))
  assert.ok(viewItems.find((item) => item.id === 'doc-fullscreen'))
  assert.equal(tabs.some((tab) => tab.id === 'layout'), false)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: { fontSize: NaN } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-text', { patch: { bold: 'false' } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.format-paragraph', { patch: { align: 'diagonal' } }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.set-block-type', { type: 'heading', level: 100 }), undefined)
  assert.deepEqual(items.find(item => item.id === 'doc-block-style').options.map(o => o.value), ['normal', 'title', 'subtitle', 'H1', 'H2', 'H3', 'H4', 'H5'])
  const quote = items.find((item) => item.id === 'doc-extra-styles').children.find(item => item.id === 'doc-quote')
  assert.deepEqual(documentCommandFromRibbon('doc.command.set-block-type', quote.params), { id: 'doc.command.set-block-type', type: 'quote', level: undefined })
  const checklist = items.find((item) => item.id === 'doc-checklist')
  assert.equal(checklist.commandId, 'doc.command.set-block-type')
  assert.deepEqual(documentCommandFromRibbon(checklist.commandId, checklist.params), { id: 'doc.command.set-block-type', type: 'checklist', level: undefined })
})

test('break menus follow body, auxiliary, table and Code capabilities while preserving their commands and view operations', () => {
  const model = createBlankDocument()
  const body = collapsedSelection({ blockId: model.blocks[0].id, offset: 0 })
  const state = documentToolbarState(model, body)
  const tableModel = createBlankDocument()
  const tableSelection = executeDocumentCommand(tableModel, collapsedSelection({ blockId: tableModel.blocks[0].id, offset: 0 }),
    { id: 'doc.command.create-table', rowCount: 2, colCount: 2 })
  const tableState = documentToolbarState(tableModel, tableSelection)
  assert.ok(tableState.table)
  const context = { state, editingDisabled: false, canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false,
    paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true, mode: 'traditional' }
  for (const [name, patch, disabled, hidden, tableDisabled] of [
    ['body', {}, false, false, false],
    ['read-only', { editingDisabled: true }, true, false, false],
    ['header/footer', { editingHeaderFooter: true }, true, false, false],
    ['auxiliary blocks', { editingAuxiliaryContent: true }, true, false, false],
    ['table cell', { state: tableState }, false, false, true],
    ['Code', { structuredBlock: { blockId: model.blocks[0].id, type: 'code', disabled: false } }, true, false, false],
    ['modern', { mode: 'modern' }, false, true, false],
  ]) {
    const tabs = createDocumentRibbon({ ...context, ...patch })
    const item = tabs.find(tab => tab.id === 'insert').groups.flatMap(group => group.items).find(item => item.id === 'doc-breaks')
    assert.equal(Boolean(item.disabled), disabled, name)
    assert.equal(Boolean(item.hidden), hidden, name)
    assert.deepEqual(item.children.slice(0, 2).map(child => child.commandId), ['doc.command.insert-page-break', 'doc.command.insert-column-break'])
    assert.equal(Boolean(item.children[0].disabled), false, name)
    assert.equal(item.children[1].disabled, tableDisabled, name)
    const sections = item.children.find(child => child.id === 'doc-section-breaks')
    assert.equal(sections.disabled, tableDisabled, name)
    assert.ok(sections.children.every(child => child.disabled === tableDisabled), name)
    assert.deepEqual(sections.children.map(child => documentCommandFromRibbon(child.commandId, child.params)),
      ['nextPage', 'continuous', 'nextColumn', 'evenPage', 'oddPage'].map(type => ({ id: 'doc.command.insert-section-break', type })))
    const view = tabs.find(tab => tab.id === 'view').groups.flatMap(group => group.items)
    assert.deepEqual(view.map(item => item.id), ['doc-page-settings', 'doc-zoom', 'doc-fullscreen'])
    assert.equal(Boolean(view.find(item => item.id === 'doc-zoom').disabled), false)
    assert.equal(Boolean(view.find(item => item.id === 'doc-fullscreen').disabled), false)
  }
})

test('paragraph Ribbon actions dispatch relative indentation and change fixed spacing to a multiplier explicitly', () => {
  const context = {
    state: { marks: {}, mixedMarks: [], paragraph: { spacingRule: 'exact', lineHeight: 24 }, mixedParagraph: ['indentFirstLine'], hasSelection: true },
    indentation: { canDecrease: true, canIncrease: false },
    canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false,
  }
  const items = createDocumentRibbon(context).find(tab => tab.id === 'home').groups.flatMap(group => group.items)
  const decrease = items.find(item => item.id === 'doc-indent-less'), increase = items.find(item => item.id === 'doc-indent-more')
  assert.equal(decrease.disabled, false)
  assert.equal(increase.disabled, true)
  assert.deepEqual(documentCommandFromRibbon(decrease.commandId, decrease.params), { id: 'doc.command.change-paragraph-indent', direction: -1 })
  assert.deepEqual(documentCommandFromRibbon(increase.commandId, increase.params), { id: 'doc.command.change-paragraph-indent', direction: 1 })
  const lineSpacing = items.find(item => item.id === 'doc-line-height')
  assert.equal(lineSpacing.value, null)
  assert.match(lineSpacing.placeholder, /24 px/)
  assert.deepEqual(documentCommandFromRibbon(lineSpacing.commandId, lineSpacing.valueParams(1.5)), { id: 'doc.command.format-paragraph', patch: { spacingRule: 'auto', lineHeight: 1.5 } })
  const denied = createDocumentRibbon({ ...context, editingDisabled: true }).find(tab => tab.id === 'home').groups.flatMap(group => group.items)
  for (const id of ['doc-indent-less', 'doc-indent-more', 'doc-line-height', 'doc-paragraph-spacing']) assert.equal(denied.find(item => item.id === id).disabled, true)
  assert.equal(documentCommandFromRibbon('doc.command.change-paragraph-indent', { direction: 0 }), undefined)
  assert.equal(documentCommandFromRibbon('doc.command.change-paragraph-indent', { direction: '1' }), undefined)
  assert.deepEqual(documentCommandFromRibbon('doc.command.format-paragraph', { patch: { indentStart: 12, indentEnd: 5, indentFirstLine: 20, hanging: 4, spacingRule: 'exact', lineHeight: 18, wordWrap: true } }), { id: 'doc.command.format-paragraph', patch: { indentStart: 12, indentEnd: 5, indentFirstLine: 20, hanging: 4, spacingRule: 'exact', lineHeight: 18, wordWrap: true } })
  for (const patch of [{ custom: 1 }, { indentStart: 101 }, { hanging: NaN }, { wordWrap: 1 }, { spacingRule: 'auto', lineHeight: 6 }]) assert.equal(documentCommandFromRibbon('doc.command.format-paragraph', { patch }), undefined)
})

test('custom imported page geometry is preserved by orientation and identified accurately', () => {
  const page = { width: 960, height: 1200, marginTop: 45, marginRight: 55, marginBottom: 65, marginLeft: 75 }
  assert.equal(documentPaperSize(page), 'Custom')
  assert.deepEqual(documentOrientedPage(page, 'landscape', 'Custom'), { ...page, width: 1200, height: 960 })
  assert.equal(documentPaperSize({ ...page, width: 1123, height: 794 }), 'A4')
  assert.equal(documentPaperSize({ ...page, width: 1056, height: 816 }), 'Letter')
  assert.equal(documentPageSettingsError(page), undefined)
})

test('page settings reject invalid dimensions and margins before a document mutation', () => {
  const page = createBlankDocument().page
  assert.match(documentPageSettingsError({ ...page, width: NaN }), /有效数字/)
  assert.match(documentPageSettingsError({ ...page, width: 50 }), /宽度/)
  assert.match(documentPageSettingsError({ ...page, marginTop: -1 }), /负数/)
  assert.match(documentPageSettingsError({ ...page, marginLeft: page.width }), /之和/)
  assert.equal(documentPageSettingsError({ ...page, marginLeft: page.width * 0.45 }), undefined)
  assert.match(documentPageSettingsError({ ...page, marginLeft: page.width * 0.51 }), /一半/)
})
