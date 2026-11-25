import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred contracts exercise mount semantics and the actual content command, not rendering snapshots.
let server, domain, utils, ribbon, commands
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-structured-popmenus', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;[domain, utils, ribbon, commands] = await Promise.all(['domain/model', 'features/document/model/document-model',
    'components/document/document-structure-ribbon', 'features/document/commands/document-commands']
    .map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
afterAll(async () => { await server?.close() })
const copy = value => structuredClone(value)
const paragraph = (id, text = '') => ({ ...utils.createParagraph(text), id })
const modelWith = blocks => ({ ...domain.createBlankDocument(), id: 'popmenu-contract', blocks })
const selectionFor = (model, id) => { const target = utils.textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === id)
  assert.ok(target); return { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } } }
const leaves = items => items.flatMap(item => item.children ? leaves(item.children) : [item])
function executeItem(model, selection, item) {
  assert.equal(!!item.disabled, false); assert.ok(!item.hidden)
  const command = ribbon.documentStructureCommandFromMenu(item.commandId, item.params); assert.ok(command)
  return commands.executeDocumentCommand(model, selection, command, { documentId: model.id, userId: 'editor' })
}
const mounts = ['layout-below', 'content-insert', 'empty-paragraph', 'quick-insert', 'quick-edit', 'quick-insert-below', 'table-below']
for (const mount of mounts) test('mount ' + mount + ' publishes real structured content at its explicit target', () => {
  const below = ['layout-below', 'content-insert', 'quick-insert-below', 'table-below'].includes(mount)
  for (const type of ['quote', 'callout', 'code', ...(mount === 'quick-edit' ? [] : ['columns'])]) {
    let source = paragraph('target', mount === 'quick-insert' || mount === 'empty-paragraph' ? '' : 'A😀B')
    let focus = source.id
    if (mount === 'table-below') {
      const table = utils.createTable(1, 1); table.rows[0].cells[0].blocks = [paragraph('cell', 'retained table')]
      source = { ...paragraph('table'), type: 'table', runs: [], metadata: { table } }; focus = 'cell'
    }
    const model = modelWith([source, paragraph('tail', 'Outside')]), before = copy(model), selection = selectionFor(model, focus)
    const context = { paragraphId: focus, afterBlockId: source.id, disabled: false, codeAllowed: true }
    const choices = ribbon.documentStructuredMountedMenuItems(context, mount)
    const item = leaves(choices).find(value => type === 'columns' ? value.params?.columnCount === 3 : value.params?.blockType === type)
    const next = executeItem(model, selection, item), result = model.blocks[below ? 1 : 0]
    assert.equal(result.type, type)
    if (below) {
      assert.deepEqual(model.blocks[0], before.blocks[0]); assert.deepEqual(model.blocks[2], before.blocks[1])
      assert.notEqual(next.focus.cellBlockId ?? next.focus.blockId, focus)
    } else {
      const nested = utils.findDocumentBlock(model, focus)
      assert.ok(nested); assert.deepEqual(nested.runs, before.blocks[0].runs)
      assert.deepEqual(model.blocks[1], before.blocks[1]); assert.equal(next.focus.cellBlockId ?? next.focus.blockId, focus)
    }
    if (type === 'columns') assert.equal(result.columns.columns.length, 3)
    else assert.ok(result.children.length > 0)
    if (type === 'code') assert.deepEqual(result.code, { language: 'plain', showLineNumbers: true, tabSize: 4, wrap: true })
  }
})

test('mount order, visible conversion choices and 2–5-column submenu remain separately reviewable', () => {
  assert.deepEqual(ribbon.DOCUMENT_STRUCTURED_PARAGRAPH_MENU_MOUNTS, mounts)
  for (const mount of mounts) {
    const choices = ribbon.documentStructuredMountedMenuItems({ paragraphId: 'p', disabled: false }, mount)
    const quick = ['quick-insert', 'quick-edit', 'quick-insert-below'].includes(mount)
    assert.deepEqual(choices.filter(item => item.params?.blockType).map(item => item.params.blockType), quick ? ['code', 'quote', 'callout'] : ['callout', 'quote', 'code'])
    if (mount !== 'quick-edit') assert.deepEqual(choices[0].children.map(item => item.params.columnCount), [2, 3, 4, 5])
  }
  for (const type of ['quote', 'callout', 'code']) {
    const choices = ribbon.documentStructuredMountedMenuItems({ paragraphId: 'p', activeBlockId: 'wrapper', activeType: type, disabled: false, codeAllowed: true }, 'quick-edit')
    assert.deepEqual(choices.filter(item => !item.hidden && item.params?.blockType).map(item => item.params.blockType), type === 'callout' ? ['callout'] : ['code', 'quote'])
    assert.equal(choices.find(item => item.params?.blockType === type).activated, true)
  }
})

test('quick-edit converts the existing container and preserves children instead of nesting a second wrapper', () => {
  const source = { ...paragraph('code'), type: 'code', runs: [], children: [paragraph('inside', 'const value = 1')],
    code: { language: 'javascript', showLineNumbers: false, tabSize: 2, wrap: false } }
  const model = modelWith([source, paragraph('tail')]), selection = selectionFor(model, 'inside'), before = copy(source.children)
  const choices = ribbon.documentStructuredMountedMenuItems({ paragraphId: 'inside', activeBlockId: 'code', activeType: 'code', disabled: false, codeAllowed: true }, 'quick-edit')
  const item = choices.find(value => value.params?.blockType === 'quote')
  assert.deepEqual(ribbon.documentStructureCommandFromMenu(item.commandId, item.params), { id: 'doc.command.convert-structured-block', blockId: 'code', blockType: 'quote' })
  executeItem(model, selection, item)
  assert.equal(model.blocks.length, 2); assert.equal(model.blocks[0].id, 'code'); assert.equal(model.blocks[0].type, 'quote')
  const converted = model.blocks[0].children
  assert.deepEqual(converted.map(child => ({ id: child.id, runs: child.runs })), before.map(child => ({ id: child.id, runs: child.runs })))
  for (const [key, value] of Object.entries(before[0].style)) assert.deepEqual(converted[0].style[key], value)
  assert.deepEqual([converted[0].style.keepLines, converted[0].style.keepNext, converted[0].style.widowControl], [true, false, true])
  assert.deepEqual(model.blocks[0].metadata.documentStructuredTypeSources.code, source.code)
})

test('column quick-edit actions bind the column group when the nearest active block is nested Code', () => {
  const group = { version: 1, gap: 16, layout: 'fixed', responsive: 'shrink', columns: ['left', 'middle', 'right'].map((id, index) => ({ id, widthRatio: index + 1,
    blocks: index === 0 ? [{ ...paragraph('code'), type: 'code', runs: [], children: [paragraph('inside', 'source')],
      code: { language: 'rust', showLineNumbers: true, tabSize: 4, wrap: true } }] : [paragraph(id + '-p', id)] })) }
  const model = modelWith([{ ...paragraph('group'), type: 'columns', runs: [], columns: group }, paragraph('tail')]), selection = selectionFor(model, 'inside')
  const context = { paragraphId: 'inside', activeBlockId: 'code', activeType: 'code', columnGroupId: 'group', columns: group, columnId: 'left',
    inColumns: true, disabled: false, codeAllowed: true }
  const choices = ribbon.documentStructuredMountedMenuItems(context, 'quick-edit'), actions = choices.filter(item => item.commandId?.includes('column') || item.id === 'doc-context-structured-remove')
  assert.deepEqual(actions.map(item => item.id), ['doc-context-column-add', 'doc-context-column-delete', 'doc-context-structured-remove', 'doc-context-column-equal'])
  assert.ok(actions.every(item => item.params.blockId === 'group'))
  const source = copy(utils.findDocumentBlock(model, 'code'))
  executeItem(model, selection, actions.find(item => item.id === 'doc-context-column-equal'))
  assert.deepEqual(model.blocks[0].columns.columns.map(column => column.widthRatio), [1, 1, 1])
  assert.deepEqual(utils.findDocumentBlock(model, 'code'), source)
})

test('column creation depends on the actual insertion container and disabled controls cannot imply target capability', () => {
  const context = { paragraphId: 'inside', afterBlockId: 'group', activeBlockId: 'group', activeType: 'columns', inColumns: true,
    belowInColumns: false, disabled: false }
  const inPlace = ribbon.documentStructuredMountedMenuItems(context, 'quick-insert').find(item => item.id.endsWith('-columns'))
  const below = ribbon.documentStructuredMountedMenuItems(context, 'quick-insert-below').find(item => item.id.endsWith('-columns'))
  assert.equal(inPlace.disabled, true); assert.equal(below.disabled, false)
  assert.ok(below.children.every(item => item.params.afterBlockId === 'group'))
  for (const mount of mounts) assert.ok(leaves(ribbon.documentStructuredMountedMenuItems({ paragraphId: 'p', disabled: true }, mount)).every(item => item.disabled))
  const code = ribbon.documentStructuredMountedMenuItems({ paragraphId: 'p', activeBlockId: 'quote', activeType: 'quote', disabled: false, codeAllowed: false }, 'quick-edit').find(item => item.params?.blockType === 'code')
  assert.equal(code.disabled, true)
})

test('More parsing carries all four Code fields and refuses malformed patches before publishing a command', () => {
  const patch = { language: 'ruby', wrap: false, showLineNumbers: false, tabSize: 2 }
  const command = ribbon.documentStructureCommandFromMenu('doc.command.update-code-settings', { blockId: 'code', patch })
  assert.deepEqual(command, { id: 'doc.command.update-code-settings', blockId: 'code', patch })
  const model = modelWith([{ ...paragraph('code'), type: 'code', runs: [], children: [paragraph('inside', 'retained')],
    code: { language: 'plaintext', wrap: true, showLineNumbers: true, tabSize: 4 } }])
  commands.executeDocumentCommand(model, selectionFor(model, 'inside'), command, { documentId: model.id, userId: 'editor' })
  assert.deepEqual(model.blocks[0].code, patch); assert.equal(utils.runText(model.blocks[0].children[0].runs), 'retained')
  for (const invalid of [{}, { tabSize: 0 }, { tabSize: 17 }, { wrap: 'false' }, { showLineNumbers: 1 }, { language: '' }, { language: 'rust', unsupported: true }]) {
    assert.equal(ribbon.documentStructureCommandFromMenu('doc.command.update-code-settings', { blockId: 'code', patch: invalid }), undefined)
  }
  const font = ribbon.documentStructureCommandFromMenu('doc.command.set-code-font-family', { blockId: 'code', fontFamily: 'Menlo' })
  assert.deepEqual(font, { id: 'doc.command.set-code-font-family', blockId: 'code', fontFamily: 'Menlo' })
  assert.equal(ribbon.documentStructureCommandFromMenu('doc.command.set-code-font-family', { blockId: 'code', fontFamily: '' }), undefined)
})

test('mounted commands retain target validation and permission failure atomicity at execution', () => {
  const model = modelWith([paragraph('p', 'retained'), paragraph('tail')])
  model.permissionRules = [{ id: 'deny', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: [] }]
  const before = copy(model), selection = selectionFor(model, 'p')
  const item = ribbon.documentStructuredMountedMenuItems({ paragraphId: 'p', disabled: false }, 'quick-insert').find(value => value.params?.blockType === 'code')
  assert.throws(() => executeItem(model, selection, item)); assert.deepEqual(model, before)
  const missing = { ...item, params: { blockType: 'code', blockIds: ['missing'] } }
  assert.throws(() => executeItem(model, selection, missing)); assert.deepEqual(model, before)
})
