import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let api

before(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

after(async () => { await server?.close() })

function documentWithText(text = 'Hello World') {
  const model = api.createBlankDocument()
  model.id = 'content-controls-doc'
  model.blocks[0].id = 'paragraph-1'
  model.blocks[0].runs[0].text = text
  model.contentControls = []
  return model
}

function createDocument(text = 'Hello World') {
  const studio = new api.OfficeStudio({ document: documentWithText(text) })
  const document = api.FUniver.newAPI(studio).getActiveDocument()
  assert.ok(document)
  return { studio, document }
}

test('text content controls preserve their range and values through the public facade, undo, and redo', () => {
  const { studio, document } = createDocument()
  const control = document.addContentControl({ id: 'name', type: 'text', blockId: 'paragraph-1', startOffset: 6, endOffset: 11, title: 'Name', placeholder: 'Enter name' })
  assert.deepEqual(control.getRange(), { blockId: 'paragraph-1', startOffset: 6, endOffset: 11 })
  assert.equal(control.getText(), 'World')
  assert.equal(document.getContentControls().length, 1)

  control.setValue('Mira')
  assert.equal(document.getText(), 'Hello Mira')
  assert.equal(control.getText(), 'Mira')
  assert.deepEqual(control.getRange(), { blockId: 'paragraph-1', startOffset: 6, endOffset: 10 })
  assert.equal(document.undo(), true)
  assert.equal(document.getText(), 'Hello World')
  assert.equal(control.getText(), 'World')
  assert.equal(document.redo(), true)
  assert.equal(document.getText(), 'Hello Mira')
  assert.equal(document.save().contentControls[0].value, 'Mira')
  studio.dispose()
})

test('ordinary text insertion moves content-control anchors and content locks reject edits atomically', () => {
  const { studio, document } = createDocument('ABCD')
  const control = document.addContentControl({ id: 'locked', type: 'text', blockId: 'paragraph-1', startOffset: 1, endOffset: 3, lock: 'contentLocked' })
  assert.throws(() => document.insertText(2, 'x'), /content is locked/)
  assert.equal(document.getText(), 'ABCD')
  assert.equal(control.getText(), 'BC')

  control.update({ lock: 'unlocked' })
  document.insertText(0, '!')
  assert.equal(document.getText(), '!ABCD')
  assert.deepEqual(control.getRange(), { blockId: 'paragraph-1', startOffset: 2, endOffset: 4 })
  assert.equal(control.getText(), 'BC')
  studio.dispose()
})

test('checkboxes and date/list controls enforce type-specific values', () => {
  const { studio, document } = createDocument('___ 2026-01-01 A')
  const checkbox = document.addContentControl({ id: 'check', type: 'checkbox', blockId: 'paragraph-1', startOffset: 0, endOffset: 3 })
  checkbox.setChecked(true)
  assert.equal(checkbox.getInfo().checked, true)
  assert.equal(checkbox.getInfo().value, 'true')
  assert.throws(() => checkbox.setValue('yes'), /does not accept text values/)

  const date = document.addContentControl({ id: 'date', type: 'date', blockId: 'paragraph-1', startOffset: 4, endOffset: 14, dateValue: '2026-01-01' })
  assert.throws(() => date.setValue('2026-02-30'), /YYYY-MM-DD/)
  date.setValue('2026-02-28')
  assert.equal(date.getInfo().dateValue, '2026-02-28')

  const list = document.addContentControl({ id: 'list', type: 'dropDownList', blockId: 'paragraph-1', startOffset: 15, endOffset: 16, value: 'A', items: [{ displayText: 'Alpha', value: 'A' }, { displayText: 'Beta', value: 'B' }] })
  assert.throws(() => list.selectItem('C'), /not present in dropdown/)
  list.selectItem('B')
  assert.equal(list.getText(), 'B')
  studio.dispose()
})

test('invalid ranges and partially overlapping controls leave the document unchanged', () => {
  const { studio, document } = createDocument('abcdefghij')
  const first = document.addContentControl({ id: 'first', type: 'text', blockId: 'paragraph-1', startOffset: 1, endOffset: 5 })
  const before = structuredClone(document.getModel())
  assert.throws(() => document.addContentControl({ id: 'crossing', type: 'text', blockId: 'paragraph-1', startOffset: 4, endOffset: 7 }), /overlap/)
  assert.throws(() => document.addContentControl({ id: 'outside', type: 'text', blockId: 'paragraph-1', startOffset: 7, endOffset: 11 }), /exceeds paragraph length/)
  assert.deepEqual(document.getModel().contentControls, before.contentControls)
  assert.equal(first.getText(), 'bcde')
  studio.dispose()
})

test('wrapper locks prevent removal, and removal is undoable', () => {
  const { studio, document } = createDocument('abc')
  const control = document.addContentControl({ id: 'locked-wrapper', type: 'text', blockId: 'paragraph-1', startOffset: 0, endOffset: 2, lock: 'sdtLocked' })
  assert.throws(() => control.remove(), /wrapper is locked/)
  control.update({ lock: 'unlocked' })
  assert.equal(control.remove(), true)
  assert.equal(document.getContentControl('locked-wrapper'), null)
  assert.equal(document.undo(), true)
  assert.equal(document.getContentControl('locked-wrapper')?.getText(), 'ab')
  studio.dispose()
})

test('group ranges retain explicit parent structure and can be removed as a tree', () => {
  const { studio, document } = createDocument('ABCD')
  const parent = document.addContentControl({ id: 'group', type: 'group', blockId: 'paragraph-1', startOffset: 0, endOffset: 4 })
  const child = document.addContentControl({ id: 'child', type: 'text', blockId: 'paragraph-1', startOffset: 1, endOffset: 3, parentId: 'group' })
  assert.equal(child.getInfo().parentId, parent.getId())
  assert.equal(parent.remove(), true)
  assert.deepEqual(document.getContentControls(), [])
  studio.dispose()
})
