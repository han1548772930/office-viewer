import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let api

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

afterAll(async () => { await server?.close() })

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
  const document = api.FOffice.newAPI(studio).getActiveDocument()
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

test('placeholder reset restores placeholder text and remains undoable', () => {
  const { studio, document } = createDocument('Mira')
  const control = document.addContentControl({
    id: 'placeholder', type: 'text', blockId: 'paragraph-1', startOffset: 0, endOffset: 4,
    placeholder: 'Enter name', showingPlaceholder: false,
  })

  assert.equal(control.resetPlaceholder(), true)
  assert.equal(control.getText(), 'Enter name')
  assert.equal(control.getInfo().showingPlaceholder, true)
  assert.equal(document.undo(), true)
  assert.equal(control.getText(), 'Mira')
  assert.equal(document.redo(), true)
  assert.equal(control.getText(), 'Enter name')
  studio.dispose()
})

test('repeating-section item duplication copies nested controls and deletion shifts following ranges', () => {
  const { studio, document } = createDocument('ABC')
  const section = document.addContentControl({
    id: 'section', type: 'repeatingSection', blockId: 'paragraph-1', startOffset: 0, endOffset: 1,
    repeatingSection: { doNotAllowInsertDelete: false },
  })
  const first = document.getContentControls().find(control => control.getType() === 'repeatingSectionItem')
  assert.ok(first)
  document.addContentControl({
    id: 'child-1', type: 'text', blockId: 'paragraph-1', startOffset: 0, endOffset: 1, parentId: first.id,
  })
  const after = document.addContentControl({
    id: 'after', type: 'text', blockId: 'paragraph-1', startOffset: 2, endOffset: 3,
  })

  assert.equal(first.duplicateRepeatingItem(), true)
  assert.equal(document.getText(), 'AABC')
  const duplicatedItem = document.getContentControls().find(control => control.getType() === 'repeatingSectionItem'
    && control.getRange().startOffset === 1)
  assert.ok(duplicatedItem)
  assert.equal(duplicatedItem.getInfo().parentId, section.id)
  assert.equal(document.getContentControls().some(control => control.getInfo().parentId === duplicatedItem.id), true)
  assert.deepEqual(after.getRange(), { blockId: 'paragraph-1', startOffset: 3, endOffset: 4 })
  assert.equal(duplicatedItem.deleteRepeatingItem(), true)
  assert.equal(document.getText(), 'ABC')
  assert.deepEqual(section.getRange(), { blockId: 'paragraph-1', startOffset: 0, endOffset: 1 })
  assert.deepEqual(after.getRange(), { blockId: 'paragraph-1', startOffset: 2, endOffset: 3 })
  studio.dispose()
})

test('picture content control replacement inserts a persistent image block and updates it in place', () => {
  const { studio, document } = createDocument('Before Choose picture After')
  const picture = document.addContentControl({
    id: 'picture', type: 'picture', blockId: 'paragraph-1', startOffset: 7, endOffset: 21, placeholder: 'Choose picture',
  })
  const suffix = document.addContentControl({
    id: 'suffix', type: 'text', blockId: 'paragraph-1', startOffset: 22, endOffset: 27,
  })

  assert.equal(picture.replacePicture({ src: 'data:image/png;base64,AA==', alt: 'First', width: 40, height: 20 }), true)
  const pictureBlockId = picture.getInfo().metadata.pictureBlockId
  let imageBlocks = document.getModel().blocks.filter(block => block.type === 'image')
  assert.equal(imageBlocks.length, 1)
  assert.equal(imageBlocks[0].id, pictureBlockId)
  assert.equal(imageBlocks[0].metadata.image.alt, 'First')
  assert.deepEqual(document.getModel().blocks.map(block => block.runs?.map(run => run.text).join('')),
    ['Before ', '', ' After'])
  assert.equal(suffix.getText(), 'After')
  assert.equal(suffix.getRange().blockId, document.getModel().blocks[2].id)
  assert.equal(document.undo(), true)
  assert.equal(document.getModel().blocks.some(block => block.type === 'image'), false)
  assert.equal(document.redo(), true)

  assert.equal(picture.replacePicture({ src: 'data:image/png;base64,AQ==', alt: 'Second', width: 80, height: 40 }), true)
  imageBlocks = document.getModel().blocks.filter(block => block.type === 'image')
  assert.equal(imageBlocks.length, 1)
  assert.equal(imageBlocks[0].id, pictureBlockId)
  assert.equal(imageBlocks[0].metadata.image.alt, 'Second')
  assert.equal(suffix.getText(), 'After')
  assert.equal(picture.resetPlaceholder(), true)
  assert.equal(document.getModel().blocks.some(block => block.type === 'image'), false)
  assert.equal(document.getText(), 'Before Choose picture After')
  assert.equal(suffix.getText(), 'After')
  assert.equal(picture.getInfo().showingPlaceholder, true)
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

test('adding a repeating section creates its initial item in the same undoable mutation', () => {
  const { studio, document } = createDocument('Repeat me')
  const section = document.addContentControl({
    id: 'repeating-section', type: 'repeatingSection', blockId: 'paragraph-1', startOffset: 0, endOffset: 9,
    placement: 'block', repeatingSection: { doNotAllowInsertDelete: false },
  })
  const items = document.getContentControls().filter(control => control.getType() === 'repeatingSectionItem')
  assert.equal(items.length, 1)
  assert.equal(items[0].getInfo().parentId, section.id)
  assert.deepEqual(items[0].getRange(), section.getRange())
  assert.equal(document.undo(), true)
  assert.equal(document.getContentControls().length, 0)
  assert.equal(document.redo(), true)
  assert.equal(document.getContentControls().filter(control => control.getType() === 'repeatingSectionItem').length, 1)
  studio.dispose()
})
