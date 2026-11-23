import assert from 'node:assert/strict'
import { afterEach, test } from 'vitest'
import { sheetSystemClipboard, sheetSystemClipboardAvailable } from '../src/components/spreadsheet/sheet-system-clipboard.ts'

// Draft only: no browser clipboard permissions have been exercised in this batch.
const originals = new Map(['navigator', 'ClipboardItem'].map(key => [key, Object.getOwnPropertyDescriptor(globalThis, key)]))
afterEach(() => {
  for (const [key, descriptor] of originals) {
    if (descriptor) Object.defineProperty(globalThis, key, descriptor)
    else delete globalThis[key]
  }
})
const install = clipboard => Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { clipboard } })
const item = contents => ({ types: Object.keys(contents), getType: async type => new Blob([contents[type]], { type }) })

test('reader availability requires a callable rich or text reader and never reads clipboard content', () => {
  let reads = 0
  for (const clipboard of [undefined, {}, { read: true }, { readText: 'reader' }, { writeText: async () => {} }]) {
    install(clipboard); assert.equal(sheetSystemClipboardAvailable(), false)
  }
  for (const clipboard of [{ read: async () => { reads++; return [] } }, { readText: async () => { reads++; return 'text' } }]) {
    install(clipboard); assert.equal(sheetSystemClipboardAvailable(), true)
  }
  assert.equal(reads, 0)
})

test('injected adapters can explicitly revoke reader availability independently of cached content', () => {
  let reads = 0, enabled = false
  const adapter = { canRead: () => enabled, read: async () => { reads++; return { text: 'external' } }, write: async () => true }
  assert.equal(sheetSystemClipboardAvailable(adapter), false); enabled = true; assert.equal(sheetSystemClipboardAvailable(adapter), true)
  assert.equal(sheetSystemClipboardAvailable({ read: async () => undefined, write: async () => true }), true)
  assert.equal(sheetSystemClipboardAvailable({ canRead: () => true, write: async () => true }), false)
  assert.equal(reads, 0)
})

test('the browser reader returns the matching plain and HTML flavours from one clipboard item', async () => {
  install({ read: async () => [item({ 'image/png': 'media' }), item({ 'text/plain': 'a\tb', 'text/html': '<table><tr><td>a</td><td>b</td></tr></table>' })] })
  assert.deepEqual(await sheetSystemClipboard.read(), { text: 'a\tb', html: '<table><tr><td>a</td><td>b</td></tr></table>' })
})

test('a successful empty or media-only read stays explicitly empty and never invokes readText', async () => {
  for (const items of [[], [item({ 'image/png': 'media' })]]) {
    let reads = 0; install({ read: async () => items, readText: async () => { reads++; return 'older data' } })
    assert.deepEqual(await sheetSystemClipboard.read(), { text: '', html: '' }); assert.equal(reads, 0)
  }
})

test('the text-only fallback preserves an empty reader result and distinguishes an absent reader', async () => {
  install({ readText: async () => '' }); assert.deepEqual(await sheetSystemClipboard.read(), { text: '' })
  install({}); assert.equal(await sheetSystemClipboard.read(), undefined)
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: undefined })
  assert.equal(await sheetSystemClipboard.read(), undefined)
})

test('an unavailable rich reader can fall back to readText while a refused final reader remains an error', async () => {
  install({ read: async () => { throw new Error('HTML unsupported') }, readText: async () => 'plain fallback' })
  assert.deepEqual(await sheetSystemClipboard.read(), { text: 'plain fallback' })
  const failure = new Error('permission refused')
  install({ read: async () => { throw failure } }); await assert.rejects(() => sheetSystemClipboard.read(), error => error === failure)
  install({ read: async () => { throw new Error('rich failed') }, readText: async () => { throw failure } })
  await assert.rejects(() => sheetSystemClipboard.read(), error => error === failure)
})

test('the browser writer includes real plain and HTML blobs in a ClipboardItem', async () => {
  const written = []
  Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: class {
    constructor(content) { this.content = content; this.types = Object.keys(content) }
  } })
  install({ write: async values => written.push(...values) })
  assert.equal(await sheetSystemClipboard.write({ plain: 'plain', html: '<b>rich</b>' }), true)
  assert.deepEqual(written[0].types, ['text/plain', 'text/html'])
  assert.equal(written[0].content['text/plain'].type, 'text/plain'); assert.equal(await written[0].content['text/plain'].text(), 'plain')
  assert.equal(written[0].content['text/html'].type, 'text/html'); assert.equal(await written[0].content['text/html'].text(), '<b>rich</b>')
})

test('unsupported rich writes use writeText, and missing writer capabilities return false', async () => {
  const written = []
  Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: class { constructor(content) { this.content = content } } })
  install({ write: async () => { throw new Error('HTML unsupported') }, writeText: async text => written.push(text) })
  assert.equal(await sheetSystemClipboard.write({ plain: 'fallback', html: '<b>fallback</b>' }), true); assert.deepEqual(written, ['fallback'])
  install({}); assert.equal(await sheetSystemClipboard.write({ plain: 'x', html: '' }), false)
  const failure = new Error('write refused'); install({ write: async () => { throw failure } })
  await assert.rejects(() => sheetSystemClipboard.write({ plain: 'x', html: '' }), error => error === failure)
})
