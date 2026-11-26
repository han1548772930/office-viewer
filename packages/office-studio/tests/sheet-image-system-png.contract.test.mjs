import assert from 'node:assert/strict'
import { File as NodeFile } from 'node:buffer'
import { afterEach, beforeEach, test, vi } from 'vitest'
import { sheetImageClipboardPngBlob } from '../src/features/sheet-images/sheet-image-system-clipboard.ts'
import { decodeSheetImageDataUrl } from '../src/features/sheet-images/sheet-image-bytes.ts'
import { sheetImageResourceFromDataUrl } from '../src/features/sheet-images/sheet-image-source.ts'
import { normalizeSheetImageResource } from '../src/features/sheet-images/sheet-image-model.ts'
import { sheetImagePng } from './fixtures/sheet-image-fixture.mjs'

// Deferred platform ownership/cancellation contracts; actual browser raster output remains unaccepted.
const gif = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
let decoded, canvases
beforeEach(() => {
  decoded = []; canvases = []
  vi.stubGlobal('Image', class {
    naturalWidth = 1; naturalHeight = 1; onload = null; onerror = null
    constructor() { decoded.push(this) }
    get src() { return this.value ?? '' }
    set src(value) { this.value = value; if (value) queueMicrotask(() => { if (this.value) this.onload?.() }) }
  })
})
afterEach(() => vi.unstubAllGlobals())
const image = () => sheetImageResourceFromDataUrl(gif, 'source.gif')
function offscreen(convert = async () => new Blob(['encoded-png'], { type: 'image/png' }), context = true) {
  vi.stubGlobal('OffscreenCanvas', class {
    constructor(width, height) { this.width = width; this.height = height; this.draws = 0; canvases.push(this) }
    getContext() { return context ? { drawImage: source => { assert.equal(source, decoded[0]); this.draws++ } } : null }
    convertToBlob(options) { assert.equal(options.type, 'image/png'); return convert() }
  })
}

test('a native PNG uses validated actual bytes and needs no decoder or canvas', async () => {
  const resource = sheetImageResourceFromDataUrl(sheetImagePng), original = structuredClone(resource), blob = await sheetImageClipboardPngBlob(resource)
  assert.equal(blob.type, 'image/png'); assert.deepEqual(new Uint8Array(await blob.arrayBuffer()), decodeSheetImageDataUrl(sheetImagePng).bytes)
  assert.equal(decoded.length, 0); assert.deepEqual(resource, original)
})

test('invalid PNG dimensions or MIME declarations are rejected before any platform allocation', async () => {
  const resource = sheetImageResourceFromDataUrl(sheetImagePng)
  await assert.rejects(() => sheetImageClipboardPngBlob({ ...resource, width: 2 }), /尺寸|实际内容/)
  await assert.rejects(() => sheetImageClipboardPngBlob({ ...resource, mimeType: 'image/jpeg' }), /MIME|实际内容/)
  assert.equal(decoded.length, 0); assert.equal(canvases.length, 0)
})

test('an already cancelled or stale clipboard task does not decode or fetch source media', async () => {
  const controller = new AbortController(); controller.abort()
  await assert.rejects(() => sheetImageClipboardPngBlob(image(), { signal: controller.signal }), error => error.name === 'AbortError')
  await assert.rejects(() => sheetImageClipboardPngBlob(image(), { current: () => false }), error => error.name === 'AbortError')
  assert.equal(decoded.length, 0); assert.equal(canvases.length, 0)
})

test('Offscreen PNG conversion clears its bitmap and decoded image after success', async () => {
  offscreen(); const resource = image(), original = structuredClone(resource), blob = await sheetImageClipboardPngBlob(resource)
  assert.equal(blob.type, 'image/png'); assert.equal(canvases[0].draws, 1)
  assert.deepEqual([canvases[0].width, canvases[0].height, decoded[0].src], [0, 0, '']); assert.deepEqual(resource, original)
})

test('a host becoming stale during asynchronous PNG encoding discards the result and releases both resources', async () => {
  let current = true
  offscreen(async () => { current = false; return new Blob(['old'], { type: 'image/png' }) })
  await assert.rejects(() => sheetImageClipboardPngBlob(image(), { current: () => current }), error => error.name === 'AbortError')
  assert.deepEqual([canvases[0].width, canvases[0].height, decoded[0].src], [0, 0, ''])
})

test('a refused canvas context or failed encoder releases the temporary bitmap and source', async () => {
  offscreen(undefined, false); await assert.rejects(() => sheetImageClipboardPngBlob(image()), /无法准备/)
  assert.deepEqual([canvases[0].width, canvases[0].height, decoded[0].src], [0, 0, ''])
  decoded = []; canvases = []; offscreen(async () => { throw new Error('encoder failed') })
  await assert.rejects(() => sheetImageClipboardPngBlob(image()), /encoder failed/)
  assert.deepEqual([canvases[0].width, canvases[0].height, decoded[0].src], [0, 0, ''])
})

test('the DOM canvas fallback releases its backing store after an asynchronous toBlob callback', async () => {
  vi.stubGlobal('OffscreenCanvas', undefined)
  vi.stubGlobal('document', { createElement: tag => {
    assert.equal(tag, 'canvas')
    const canvas = { width: 0, height: 0, getContext: () => ({ drawImage: () => {} }),
      toBlob: (callback, mimeType) => { assert.equal(mimeType, 'image/png'); queueMicrotask(() => callback(new Blob(['png'], { type: mimeType }))) } }
    canvases.push(canvas); return canvas
  } })
  assert.equal((await sheetImageClipboardPngBlob(image())).type, 'image/png')
  assert.deepEqual([canvases[0].width, canvases[0].height, decoded[0].src], [0, 0, ''])
})

test('remote clipboard media uses fetched bytes and their actual MIME rather than cached URL hints', async () => {
  const bytes = decodeSheetImageDataUrl(sheetImagePng).bytes, source = 'https://example.test/source.gif'
  vi.stubGlobal('File', NodeFile)
  vi.stubGlobal('FileReader', class {
    result = null; onload = null; onerror = null; onabort = null
    readAsArrayBuffer(file) { file.arrayBuffer().then(value => { this.result = value; this.onload?.() }) }
    abort() { this.onabort?.() }
  })
  let fetched
  vi.stubGlobal('fetch', async (url, options) => { fetched = { url, options }; return new Response(bytes, { status: 200 }) })
  const controller = new AbortController(), resource = normalizeSheetImageResource({ ...image(), source }), before = structuredClone(resource),
    blob = await sheetImageClipboardPngBlob(resource, { signal: controller.signal, current: () => true })
  assert.equal(fetched.url, source); assert.equal(fetched.options.signal, controller.signal)
  assert.equal(blob.type, 'image/png'); assert.deepEqual(new Uint8Array(await blob.arrayBuffer()), bytes)
  assert.deepEqual(resource, before); assert.equal(decoded[0].src, '')
})
