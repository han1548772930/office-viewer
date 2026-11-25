import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred DOC-22 API/transaction drafts. No checks or browsers have been run.
const fixturePath = '/__custom-xml-candidate-test__'
const fixture = `<!doctype html><html><head><meta charset="utf-8"></head><body><script type="module">
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph } from '/packages/office-studio/src/features/document/model/document-model'
const copy = value => JSON.parse(JSON.stringify(value))
function make(bindings) {
  const model = createBlankDocument(); model.id = 'xml-candidate'; model.name = 'XML candidate'
  model.blocks = ['origin', 'alias'].map(id => { const p = createParagraph('old'); p.id = id + '-p'; return p })
  model.contentControls = ['origin', 'alias'].map((id, index) => ({
    id, type: 'richText', placement: 'inline', blockId: id + '-p', startOffset: 0, endOffset: 3,
    dataBinding: bindings?.[index] ?? { storeItemID: '{STORE}', xpath: index ? '/root/value/text()' : '/root/value', cachedResult: 'old', dirty: false },
  }))
  model.customXmlParts = [{ partName: 'customXml/item.xml', storeItemID: '{STORE}', xml: '<root><value>old</value></root>' },
    { partName: 'customXml/untouched.xml', storeItemID: '{UNTOUCHED}', xml: '<untouched attr="retained"/>' }]
  return model
}
function open(model) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' }); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id), dispose: () => studio.dispose() }
}
function state(opened) {
  const unit = opened.studio.editor.history.units[opened.document.id]
  return { model: copy(opened.document.getModel()), revision: opened.document.getModelRevision(),
    undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0 }
}
function attempt(work) { try { work(); return null } catch (error) { return { code: error.code, message: error.message } } }
window.xmlCandidateFixture = { copy, make, open, state, attempt, OfficeStudio }
</script></body></html>`

let server, browser, page, baseURL, pageErrors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'custom-xml-candidate-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
      try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) }
      catch (error) { next(error) }
    }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('XML candidate fixture did not allocate a port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  pageErrors = []; page = await browser.newPage(); page.on('pageerror', error => pageErrors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => !!window.xmlCandidateFixture)
})
afterEach(async () => { await page?.close(); assert.deepEqual(pageErrors, []) })

test('constructor and createDocument keep optional dormant bindings, and reject activation against an empty store list atomically', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlCandidateFixture
    return [{}, { xpath: '/root/value' }, { storeItemID: '{STORE}' }, { xpath: '', storeItemID: '{STORE}' }].map(binding => {
      const model = f.make([binding, binding]); model.customXmlParts = []
      const opened = f.open(model)
      try {
        const before = f.state(opened), created = opened.studio.createDocument({ ...f.copy(model), id: 'dormant-created' }, false)
        const error = f.attempt(() => opened.document.mutate('Activate binding', 'doc.command.xml-draft', candidate => {
          candidate.contentControls[0].dataBinding = { xpath: '/root/value', storeItemID: '{STORE}' }
        }))
        return { binding, before, after: f.state(opened), error, created: f.copy(created.contentControls[0].dataBinding) }
      } finally { opened.dispose() }
    })
  })
  for (const item of observed) {
    assert.deepEqual(item.created, item.binding)
    assert.deepEqual(item.before.model.contentControls[0].dataBinding, item.binding)
    assert.equal(item.error?.code, 'DOC_CUSTOM_XML_STORE_NOT_FOUND')
    assert.deepEqual(item.after, item.before)
  }
})

test('generic source mutation refreshes XPath aliases, caches, word count and XML in one reversible history entry', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlCandidateFixture, opened = f.open(f.make()); let reopened
    try {
      const before = f.state(opened), value = 'new & 😀'
      const changed = opened.document.mutate('Change XML source', 'doc.command.xml-draft', candidate => {
        candidate.customXmlParts[0].xml = '<root><value>new &amp; 😀</value></root>'
      })
      const after = f.state(opened)
      opened.document.undo(); const undone = f.state(opened)
      opened.document.redo(); const redone = f.state(opened)
      reopened = f.open(f.copy(opened.document.saveModel()))
      return { before, changed, after, undone, redone, reopened: f.state(reopened), value }
    } finally { opened.dispose(); reopened?.dispose() }
  })
  assert.equal(observed.changed, true); assert.equal(observed.after.undo, 1)
  for (const control of observed.after.model.contentControls) {
    assert.equal(control.dataBinding.cachedResult, observed.value); assert.equal(control.dataBinding.dirty, false)
    const block = observed.after.model.blocks.find(block => block.id === control.blockId)
    assert.equal(block.runs.map(run => run.text).join(''), observed.value)
  }
  assert.equal(observed.after.model.wordCount, observed.value.replace(/\s/g, '').length * 2)
  assert.deepEqual(observed.after.model.customXmlParts[1], observed.before.model.customXmlParts[1])
  assert.deepEqual(observed.undone.model.customXmlParts, observed.before.model.customXmlParts)
  assert.deepEqual(observed.undone.model.contentControls, observed.before.model.contentControls)
  assert.deepEqual(observed.redone.model.customXmlParts, observed.after.model.customXmlParts)
  assert.deepEqual(observed.reopened.model.contentControls, observed.after.model.contentControls)
  assert.deepEqual(observed.reopened.model.customXmlParts, observed.after.model.customXmlParts)
  assert.equal(observed.reopened.undo, 0)
})

test('source mutation rejects a locked or denied alias without publishing source, text, cache or history', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlCandidateFixture
    return ['lock', 'permission'].map(scenario => {
      const model = f.make()
      if (scenario === 'lock') model.contentControls[1].lock = 'contentLocked'
      else model.permissionRules = [{ id: 'deny-alias', scope: 'paragraph', targetId: 'alias-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
      const opened = f.open(model)
      try {
        const before = f.state(opened)
        const error = f.attempt(() => opened.document.mutate('Change XML source', 'doc.command.xml-draft', candidate => {
          candidate.customXmlParts[0].xml = '<root><value>changed</value></root>'
        }))
        return { error, before, after: f.state(opened) }
      } finally { opened.dispose() }
    })
  })
  for (const item of observed) { assert.ok(item.error); assert.deepEqual(item.after, item.before) }
})

test('explicit stores cannot become legacy or dangling, while one mutation can unbind peers and clear sources', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlCandidateFixture, opened = f.open(f.make())
    try {
      const before = f.state(opened)
      const missing = f.attempt(() => opened.document.mutate('Remove sources', 'doc.command.xml-draft', model => { delete model.customXmlParts }))
      const dangling = f.attempt(() => opened.document.mutate('Clear sources', 'doc.command.xml-draft', model => { model.customXmlParts = [] }))
      const rejected = f.state(opened)
      const cleared = opened.document.mutate('Unbind and clear', 'doc.command.xml-draft', model => {
        for (const control of model.contentControls) delete control.dataBinding
        model.customXmlParts = []
      })
      const after = f.state(opened); opened.document.undo()
      return { before, missing, dangling, rejected, cleared, after, undone: f.state(opened) }
    } finally { opened.dispose() }
  })
  assert.equal(observed.missing?.code, 'DOC_CUSTOM_XML_INVALID_SOURCE')
  assert.equal(observed.dangling?.code, 'DOC_CUSTOM_XML_STORE_NOT_FOUND')
  assert.deepEqual(observed.rejected, observed.before)
  assert.equal(observed.cleared, true); assert.deepEqual(observed.after.model.customXmlParts, [])
  assert.equal(observed.after.undo, 1)
  assert.deepEqual(observed.undone.model.contentControls, observed.before.model.contentControls)
  assert.deepEqual(observed.undone.model.customXmlParts, observed.before.model.customXmlParts)
})

test('same-value setter refreshes stale mapped text once, returns the facade and isolates getter objects', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlCandidateFixture, model = f.make(); model.customXmlParts[0].xml = '<root><value>from XML</value></root>'
    const opened = f.open(model)
    try {
      const source = opened.document.getCustomXmlParts(), returned = opened.document.setCustomXmlParts(source)
      const first = f.state(opened); source[0].xml = '<damaged>'
      const getter = opened.document.getCustomXmlParts(); getter[0].storeItemID = 'damaged'
      opened.document.setCustomXmlParts(opened.document.getCustomXmlParts())
      return { same: returned === opened.document, first, second: f.state(opened), original: f.copy(model.customXmlParts) }
    } finally { opened.dispose() }
  })
  assert.equal(observed.same, true); assert.equal(observed.first.undo, 1)
  assert.equal(observed.first.model.contentControls[0].dataBinding.cachedResult, 'from XML')
  assert.equal(observed.second.undo, 1)
  assert.deepEqual(observed.second.model.customXmlParts, observed.original)
})

test('constructor and candidate reject real DTDs or missing graph targets but accept declaration text in comments and CDATA', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlCandidateFixture, model = f.make([{}, {}])
    const bad = f.copy(model); bad.customXmlParts[0].xml = '<!DOCTYPE root [<!ENTITY value "unsafe">]><root>&value;</root>'
    const initialError = f.attempt(() => { const studio = new f.OfficeStudio({ document: bad }); studio.dispose() })
    const opened = f.open(model)
    try {
      const before = f.state(opened), graphError = f.attempt(() => opened.document.mutate('Bad graph', 'doc.command.xml-draft', candidate => {
        candidate.customXmlParts[0].relationshipsXml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="urn:payload" Target="missing.bin"/></Relationships>'
      }))
      const rejected = f.state(opened)
      const accepted = opened.document.mutate('Literal declarations', 'doc.command.xml-draft', candidate => {
        candidate.customXmlParts[0].xml = '<!-- literal <!DOCTYPE root> --><root><![CDATA[<!DOCTYPE literal>]]><?keep x?></root>'
      })
      return { initialError, graphError, before, rejected, accepted, after: f.state(opened) }
    } finally { opened.dispose() }
  })
  assert.ok(observed.initialError); assert.equal(observed.graphError?.code, 'DOC_CUSTOM_XML_MISSING_PART')
  assert.deepEqual(observed.rejected, observed.before); assert.equal(observed.accepted, true)
  assert.ok(observed.after.model.customXmlParts[0].xml.includes('<![CDATA[<!DOCTYPE literal>]]>'))
})
