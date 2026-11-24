import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred DOC-11 integration drafts. Real browser XPath is used; these drafts have not run.
// They do not replace Office-file, Workbench/IME, renderer or large-document acceptance.
const fixturePath = '/__document-custom-xml-bindings-test__'
const fixture = `<!doctype html><html><head><meta charset="utf-8"><title>Custom XML binding drafts</title></head><body><script type="module">
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import * as modelTools from '/packages/office-studio/src/features/document/model/document-model'
import * as controls from '/packages/office-studio/src/features/document/services/content-controls'
import * as bindings from '/packages/office-studio/src/features/document/services/content-control-bindings'
import * as interaction from '/packages/office-studio/src/features/document/controllers/content-control-interaction'
import * as editor from '/packages/office-studio/src/api/document-editor-commands'
import * as search from '/packages/office-studio/src/features/document/services/document-search'
const copy = value => JSON.parse(JSON.stringify(value))
const permission = { documentId: 'custom-xml-browser', userId: 'editor' }
function paragraph(id, text) { const block = modelTools.createParagraph(text); block.id = id; return block }
function make(specs, xml = '<root><value>old</value></root>') {
  const model = createBlankDocument(); model.id = permission.documentId
  model.blocks = specs.map(spec => paragraph(spec.id + '-p', spec.text ?? 'old')); model.contentControls = []
  for (const spec of specs) {
    const { id, text = 'old', xpath = '/root/value', prefixMappings, storeItemID = '{STORE}', cache = text, ...patch } = spec
    controls.createDocumentContentControl(model, { id, type: 'richText', blockId: id + '-p', startOffset: 0, endOffset: text.length, value: text,
      metadata: { retained: { id } }, dataBinding: { storeItemID, xpath, ...(prefixMappings ? { prefixMappings } : {}), cachedResult: cache, dirty: false }, ...patch })
  }
  model.customXmlParts = [{ partName: 'customXml/item1.xml', storeItemID: '{STORE}', xml, propertiesPartName: 'customXml/itemProps1.xml',
    propertiesXml: '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{STORE}"/>' },
    { partName: 'customXml/item2.xml', storeItemID: '{UNTOUCHED}', xml: '<untouched attr="retained"> keep </untouched>' }]
  return model
}
function open(model) {
  const studio = new OfficeStudio({ document: model, currentUserId: permission.userId }); studio.activateDocument(model.id)
  const document = new FDocument(studio, model.id), events = []
  const subscription = studio.events.addEvent('CommandExecuted', event => { if (event.kind === 'document') events.push(event.id) })
  return { studio, document, events, dispose: () => { subscription.dispose(); studio.dispose() } }
}
const info = (model, id) => model.contentControls?.find(control => control.id === id)
const text = (model, id) => controls.contentControlText(model, info(model, id))
const history = opened => ({ undo: opened.studio.editor.history.units[opened.document.id]?.undo.length ?? 0,
  redo: opened.studio.editor.history.units[opened.document.id]?.redo.length ?? 0 })
const xmlValues = model => Array.from(new DOMParser().parseFromString(model.customXmlParts[0].xml, 'application/xml').documentElement.children)
  .map(element => element.textContent)
const selection = (id, start, end = start) => ({ anchor: { blockId: id, offset: start }, focus: { blockId: id, offset: end } })
window.xmlBindingFixture = { copy, permission, paragraph, make, open, info, text, history, xmlValues, selection, controls, bindings, interaction, editor, search }
</script></body></html>`

let server, browser, page, baseURL, pageErrors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-custom-xml-binding-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
      try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) }
      catch (error) { next(error) }
    }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('The custom-XML fixture did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  pageErrors = []; page = await browser.newPage(); page.on('pageerror', error => pageErrors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => !!window.xmlBindingFixture)
})
afterEach(async () => { await page?.close(); assert.deepEqual(pageErrors, []) })

test('namespace and leaf-text aliases update a header peer, XML and caches in one undoable facade transaction', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, xml = '<p:person xmlns:p="urn:person"><p:name>old</p:name></p:person>'
    const model = f.make([{ id: 'origin', temporary: true, xpath: '/p:person/p:name', prefixMappings: 'xmlns:p="urn:person"' },
      { id: 'header', xpath: '/n:person/n:name/text()', prefixMappings: "xmlns:n='urn:person'" }], xml)
    const header = model.blocks.splice(1, 1)[0]
    model.sectionHeaderFooters = { first: { header: 'old', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
      differentFirstPage: false, differentOddEven: false, linkToPrevious: false,
      storyContent: { header: { blocks: [header], runs: f.copy(header.runs), paragraphs: [{ startOffset: 0, endOffset: 3, style: {} }] } } } }
    const opened = f.open(model); let reopened
    try {
      const before = f.copy(opened.document.getModel()), value = 'Mira &😀<'
      const changed = opened.document.getContentControl('origin').setValue(value), after = f.copy(opened.document.getModel()), stack = f.history(opened)
      opened.document.undo(); const undone = f.copy(opened.document.getModel())
      opened.document.redo(); const redone = f.copy(opened.document.getModel())
      const saved = f.copy(opened.document.save()); reopened = f.open(saved)
      return { changed, before, after, undone, redone, stack, values: f.xmlValues(after), headerText: f.text(after, 'header'),
        reopened: f.copy(reopened.document.getModel()), reopenedStack: f.history(reopened), errors: opened.studio.getEventListenerErrors().length }
    } finally { opened.dispose(); reopened?.dispose() }
  })
  assert.equal(observed.changed, true); assert.deepEqual(observed.stack, { undo: 1, redo: 0 })
  assert.deepEqual(observed.values, ['Mira &😀<']); assert.equal(observed.headerText, 'Mira &😀<')
  assert.equal(observed.after.contentControls.some(control => control.id === 'origin'), false)
  const header = observed.after.contentControls.find(control => control.id === 'header')
  assert.equal(header.dataBinding.cachedResult, 'Mira &😀<'); assert.equal(header.dataBinding.dirty, false)
  assert.deepEqual(header.metadata.retained, { id: 'header' })
  assert.equal(observed.after.customXmlParts[0].propertiesXml, observed.before.customXmlParts[0].propertiesXml)
  assert.deepEqual(observed.after.customXmlParts[1], observed.before.customXmlParts[1])
  assert.deepEqual(observed.undone.contentControls, observed.before.contentControls); assert.deepEqual(observed.undone.customXmlParts, observed.before.customXmlParts)
  assert.deepEqual(observed.redone.contentControls, observed.after.contentControls); assert.deepEqual(observed.redone.customXmlParts, observed.after.customXmlParts)
  assert.deepEqual(observed.reopened.contentControls, observed.after.contentControls); assert.deepEqual(observed.reopened.customXmlParts, observed.after.customXmlParts)
  assert.deepEqual(observed.reopenedStack, { undo: 0, redo: 0 }); assert.equal(observed.errors, 0)
})

test('ordinary input maps a UTF-16 caret through same-paragraph XPath aliases after temporary source removal', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, model = f.make([],'<root><value>A</value></root>')
    model.blocks = [f.paragraph('p', 'A / A / A')]
    for (const [id, start, xpath, temporary] of [['first', 0, '/root/value', false], ['second', 4, '/root/value/text()', false], ['origin', 8, '//*[local-name()="value"]', true]]) {
      f.controls.createDocumentContentControl(model, { id, type: 'richText', blockId: 'p', startOffset: start, endOffset: start + 1, value: 'A', temporary,
        dataBinding: { storeItemID: '{STORE}', xpath, cachedResult: 'A', dirty: false } })
    }
    const selection = f.editor.applyDocumentEditorCommand(model, { unitId: model.id, selection: f.selection('p', 8, 9),
      command: { id: 'doc.command.insert-text', text: '中😀', activeContentControlId: 'origin' } }, f.permission)
    return { model: f.copy(model), selection, values: f.xmlValues(model) }
  })
  assert.equal(observed.model.blocks[0].runs.map(run => run.text).join(''), '中😀 / 中😀 / 中😀')
  assert.equal(observed.selection.anchor.offset, 15); assert.equal(observed.selection.focus.offset, 15)
  assert.equal(observed.model.contentControls.some(control => control.id === 'origin'), false)
  assert.deepEqual(observed.values, ['中😀'])
  for (const control of observed.model.contentControls) { assert.equal(control.dataBinding.cachedResult, '中😀'); assert.equal(control.dataBinding.dirty, false) }
})

test('a checkbox peer reads XML 1 without rewriting its lexical scalar to true', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, opened = f.open(f.make([{ id: 'origin', text: '0' }, { id: 'check', type: 'checkbox', text: '☐', checked: false }], '<root><value>0</value></root>'))
    try {
      opened.document.getContentControl('origin').setValue('1')
      const model = f.copy(opened.document.getModel()); return { model, values: f.xmlValues(model), checkedText: f.text(model, 'check') }
    } finally { opened.dispose() }
  })
  assert.deepEqual(observed.values, ['1']); assert.equal(observed.checkedText, '☒')
  const checkbox = observed.model.contentControls.find(control => control.id === 'check')
  assert.equal(checkbox.checked, true); assert.equal(checkbox.value, 'true')
  assert.equal(checkbox.dataBinding.cachedResult, '1'); assert.equal(checkbox.dataBinding.dirty, false)
})

test('explicit temporary value commands recover their caret through surviving XPath aliases and rematched predicates', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture
    return [false, true].map(predicate => {
      const xml = predicate ? '<root><item>A</item><item>A</item></root>' : '<root><value>A</value></root>'
      const model = f.make([], xml); model.blocks = [f.paragraph('p', 'A / A / A')]
      const specs = predicate ? [['first', 0, '/root/item[1]'], ['second', 4, '/root/item[.="A"]'], ['origin', 8, '/root/item[1]/text()']]
        : [['first', 0, '/root/value'], ['second', 4, '/root/value/text()'], ['origin', 8, '//*[local-name()="value"]']]
      for (const [id, start, xpath] of specs) f.controls.createDocumentContentControl(model, { id, type: 'richText', blockId: 'p',
        startOffset: start, endOffset: start + 1, value: 'A', temporary: id === 'origin',
        dataBinding: { storeItemID: '{STORE}', xpath, cachedResult: 'A', dirty: false } })
      const opened = f.open(model)
      try {
        const source = f.copy(opened.document.getModel()), value = '中😀', command = { id: 'doc.command.set-content-control-value', params: { controlId: 'origin', value } }
        opened.document.getContentControl('origin').setValue(value)
        const current = opened.document.getModel(), request = { source, controlId: 'origin', selection: f.selection('p', 8, 9), command }
        return { predicate, text: current.blocks[0].runs.map(run => run.text).join(''), values: f.xmlValues(current), stack: f.history(opened),
          caret: f.interaction.documentContentControlSelectionAfterValueEdit(request, current), removed: !f.info(current, 'origin') }
      } finally { opened.dispose() }
    })
  })
  for (const sample of observed) {
    const expected = sample.predicate ? '中😀 / A / 中😀' : '中😀 / 中😀 / 中😀'
    assert.equal(sample.text, expected); assert.equal(sample.caret.anchor.offset, expected.length); assert.equal(sample.caret.focus.offset, expected.length)
    assert.equal(sample.removed, true); assert.deepEqual(sample.stack, { undo: 1, redo: 0 })
    assert.deepEqual(sample.values, sample.predicate ? ['中😀', 'A'] : ['中😀'])
  }
})

test('explicit temporary clearing maps past a surviving leaf-text alias prompt', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, model = f.make([], '<root><value>A</value></root>'); model.blocks = [f.paragraph('p', 'A / A')]
    for (const [id, start, xpath] of [['peer', 0, '/root/value/text()'], ['origin', 4, '/root/value']]) f.controls.createDocumentContentControl(model,
      { id, type: 'richText', blockId: 'p', startOffset: start, endOffset: start + 1, value: 'A', temporary: id === 'origin', placeholder: 'Prompt',
        dataBinding: { storeItemID: '{STORE}', xpath, cachedResult: 'A', dirty: false } })
    const opened = f.open(model)
    try {
      const source = f.copy(opened.document.getModel()), command = { id: 'doc.command.set-content-control-value', params: { controlId: 'origin', value: '' } }
      opened.document.getContentControl('origin').setValue('')
      const current = opened.document.getModel()
      return { text: current.blocks[0].runs.map(run => run.text).join(''), values: f.xmlValues(current), peer: f.copy(f.info(current, 'peer')),
        caret: f.interaction.documentContentControlSelectionAfterValueEdit({ source, controlId: 'origin', selection: f.selection('p', 4, 5), command }, current) }
    } finally { opened.dispose() }
  })
  assert.equal(observed.text, 'Prompt / '); assert.deepEqual(observed.values, ['']); assert.equal(observed.peer.showingPlaceholder, true)
  assert.equal(observed.caret.anchor.offset, observed.text.length); assert.equal(observed.caret.focus.offset, observed.text.length)
  assert.equal(observed.peer.dataBinding.cachedResult, ''); assert.equal(observed.peer.dataBinding.dirty, false)
})

test('explicit temporary placeholder reset maps its prompt past a surviving leaf-text alias prompt', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, model = f.make([], '<root><value>A</value></root>'); model.blocks = [f.paragraph('p', 'A / A')]
    for (const [id, start, xpath, placeholder] of [['peer', 0, '/root/value/text()', 'Peer😀'], ['origin', 4, '/root/value', 'Origin😀']]) {
      f.controls.createDocumentContentControl(model, { id, type: 'richText', blockId: 'p', startOffset: start, endOffset: start + 1, value: 'A',
        temporary: id === 'origin', placeholder, dataBinding: { storeItemID: '{STORE}', xpath, cachedResult: 'A', dirty: false } })
    }
    const opened = f.open(model)
    try {
      const source = f.copy(opened.document.getModel()), command = { id: 'doc.command.reset-content-control-placeholder', params: { controlId: 'origin' } }
      const changed = opened.document.getContentControl('origin').resetPlaceholder(), current = opened.document.getModel()
      return { changed, removed: !f.info(current, 'origin'), text: current.blocks[0].runs.map(run => run.text).join(''), values: f.xmlValues(current),
        peer: f.copy(f.info(current, 'peer')), stack: f.history(opened),
        caret: f.interaction.documentContentControlSelectionAfterValueEdit({ source, controlId: 'origin', selection: f.selection('p', 4, 5), command }, current) }
    } finally { opened.dispose() }
  })
  assert.equal(observed.changed, true); assert.equal(observed.removed, true); assert.equal(observed.text, 'Peer😀 / Origin😀')
  assert.deepEqual(observed.values, ['']); assert.equal(observed.peer.showingPlaceholder, true)
  assert.deepEqual([observed.peer.startOffset, observed.peer.endOffset], [0, 'Peer😀'.length])
  assert.equal(observed.peer.dataBinding.cachedResult, ''); assert.equal(observed.peer.dataBinding.dirty, false)
  assert.equal(observed.caret.anchor.offset, observed.text.length); assert.equal(observed.caret.focus.offset, observed.text.length)
  assert.deepEqual(observed.stack, { undo: 1, redo: 0 })
})

test('clearing XML restores a leaf-text alias prompt and retains empty scalar caches through temporary reset', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, opened = f.open(f.make([{ id: 'origin', temporary: true, placeholder: 'Origin prompt' },
      { id: 'alias', xpath: '/root/value/text()', placeholder: '', placeholderDocPart: { name: 'Part', text: 'Unused prompt' } }]))
    try {
      const before = f.copy(opened.document.getModel())
      const changed = opened.document.getContentControl('origin').resetPlaceholder(), after = f.copy(opened.document.getModel()), stack = f.history(opened)
      opened.document.undo(); const undone = f.copy(opened.document.getModel())
      opened.document.redo(); const redone = f.copy(opened.document.getModel())
      return { changed, before, after, undone, redone, stack, values: f.xmlValues(after), aliasText: f.text(after, 'alias') }
    } finally { opened.dispose() }
  })
  assert.equal(observed.changed, true); assert.deepEqual(observed.values, ['']); assert.equal(observed.aliasText, '')
  assert.equal(observed.after.contentControls.some(control => control.id === 'origin'), false)
  assert.equal(observed.after.blocks[0].runs.map(run => run.text).join(''), 'Origin prompt')
  const alias = observed.after.contentControls.find(control => control.id === 'alias')
  assert.equal(alias.showingPlaceholder, true); assert.equal(alias.dataBinding.cachedResult, ''); assert.equal(alias.dataBinding.dirty, false)
  assert.deepEqual(observed.stack, { undo: 1, redo: 0 }); assert.deepEqual(observed.undone.contentControls, observed.before.contentControls)
  assert.deepEqual(observed.undone.customXmlParts, observed.before.customXmlParts); assert.deepEqual(observed.redone.customXmlParts, observed.after.customXmlParts)
})

test('explicit XML refresh uses saved list values and ISO dates while retaining the exact supplied XML', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, model = f.make([
      { id: 'list', type: 'dropDownList', text: 'Alpha', value: 'A', xpath: '/root/choice', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Alpha' }] },
      { id: 'date', type: 'date', text: '2026-11-24', value: '2026-11-24', dateValue: '2026-11-24', dateFormat: 'dd MMM yyyy', dateLocale: 'en-US', xpath: '/root/date' },
    ], '<root><choice>A</choice><date>2026-11-24</date></root>')
    const opened = f.open(model)
    try {
      const parts = f.copy(opened.document.getModel().customXmlParts), xml = '<root>\n <choice>B</choice><date>2026-11-30</date>\n</root>'
      parts[0].xml = xml
      const beforeRevision = opened.studio.getDocumentRevision(model.id)
      const fluent = opened.document.setCustomXmlParts(parts) === opened.document, after = f.copy(opened.document.getModel()), stack = f.history(opened)
      const revision = opened.studio.getDocumentRevision(model.id), sameFluent = opened.document.setCustomXmlParts(f.copy(parts)) === opened.document
      const values = opened.document.getCustomXmlParts(); values[0].xml = '<caller-change/>'
      return { fluent, sameFluent, after, stack, finalStack: f.history(opened), text: [f.text(after, 'list'), f.text(after, 'date')], xml,
        beforeRevision, revision, finalRevision: opened.studio.getDocumentRevision(model.id), isolated: opened.document.getCustomXmlParts()[0].xml }
    } finally { opened.dispose() }
  })
  assert.equal(observed.fluent, true); assert.equal(observed.sameFluent, true); assert.deepEqual(observed.stack, { undo: 1, redo: 0 })
  assert.equal(observed.revision, observed.beforeRevision + 1); assert.deepEqual(observed.finalStack, observed.stack)
  assert.deepEqual(observed.text, ['Alpha', '30 Nov 2026']); assert.equal(observed.after.customXmlParts[0].xml, observed.xml)
  assert.equal(observed.after.contentControls.find(control => control.id === 'list').value, 'B')
  assert.equal(observed.after.contentControls.find(control => control.id === 'date').dateValue, '2026-11-30')
  assert.equal(observed.finalRevision, observed.revision); assert.equal(observed.isolated, observed.xml)
})

test('internal refresh ignores changed UI sources when the XML source has been replaced', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, before = f.make([{ id: 'origin' }, { id: 'peer', xpath: '/root/value/text()' }]), draft = f.copy(before)
    f.controls.setDocumentContentControlValue(draft, 'origin', 'stale UI override')
    draft.customXmlParts[0].xml = '<root><value>XML authority</value></root>'
    f.bindings.synchronizeDocumentContentControlBindings(before, draft, { refreshCustomXml: true, permissionContext: f.permission })
    return { text: [f.text(draft, 'origin'), f.text(draft, 'peer')], model: f.copy(draft), values: f.xmlValues(draft) }
  })
  assert.deepEqual(observed.text, ['XML authority', 'XML authority']); assert.deepEqual(observed.values, ['XML authority'])
  for (const control of observed.model.contentControls) { assert.equal(control.dataBinding.cachedResult, 'XML authority'); assert.equal(control.dataBinding.dirty, false) }
})

test('missing stores, source nodes, malformed XML and invalid XPath roll back the facade source edit', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, cases = [
      { patch: { storeItemID: '{MISSING}' }, code: 'DOC_CUSTOM_XML_STORE_NOT_FOUND' },
      { patch: { xpath: '/root/missing' }, code: 'DOC_CUSTOM_XML_NODE_NOT_FOUND' },
      { patch: { xpath: '/root/value[' }, code: 'DOC_CUSTOM_XML_INVALID_XPATH' },
      { xml: '<root>', code: 'DOC_CUSTOM_XML_INVALID_XML' },
    ]
    return cases.map(sample => {
      const opened = f.open(f.make([{ id: 'origin', temporary: true, ...sample.patch }, { id: 'peer' }], sample.xml))
      try {
        const before = f.copy(opened.document.getModel()), revision = opened.studio.getDocumentRevision(opened.document.id); let error
        try { opened.document.getContentControl('origin').setValue('new') } catch (failure) { error = { code: failure.code, message: failure.message } }
        return { before, after: f.copy(opened.document.getModel()), stack: f.history(opened), events: [...opened.events], error, expected: sample.code,
          revision, finalRevision: opened.studio.getDocumentRevision(opened.document.id) }
      } finally { opened.dispose() }
    })
  })
  for (const sample of observed) {
    assert.equal(sample.error?.code, sample.expected); assert.deepEqual(sample.after, sample.before)
    assert.deepEqual(sample.stack, { undo: 0, redo: 0 }); assert.deepEqual(sample.events, []); assert.equal(sample.finalRevision, sample.revision)
  }
})

test('distinct XPath aliases with conflicting replace-all scalars reject the whole originating batch', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, model = f.make([{ id: 'origin', text: 'A1' }, { id: 'peer', text: 'A2', xpath: '/root/value/text()' }], '<root><value>A1</value></root>')
    const before = f.copy(model); let code
    try { f.search.replaceDocumentMatches(model, f.search.findDocumentMatches(model, 'A'), 'B', f.permission) } catch (error) { code = error.code }
    return { before, after: f.copy(model), code }
  })
  assert.equal(observed.code, 'DOC_CUSTOM_XML_VALUE_CONFLICT'); assert.deepEqual(observed.after, observed.before)
})

test('later peer locks, permissions and invalid list values also roll back XML and temporary source removal', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture
    return ['lock', 'permission', 'list'].map(scenario => {
      const model = f.make([{ id: 'origin', temporary: true }, { id: 'first', xpath: '/root/value/text()' },
        { id: 'last', ...(scenario === 'lock' ? { lock: 'contentLocked' } : scenario === 'list' ? { type: 'dropDownList', value: 'old', items: [{ value: 'old' }] } : {}) }])
      if (scenario === 'permission') model.permissionRules = [{ id: 'deny-last', scope: 'paragraph', targetId: 'last-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
      const opened = f.open(model)
      try {
        const before = f.copy(opened.document.getModel()); let message
        try { opened.document.getContentControl('origin').setValue('new') } catch (error) { message = error.message }
        return { scenario, message, before, after: f.copy(opened.document.getModel()), stack: f.history(opened), events: [...opened.events] }
      } finally { opened.dispose() }
    })
  })
  for (const sample of observed) {
    assert.match(sample.message ?? '', /locked|permission denied|dropdown/); assert.deepEqual(sample.after, sample.before)
    assert.deepEqual(sample.stack, { undo: 0, redo: 0 }); assert.deepEqual(sample.events, [])
  }
})

test('cache-only XML peer writes cannot bypass a denied paragraph', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, model = f.make([{ id: 'origin' }, { id: 'cache', text: 'same', cache: 'stale', xpath: '/root/other' }],
      '<root><value>old</value><other>same</other></root>')
    model.permissionRules = [{ id: 'deny-cache', scope: 'paragraph', targetId: 'cache-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
    const opened = f.open(model)
    try {
      const before = f.copy(opened.document.getModel()); let message
      try { opened.document.getContentControl('origin').setValue('new') } catch (error) { message = error.message }
      return { before, after: f.copy(opened.document.getModel()), message, stack: f.history(opened) }
    } finally { opened.dispose() }
  })
  assert.match(observed.message ?? '', /permission denied/); assert.deepEqual(observed.after, observed.before)
  assert.deepEqual(observed.stack, { undo: 0, redo: 0 })
})

test('a source predicate that stops matching retains its written node for this transaction', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, opened = f.open(f.make([{ id: 'origin', xpath: '/root/value[.="old"]' }, { id: 'peer', xpath: '/root/value/text()' }]))
    try {
      opened.document.getContentControl('origin').setValue('new')
      const model = f.copy(opened.document.getModel()); return { model, text: [f.text(model, 'origin'), f.text(model, 'peer')], values: f.xmlValues(model), stack: f.history(opened) }
    } finally { opened.dispose() }
  })
  assert.deepEqual(observed.text, ['new', 'new']); assert.deepEqual(observed.values, ['new']); assert.deepEqual(observed.stack, { undo: 1, redo: 0 })
  for (const control of observed.model.contentControls) { assert.equal(control.dataBinding.cachedResult, 'new'); assert.equal(control.dataBinding.dirty, false) }
})

test('a source predicate that selects a second node updates the source display from that new key without changing the second XML node', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, opened = f.open(f.make([{ id: 'origin', xpath: '/root/item[.="old"]' },
      { id: 'first', xpath: '/root/item[1]/text()' }, { id: 'second', xpath: '/root/item[2]' }], '<root><item>old</item><item>old</item></root>'))
    try {
      opened.document.getContentControl('origin').setValue('new')
      const model = f.copy(opened.document.getModel()); return { model, text: ['origin', 'first', 'second'].map(id => f.text(model, id)), values: f.xmlValues(model), stack: f.history(opened) }
    } finally { opened.dispose() }
  })
  assert.deepEqual(observed.values, ['new', 'old']); assert.deepEqual(observed.text, ['old', 'new', 'old'])
  assert.deepEqual(observed.stack, { undo: 1, redo: 0 })
  assert.equal(observed.model.contentControls.find(control => control.id === 'origin').dataBinding.cachedResult, 'old')
})

test('an unmatched non-source XPath retains its prior content and cache while the actual source is synchronized', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture, opened = f.open(f.make([{ id: 'origin' }, { id: 'unmatched', text: 'retained', xpath: '/root/missing' }]))
    try {
      const prior = f.copy(opened.document.getContentControl('unmatched').getInfo())
      opened.document.getContentControl('origin').setValue('new')
      const model = f.copy(opened.document.getModel()); return { prior, current: f.info(model, 'unmatched'), text: f.text(model, 'unmatched'), values: f.xmlValues(model) }
    } finally { opened.dispose() }
  })
  assert.deepEqual(observed.current, observed.prior); assert.equal(observed.text, 'retained'); assert.deepEqual(observed.values, ['new'])
})

test('a failed XML-source refresh preserves the old parts, peer values and history', async () => {
  const observed = await page.evaluate(() => {
    const f = window.xmlBindingFixture
    return ['xml', 'list', 'permission'].map(scenario => {
      const model = f.make([{ id: 'origin' }, { id: 'last', ...(scenario === 'list' ? { type: 'dropDownList', value: 'old', items: [{ value: 'old' }] } : {}) }])
      if (scenario === 'permission') model.permissionRules = [{ id: 'deny-last', scope: 'paragraph', targetId: 'last-p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
      const opened = f.open(model)
      try {
        const before = f.copy(opened.document.getModel()), parts = f.copy(before.customXmlParts); parts[0].xml = scenario === 'xml' ? '<root>' : '<root><value>new</value></root>'
        let error
        try { opened.document.setCustomXmlParts(parts) } catch (failure) { error = { code: failure.code, message: failure.message } }
        return { before, after: f.copy(opened.document.getModel()), error, scenario, stack: f.history(opened), events: [...opened.events] }
      } finally { opened.dispose() }
    })
  })
  for (const sample of observed) {
    assert.match(sample.error?.message ?? '', /XML|dropdown|permission denied/)
    assert.deepEqual(sample.after, sample.before); assert.deepEqual(sample.stack, { undo: 0, redo: 0 }); assert.deepEqual(sample.events, [])
  }
})
