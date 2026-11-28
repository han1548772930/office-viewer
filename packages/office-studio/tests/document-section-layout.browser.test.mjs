import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred Canvas/antdv-next/public facade drafts, 2026-11-22 through 2026-11-24. No browser run.
// This fixture covers the OfficeStudio field-cache adapter; production app wiring and real Windows IME remain unverified.
const fixturePath = '/__document-section-layout-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Docs 节布局</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, createRun, normalizeDocument, runText } from '/packages/office-studio/src/features/document/model/document-model'
import { EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious } from '/packages/office-studio/src/features/document/services/section-header-footer'
import { DOCUMENT_STORY_KEYS, documentStoryDefaultStyle } from '/packages/office-studio/src/features/document/model/document-story'
import { documentTextFields, setDocumentTextFields } from '/packages/office-studio/src/features/document/model/document-text-fields'
import { hydrateDocumentTocsFromNativeFields } from '/packages/office-studio/src/features/document/toc/document-native-toc'
import { documentTocHighlightLayout } from '/packages/office-studio/src/features/document/toc/document-toc-layout'
import { replaceStoryRange, storyBuffer } from '/packages/office-studio/src/features/document/services/story-editor'
import { createTextMeasurer, layoutDocument } from '/packages/office-studio/src/features/document/layout/document-layout'
import { sceneOffset } from '/packages/office-studio/src/features/document/rendering/document-scene'
import { WorkbenchDisplayKey } from '/packages/office-studio/src/platform/workbench-display'
import { createOfficeMessages, OfficeDirectionKey, OfficeLocaleKey, OfficeTranslatorKey } from '/packages/office-studio/src/i18n/office-i18n'
const copy = value => value === undefined ? undefined : JSON.parse(JSON.stringify(value))
const initial = createBlankDocument(); initial.id = 'section-layout-browser'; initial.name = '节布局编辑'
initial.page = { width: 440, height: 360, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 }
initial.sectionSetup = { columnCount: 2, columnGap: 24, columnSeparator: 'line' }
const first = createParagraph('列'.repeat(240), { fontSize: 12 }); first.id = 'body'
const boundary = { ...createParagraph(), id: 'section-two', type: 'sectionBreak', runs: [], sectionBreak: { type: 'evenPage', setup: { width: 620, height: 420, pageNumberStart: 3 } } }
const last = createParagraph('Second section', { fontSize: 12 }); last.id = 'after'
for (const block of [first, last]) Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
initial.blocks = [first, boundary, last]
initial.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: 'Odd', evenPageHeader: 'Even', differentOddEven: true } }
let studio = new OfficeStudio({ document: normalizeDocument(initial), currentUserId: 'local-user' })
studio.activateDocument(initial.id)
const documentApi = shallowRef(new FDocument(studio, initial.id)), model = ref(documentApi.value.getModel())
const fieldCacheAdapter = shallowRef(studio.getDocumentFieldCacheAdapter(initial.id)), standalone = ref(false), workbenchGeneration = ref(0)
const dark = ref(false), variant = ref('classic'), locale = ref('zh-CN'), notices = [], commands = [], cacheEvents = [], contentEvents = []
const primary = ref('#3064c7'), mountedWorkbench = ref(null), display = ref({ chrome: 'full', zoom: 100 })
let reentrantReplacements = 0
let cancelTocCommands = false
let stopWatching
function observe() {
  stopWatching?.()
  stopWatching = watch(() => studio.editor.document.model, value => { model.value = value }, { flush: 'sync' })
  studio.events.addEvent('CommandExecuted', event => commands.push(copy(event)))
  studio.events.addEvent('BeforeCommandExecute', event => {
    if (cancelTocCommands && ['doc.command.insert-table-of-contents', 'doc.command.update-table-of-contents'].includes(event.id)) event.cancel = true
  })
}
observe()
const bodySelection = () => ({ anchor: { blockId: 'body', offset: 0 }, focus: { blockId: 'body', offset: 0 } })
const layout = () => layoutDocument(model.value, createTextMeasurer())
function resetDocumentHost(source, useStandalone = false, prepare) {
  const next = new OfficeStudio({ document: normalizeDocument(source), currentUserId: 'local-user' }), api = new FDocument(next, source.id)
  next.activateDocument(source.id)
  prepare?.(api)
  const unit = next.editor.history.units[source.id]
  const before = { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, dirty: next.editor.history.isDirty(source.id), input: copy(api.getInputState()) }
  stopWatching?.(); studio.dispose(); studio = next
  documentApi.value = api; model.value = api.getModel(); fieldCacheAdapter.value = studio.getDocumentFieldCacheAdapter(source.id)
  standalone.value = useStandalone; workbenchGeneration.value++; commands.length = 0; cacheEvents.length = 0; contentEvents.length = 0; observe()
  return before
}
function position(pageIndex, x, y) {
  const scroll = document.querySelector('.document-scroll'), bounds = scroll.getBoundingClientRect(), current = layout()
  scroll.scrollTop = Math.max(0, y - 120); scroll.dispatchEvent(new Event('scroll'))
  const offset = sceneOffset(current, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 100 })
  return { page: pageIndex, x: bounds.left + offset.x + x, y: bounds.top + offset.y + y }
}
function liveCanvasLayout() {
  const current = mountedWorkbench.value?.$refs?.canvasRef?.getLayout()
  if (!current) throw new Error('Mounted Docs Canvas layout is not available')
  return current
}
function liveTocTab(blockId) {
  const current = liveCanvasLayout()
  for (const line of current.lines.filter(line => line.target.blockId === blockId)) {
    const glyph = line.glyphs.find(glyph => glyph.tab)
    if (glyph) return { current, line, glyph }
  }
  throw new Error('Mounted TOC row has no right tab')
}
function canvasTabPoint(blockId, fraction, top = false) {
  const { current, line, glyph } = liveTocTab(blockId), scroll = document.querySelector('.document-scroll')
  const scale = display.value.zoom / 100, x = glyph.x + glyph.width * fraction, y = top ? line.y + 2 : line.y + line.height / 2
  scroll.scrollTop = Math.max(0, y * scale - 120); scroll.dispatchEvent(new Event('scroll'))
  const bounds = scroll.getBoundingClientRect(), offset = sceneOffset(current, { width: scroll.clientWidth, height: scroll.clientHeight,
    scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: display.value.zoom })
  return { x: bounds.left + offset.x + x * scale, y: bounds.top + offset.y + y * scale, start: glyph.offset, end: glyph.end }
}
function liveTocGeometry() {
  const current = liveCanvasLayout(), input = documentApi.value.getInputState()
  const highlight = input?.valid && input.location.kind === 'body' ? documentTocHighlightLayout(model.value, current, input.location.selection) : undefined
  return { zoom: display.value.zoom, highlight: highlight && copy(highlight),
    rows: model.value.blocks.filter(block => block.metadata.documentToc?.role === 'entry').map(block => {
      const { line, glyph } = liveTocTab(block.id)
      const field = documentTextFields(block).find(field => field.field === 'PAGEREF')
      const digits = current.lines.filter(line => line.target.blockId === block.id).flatMap(line => line.glyphs)
        .filter(glyph => glyph.offset >= field.startOffset && glyph.end <= field.endOffset)
      const column = current.pages[line.page].sections.find(section => section.fragment === line.sectionFragment).columns[line.column ?? 0]
      return { id: block.id, start: glyph.offset, end: glyph.end, width: glyph.width, leader: glyph.tab.leader, leaderText: glyph.tab.leaderText,
        page: line.page, x: line.x, columnX: column.x, columnWidth: column.width, pageRight: digits.at(-1).x + digits.at(-1).width }
    }) }
}
async function sampleCanvasContext(blockId) {
  const point = canvasTabPoint(blockId, 0.5, true)
  await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  const canvas = document.querySelector('.document-surface canvas'), bounds = canvas?.getBoundingClientRect()
  if (!canvas || !bounds?.width || !bounds.height) throw new Error('Docs Canvas is not ready for a pixel sample')
  const sx = Math.floor((point.x - bounds.left) * canvas.width / bounds.width), sy = Math.floor((point.y - bounds.top) * canvas.height / bounds.height)
  if (sx < 0 || sy < 0 || sx >= canvas.width || sy >= canvas.height) throw new Error('TOC context probe is outside the rendered viewport')
  const scratch = document.createElement('canvas'); scratch.width = scratch.height = 1
  const context = scratch.getContext('2d')
  if (!context) throw new Error('Pixel-sampling Canvas2D context is unavailable')
  context.drawImage(canvas, sx, sy, 1, 1, 0, 0, 1, 1)
  return { rgba: [...context.getImageData(0, 0, 1, 1).data], backend: document.querySelector('.backend-announcement')?.textContent, dpr: window.devicePixelRatio }
}
window.documentSectionFixture = {
  snapshot: () => copy(model.value), input: () => copy(documentApi.value.getInputState()),
  section: (sectionId = 'section-two') => copy(resolveSectionHeaderFooter(model.value, sectionId)),
  links: (sectionId = 'section-two') => {
    const section = resolveSectionHeaderFooter(model.value, sectionId)
    return Object.fromEntries(DOCUMENT_STORY_KEYS.map(story => [story, sectionStoryLinkedToPrevious(section, story)]))
  },
  text: () => runText(model.value.blocks.find(block => block.id === 'body').runs),
  tocText: () => model.value.blocks.filter(block => block.metadata.documentToc).map(block => runText(block.runs)),
  selectBody: (blockId, start = 0, end = start) => documentApi.value.setInputSelection({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } }),
  pointInBlock: blockId => {
    const line = layout().lines.find(line => line.target.blockId === blockId && line.glyphs.length)
    if (!line) throw new Error('Requested block has no Canvas line')
    const glyph = line.glyphs[0]
    return position(line.page, glyph.x + 1, line.y + line.height / 2)
  },
  canvasReady: blockId => {
    const current = mountedWorkbench.value?.$refs?.canvasRef?.getLayout(), block = model.value.blocks.find(block => block.id === blockId)
    return !!block && !!current && current.lines.filter(line => line.target.blockId === blockId).flatMap(line => line.glyphs).map(glyph => glyph.text).join('') === runText(block.runs)
      && ['webgpu', 'canvas2d'].includes(document.querySelector('.backend-announcement')?.textContent)
  },
  canvasTabPoint, canvasGeometry: liveTocGeometry, sampleCanvasContext,
  zoom: value => { display.value = { ...display.value, zoom: value } },
  bodyFields: () => copy(documentTextFields(model.value.blocks.find(block => block.id === 'body'))),
  remount: () => { workbenchGeneration.value++ },
  redo: () => documentApi.value.redo(),
  cacheNotifications: () => ({ cache: cacheEvents.length, content: contentEvents.length }),
  openFieldCommands: () => commands.filter(event => event.id === 'doc.command.refresh-header-footer-fields' && event.params?.command?.onOpen === true).length,
  seedFieldDocument: useStandalone => {
    const source = createBlankDocument(); source.id = 'field-cache-browser'; source.name = '字段缓存'
    const paragraph = createParagraph('old tail'); paragraph.id = 'body'
    setDocumentTextFields(paragraph, [{ id: 'date-year', field: 'DATE', instruction: 'DATE \\\\@ "yyyy"', cachedResult: 'old', startOffset: 0, endOffset: 3 }])
    source.blocks = [paragraph]
    return resetDocumentHost(source, useStandalone, api => {
      api.executeCommand({ id: 'doc.command.insert-text', text: '-draft' }, { anchor: { blockId: 'body', offset: 8 }, focus: { blockId: 'body', offset: 8 } })
      api.undo()
      api.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, { anchor: { blockId: 'body', offset: 0 }, focus: { blockId: 'body', offset: 0 } })
    })
  },
  seedNativeToc: unresolved => {
    const source = createBlankDocument(); source.id = 'native-toc-browser'; source.name = '原生目录'
    source.page = { width: 440, height: 360, marginLeft: 40, marginRight: 40, marginTop: 40, marginBottom: 40 }
    const first = createParagraph(); first.id = 'toc-one'; first.runs = [createRun('Alpha\\t', { fontFamily: 'Georgia', italic: true }), createRun('9', { bold: true, color: '#aa0022' })]
    const second = createParagraph('Beta\\t99'); second.id = 'toc-two'
    for (const row of [first, second]) row.metadata.source = { details: { style: { tabStops: [{ alignment: 'right', leader: 'hyphen', positionTwips: 4800 }] } } }
    const heading = createParagraph('Revised heading'); heading.id = 'body'; heading.type = 'heading'; heading.level = 1
    const other = createParagraph('Beta'); other.id = 'after'; other.type = 'heading'; other.level = 2
    heading.metadata.source = { details: { bookmarks: unresolved ? [] : [{ name: '_TocFirst', childIndex: 0 }] } }
    other.metadata.source = { details: { bookmarks: [{ name: '_TocSecond', childIndex: 0 }] } }
    setDocumentTextFields(first, [
      { id: 'native-toc', field: 'TOC', instruction: 'TOC \\\\o "1-3" \\\\h \\\\z \\\\u', cachedResult: 'Alpha\\t9\\nBeta\\t99', startOffset: 0, endBlockId: second.id, endOffset: 7, containsNestedFields: true },
      { id: 'page-one', field: 'PAGEREF', instruction: 'PAGEREF _TocFirst \\\\h', cachedResult: '9', startOffset: 6, endOffset: 7 },
    ])
    setDocumentTextFields(second, [{ id: 'page-two', field: 'PAGEREF', instruction: 'PAGEREF _TocSecond \\\\h', cachedResult: '99', startOffset: 5, endOffset: 7 }])
    source.blocks = [first, second, heading, { ...createParagraph(), id: 'break', type: 'pageBreak', runs: [] }, other]
    source.contentControls = [{ id: 'wrapper', type: 'group', blockId: first.id, startOffset: 0, endOffset: 7, lock: unresolved ? 'unlocked' : 'sdtLocked' },
      { id: 'page-control', type: 'richText', parentId: 'wrapper', blockId: first.id, startOffset: 6, endOffset: 7 }]
    source.comments = { page: { id: 'page', messages: [], anchor: [{ blockId: first.id, startOffset: 6, endOffset: 7, text: '9' }], referenceText: '9', updatedAt: 'old' } }
    hydrateDocumentTocsFromNativeFields(source)
    return resetDocumentHost(source)
  },
  seedTocInsertion: (level = 1) => {
    const source = createBlankDocument(); source.id = 'toc-insertion-browser'; source.name = '目录插入'
    source.blocks = [{ ...createParagraph('BeforeAfter'), id: 'body' },
      { ...createParagraph('Alpha'), id: 'heading-alpha', type: 'heading', level },
      { ...createParagraph('Beta'), id: 'heading-beta', type: 'heading', level: level + 1 }]
    return resetDocumentHost(source, false, api => api.setInputSelection({ anchor: { blockId: 'body', offset: 6 }, focus: { blockId: 'body', offset: 6 } }))
  },
  cancelTocCommands: value => { cancelTocCommands = value },
  history: () => { const unit = studio.editor.history.units[documentApi.value.id]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, dirty: studio.editor.history.isDirty(documentApi.value.id) } },
  geometry: () => layout().pages.map(page => ({ x: page.x, y: page.y, width: page.width, height: page.height, number: page.pageNumber, blank: page.isParityBlank, sectionId: page.sectionId })),
  storyGlyphs: (pageIndex, area = 'header') => layout().pageHeaderFooters[pageIndex][area + 'LayoutLines'].flatMap(line => line.glyphs)
    .map(glyph => ({ text: glyph.text, start: glyph.offset, end: glyph.end, color: glyph.run.color, bold: glyph.run.bold, italic: glyph.run.italic, underline: glyph.run.underline, href: glyph.run.href })),
  pointInColumn: column => {
    const line = layout().lines.find(line => line.target.blockId === 'body' && line.column === column && line.glyphs.length)
    if (!line) throw new Error('Requested column has no text')
    const glyph = line.glyphs[0]
    return { ...position(line.page, glyph.x + 1, line.y + line.height / 2), offset: glyph.offset }
  },
  pageHeaderPoint: pageIndex => { const page = layout().pages[pageIndex]; return position(pageIndex, page.x + page.width / 2, page.y + 12) },
  setup: (sectionId, patch) => documentApi.value.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch }, bodySelection()),
  permissions: rules => documentApi.value.executeCommand({ id: 'doc.command.set-document-permission-rules', rules }, bodySelection()),
  unlinkEvenStories: () => documentApi.value.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: { storyLinkToPrevious: { evenPageHeader: false, evenPageFooter: false } } }, bodySelection()),
  selectStory: (story, start = 0, end = start) => documentApi.value.setInputSelection(bodySelection(), { sectionId: 'section-two', story, start, end }),
  seedMixedPageField: () => {
    documentApi.value.executeCommand({ id: 'doc.command.set-section-setup', sectionId: 'section-two', patch: { pageNumberStart: 19 } }, bodySelection())
    documentApi.value.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: { storyLinkToPrevious: { evenPageHeader: false } } }, bodySelection())
    const source = documentApi.value.getModel(), section = resolveSectionHeaderFooter(source, 'section-two'), defaults = documentStoryDefaultStyle(source)
    const field = { field: 'PAGE', instruction: 'PAGE', cachedResult: '19', metadata: { retained: { source: 'mixed-cache' }, officeStoryFieldCache: { version: 1, runs: [
      createRun('1', { ...defaults.text, bold: true, color: '#aa0000' }), createRun('9', { ...defaults.text, italic: true, color: '#0000aa', href: 'https://example.org/field' }),
    ] } } }
    return documentApi.value.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: replaceStoryRange(section, 'evenPageHeader', 0, storyBuffer(section, 'evenPageHeader').value.length, field, source) }, bodySelection())
  },
  replaceOnNextStorySelection: () => {
    const api = documentApi.value
    const subscription = api.onInputStateChanged(event => {
      const state = event.state
      if (!state?.valid || state.location.kind !== 'story' || state.location.range.story !== 'evenPageHeader' || state.location.range.start !== 0 || state.location.range.end !== 1) return
      subscription.dispose(); reentrantReplacements++
      const source = api.getModel(), section = resolveSectionHeaderFooter(source, 'section-two')
      api.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'section-two', patch: replaceStoryRange(section, 'evenPageHeader', 0, storyBuffer(section, 'evenPageHeader').value.length, 'Observer replacement', source) }, bodySelection())
    })
  },
  reentrantReplacements: () => reentrantReplacements,
  appearance: (isDark, ribbonVariant, colorPrimary) => { dark.value = isDark; variant.value = ribbonVariant; if (colorPrimary) primary.value = colorPrimary },
  saveReopen: () => {
    const saved = copy(documentApi.value.save()), next = new OfficeStudio({ document: saved, currentUserId: 'local-user' })
    stopWatching?.(); studio.dispose(); studio = next; studio.activateDocument(saved.id)
    documentApi.value = new FDocument(studio, saved.id); model.value = documentApi.value.getModel()
    fieldCacheAdapter.value = studio.getDocumentFieldCacheAdapter(saved.id); workbenchGeneration.value++; observe()
    return saved
  },
  notices: () => [...notices], commands: () => copy(commands),
}
const direction = computed(() => 'ltr'), messages = createOfficeMessages(), translate = (key, fallback) => messages[locale.value][key] ?? fallback
const app = createApp({ render: () => h(ConfigProvider, { direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: primary.value } } }, {
  default: () => h(DocumentWorkbench, { ref: mountedWorkbench, key: workbenchGeneration.value, modelValue: model.value, documentApi: standalone.value ? undefined : documentApi.value,
    fieldCacheAdapter: standalone.value ? fieldCacheAdapter.value : undefined, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { model.value = value }, onDocumentCacheChange: value => cacheEvents.push(copy(value)), onDocumentChange: value => contentEvents.push(copy(value)),
    onUndo: () => documentApi.value.undo(), onRedo: () => documentApi.value.redo(), onNotify: message => notices.push(message) }),
}) })
app.provide(OfficeLocaleKey, locale); app.provide(OfficeDirectionKey, direction); app.provide(OfficeTranslatorKey, translate)
app.provide(WorkbenchDisplayKey, display)
app.mount('#app')
window.addEventListener('pagehide', () => { stopWatching?.(); app.unmount(); studio.dispose() }, { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-section-layout-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const inputState = () => page.evaluate(() => window.documentSectionFixture.input())
async function renderedNativeToc() {
  await expect.poll(() => page.evaluate(() => ['toc-one', 'toc-two'].every(id => window.documentSectionFixture.canvasReady(id)))).toBe(true)
  return page.evaluate(() => window.documentSectionFixture.canvasGeometry())
}
function assertNativeTocStops(geometry) {
  assert.equal(geometry.rows.length, 2)
  for (const row of geometry.rows) {
    assert.equal(row.end - row.start, 1); assert.equal(row.leader, 'dashes'); assert.ok(row.width > 32)
    assert.ok(row.leaderText.length > 0); assert.ok([...row.leaderText].every(character => character === '-'))
    assert.ok(Math.abs(row.pageRight - row.columnX - 320) < 0.001)
  }
}

async function tocInsertMenu(mode) {
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  const selector = '[data-ribbon-item="doc-insert-table-of-contents"]'
  let entry = page.locator('.ribbon-visible-groups ' + selector + ', .ribbon-grid-group:not(.is-hidden) ' + selector)
  if (!await entry.count()) {
    await page.locator('.ribbon-more-button:visible, .ribbon-grid-more-button:visible').click()
    entry = page.locator('.ribbon-overflow-panel:visible ' + selector)
  }
  await entry.getByRole('button').click()
  await page.getByRole('menuitem', { name: mode === 'custom' ? '自定义目录…' : '自动目录', exact: true }).click()
}

test.each(['classic', 'grid'])('custom TOC insertion retains a canceled draft and publishes its canonical command, history and snapshot in %s', async variant => {
  await page.evaluate(layout => { window.documentSectionFixture.seedTocInsertion(); window.documentSectionFixture.appearance(false, layout) }, variant)
  const before = await page.evaluate(() => window.documentSectionFixture.snapshot())
  await tocInsertMenu('custom')
  const dialog = page.getByRole('dialog', { name: '插入目录', exact: true })
  await expect(dialog).toBeVisible()
  const pages = dialog.getByRole('checkbox', { name: '显示页码', exact: true })
  await pages.uncheck()
  await page.evaluate(() => window.documentSectionFixture.cancelTocCommands(true))
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeVisible()
  await expect(dialog.getByRole('alert')).toBeVisible()
  await expect(pages).not.toBeChecked()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.snapshot()), before)
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, 0)
  await page.evaluate(() => window.documentSectionFixture.cancelTocCommands(false))
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha', 'Beta'])
  const inserted = await page.evaluate(() => window.documentSectionFixture.snapshot())
  assert.equal(inserted.blocks.filter(block => block.metadata.documentToc?.role === 'title').length, 0)
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, 1)
  const events = await page.evaluate(() => window.documentSectionFixture.commands())
  assert.equal(events.filter(event => event.id === 'doc.command.insert-table-of-contents').length, 1)
  assert.equal(events.filter(event => event.id === 'doc.command.paste-blocks').length, 0)
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect(bridge).toBeFocused(); await bridge.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.snapshot())).toEqual(before)
  await bridge.press('Control+Shift+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.snapshot())).toEqual(inserted)
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha', 'Beta'])
})

test('a TOC dialog cannot apply its draft to a different caret and does not restore the old input on cancellation', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedTocInsertion())
  await tocInsertMenu('custom')
  const dialog = page.getByRole('dialog', { name: '插入目录', exact: true }), before = await page.evaluate(() => window.documentSectionFixture.snapshot())
  await page.evaluate(() => window.documentSectionFixture.selectBody('body', 1))
  await expect(dialog.getByRole('button', { name: /应\s*用/ })).toBeDisabled()
  await dialog.getByRole('button', { name: /取\s*消/ }).click()
  await expect(dialog).toBeHidden()
  assert.equal((await inputState()).location.selection.focus.offset, 1)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.snapshot()), before)
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, 0)
})

test('custom TOC settings open for level-four-only headings and retain a rejected level-three draft for correction', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedTocInsertion(4))
  await tocInsertMenu('custom')
  const dialog = page.getByRole('dialog', { name: '插入目录', exact: true })
  await expect(dialog).toBeVisible()
  await dialog.getByRole('checkbox', { name: '显示页码', exact: true }).uncheck()
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeVisible()
  await expect(dialog.getByRole('alert')).toContainText('标题')
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, 0)
  await dialog.getByRole('combobox').first().click()
  await page.locator('.ant-select-dropdown:visible').getByText('4', { exact: true }).click()
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha'])
})

test('a same-ID Studio replacement closes the old TOC draft without changing the replacement document', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedTocInsertion())
  await tocInsertMenu('custom')
  const dialog = page.getByRole('dialog', { name: '插入目录', exact: true })
  await expect(dialog).toBeVisible()
  await page.evaluate(() => window.documentSectionFixture.seedTocInsertion())
  await expect(dialog).toBeHidden()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.tocText()), [])
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, 0)
})

test('an update command rejected by BeforeCommand keeps the real TOC dialog available for retry', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedNativeToc(false))
  await renderedNativeToc()
  await page.evaluate(() => window.documentSectionFixture.selectBody('toc-one', 6))
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '更新目录', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '更新目录', exact: true }), before = await page.evaluate(() => window.documentSectionFixture.snapshot())
  await page.evaluate(() => window.documentSectionFixture.cancelTocCommands(true))
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeVisible()
  await expect(dialog.getByRole('alert')).toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.snapshot()), before)
  await page.evaluate(() => window.documentSectionFixture.cancelTocCommands(false))
  await dialog.getByRole('button', { name: /应\s*用/ }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t1', 'Beta\t2'])
})

test('Canvas input in the second column retains logical positions through edit, keyboard history and internal reopen', async () => {
  const before = await page.evaluate(() => window.documentSectionFixture.text())
  const point = await page.evaluate(() => window.documentSectionFixture.pointInColumn(1))
  await page.mouse.click(point.x, point.y)
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect(bridge).toBeFocused(); await expect(bridge).toHaveClass(/ant-input/)
  await expect.poll(async () => (await inputState())?.location.selection.focus.offset).toBe(point.offset)
  await page.keyboard.insertText('编')
  const edited = before.slice(0, point.offset) + '编' + before.slice(point.offset)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(edited)
  await page.keyboard.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(before)
  await page.keyboard.press('Control+Shift+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(edited)
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(edited)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('standalone on-open field caches share the host redo/savepoint and emit only cache notifications across remounts', async () => {
  const before = await page.evaluate(() => window.documentSectionFixture.seedFieldDocument(true))
  assert.equal(before.input.marks.bold, true)
  const year = await page.evaluate(() => String(new Date().getFullYear()))
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(year + ' tail')
  const { input: _input, ...history } = before
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
  assert.deepEqual((await inputState()).marks, before.input.marks)
  assert.equal((await page.evaluate(() => window.documentSectionFixture.snapshot())).modified, false)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.cacheNotifications()), { cache: 1, content: 0 })
  assert.equal(await page.evaluate(() => window.documentSectionFixture.openFieldCommands()), 1)
  await page.evaluate(() => window.documentSectionFixture.remount())
  await expect(page.getByRole('textbox', { name: '文档画布编辑区', exact: true })).toBeVisible()
  assert.equal(await page.evaluate(() => window.documentSectionFixture.openFieldCommands()), 1)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
  await page.evaluate(() => window.documentSectionFixture.redo())
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(year + ' tail-draft')
  assert.equal((await page.evaluate(() => window.documentSectionFixture.bodyFields()))[0].cachedResult, year)
})

test('F9 restores an edited native body result as one undoable command and ignores a composing key event', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedFieldDocument(false))
  const year = await page.evaluate(() => String(new Date().getFullYear())), bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(year + ' tail')
  await page.evaluate(() => window.documentSectionFixture.selectBody('body', 2))
  await expect.poll(async () => (await inputState())?.location.selection.focus.offset).toBe(2)
  await bridge.focus(); await page.keyboard.insertText('x')
  const manual = year.slice(0, 2) + 'x' + year.slice(2) + ' tail'
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(manual)
  const history = await page.evaluate(() => window.documentSectionFixture.history())
  await bridge.evaluate(element => element.dispatchEvent(new KeyboardEvent('keydown', { key: 'F9', isComposing: true, bubbles: true })))
  assert.equal(await page.evaluate(() => window.documentSectionFixture.text()), manual)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
  await bridge.press('F9')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(year + ' tail')
  const refreshed = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(refreshed.undo, history.undo + 1)
  assert.equal((await page.evaluate(() => window.documentSectionFixture.bodyFields()))[0].id, 'date-year')
  await expect(bridge).toBeFocused(); await bridge.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(manual)
  await bridge.press('Control+Shift+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toBe(year + ' tail')
})

test('mounted TOC Canvas tab halves preserve one UTF-16 offset and numeric typing marks through zoom and Ctrl navigation', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedNativeToc(false))
  assertNativeTocStops(await renderedNativeToc())
  const before = await page.evaluate(() => window.documentSectionFixture.snapshot())
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  for (const zoom of [100, 150]) {
    await page.evaluate(value => window.documentSectionFixture.zoom(value), zoom)
    await expect(page.locator('.zoom-value')).toContainText(zoom + '%')
    assertNativeTocStops(await renderedNativeToc())
    const left = await page.evaluate(() => window.documentSectionFixture.canvasTabPoint('toc-one', 0.25))
    await page.mouse.click(left.x, left.y)
    await expect(bridge).toBeFocused()
    await expect.poll(async () => (await inputState()).location.selection.focus.offset).toBe(left.start)
    const right = await page.evaluate(() => window.documentSectionFixture.canvasTabPoint('toc-one', 0.75))
    await page.mouse.click(right.x, right.y)
    await expect.poll(async () => (await inputState()).location.selection.focus.offset).toBe(right.end)
    await expect.poll(async () => (await inputState()).marks.bold).toBe(true)
    assert.equal((await inputState()).marks.color, '#aa0022')
    await bridge.press('ArrowRight')
    await expect.poll(async () => (await inputState()).location.selection.focus.offset).toBe(right.end + 1)
  }
  const point = await page.evaluate(() => window.documentSectionFixture.canvasTabPoint('toc-one', 0.5))
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  await expect.poll(async () => (await inputState()).location.selection.focus.blockId).toBe('body')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.canvasGeometry().highlight?.tocId ?? null)).toBeNull()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.snapshot()), before)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('real Canvas context pixels follow antdv-next primary/theme changes and clear for story or body selection outside the TOC', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedNativeToc(false))
  await renderedNativeToc()
  await page.evaluate(() => window.documentSectionFixture.selectBody('body', 0))
  const before = await page.evaluate(() => window.documentSectionFixture.snapshot())
  const unselected = await page.evaluate(() => window.documentSectionFixture.sampleCanvasContext('toc-one'))
  assert.ok(['webgpu', 'canvas2d'].includes(unselected.backend)); assert.deepEqual(unselected.rgba, [255, 255, 255, 255])
  await page.evaluate(() => window.documentSectionFixture.selectBody('toc-one', 6))
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.canvasGeometry().highlight?.tocId)).toBe('native-toc')
  const selected = await page.evaluate(() => window.documentSectionFixture.sampleCanvasContext('toc-one'))
  assert.notDeepEqual(selected.rgba, unselected.rgba); assert.equal(selected.rgba[3], 255)
  await page.evaluate(() => window.documentSectionFixture.appearance(true, 'grid', '#b42850'))
  const themed = await page.evaluate(() => window.documentSectionFixture.sampleCanvasContext('toc-one'))
  assert.notDeepEqual(themed.rgba, selected.rgba); assert.notDeepEqual(themed.rgba, unselected.rgba)
  const header = await page.evaluate(() => window.documentSectionFixture.pageHeaderPoint(0))
  await page.mouse.dblclick(header.x, header.y)
  const story = page.getByLabel('编辑页眉', { exact: true })
  await expect(story).toBeFocused()
  await expect.poll(async () => (await inputState()).location.kind).toBe('story')
  assert.deepEqual((await page.evaluate(() => window.documentSectionFixture.sampleCanvasContext('toc-one'))).rgba, unselected.rgba)
  await story.press('Escape')
  await expect(page.getByLabel('文档文字输入', { exact: true })).toBeFocused()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.canvasGeometry().highlight?.tocId)).toBe('native-toc')
  assert.notDeepEqual((await page.evaluate(() => window.documentSectionFixture.sampleCanvasContext('toc-one'))).rgba, unselected.rgba)
  const body = await page.evaluate(() => window.documentSectionFixture.pointInBlock('body'))
  await page.mouse.click(body.x, body.y)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.canvasGeometry().highlight?.tocId ?? null)).toBeNull()
  assert.deepEqual((await page.evaluate(() => window.documentSectionFixture.sampleCanvasContext('toc-one'))).rgba, unselected.rgba)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.snapshot()), before)
})

test('TOC page-only dialog recomputes real Canvas gaps through keyboard undo/redo and internal save/reopen', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedNativeToc(false))
  const before = await renderedNativeToc(); assertNativeTocStops(before)
  await page.evaluate(() => window.documentSectionFixture.selectBody('toc-one', 6))
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '更新目录', exact: true }).click()
  const dialog = page.locator('.ant-modal:visible')
  await expect(dialog.locator('.ant-form')).toBeVisible()
  await dialog.getByRole('button', { name: /应\s*用/ }).click(); await expect(dialog).toBeHidden()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t1', 'Beta\t2'])
  const updated = await renderedNativeToc(); assertNativeTocStops(updated)
  assert.ok(updated.rows[1].width > before.rows[1].width + 3)
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect(bridge).toBeFocused(); await bridge.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t9', 'Beta\t99'])
  const undone = await renderedNativeToc(); assertNativeTocStops(undone)
  assert.ok(Math.abs(undone.rows[1].width - before.rows[1].width) < 0.001)
  await bridge.press('Control+Shift+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t1', 'Beta\t2'])
  assertNativeTocStops(await renderedNativeToc())
  const saved = await page.evaluate(() => window.documentSectionFixture.saveReopen())
  assert.equal(JSON.stringify(saved).includes('"leaderText"'), false)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t1', 'Beta\t2'])
  const reopened = await renderedNativeToc(); assertNativeTocStops(reopened)
  assert.ok(Math.abs(reopened.rows[1].width - updated.rows[1].width) < 0.001)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('native TOC full-update dialog preserves numeric anchors and formatting, keyboard history, navigation and internal reopen', async () => {
  await page.evaluate(() => { window.documentSectionFixture.seedNativeToc(false); window.documentSectionFixture.appearance(true, 'grid') })
  await page.evaluate(() => window.documentSectionFixture.selectBody('toc-one', 7, 6))
  const ribbon = page.getByTestId('workbench-ribbon'), update = ribbon.getByRole('button', { name: '更新目录', exact: true })
  await expect(update).toBeEnabled(); await update.click()
  const dialog = page.locator('.ant-modal:visible')
  await expect(dialog.locator('.ant-form')).toBeVisible(); await dialog.getByRole('combobox').click()
  await page.locator('.ant-select-dropdown:visible').getByText('更新整个目录', { exact: true }).click()
  const apply = dialog.getByRole('button', { name: /应\s*用/ })
  await expect(apply).toBeEnabled(); await apply.click(); await expect(dialog).toBeHidden()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Revised heading\t1', 'Beta\t2'])
  const updatedCanvas = await renderedNativeToc(); assertNativeTocStops(updatedCanvas)
  assert.ok(Math.abs(updatedCanvas.rows[1].x - updatedCanvas.rows[1].columnX - 18) < 0.001)
  const updated = await page.evaluate(() => window.documentSectionFixture.snapshot()), start = 'Revised heading\t'.length
  const control = updated.contentControls.find(control => control.id === 'page-control')
  assert.deepEqual([control.startOffset, control.endOffset], [start, start + 1])
  assert.deepEqual(updated.comments.page.anchor, [{ blockId: 'toc-one', startOffset: start, endOffset: start + 1, text: '1' }])
  assert.equal(updated.blocks[0].runs.find(run => run.text === '1').bold, true)
  assert.equal(updated.blocks[0].runs.find(run => run.text === '1').color, '#aa0022')
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect(bridge).toBeFocused(); await bridge.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t9', 'Beta\t99'])
  assertNativeTocStops(await renderedNativeToc())
  await bridge.press('Control+Shift+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Revised heading\t1', 'Beta\t2'])
  assertNativeTocStops(await renderedNativeToc())
  const point = await page.evaluate(() => window.documentSectionFixture.pointInBlock('toc-one'))
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  await expect.poll(async () => (await inputState())?.location.selection.focus.blockId).toBe('body')
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Revised heading\t1', 'Beta\t2'])
  assertNativeTocStops(await renderedNativeToc())
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('unresolved native TOC targets disable update while antdv-next Ribbon and Layout menu keep deletion executable', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedNativeToc(true))
  await page.evaluate(() => window.documentSectionFixture.selectBody('toc-one', 0))
  const ribbon = page.getByTestId('workbench-ribbon')
  await expect(ribbon.getByRole('button', { name: '更新目录', exact: true })).toBeDisabled()
  await expect(ribbon.getByRole('button', { name: '删除目录', exact: true })).toBeEnabled()
  const point = await page.evaluate(() => window.documentSectionFixture.pointInBlock('toc-one'))
  await page.mouse.click(point.x, point.y, { button: 'right' })
  const menu = page.locator('.document-context-menu .ant-menu')
  await expect(menu.getByRole('menuitem', { name: '更新目录', exact: true })).toHaveClass(/ant-menu-item-disabled/)
  const remove = menu.getByRole('menuitem', { name: '删除目录', exact: true })
  await expect(remove).not.toHaveClass(/ant-menu-item-disabled/); await remove.click()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual([])
  const history = await page.evaluate(() => window.documentSectionFixture.history()); assert.equal(history.undo, 1)
  const bridge = page.getByLabel('文档文字输入', { exact: true })
  await expect(bridge).toBeFocused(); await bridge.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.tocText())).toEqual(['Alpha\t9', 'Beta\t99'])
})

test('mixed page geometry stays clickable after a section change and theme/Ribbon switching', async () => {
  await page.evaluate(() => { window.documentSectionFixture.setup('first', { columnCount: 1, marginLeft: 50, marginRight: 50 }); window.documentSectionFixture.appearance(true, 'grid') })
  await page.setViewportSize({ width: 980, height: 760 })
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  assert.equal(geometry[0].width, 440)
  assert.equal(geometry.at(-1).width, 620)
  assert.equal(geometry[0].x - geometry.at(-1).x, 90)
  const point = await page.evaluate(() => window.documentSectionFixture.pointInColumn(0))
  await page.mouse.click(point.x, point.y)
  await expect(page.getByLabel('文档文字输入', { exact: true })).toBeFocused()
  await expect.poll(async () => (await inputState())?.location.selection.focus.offset).toBe(point.offset)
  await page.keyboard.insertText('改')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.text())).toMatch(/^改列/)
})

test('parity filler pages reject story activation and logical even pages open the even-header story', async () => {
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry()), blank = geometry.findIndex(page => page.blank)
  assert.ok(blank > 0)
  const point = await page.evaluate(index => window.documentSectionFixture.pageHeaderPoint(index), blank)
  await page.mouse.dblclick(point.x, point.y)
  await expect(page.getByLabel('编辑页眉', { exact: true })).toHaveCount(0)
  const actual = geometry.findIndex((entry, index) => index > blank && !entry.blank), header = await page.evaluate(index => window.documentSectionFixture.pageHeaderPoint(index), actual)
  assert.equal(geometry[actual].number, 4)
  await page.mouse.dblclick(header.x, header.y)
  await expect(page.getByLabel('编辑页眉', { exact: true })).toBeFocused()
  await expect(page.getByLabel('编辑页眉', { exact: true })).toHaveValue('Even')
})

test('section setup survives save/reopen with the same page numbers, filler pages and physical rectangles', async () => {
  await page.evaluate(() => window.documentSectionFixture.setup('section-two', { width: 500, height: 460, marginTop: 50, pageNumberStart: 9 }))
  const before = await page.evaluate(() => window.documentSectionFixture.geometry())
  const saved = await page.evaluate(() => window.documentSectionFixture.saveReopen())
  const after = await page.evaluate(() => window.documentSectionFixture.geometry())
  assert.deepEqual(after, before)
  assert.equal(saved.blocks.find(block => block.id === 'section-two').sectionBreak.setup.pageNumberStart, 9)
  assert.equal(after.at(-1).number, 10)
  assert.equal(after.at(-1).height, 460)
})

async function openSecondSectionHeader() {
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  const index = geometry.findIndex(entry => entry.sectionId === 'section-two' && !entry.blank)
  assert.ok(index >= 0)
  const point = await page.evaluate(index => window.documentSectionFixture.pageHeaderPoint(index), index)
  await page.mouse.dblclick(point.x, point.y)
  await expect(page.getByLabel('编辑页眉', { exact: true })).toBeFocused()
  return index
}

test('header/footer navigation shows independent associations and Canvas edits survive undo and internal reopen', async () => {
  await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true }), link = page.getByRole('checkbox', { name: '链接到前一节', exact: true })
  await expect(header).toHaveValue('Even'); await expect(header).toHaveJSProperty('readOnly', true)
  const linkedHistory = await page.evaluate(() => window.documentSectionFixture.history())
  await page.keyboard.press('Control+b')
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), linkedHistory)
  await expect(link).toBeChecked(); await link.uncheck()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.links().evenPageHeader)).toBe(false)
  await expect(header).toHaveJSProperty('readOnly', false); await expect(header).toBeFocused()
  await page.keyboard.insertText('Own ')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Own Even')
  await page.getByRole('button', { name: '转到页脚', exact: true }).click()
  const footer = page.getByLabel('编辑页脚', { exact: true })
  await expect(footer).toBeFocused(); await expect(footer).toHaveJSProperty('readOnly', true); await expect(link).toBeChecked()
  await link.uncheck(); await expect(footer).toHaveJSProperty('readOnly', false); await expect(footer).toBeFocused()
  await page.keyboard.insertText('Own footer')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Own footer')
  const links = await page.evaluate(() => window.documentSectionFixture.links())
  assert.deepEqual(links, { header: true, footer: true, firstPageHeader: true, firstPageFooter: true, evenPageHeader: false, evenPageFooter: false })
  await page.getByRole('button', { name: '转到页眉', exact: true }).click()
  await expect(header).toHaveValue('Own Even'); await expect(link).not.toBeChecked()
  await link.check(); await expect(header).toHaveValue('Even'); await expect(header).toHaveJSProperty('readOnly', true)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Own footer')
  await expect(header).toBeFocused(); await page.keyboard.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Own Even')
  const saved = await page.evaluate(() => window.documentSectionFixture.saveReopen())
  assert.equal(saved.sectionHeaderFooters['section-two'].storyLinkToPrevious.evenPageHeader, false)
  assert.equal(saved.sectionHeaderFooters['section-two'].storyLinkToPrevious.evenPageFooter, false)
  await page.evaluate(() => window.documentSectionFixture.appearance(true, 'grid'))
  await openSecondSectionHeader(); await expect(header).toHaveValue('Own Even'); await expect(link).not.toBeChecked()
  await expect(header).toHaveClass(/ant-input/)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Own footer')
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('first-page options retarget the public input story and keep variant associations separate', async () => {
  await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true }), first = page.getByRole('checkbox', { name: '首页不同', exact: true })
  const link = page.getByRole('checkbox', { name: '链接到前一节', exact: true })
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('evenPageHeader')
  await first.check()
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('firstPageHeader')
  await expect(header).toHaveValue(''); await expect(link).toBeChecked()
  await link.uncheck(); await expect(header).toHaveJSProperty('readOnly', false); await expect(header).toBeFocused()
  await page.keyboard.insertText('Cover')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().firstPageHeader)).toBe('Cover')
  let links = await page.evaluate(() => window.documentSectionFixture.links())
  assert.equal(links.firstPageHeader, false); assert.equal(links.evenPageHeader, true); assert.equal(links.header, true)
  await first.uncheck()
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('evenPageHeader')
  await expect(header).toHaveValue('Even'); await expect(header).toHaveJSProperty('readOnly', true); await expect(link).toBeChecked()
  await first.check()
  await expect.poll(async () => (await inputState())?.location.range.story).toBe('firstPageHeader')
  await expect(header).toHaveValue('Cover'); await expect(link).not.toBeChecked()
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await openSecondSectionHeader(); await expect(header).toHaveValue('Cover')
  links = await page.evaluate(() => window.documentSectionFixture.links())
  assert.equal(links.firstPageHeader, false); assert.equal(links.evenPageHeader, true)
})

test('Canvas live field glyphs keep mixed cache formats through a story shortcut, undo/redo and internal reopen', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField())
  const index = await openSecondSectionHeader(), header = page.getByLabel('编辑页眉', { exact: true })
  const geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  assert.equal(geometry[index].number, 20)
  await expect(header).toHaveValue('\uFFFC'); await expect(header).toHaveJSProperty('readOnly', false)
  const before = await page.evaluate(() => window.documentSectionFixture.section().pageFields)
  const glyphs = () => page.evaluate(index => window.documentSectionFixture.storyGlyphs(index), index)
  assert.deepEqual((await glyphs()).map(glyph => [glyph.text, glyph.start, glyph.end, glyph.color, glyph.bold, glyph.italic]), [
    ['2', 0, 1, '#aa0000', true, false], ['0', 0, 1, '#0000aa', false, true],
  ])
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageHeader', 0, 1))
  await expect.poll(async () => (await inputState())?.location.range.end).toBe(1); await expect(header).toBeFocused()
  await page.keyboard.press('Control+b')
  await expect.poll(async () => (await glyphs()).map(glyph => glyph.bold)).toEqual([true, true])
  const formatted = await page.evaluate(() => window.documentSectionFixture.section().pageFields)
  assert.equal(formatted[0].cachedResult, '19'); assert.equal(formatted[0].instruction, 'PAGE')
  assert.equal(formatted[0].metadata.officeStoryFieldCache.runs.map(run => run.text).join(''), '19')
  assert.equal((await glyphs())[1].href, 'https://example.org/field')
  await expect(header).toHaveValue('\uFFFC')
  await page.keyboard.press('Control+z')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().pageFields)).toEqual(before)
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageHeader', 0, 1)); await expect(header).toBeFocused()
  await page.keyboard.press('Control+y')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().pageFields)).toEqual(formatted)
  await page.evaluate(() => window.documentSectionFixture.saveReopen())
  await page.evaluate(() => window.documentSectionFixture.appearance(true, 'grid')); await openSecondSectionHeader()
  await expect(header).toHaveValue('\uFFFC'); await expect(header).toHaveClass(/ant-input/)
  assert.deepEqual((await glyphs()).map(glyph => [glyph.text, glyph.color, glyph.bold, glyph.italic]), [
    ['2', '#aa0000', true, false], ['0', '#0000aa', true, true],
  ])
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.section().pageFields), formatted)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), { undo: 0, redo: 0, dirty: false })
})

test('contextual PAGE and NUMPAGES insertion writes the current layout values into new field caches', async () => {
  const index = await openSecondSectionHeader(), geometry = await page.evaluate(() => window.documentSectionFixture.geometry())
  await page.getByRole('checkbox', { name: '链接到前一节', exact: true }).uncheck()
  const pageButton = page.getByRole('button', { name: '页码', exact: true })
  await expect(pageButton).toBeEnabled(); await pageButton.click()
  await page.getByRole('menuitem', { name: '当前页码', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().pageFields.find(field => field.field === 'PAGE')?.cachedResult)).toBe(String(geometry[index].number))
  await pageButton.click(); await page.getByRole('menuitem', { name: '总页数', exact: true }).click()
  const fields = await page.evaluate(() => window.documentSectionFixture.section().pageFields.filter(field => field.story === 'evenPageHeader'))
  assert.deepEqual(fields.map(field => [field.field, field.instruction, field.cachedResult]), [
    ['PAGE', 'PAGE', String(geometry[index].number)], ['NUMPAGES', 'NUMPAGES', String(geometry.length)],
  ])
  await expect(page.getByLabel('编辑页眉', { exact: true })).toHaveValue('\uFFFC\uFFFCEven')
})

for (const [shortcut, property] of [['Control+i', 'italic'], ['Control+u', 'underline']]) test(`story ${shortcut} updates field formats and ignores composition key events`, async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true })
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageHeader', 0, 1)); await expect(header).toBeFocused()
  const before = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => document.querySelector('textarea[aria-label="编辑页眉"]').dispatchEvent(new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, isComposing: true, bubbles: true })))
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), before)
  await page.keyboard.press(shortcut)
  await expect.poll(() => page.evaluate(property => window.documentSectionFixture.section().pageFields[0].metadata.officeStoryFieldCache.runs.map(run => run[property]), property)).toEqual([true, true])
  assert.equal((await page.evaluate(() => window.documentSectionFixture.section())).pageFields[0].cachedResult, '19')
  await expect(header).toHaveValue('\uFFFC')
  assert.equal((await page.evaluate(() => window.documentSectionFixture.history())).undo, before.undo + 1)
})

const deniedEvenHeader = [{ id: 'deny-current-header', scope: 'entity', targetId: 'section-two:evenPageHeader', action: 'edit', effect: 'deny', userIds: [] }]

test('read-only header retains section position controls and can navigate to an editable footer', async () => {
  await page.evaluate(rules => { window.documentSectionFixture.unlinkEvenStories(); window.documentSectionFixture.permissions(rules) }, deniedEvenHeader)
  await openSecondSectionHeader()
  const header = page.getByLabel('编辑页眉', { exact: true })
  await expect(header).toHaveJSProperty('readOnly', true)
  const beforeShortcut = await page.evaluate(() => window.documentSectionFixture.history())
  await page.keyboard.press('Control+b')
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), beforeShortcut)
  await expect(page.getByRole('button', { name: '页码', exact: true })).toBeDisabled()
  await expect(page.getByRole('checkbox', { name: '首页不同', exact: true })).toBeDisabled()
  const margin = page.locator('[data-control-kind="position"]:visible').getByRole('spinbutton').first()
  await expect(margin).toBeEnabled(); await margin.fill('32'); await margin.press('Tab')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.snapshot().blocks.find(block => block.id === 'section-two').sectionBreak.setup.marginHeader)).toBe(32)
  const navigate = page.getByRole('button', { name: '转到页脚', exact: true })
  await expect(navigate).toBeEnabled(); await navigate.click()
  const footer = page.getByLabel('编辑页脚', { exact: true })
  await expect(footer).toHaveJSProperty('readOnly', false); await expect(footer).toBeFocused()
  await page.keyboard.insertText('Allowed footer')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('Allowed footer')
  assert.equal((await page.evaluate(() => window.documentSectionFixture.section())).evenPageHeader, 'Even')
})

test('narrow classic More keeps read-only story navigation available with the shared dark theme', async () => {
  await page.evaluate(rules => { window.documentSectionFixture.unlinkEvenStories(); window.documentSectionFixture.permissions(rules); window.documentSectionFixture.appearance(true, 'classic') }, deniedEvenHeader)
  await page.setViewportSize({ width: 240, height: 900 }); await openSecondSectionHeader()
  const more = page.locator('.ribbon-more-button')
  await expect(more).toBeVisible(); await more.click()
  const overflow = page.locator('.ribbon-overflow-panel')
  await expect(overflow).toBeVisible()
  const navigate = overflow.getByRole('button', { name: '转到页脚', exact: true })
  await expect(navigate).toBeEnabled(); await navigate.click()
  const footer = page.getByLabel('编辑页脚', { exact: true })
  await expect(footer).toHaveJSProperty('readOnly', false); await expect(footer).toBeFocused()
  await page.keyboard.insertText('More footer')
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageFooter)).toBe('More footer')
})

test('date and page popups close on story or model changes and reject queued actions from their old scope', async () => {
  await page.evaluate(() => window.documentSectionFixture.unlinkEvenStories()); await openSecondSectionHeader()
  const date = page.getByRole('button', { name: '日期和时间', exact: true }), dialog = page.getByRole('dialog', { name: '日期和时间', exact: true })
  const before = await page.evaluate(() => window.documentSectionFixture.section()), history = await page.evaluate(() => window.documentSectionFixture.history())
  await date.click(); await expect(dialog).toBeVisible()
  await page.evaluate(() => window.documentSectionFixture.selectStory('evenPageFooter'))
  await expect(dialog).not.toBeVisible(); await expect(page.getByLabel('编辑页脚', { exact: true })).toBeFocused()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.section()), before)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
  await date.click(); await expect(dialog).toBeVisible()
  await page.evaluate(() => {
    const oldOk = document.querySelector('.ant-modal [role="button"].ant-btn-primary, .ant-modal button.ant-btn-primary')
    if (!oldOk) throw new Error('Date modal has no OK button')
    window.documentSectionFixture.setup('section-two', { columnGap: 32 }); oldOk.click()
  })
  await expect(dialog).not.toBeVisible()
  let currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(currentHistory.undo, history.undo + 1)
  assert.deepEqual((await page.evaluate(() => window.documentSectionFixture.section())).pageFields, before.pageFields)
  await page.getByRole('button', { name: '页码', exact: true }).click()
  await expect(page.getByRole('menuitem', { name: '当前页码', exact: true })).toBeVisible()
  await page.evaluate(() => {
    const oldItem = Array.from(document.querySelectorAll('.ant-menu-item')).find(item => item.textContent.trim() === '当前页码')
    if (!oldItem) throw new Error('Page popup has no PAGE item')
    window.documentSectionFixture.setup('section-two', { columnGap: 36 }); oldItem.click()
  })
  await expect(page.getByRole('menuitem', { name: '当前页码', exact: true })).not.toBeVisible()
  currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(currentHistory.undo, history.undo + 2)
  assert.deepEqual((await page.evaluate(() => window.documentSectionFixture.section())).pageFields, before.pageFields)
})

test('a synchronous public selection observer replaces a story without letting an older date action overwrite it', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  await page.getByRole('button', { name: '日期和时间', exact: true }).click()
  await expect(page.getByRole('dialog', { name: '日期和时间', exact: true })).toBeVisible()
  const history = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => {
    window.documentSectionFixture.replaceOnNextStorySelection()
    const header = document.querySelector('textarea[aria-label="编辑页眉"]'), oldOk = document.querySelector('.ant-modal button.ant-btn-primary')
    if (!header || !oldOk) throw new Error('Reentrancy fixture is missing its story input or date button')
    header.setSelectionRange(0, 1); oldOk.click()
  })
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.reentrantReplacements())).toBe(1)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Observer replacement')
  const section = await page.evaluate(() => window.documentSectionFixture.section()), currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.deepEqual(section.pageFields?.filter(field => field.story === 'evenPageHeader') ?? [], [])
  assert.equal(currentHistory.undo, history.undo + 1)
})

test('an older date popup cannot rebind to a native story selection that has not yet emitted its select event', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  await page.getByRole('button', { name: '日期和时间', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '日期和时间', exact: true })
  await expect(dialog).toBeVisible()
  const section = await page.evaluate(() => window.documentSectionFixture.section()), history = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => {
    const header = document.querySelector('textarea[aria-label="编辑页眉"]'), oldOk = document.querySelector('.ant-modal button.ant-btn-primary')
    if (!header || !oldOk) throw new Error('Popup fixture is missing its story input or date button')
    header.setSelectionRange(0, 1); oldOk.click()
  })
  await expect(dialog).not.toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.section()), section)
  assert.deepEqual(await page.evaluate(() => window.documentSectionFixture.history()), history)
})

test('a story format shortcut rejects the old source after a synchronous public selection observer replaces it', async () => {
  await page.evaluate(() => window.documentSectionFixture.seedMixedPageField()); await openSecondSectionHeader()
  const history = await page.evaluate(() => window.documentSectionFixture.history())
  await page.evaluate(() => {
    window.documentSectionFixture.replaceOnNextStorySelection()
    const header = document.querySelector('textarea[aria-label="编辑页眉"]')
    if (!header) throw new Error('Shortcut fixture has no story input')
    header.setSelectionRange(0, 1)
    header.dispatchEvent(new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, bubbles: true, cancelable: true }))
  })
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.reentrantReplacements())).toBe(1)
  await expect.poll(() => page.evaluate(() => window.documentSectionFixture.section().evenPageHeader)).toBe('Observer replacement')
  const section = await page.evaluate(() => window.documentSectionFixture.section()), currentHistory = await page.evaluate(() => window.documentSectionFixture.history())
  assert.equal(currentHistory.undo, history.undo + 1)
  assert.deepEqual(section.pageFields?.filter(field => field.story === 'evenPageHeader') ?? [], [])
})
