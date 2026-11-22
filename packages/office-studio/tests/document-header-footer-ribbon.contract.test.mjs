import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let createDocumentRibbon

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  ;({ createDocumentRibbon } = await server.ssrLoadModule('/packages/office-studio/src/components/document/document-ribbon.ts'))
})

afterAll(async () => { await server?.close() })

test('header/footer popup component compiles through its normal Vue module entry', async () => {
  const module = await server.ssrLoadModule('/packages/office-studio/src/components/document/DocumentHeaderFooterRibbonControl.vue')
  assert.equal(typeof module.default.setup, 'function')
  assert.equal(typeof module.default.ssrRender, 'function')
})

const toolbarState = { marks: {}, mixedMarks: [], paragraph: {}, hasSelection: false }
const headerFooterState = {
  visible: true,
  area: 'header',
  disabled: false,
  canLinkToPrevious: true,
  linkedToPrevious: false,
  differentFirstPage: false,
  differentOddEven: false,
  marginHeader: 24,
  marginFooter: 24,
}

test('header/footer contextual ribbon follows Univer groups and carries live section state', () => {
  const tabs = createDocumentRibbon({
    state: toolbarState,
    canUndo: true,
    canRedo: false,
    outline: false,
    focus: false,
    fullscreen: false,
    paper: 'A4',
    orientation: 'portrait',
    zoom: 100,
    contentControlsAvailable: true,
    mode: 'traditional',
    headerFooter: headerFooterState,
  })
  const tab = tabs.find(candidate => candidate.id === 'header-footer')
  assert.equal(tab?.contextual, true)
  assert.deepEqual(tab?.groups[0]?.items.map(item => item.id), [
    'doc-header-footer-insertion',
    'doc-header-footer-navigation',
    'doc-header-footer-options',
    'doc-header-footer-position',
    'doc-header-footer-close',
  ])
  for (const item of tab.groups[0].items.slice(0, 4)) {
    assert.equal(item.popupCommandId, 'doc.operation.header-footer-control')
    assert.equal(item.popupProps.state, headerFooterState)
  }
  assert.equal(tab.groups[0].items[4].commandId, 'doc.operation.close-header-footer')
})

test('header/footer contextual ribbon is absent outside an active header/footer edit area', () => {
  const tabs = createDocumentRibbon({
    state: toolbarState,
    canUndo: true,
    canRedo: false,
    outline: false,
    focus: false,
    fullscreen: false,
    paper: 'A4',
    orientation: 'portrait',
    zoom: 100,
    contentControlsAvailable: true,
    mode: 'traditional',
    headerFooter: { ...headerFooterState, visible: false },
  })
  assert.equal(tabs.some(candidate => candidate.id === 'header-footer'), false)
})
