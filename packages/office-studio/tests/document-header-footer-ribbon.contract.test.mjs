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
  assert.deepEqual(tab.groups.map(group => group.items.length), [1, 1, 1, 1, 1])
  const items = tab.groups.flatMap(group => group.items)
  assert.deepEqual(items.map(item => item.id), [
    'doc-header-footer-insertion',
    'doc-header-footer-navigation',
    'doc-header-footer-options',
    'doc-header-footer-position',
    'doc-header-footer-close',
  ])
  for (const item of items.slice(0, 4)) {
    assert.equal(item.gridLayout.column, 1)
    assert.equal(item.popupCommandId, 'doc.operation.header-footer-control')
    assert.equal(item.popupProps.state, headerFooterState)
  }
  assert.equal(items[4].commandId, 'doc.command.close-header-footer')
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

// Deferred permission presentation matrix, 2026-11-22. Not executed.
for (const [name, state, disabled] of [
  ['story denied and section editable', { ...headerFooterState, disabled: true, sectionDisabled: false }, [true, false, true, false, false]],
  ['story linked and section editable', { ...headerFooterState, contentDisabled: true, sectionDisabled: false }, [true, false, false, false, false]],
  ['section denied', { ...headerFooterState, disabled: true, sectionDisabled: true }, [true, false, true, true, false]],
]) test(`contextual ribbon exposes ${name} consistently to classic, grid and overflow presenters`, () => {
  const tabs = createDocumentRibbon({
    state: toolbarState, canUndo: true, canRedo: false, outline: false, focus: false, fullscreen: false,
    paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true, mode: 'traditional', headerFooter: state,
  })
  const items = tabs.find(tab => tab.id === 'header-footer').groups.flatMap(group => group.items)
  assert.deepEqual(items.map(item => Boolean(item.disabled)), disabled)
  for (const item of items.slice(0, 4)) assert.equal(item.popupProps.state, state)
})
