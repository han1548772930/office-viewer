import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'

// Real Ant Button/Tooltip + Canvas/history drafts; execution remains deferred.
// Synthetic composition events below do not validate a native Windows IME.
const fixturePath = '/__document-heading-collapse__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, KeepAlive, ref, shallowRef, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FDocument } from '/packages/office-studio/src/api/facade'
import { createBlankDocument } from '/packages/office-studio/src/domain/model'
import { createParagraph, createTable, collapsedSelection, runText } from '/packages/office-studio/src/features/document/model/document-model'
import { documentHeadingVisibility } from '/packages/office-studio/src/features/document/headings/heading-collapse'
import DocumentWorkbench from '/packages/office-studio/src/components/DocumentWorkbench.vue'
const paragraph = (id,text) => ({ ...createParagraph(text), id })
const heading = (id,text,level=1) => ({ ...paragraph(id,text), type:'heading',level })
const initial = createBlankDocument();initial.id='heading-browser';initial.name='标题折叠';initial.metadata={documentFlavor:2,documentFlavorContract:'univer'}
initial.blocks=[heading('parent','父标题'),paragraph('intro','父标题正文'),heading('child','子标题',2),paragraph('child-body','子标题正文'),heading('next','下一标题'),paragraph('next-body','下一标题正文')]
let studio=new OfficeStudio({document:initial,currentUserId:'editor'});studio.activateDocument(initial.id)
const api=shallowRef(new FDocument(studio,initial.id)),model=shallowRef(api.value.getModel()),workbench=ref(),dark=ref(false),variant=ref('classic'),visible=ref(true)
let stop
function observe(){stop?.();stop=watch(()=>studio.editor.document.model,value=>{model.value=value},{flush:'sync'})}observe()
api.value.setInputSelection(collapsedSelection({blockId:'child-body',offset:2}))
window.headingFixture={
  snapshot:()=>JSON.parse(JSON.stringify(api.value.getModel())),
  hidden:()=>[...documentHeadingVisibility(api.value.getModel()).hiddenOwners.entries()],
  input:()=>api.value.getInputState(),text:()=>api.value.getModel().blocks.map(block=>runText(block.runs)),
  depth:()=>({undo:studio.editor.history.units[api.value.id]?.undo.length??0,redo:studio.editor.history.units[api.value.id]?.redo.length??0}),
  lines:()=>workbench.value?.$.setupState.canvasRef?.getLayout().lines.map(line=>line.target.cellBlockId??line.target.blockId),
  select:id=>api.value.setInputSelection(collapsedSelection({blockId:id,offset:0})),
  fold:(id,collapsed)=>api.value.setHeadingCollapsed(id,collapsed),
  undo:()=>api.value.undo(),redo:()=>api.value.redo(),appearance:(isDark,layout)=>{dark.value=isDark;variant.value=layout},
  readonly:()=>{api.value.executeCommand({id:'doc.command.set-document-permission-rules',rules:[{id:'readonly',scope:'document',action:'edit',effect:'deny',userIds:[]}]},collapsedSelection({blockId:'parent',offset:0}))},
  show:value=>{visible.value=value},
  nested:()=>{
    const next=createBlankDocument();next.id=initial.id;next.name='嵌套大纲';next.metadata={documentFlavor:2,documentFlavorContract:'univer'}
    const title=paragraph('real-title','作者Title');title.metadata.namedStyleType='title'
    const table=createTable(1,1);table.rows[0].cells[0].blocks=[heading('cell-heading','表内标题',2),paragraph('cell-body','表内正文')]
    next.blocks=[title,heading('parent','父标题'),{...paragraph('nested-table',''),type:'table',runs:[],metadata:{table}},
      {...paragraph('columns',''),type:'columns',runs:[],columns:{version:1,layout:'fixed',responsive:'shrink',gap:20,columns:[
        {id:'left',widthRatio:1,blocks:[heading('column-heading','栏内标题',2),paragraph('column-body','栏内正文')]},
        {id:'right',widthRatio:1,blocks:[paragraph('right-body','右栏正文')]}]}},heading('next','下一标题')]
    stop?.();studio.dispose();studio=new OfficeStudio({document:next,currentUserId:'editor'});studio.activateDocument(next.id)
    api.value=new FDocument(studio,next.id);model.value=api.value.getModel();observe();api.value.setInputSelection(collapsedSelection({blockId:'parent',offset:0}))
  },
  saveReopen:()=>{const snapshot=api.value.save();stop?.();studio.dispose();studio=new OfficeStudio({document:snapshot,currentUserId:'editor'});studio.activateDocument(initial.id);api.value=new FDocument(studio,initial.id);model.value=api.value.getModel();observe();api.value.setInputSelection(collapsedSelection({blockId:'parent',offset:0}));return snapshot},
}
const app=createApp({render:()=>h(ConfigProvider,{theme:{algorithm:dark.value?theme.darkAlgorithm:theme.defaultAlgorithm}},{
 default:()=>h(KeepAlive,null,{default:()=>visible.value?h(DocumentWorkbench,{ref:workbench,modelValue:model.value,documentApi:api.value,ribbonVariant:variant.value,canUndo:studio.editor.canUndo,canRedo:studio.editor.canRedo,
  'onUpdate:modelValue':value=>{model.value=value},onUndo:()=>api.value.undo(),onRedo:()=>api.value.redo()}):null}),
})});app.mount('#app');window.addEventListener('pagehide',()=>{stop?.();app.unmount();studio.dispose()},{once:true})
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'heading-collapse-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Heading fixture did not allocate a port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.locator('[data-heading-collapse="parent"]')).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { await page?.close() })

test.each(['classic', 'grid'])('the %s Ant fold button hides actual Canvas lines and undo restores the original hidden caret', async variant => {
  await page.evaluate(value => window.headingFixture.appearance(value === 'grid', value), variant)
  const source = await page.evaluate(() => window.headingFixture.text())
  const parent = page.locator('[data-heading-collapse="parent"]')
  await parent.click()
  await expect(parent).toHaveAttribute('aria-expanded', 'false')
  await expect.poll(() => page.evaluate(() => window.headingFixture.lines().includes('child-body'))).toBe(false)
  assert.equal((await page.evaluate(() => window.headingFixture.input())).location.selection.focus.blockId, 'parent')
  assert.deepEqual(await page.evaluate(() => window.headingFixture.depth()), { undo: 1, redo: 0 })
  await page.evaluate(() => window.headingFixture.undo())
  await expect(parent).toHaveAttribute('aria-expanded', 'true')
  await expect.poll(() => page.evaluate(() => window.headingFixture.lines().includes('child-body'))).toBe(true)
  assert.equal((await page.evaluate(() => window.headingFixture.input())).location.selection.focus.blockId, 'child-body')
  await page.evaluate(() => window.headingFixture.redo())
  await expect(parent).toHaveAttribute('aria-expanded', 'false')
  assert.deepEqual(await page.evaluate(() => window.headingFixture.text()), source)
  assert.deepEqual(errors, [])
})

test('keyboard activation, nested state and public save/reopen preserve the source while refreshing button owners', async () => {
  const child = page.locator('[data-heading-collapse="child"]')
  await child.focus(); await child.press('Enter')
  await expect(child).toHaveAttribute('aria-expanded', 'false')
  await page.locator('[data-heading-collapse="parent"]').click()
  await expect(child).toHaveCount(0)
  const source = await page.evaluate(() => window.headingFixture.text())
  await page.evaluate(() => window.headingFixture.saveReopen())
  const parent = page.locator('[data-heading-collapse="parent"]')
  await expect(parent).toHaveAttribute('aria-expanded', 'false')
  await parent.click()
  await expect(child).toHaveAttribute('aria-expanded', 'false')
  await expect.poll(() => page.evaluate(() => window.headingFixture.lines().includes('child-body'))).toBe(false)
  await child.click()
  await expect.poll(() => page.evaluate(() => window.headingFixture.lines().includes('child-body'))).toBe(true)
  assert.deepEqual(await page.evaluate(() => window.headingFixture.text()), source)
  assert.deepEqual(errors, [])
})

test('composition and readonly state disable the actual button, and hidden workbench has no active fold control', async () => {
  const parent = page.locator('[data-heading-collapse="parent"]'), bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.dispatchEvent('compositionstart', { data: '' })
  await expect(parent).toBeDisabled()
  await expect(page.locator('[data-outline-item="child"]')).toBeDisabled()
  await bridge.dispatchEvent('compositionend', { data: '' })
  await expect(parent).toBeEnabled()
  await page.evaluate(() => window.headingFixture.show(false))
  await expect(parent).toHaveCount(0)
  await page.evaluate(() => window.headingFixture.show(true))
  await expect(parent).toBeVisible()
  await page.evaluate(() => window.headingFixture.readonly())
  await expect(parent).toBeDisabled()
  await page.locator('[data-outline-item="child"]').click()
  assert.equal((await page.evaluate(() => window.headingFixture.input())).location.selection.focus.blockId, 'child')
  assert.deepEqual(errors, [])
})

test('outline navigation uses real Title and cell/column addresses while hidden targets resolve to the visible parent', async () => {
  await page.evaluate(() => window.headingFixture.nested())
  await expect(page.locator('.document-outline__item.is-title')).toHaveCount(1)
  await page.locator('[data-outline-item="real-title"]').click()
  await expect.poll(() => page.evaluate(() => window.headingFixture.input().location.selection.focus.blockId)).toBe('real-title')
  const before = await page.evaluate(() => window.headingFixture.snapshot().blocks)
  await page.locator('[data-outline-item="cell-heading"]').click()
  assert.deepEqual((await page.evaluate(() => window.headingFixture.input())).location.selection.focus, { blockId: 'nested-table', row: 0, col: 0, cellBlockId: 'cell-heading', offset: 0 })
  await expect(page.locator('[data-outline-item="cell-heading"]')).toHaveAttribute('aria-current', 'location')
  await page.locator('[data-outline-item="column-heading"]').click()
  assert.equal((await page.evaluate(() => window.headingFixture.input())).location.selection.focus.blockId, 'column-heading')
  assert.deepEqual(await page.evaluate(() => window.headingFixture.snapshot().blocks), before)
  await page.evaluate(() => window.headingFixture.fold('parent', true))
  await expect.poll(() => page.evaluate(() => window.headingFixture.lines().includes('cell-heading'))).toBe(false)
  await page.locator('[data-outline-item="cell-heading"]').click()
  assert.equal((await page.evaluate(() => window.headingFixture.input())).location.selection.focus.blockId, 'parent')
  assert.equal((await page.evaluate(() => window.headingFixture.snapshot())).blocks.find(block => block.id === 'parent').metadata.headingCollapsed, true)
  assert.deepEqual(await page.evaluate(() => window.headingFixture.depth()), { undo: 1, redo: 0 })
  assert.deepEqual(errors, [])
})
