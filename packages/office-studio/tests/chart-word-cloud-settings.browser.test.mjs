import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'

// Browser/UI drafts only; no server, browser, compiler, runner or sample acceptance ran.
const fixturePath = '/__word-cloud-settings-resource-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"></head><body><div id="app"></div><script type="module">
import { createApp, h, nextTick, ref } from 'vue'
import { Button, ConfigProvider, Form, theme } from 'antdv-next'
import ChartWordCloudSettings from '/packages/office-studio/src/features/charts/views/ChartWordCloudSettings.vue'
import { createChartWordCloudCanvas, drawChartWordCloudMaskImage } from '/packages/office-studio/src/features/charts/chart-word-cloud-mask.ts'
const empty = () => ({ type:'wordCloud',source:{range:{startRow:0,startCol:0,endRow:1,endCol:1},orientation:'columns'},position:{x:0,y:0},size:{width:560,height:400},options:{} })
const config=ref(empty()),active=ref(true),session=ref(1),loading=ref(false),dark=ref(false),updates=[]
const app=createApp({setup:()=>()=>h(ConfigProvider,{theme:{algorithm:dark.value?theme.darkAlgorithm:theme.defaultAlgorithm}},
  {default:()=>h(Form,{layout:'vertical'}, {default:()=>[
    h(ChartWordCloudSettings,{config:config.value,active:active.value,sessionKey:session.value,
      'onUpdate:config':value=>{config.value=value;updates.push(value)},'onResource-loading':value=>loading.value=value}),
    h(Button,{disabled:loading.value,'aria-label':'应用词云草稿'},()=> '应用')
  ]})})})
app.mount('#app')
let deferredReader,continueReading,originalReader=window.FileReader
window.wordCloudSettingsFixture={config,loading,updates,close:async()=>{active.value=false;await nextTick()},
  replace:async()=>{config.value=empty();session.value++;await nextTick()},
  dark:async()=>{dark.value=!dark.value;await nextTick()},dispose:()=>app.unmount(),
  delayRead:()=>{window.FileReader=class extends originalReader{readAsDataURL(file){deferredReader=this;continueReading=()=>super.readAsDataURL(file)}}},
  finishRead:()=>{continueReading?.();window.FileReader=originalReader},
  released:()=>deferredReader?.onload===null&&deferredReader?.onerror===null&&deferredReader?.onabort===null,
  image:()=>{const c=document.createElement('canvas');c.width=64;c.height=64;const d=c.getContext('2d');d.fillStyle='#000';d.beginPath();d.moveTo(32,2);d.lineTo(62,62);d.lineTo(2,62);d.closePath();d.fill();return c.toDataURL('image/png')},
  decode:async()=>{const owner=createChartWordCloudCanvas(64,64);try{const options=config.value.options.wordCloud;
    await drawChartWordCloudMaskImage(owner,options.maskImageId,{resolveImageResource:()=>options.maskImageResource});
    return {center:owner.context.getImageData(32,40,1,1).data[3],corner:owner.context.getImageData(1,1,1,1).data[3]}
  }finally{owner.release()}}
}
</script></body></html>`
let server, browser, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), { name: 'word-cloud-settings-resource-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
      try { response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture)) }
      catch (error) { next(error) }
    }) } }], resolve: { alias: { vue: resolve(process.cwd(), 'packages/office-studio/node_modules/vue'),
      'antdv-next': resolve(process.cwd(), 'packages/office-studio/node_modules/antdv-next') } },
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate a port')
  url = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
})
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.goto(url); await page.waitForFunction(() => Boolean(window.wordCloudSettingsFixture))
  await page.getByRole('combobox', { name: '词云轮廓来源' }).click()
  await page.getByText('图片资源', { exact: true }).last().click()
})
afterEach(async () => {
  await page?.evaluate(() => window.wordCloudSettingsFixture?.dispose())
  await page?.close(); assert.deepEqual(errors, [])
})
afterAll(async () => { await browser?.close(); await server?.close() })
async function upload() {
  const source = await page.evaluate(() => window.wordCloudSettingsFixture.image())
  await page.locator('input[type=file]').setInputFiles({ name: 'triangle.png', mimeType: 'image/png', buffer: Buffer.from(source.split(',')[1], 'base64') })
}

test('antdv-next Upload stores a detached real PNG resource and actual bitmap loading uses its ID', async () => {
  const requests = []; page.on('request', request => requests.push(request.url()))
  await upload()
  await page.waitForFunction(() => !window.wordCloudSettingsFixture.loading.value && window.wordCloudSettingsFixture.config.value.options.wordCloud?.maskImageResource)
  const result = await page.evaluate(async () => {
    const f = window.wordCloudSettingsFixture, options = f.config.value.options.wordCloud
    return { sameId: options.maskImageId === options.maskImageResource.id, sourceType: options.maskImageResource.sourceType,
      image: await f.decode(), count: f.updates.length }
  })
  assert.equal(result.sameId, true); assert.equal(result.sourceType, 'dataUrl')
  assert.equal(result.image.center, 255); assert.equal(result.image.corner, 0); assert.equal(result.count, 1)
  assert.equal(requests.some(value => /triangle\.png$/.test(value)), false)
  assert.equal(await page.getByRole('button', { name: '应用词云草稿' }).isDisabled(), false)
})

for (const action of ['close', 'replace']) test(`closing/replacing a session cancels the FileReader before ${action} can receive stale bytes`, async () => {
  await page.evaluate(() => window.wordCloudSettingsFixture.delayRead()); await upload()
  await page.waitForFunction(() => window.wordCloudSettingsFixture.loading.value)
  assert.equal(await page.getByRole('button', { name: '应用词云草稿' }).isDisabled(), true)
  await page.evaluate(async value => { const f = window.wordCloudSettingsFixture; await f[value](); f.finishRead() }, action)
  await page.waitForFunction(() => !window.wordCloudSettingsFixture.loading.value)
  const result = await page.evaluate(() => { const f = window.wordCloudSettingsFixture; return { updates: f.updates.length,
    image: f.config.value.options.wordCloud?.maskImageResource, released: f.released() } })
  assert.equal(result.updates, 0); assert.equal(result.image, undefined); assert.equal(result.released, true)
})

test('resource input retains ConfigProvider theme and focus while the live draft is edited', async () => {
  const input = page.getByRole('textbox', { name: '词云遮罩图片资源' })
  await input.fill('./mask.png'); await input.focus()
  const light = await input.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.wordCloudSettingsFixture.dark())
  const dark = await input.evaluate(element => getComputedStyle(element).backgroundColor)
  assert.notEqual(dark, light)
  assert.equal(await input.evaluate(element => document.activeElement === element), true)
  assert.equal(await input.inputValue(), './mask.png')
})
