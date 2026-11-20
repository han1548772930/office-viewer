import assert from 'node:assert/strict'
import { test } from '@playwright/test'

test('Sheets startup canvas exposes its actual backend and renders visible grid pixels', async ({ page }, testInfo) => {
  const pageErrors = []
  page.on('pageerror', (error) => pageErrors.push(error.message))

  await page.goto('/', { waitUntil: 'networkidle' })
  const grid = page.locator('[role="grid"][aria-label="工作表画布"]')
  await grid.waitFor({ state: 'visible' })
  const canvas = page.locator('canvas[data-render-surface="true"]')
  await canvas.waitFor({ state: 'attached' })
  await page.waitForFunction(() => {
    const target = document.querySelector('canvas[data-render-surface="true"]')
    return target instanceof HTMLCanvasElement && target.width > 0 && target.height > 0 && Boolean(target.dataset.renderBackend)
  })

  const surface = await canvas.evaluate((target) => ({
    backend: target.dataset.renderBackend,
    fallbackReason: target.dataset.renderBackendReason ?? null,
    width: target.width,
    height: target.height,
    cssWidth: target.getBoundingClientRect().width,
    cssHeight: target.getBoundingClientRect().height,
  }))
  const gpu = await page.evaluate(async () => {
    if (!navigator.gpu) return { available: false, adapter: false }
    const adapter = await navigator.gpu.requestAdapter({ powerPreference: 'low-power' })
    return { available: true, adapter: Boolean(adapter) }
  })
  assert.ok(surface.width >= surface.cssWidth)
  assert.ok(surface.height >= surface.cssHeight)
  assert.ok(['webgpu', 'canvas2d'].includes(surface.backend))

  const screenshot = await canvas.screenshot()
  const pixels = await page.evaluate(async (base64) => {
    const image = new Image()
    image.src = `data:image/png;base64,${base64}`
    await image.decode()
    const reader = document.createElement('canvas')
    reader.width = image.width
    reader.height = image.height
    const context = reader.getContext('2d')
    if (!context) throw new Error('Canvas2D is unavailable for screenshot pixel inspection')
    context.drawImage(image, 0, 0)
    const data = context.getImageData(0, 0, reader.width, reader.height).data
    let nonWhitePixels = 0
    const colors = new Set()
    for (let offset = 0; offset < data.length; offset += 4) {
      const color = `${data[offset]},${data[offset + 1]},${data[offset + 2]}`
      if (color !== '255,255,255') nonWhitePixels += 1
      if (colors.size < 512) colors.add(color)
    }
    return { width: reader.width, height: reader.height, nonWhitePixels, distinctColors: colors.size }
  }, screenshot.toString('base64'))

  assert.ok(pixels.nonWhitePixels > 1_000, `expected visible grid pixels, got ${pixels.nonWhitePixels}`)
  assert.ok(pixels.distinctColors > 5, `expected grid lines and selection colors, got ${pixels.distinctColors}`)

  await grid.focus()
  await page.keyboard.press('9')
  const editor = page.locator('.canvas-cell-input')
  await editor.waitFor({ state: 'visible' })
  assert.equal(await editor.inputValue(), '9')
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => document.querySelector('.formula-input')?.value === '9')
  await page.waitForTimeout(50)
  const editedScreenshot = await canvas.screenshot()
  assert.notDeepEqual(editedScreenshot, screenshot, 'editing a cell should change the Canvas pixels')

  await testInfo.attach('sheets-canvas-render-report.json', {
    body: Buffer.from(JSON.stringify({ backend: surface.backend, fallbackReason: surface.fallbackReason, gpu, surface, pixels, pageErrors }, null, 2)),
    contentType: 'application/json',
  })
  await testInfo.attach('sheets-canvas.png', { body: screenshot, contentType: 'image/png' })
  await testInfo.attach('sheets-canvas-after-edit.png', { body: editedScreenshot, contentType: 'image/png' })
  assert.deepEqual(pageErrors, [])
})
