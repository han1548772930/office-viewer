import assert from 'node:assert/strict'
import { afterEach, test } from 'node:test'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true }, appType: 'custom' })
const { createCanvas2DDriver } = await vite.ssrLoadModule('/packages/office-studio/src/rendering/canvas2d.ts')
const { createRenderSurface } = await vite.ssrLoadModule('/packages/office-studio/src/rendering/createRenderSurface.ts')

const originalDocument = globalThis.document
const originalWindow = globalThis.window

class FakeCanvas {
  width = 300
  height = 150
  style = {}
  dataset = {}
  attributes = new Map()
  parentNode

  setAttribute(name, value) {
    this.attributes.set(name, value)
  }

  getContext() {
    return null
  }

  replaceWith(replacement) {
    const parent = this.parentNode
    if (!parent) return
    const index = parent.children.indexOf(this)
    if (index < 0) return
    parent.children[index] = replacement
    replacement.parentNode = parent
    this.parentNode = undefined
  }

  remove() {
    const parent = this.parentNode
    if (!parent) return
    const index = parent.children.indexOf(this)
    if (index >= 0) parent.children.splice(index, 1)
    this.parentNode = undefined
  }
}

function installFakeDom() {
  const createdCanvases = []
  globalThis.document = {
    createElement(tagName) {
      assert.equal(tagName, 'canvas')
      const canvas = new FakeCanvas()
      createdCanvases.push(canvas)
      return canvas
    },
  }
  globalThis.window = { devicePixelRatio: 1 }
  return createdCanvases
}

afterEach(() => {
  globalThis.document = originalDocument
  globalThis.window = originalWindow
})

test.after(() => vite.close())

test('Canvas2D uses the effective rounded backing-store scale after resize', () => {
  const transforms = []
  const fills = []
  const context = {
    setTransform(...values) { transforms.push(values) },
    fillRect(...values) { fills.push(values) },
    save() {},
    restore() {},
  }
  const canvas = {
    width: 300,
    height: 150,
    getContext(kind, options) {
      assert.equal(kind, '2d')
      assert.deepEqual(options, { alpha: false })
      return context
    },
  }
  const driver = createCanvas2DDriver(canvas)

  driver.resize(100.25, 40.25, 1.5)
  assert.equal(canvas.width, 150)
  assert.equal(canvas.height, 60)
  driver.render({ width: 100.25, height: 40.25, background: '#eee', commands: [] })

  assert.deepEqual(transforms, [[150 / 100.25, 0, 0, 60 / 40.25, 0, 0]])
  assert.deepEqual(fills, [[0, 0, 100.25, 40.25]])
  driver.dispose()
  assert.equal(canvas.width, 1)
  assert.equal(canvas.height, 1)
})

test('surface replaces a lost WebGPU canvas, preserves its CSS size and redraws the last scene', async () => {
  installFakeDom()
  const host = {
    children: [],
    appendChild(canvas) {
      this.children.push(canvas)
      canvas.parentNode = this
    },
  }
  let loseWebGPU
  let gpuDisposed = 0
  let fallbackDisposed = 0
  const gpuResizeCalls = []
  const fallbackResizeCalls = []
  const fallbackScenes = []
  const backendChanges = []
  const scene = { width: 321.5, height: 180.25, commands: [] }
  const surface = await createRenderSurface(host, {
    onBackendChange: (backend, reason) => backendChanges.push([backend, reason]),
  }, {
    async webgpu(_canvas, onLost) {
      loseWebGPU = onLost
      return {
        backend: 'webgpu',
        resize: (...args) => gpuResizeCalls.push(args),
        render() {},
        dispose() { gpuDisposed += 1 },
      }
    },
    canvas2d(canvas) {
      return {
        backend: 'canvas2d',
        resize: (...args) => fallbackResizeCalls.push(args),
        render: value => fallbackScenes.push(value),
        dispose() { fallbackDisposed += 1 },
      }
    },
  })

  surface.resize(scene.width, scene.height, 2.5)
  surface.render(scene)
  const originalCanvas = surface.canvas
  loseWebGPU('device lost')

  assert.equal(surface.backend, 'canvas2d')
  assert.notEqual(surface.canvas, originalCanvas)
  assert.deepEqual([surface.canvas.style.width, surface.canvas.style.height], ['321.5px', '180.25px'])
  assert.deepEqual(gpuResizeCalls.at(-1), [scene.width, scene.height, 2.5])
  assert.deepEqual(fallbackResizeCalls.at(-1), [scene.width, scene.height, 2.5])
  assert.equal(fallbackScenes.length, 1)
  assert.equal(fallbackScenes[0], scene)
  assert.equal(gpuDisposed, 1)
  assert.deepEqual(backendChanges, [['webgpu', undefined], ['canvas2d', 'device lost']])

  surface.dispose()
  assert.equal(fallbackDisposed, 1)
  assert.deepEqual(host.children, [])
})

test('surface reports Canvas2D and paints the first scene when WebGPU initialization fails', async () => {
  installFakeDom()
  const host = {
    children: [],
    appendChild(canvas) {
      this.children.push(canvas)
      canvas.parentNode = this
    },
  }
  const fallbackScenes = []
  const backendChanges = []
  const scene = { width: 120, height: 80, background: '#ffffff', commands: [{ kind: 'rect', x: 8, y: 9, width: 20, height: 18, fill: '#f00' }] }
  const surface = await createRenderSurface(host, {
    onBackendChange: (backend, reason) => backendChanges.push([backend, reason]),
  }, {
    async webgpu() { throw new Error('没有可用的 WebGPU 适配器') },
    canvas2d(canvas) {
      return {
        backend: 'canvas2d',
        resize() {},
        render: value => fallbackScenes.push(value),
        dispose() {},
      }
    },
  })

  assert.equal(surface.backend, 'canvas2d')
  assert.equal(surface.canvas.dataset.renderBackend, 'canvas2d')
  assert.equal(surface.canvas.dataset.renderBackendReason, '没有可用的 WebGPU 适配器')
  assert.deepEqual(backendChanges, [['canvas2d', '没有可用的 WebGPU 适配器']])
  surface.render(scene)
  assert.deepEqual(fallbackScenes, [scene])
  surface.dispose()
})
