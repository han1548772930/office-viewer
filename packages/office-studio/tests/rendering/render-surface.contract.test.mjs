import assert from 'node:assert/strict'
import { afterAll, afterEach, test } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
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

afterAll(() => vite.close())

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

function createHost() {
  return {
    children: [],
    appendChild(canvas) { this.children.push(canvas); canvas.parentNode = this },
  }
}

test('surface falls back when the first WebGPU resize fails and releases the GPU driver', async () => {
  installFakeDom()
  const host = createHost()
  const backendChanges = [], fallbackSizes = []
  let gpuDisposals = 0
  const surface = await createRenderSurface(host, {
    onBackendChange: (backend, reason) => backendChanges.push([backend, reason]),
  }, {
    async webgpu() {
      return { backend: 'webgpu', resize() { throw new Error('initial configure failed') }, render() {}, dispose() { gpuDisposals += 1 } }
    },
    canvas2d() {
      return { backend: 'canvas2d', resize: (...args) => fallbackSizes.push(args), render() {}, dispose() {} }
    },
  })

  assert.equal(surface.backend, 'canvas2d')
  assert.equal(gpuDisposals, 1)
  assert.equal(host.children.length, 1)
  assert.deepEqual(fallbackSizes, [[1, 1, 1]])
  assert.equal(backendChanges.length, 1)
  assert.equal(backendChanges[0][0], 'canvas2d')
  assert.match(backendChanges[0][1], /initial configure failed/)
  surface.dispose()
})

test('hidden surfaces defer fallback drawing until they have a nonzero viewport', async () => {
  installFakeDom()
  let loseWebGPU
  const fallbackSizes = [], fallbackScenes = []
  const surface = await createRenderSurface(createHost(), {}, {
    async webgpu(_canvas, onLost) {
      loseWebGPU = onLost
      return { backend: 'webgpu', resize() {}, render() {}, dispose() {} }
    },
    canvas2d() {
      return { backend: 'canvas2d', resize: (...args) => fallbackSizes.push(args), render: scene => fallbackScenes.push(scene), dispose() {} }
    },
  })
  const scene = { width: 500, height: 300, commands: [] }
  surface.resize(500, 300, 2)
  surface.render(scene)
  surface.resize(0, 0, 2)
  loseWebGPU('device lost while tab hidden')

  assert.equal(surface.backend, 'canvas2d')
  assert.equal(surface.canvas.style.width, '0px')
  assert.deepEqual(fallbackSizes, [])
  assert.deepEqual(fallbackScenes, [])
  surface.resize(500, 300, 2)
  surface.render(scene)
  assert.deepEqual(fallbackSizes, [[500, 300, 2]])
  assert.deepEqual(fallbackScenes, [scene])
  surface.dispose()
})

test('surface removes its canvas and driver when even Canvas2D initialization fails', async () => {
  installFakeDom()
  const host = createHost()
  let disposals = 0
  await assert.rejects(createRenderSurface(host, { preferredBackend: 'canvas2d' }, {
    async webgpu() { throw new Error('WebGPU should not be tried') },
    canvas2d() {
      return { backend: 'canvas2d', resize() { throw new Error('canvas sizing failed') }, render() {}, dispose() { disposals += 1 } }
    },
  }), /canvas sizing failed/)
  assert.equal(disposals, 1)
  assert.deepEqual(host.children, [])
})

test('surface never passes nonfinite viewport sizes or DPR to its driver', async () => {
  installFakeDom()
  const sizes = [], scenes = []
  const surface = await createRenderSurface(createHost(), { preferredBackend: 'canvas2d' }, {
    async webgpu() { throw new Error('WebGPU should not be tried') },
    canvas2d() { return { backend: 'canvas2d', resize: (...args) => sizes.push(args), render: value => scenes.push(value), dispose() {} } },
  })
  surface.resize(Number.NaN, Number.POSITIVE_INFINITY, Number.NaN)
  surface.render({ width: 100, height: 100, commands: [] })
  assert.equal(surface.canvas.style.width, '0px')
  assert.deepEqual(sizes, [[1, 1, 1]])
  assert.deepEqual(scenes, [])
  surface.resize(200, 100, Number.NaN)
  assert.deepEqual(sizes.at(-1), [200, 100, 1])
  surface.dispose()
})

test('failed fallback disposes the old GPU driver exactly once and removes the replacement canvas', async () => {
  installFakeDom()
  const host = createHost()
  let loseWebGPU, gpuDisposals = 0
  const surface = await createRenderSurface(host, {}, {
    async webgpu(_canvas, onLost) {
      loseWebGPU = onLost
      return { backend: 'webgpu', resize() {}, render() {}, dispose() { gpuDisposals += 1 } }
    },
    canvas2d() { throw new Error('Canvas2D unavailable') },
  })
  assert.throws(() => loseWebGPU('lost'), /Canvas2D unavailable/)
  assert.equal(gpuDisposals, 1)
  assert.deepEqual(host.children, [])
  surface.dispose()
  loseWebGPU('late device lost notification')
  assert.equal(gpuDisposals, 1)
})
