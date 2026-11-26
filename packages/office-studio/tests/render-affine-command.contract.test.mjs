import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffinePoint } from '../src/domain/drawing-affine.ts'
import { clipRenderTexturePolygon, renderLocalClipPolygon, renderLocalViewportBounds, renderProjectedVertices,
  renderRectPoints } from '../src/rendering/affine-command-geometry.ts'
import { imageCommandGeometry } from '../src/rendering/image-geometry.ts'
import { createCanvas2DDriver } from '../src/rendering/canvas2d.ts'

// Deferred math/driver contracts. The recording context is not a browser backend or performance measurement.
const close = (actual, expected, epsilon = 1e-8) => { for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < epsilon, `${key}: ${actual[key]} / ${expected[key]}`) }
const uv = [[0, 0], [1, 0], [1, 1], [0, 1]]
const points = [0, 0, 100, 0, 100, 50, 0, 50]

test('local rectangular clipping interpolates UV instead of stretching the retained image area', () => {
  const clipped = clipRenderTexturePolygon(renderProjectedVertices(points, uv), renderRectPoints({ x: 25, y: 10, width: 50, height: 30 }))
  assert.equal(clipped.length, 4)
  for (const vertex of clipped) close(vertex, { u: vertex.x / 100, v: vertex.y / 50 })
  assert.deepEqual(new Set(clipped.map(vertex => vertex.x)), new Set([25, 75]))
  assert.deepEqual(new Set(clipped.map(vertex => vertex.y)), new Set([10, 40]))
})

for (const matrix of [
  { a: 0, b: 1, c: -1, d: 0, e: 150, f: 30 },
  { a: -1, b: 0.25, c: 0.2, d: 1, e: 200, f: 10 },
  { a: 1, b: 0.5, c: 0.25, d: 1, e: 15, f: 10 },
]) test(`affine local clip retains winding and UV under matrix ${JSON.stringify(matrix)}`, () => {
  const command = { matrix, localClip: { x: 25, y: 10, width: 50, height: 30 } }
  const clipped = clipRenderTexturePolygon(renderProjectedVertices(points, uv, matrix), renderLocalClipPolygon(command))
  assert.equal(clipped.length, 4)
  for (const vertex of clipped) close(vertex, drawingAffinePoint(matrix, vertex.u * 100, vertex.v * 50))
  for (const vertex of clipped) { assert.ok(vertex.u >= 0.25 - 1e-8 && vertex.u <= 0.75 + 1e-8); assert.ok(vertex.v >= 0.2 - 1e-8 && vertex.v <= 0.8 + 1e-8) }
})

test('an empty or degenerate local clip emits no polygon while an absent local clip has no restriction', () => {
  assert.equal(renderLocalClipPolygon({}), undefined)
  for (const width of [0, -1]) assert.deepEqual(renderLocalClipPolygon({ localClip: { x: 0, y: 0, width, height: 10 } }), [])
  assert.deepEqual(clipRenderTexturePolygon(renderProjectedVertices(points, uv), [{ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 2, y: 2 }]), [])
})

test('relative polygon area survives large world translations and retains normalized texture coordinates', () => {
  const matrix = { a: -1, b: 0, c: 0, d: 1, e: 1e12, f: 1e12 }
  const clipped = clipRenderTexturePolygon(renderProjectedVertices(points, uv, matrix), renderLocalClipPolygon({ matrix, localClip: { x: 25, y: 10, width: 50, height: 30 } }))
  assert.equal(clipped.length, 4)
  for (const vertex of clipped) close(vertex, drawingAffinePoint(matrix, vertex.u * 100, vertex.v * 50), 1e-3)
})

test('inverse viewport bounds conservatively include all visible local text corners under rotation and shear', () => {
  const viewport = { x: 40, y: 20, width: 300, height: 100 }, matrix = { a: 1, b: 0.5, c: -0.3, d: 1, e: 150, f: 30 }
  const local = renderLocalViewportBounds(viewport, matrix)
  assert.ok(local.width > 0 && local.height > 0)
  const projected = renderRectPoints(local, matrix)
  assert.ok(Math.min(...projected.map(point => point.x)) <= viewport.x); assert.ok(Math.max(...projected.map(point => point.x)) >= viewport.x + viewport.width)
  assert.ok(Math.min(...projected.map(point => point.y)) <= viewport.y); assert.ok(Math.max(...projected.map(point => point.y)) >= viewport.y + viewport.height)
})

test('image rotation/flip precede one matrix application and its transformed clip shares the same scene coordinates', () => {
  const matrix = { a: 1, b: 0.25, c: 0.3, d: 1, e: 100, f: 30 }, command = { kind: 'image', x: 0, y: 0, width: 100, height: 50,
    rotation: 90, flipX: true, source: { width: 100, height: 50 }, matrix, localClip: { x: 0, y: 0, width: 100, height: 50 } }
  const geometry = imageCommandGeometry(command), plain = imageCommandGeometry({ ...command, matrix: undefined })
  for (let index = 0; index < 4; index++) close({ x: geometry.points[index * 2], y: geometry.points[index * 2 + 1] },
    drawingAffinePoint(matrix, plain.points[index * 2], plain.points[index * 2 + 1]))
  assert.deepEqual(renderProjectedVertices(geometry.points, geometry.uv).map(({ x, y }) => [x, y]).flat(), geometry.points)
})

test('Canvas2D establishes scene clip, affine matrix and local clip in that order and does not repeat the image matrix', () => {
  const calls = [], methods = ['setTransform', 'fillRect', 'save', 'restore', 'beginPath', 'rect', 'clip', 'transform', 'translate', 'rotate', 'scale', 'drawImage']
  const context = Object.fromEntries(methods.map(name => [name, (...args) => calls.push([name, ...args])]))
  const canvas = { width: 100, height: 50, getContext: kind => kind === '2d' ? context : null }, driver = createCanvas2DDriver(canvas)
  const command = { kind: 'image', x: 0, y: 0, width: 100, height: 50, source: { width: 100, height: 50 },
    matrix: { a: -1, b: 0, c: 0.2, d: 1, e: 200, f: 30 }, clip: { x: 20, y: 10, width: 300, height: 150 }, localClip: { x: 10, y: 5, width: 80, height: 40 } }
  driver.render({ width: 400, height: 200, commands: [command] })
  const relevant = calls.filter(([method]) => ['rect', 'transform', 'drawImage'].includes(method))
  assert.deepEqual(relevant.slice(0, 3), [['rect', 20, 10, 300, 150], ['transform', -1, 0, 0.2, 1, 200, 30], ['rect', 10, 5, 80, 40]])
  assert.equal(calls.filter(([method]) => method === 'transform').length, 1); assert.equal(calls.filter(([method]) => method === 'drawImage').length, 1)
  driver.dispose(); assert.equal(canvas.width, 1); assert.equal(canvas.height, 1)
})
