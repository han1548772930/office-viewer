import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffinePoint } from '../src/domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT as H, GRID_HEADER_WIDTH as W } from '../src/rendering/spreadsheet/layout.ts'
import { chartDrawingKeyboardTranslation, chartDrawingLocalDelta, chartDrawingLocalPoint, chartDrawingOutlineCommands,
  chartDrawingViewportBounds, chartDrawingViewportPoint, chartDrawingWorldPoint } from '../src/features/charts/chart-drawing-affine.ts'

// Deferred contracts. This file has not been run; no server, browser, type check, build or benchmark was started.
const close = (actual, expected) => {
  for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < 1e-8, `${key}: ${actual[key]} !== ${expected[key]}`)
}
const legacy = () => ({ position: { x: W + 80, y: H + 90 }, size: { width: 480, height: 300 } })
const affine = () => ({ ...legacy(), drawingGeometry: { width: 240, height: 120,
  matrix: { a: 0.8, b: 0.6, c: -0.3, d: 1.2, e: 80, f: 90 } } })
const view = { zoom: 125, scrollLeft: 70, scrollTop: 50 }

test('legacy renderer coordinates round-trip through the worksheet and viewport at every supported zoom', () => {
  const chart = legacy(), local = { x: 123, y: 87 }
  assert.deepEqual(chartDrawingWorldPoint(chart, local), { x: W + 203, y: H + 177 })
  for (const zoom of [10, 50, 100, 175, 400]) {
    const viewport = { ...view, zoom }, point = chartDrawingViewportPoint(chart, local, viewport)
    close(chartDrawingLocalPoint(chart, point, viewport), local)
  }
})

for (const [name, matrix] of [
  ['shear', { a: 1.5, b: 0.25, c: 0.7, d: 1.2, e: 100, f: 120 }],
  ['reflection', { a: -1.5, b: 0.25, c: 0.4, d: 1.2, e: 500, f: 120 }],
  ['quarter rotation', { a: 0, b: 2, c: -1, d: 0, e: 500, f: 120 }],
]) test(`${name} retains author local extent while all four frozen-pane coordinate pairs round-trip`, () => {
  const chart = affine(); chart.drawingGeometry.matrix = matrix
  const original = structuredClone(chart), local = { x: 216, y: 225 }
  const expected = drawingAffinePoint(matrix, 108, 90)
  close(chartDrawingWorldPoint(chart, local), { x: W + expected.x, y: H + expected.y })
  for (const scrollLeft of [0, 70]) for (const scrollTop of [0, 50]) {
    const viewport = { ...view, scrollLeft, scrollTop }
    close(chartDrawingLocalPoint(chart, chartDrawingViewportPoint(chart, local, viewport), viewport), local)
  }
  assert.deepEqual(chart, original)
})

test('title, plot and content drag deltas use renderer pixels rather than the transformed display size', () => {
  const chart = affine(), first = { x: 40, y: 30 }, last = { x: 85, y: 75 }
  for (const viewport of [view, { zoom: 200, scrollLeft: 0, scrollTop: 50 }]) {
    close(chartDrawingLocalDelta(chart, chartDrawingViewportPoint(chart, first, viewport),
      chartDrawingViewportPoint(chart, last, viewport), viewport), { x: 45, y: 45 })
  }
})

test('author floating controls use the four projected corners of a local rectangle', () => {
  const chart = affine(), rect = { x: 20, y: 30, width: 100, height: 40 }
  const points = [[20, 30], [120, 30], [120, 70], [20, 70]].map(([x, y]) => chartDrawingViewportPoint(chart, { x, y }, view))
  const xs = points.map(point => point.x), ys = points.map(point => point.y)
  close(chartDrawingViewportBounds(chart, rect, view), { x: Math.min(...xs), y: Math.min(...ys),
    width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) })
})

test('rotated author outlines clip their line in local space before projection and retain common drawing ownership', () => {
  const chart = affine(), source = { kind: 'line', x1: -50, y1: 75, x2: 250, y2: 75,
    color: '#123456', width: 2, clip: { x: 50, y: 40, width: 100, height: 70 } }
  const commands = chartDrawingOutlineCommands(chart, 'owned-chart', [source])
  assert.equal(commands.length, 1)
  const command = commands[0], first = chartDrawingWorldPoint(chart, { x: 50, y: 75 }), last = chartDrawingWorldPoint(chart, { x: 150, y: 75 })
  close(command, { x1: first.x, y1: first.y, x2: last.x, y2: last.y })
  assert.deepEqual(command.drawingRef, { kind: 'chart', id: 'owned-chart' }); assert.equal(command.drawingOverlay, true)
  assert.equal(command.clip, undefined); assert.equal(command.color, source.color); assert.equal(command.width, source.width)
  assert.deepEqual(chartDrawingOutlineCommands(chart, 'owned-chart', [{ ...source, y1: 10, y2: 10 }]), [])
})

test('parallel and point outlines honor local clip boundaries without leaking or dividing by zero', () => {
  const chart = legacy(), clip = { x: 10, y: 20, width: 30, height: 40 }, line = { kind: 'line', color: '#000', width: 1, clip }
  const commands = chartDrawingOutlineCommands(chart, 'c', [
    { ...line, x1: 10, y1: 0, x2: 10, y2: 100 },
    { ...line, x1: 11, y1: 21, x2: 11, y2: 21 },
    { ...line, x1: 9, y1: 21, x2: 9, y2: 21 },
  ])
  assert.equal(commands.length, 2)
  close(commands[0], { x1: W + 90, y1: H + 110, x2: W + 90, y2: H + 150 })
  assert.throws(() => chartDrawingOutlineCommands(chart, 'c', [{ kind: 'rect', x: 0, y: 0, width: 1, height: 1 }]), TypeError)
})

test('object nudges use worksheet pixels and preserve reflection, shear and renderer dimensions', () => {
  const chart = affine(), original = structuredClone(chart), extent = { width: 2000, height: 1000 }
  for (const [key, e, f] of [['ArrowLeft', -1, 0], ['ArrowRight', 1, 0], ['ArrowUp', 0, -1], ['ArrowDown', 0, 1]]) {
    assert.deepEqual(chartDrawingKeyboardTranslation(chart, key, extent), { a: 1, b: 0, c: 0, d: 1, e, f })
    assert.deepEqual(chartDrawingKeyboardTranslation(chart, key, extent, true), { a: 1, b: 0, c: 0, d: 1, e: e * 10, f: f * 10 })
  }
  assert.equal(chartDrawingKeyboardTranslation(chart, 'Enter', extent), undefined); assert.deepEqual(chart, original)
})

test('object nudges clamp complete bounds at the worksheet edge and round partial deltas to six decimal places', () => {
  const chart = { position: { x: W, y: H }, size: { width: 100, height: 80 } }, extent = { width: 100, height: 80 }
  for (const key of ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown']) assert.equal(chartDrawingKeyboardTranslation(chart, key, extent), undefined)
  chart.position.x += 0.12345678; chart.position.y += 10
  close(chartDrawingKeyboardTranslation(chart, 'ArrowLeft', { width: 200, height: 200 }), { e: -0.123457, f: 0 })
})
