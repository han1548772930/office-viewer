import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSparklineSpec } from '../src/features/sparkline/sparkline-model.ts'
import { createSparklineCommands, readSparklineSamples, readSparklineValues, sparklineCommandsForCell } from '../src/features/sparkline/sparkline-render.ts'

// Deferred Canvas draw/data contracts. Neither renderer backend, browser, large data nor Office sample was executed.
const bounds = { x: 10, y: 20, width: 120, height: 40 }
const sheet = cells => ({ id: 'data', name: 'Data', rowCount: 10, columnCount: 10, showGridlines: true, cells })
const spec = (settings = {}, sourceRange = 'B1:E1') => createSparklineSpec({ targetRange: 'G1', sourceRange, type: 'line', color: '#246fce', ...settings })
const numeric = value => ({ value: String(value), rawValue: value, dataType: 'number' })
const lines = commands => commands.filter(command => command.kind === 'line' && command.color === '#246fce')
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < .00001, `${actual} !== ${expected}`)
const center = polygon => ({ x: polygon.points.reduce((sum, point) => sum + point.x, 0) / polygon.points.length,
  y: polygon.points.reduce((sum, point) => sum + point.y, 0) / polygon.points.length })
function finiteGeometry(commands) {
  for (const command of commands) {
    for (const key of ['x', 'y', 'x1', 'x2', 'y1', 'y2', 'width', 'height']) if (typeof command[key] === 'number') assert.ok(Number.isFinite(command[key]))
    for (const point of command.points ?? []) assert.ok(Number.isFinite(point.x) && Number.isFinite(point.y))
    assert.ok(command.clip.width >= 0 && command.clip.height >= 0)
  }
}

for (const [emptyCells, nonNumericCells, expected] of [
  ['gap', 'zero', [2, 'gap', 0, 4]], ['zero', 'connect', [2, 0, 'connect', 4]], ['connect', 'gap', [2, 'connect', 'gap', 4]],
]) test(`empty and numeric-looking text use independent policies: ${emptyCells}/${nonNumericCells}`, () => {
  const source = sheet({ B1: numeric(2), C1: { value: '', rawValue: null, dataType: 'blank' }, D1: { value: '123', rawValue: '123', dataType: 'text' }, E1: numeric(4) })
  assert.deepEqual(readSparklineSamples(spec({ emptyCells, nonNumericCells }), source), expected)
  assert.deepEqual(readSparklineValues(spec({ emptyCells, nonNumericCells }), source), expected.map(value => typeof value === 'number' ? value : null))
})

test('formula caches, booleans and declared numeric text retain scalar semantics without parsing displays', () => {
  const source = sheet({
    B1: { value: 'display is not a number', rawValue: null, formula: '=2+7', cachedValue: 9, cachedDataType: 'number' },
    C1: { value: 'TRUE', rawValue: true, dataType: 'boolean' }, D1: { value: 'not used', rawValue: '0', dataType: 'b' },
    E1: { value: 'rounded display', rawValue: '12.5', dataType: 'n' },
  })
  assert.deepEqual(readSparklineSamples(spec(), source), [9, 1, 0, 12.5])
  source.cells.B1.cachedDataType = 'text'; source.cells.B1.cachedValue = '9'
  assert.deepEqual(readSparklineSamples(spec({ nonNumericCells: 'connect' }), source), ['connect', 1, 0, 12.5])
})

test('missing cells, errors, nonfinite results and untyped text obey their own missing-value category', () => {
  const source = sheet({ C1: { value: '#DIV/0!', rawValue: '#DIV/0!', dataType: 'error' }, D1: numeric(Infinity), E1: { value: '7', rawValue: '7' } })
  assert.deepEqual(readSparklineSamples(spec({ emptyCells: 'zero', nonNumericCells: 'gap' }), source), [0, 'gap', 'gap', 'gap'])
})

for (const hidden of [
  { hiddenColumns: [2] }, { outlineHiddenColumns: [2] }, { columnHiddenSpans: [{ start: 2, end: 2, value: true }] },
  { outlineHiddenColumnSpans: [{ start: 2, end: 2, value: true }] },
]) test(`hidden source columns are excluded until explicitly included: ${JSON.stringify(hidden)}`, () => {
  const source = { ...sheet({ B1: numeric(1), C1: numeric(2), D1: numeric(3), E1: numeric(4) }), ...hidden }
  assert.deepEqual(readSparklineSamples(spec(), source), [1, 3, 4])
  assert.deepEqual(readSparklineSamples(spec({ containHiddenCells: true }), source), [1, 2, 3, 4])
})

for (const hidden of [
  { hiddenRows: [1] }, { outlineHiddenRows: [1] }, { filteredRows: [1] }, { rowHiddenSpans: [{ start: 1, end: 1, value: true }] },
  { outlineHiddenRowSpans: [{ start: 1, end: 1, value: true }] },
]) test(`manual, outline, filter and span row hides retain one include-hidden switch: ${JSON.stringify(hidden)}`, () => {
  const source = { ...sheet({ B1: numeric(1), B2: numeric(2), B3: numeric(3), B4: numeric(4) }), ...hidden }
  assert.deepEqual(readSparklineSamples(spec({}, 'B1:B4'), source), [1, 3, 4])
  assert.deepEqual(readSparklineSamples(spec({ containHiddenCells: true }, 'B1:B4'), source), [1, 2, 3, 4])
})

test('quoted cross-sheet sources resolve by name/id and absent sources do not fall back to local data', () => {
  const local = sheet({ B1: numeric(1), C1: numeric(2) }), other = { ...sheet({ B1: numeric(10), C1: numeric(20) }), id: 'other', name: "Data, O'Neil" }
  assert.deepEqual(readSparklineSamples(spec({}, "'Data, O''Neil'!$B$1:$C$1"), local, { sheets: [local, other] }), [10, 20])
  assert.deepEqual(readSparklineSamples(spec({}, 'other!B1:C1'), local, { sheets: [local, other] }), [10, 20])
  assert.deepEqual(readSparklineSamples(spec({}, 'Missing!B1:C1'), local, { sheets: [local, other] }), [])
})

test('source permission is rechecked before returning a previously cached sample', () => {
  const local = sheet({ B1: numeric(1), C1: numeric(2) }), item = spec({}, 'B1:C1')
  let readable = true, calls = 0
  const context = { sampleCache: new Map(), canReadSource: () => { calls += 1; return readable } }
  assert.deepEqual(readSparklineSamples(item, local, context), [1, 2]); assert.equal(context.sampleCache.size, 1)
  readable = false; assert.deepEqual(readSparklineSamples(item, local, context), []); assert.equal(calls, 2)
  assert.deepEqual(sparklineCommandsForCell(local, { value: '', sparkline: item }, bounds, context), [])
})

test('group extrema include readable cross-sheet members and cannot reuse values after source permission changes', () => {
  const own = spec({ groupId: 'g', extremumMin: { type: 'individual' }, extremumMax: { type: 'group' } }, 'B1:C1')
  const peer = spec({ groupId: 'g', extremumMin: { type: 'individual' }, extremumMax: { type: 'group' } }, 'other!B1:C1')
  const local = sheet({ B1: numeric(0), C1: numeric(10), G1: { value: '', sparkline: own }, G2: { value: '', sparkline: peer } })
  const other = { ...sheet({ B1: numeric(0), C1: numeric(100) }), id: 'other', name: 'Other' }
  let peerAllowed = true
  const context = { sheets: [local, other], sampleCache: new Map(), groupCache: new Map(), canReadSource: source => source.id !== 'other' || peerAllowed }
  const shared = lines(sparklineCommandsForCell(local, local.cells.G1, bounds, context))[0]
  near(shared.y2, 56 - 32 * .1)
  peerAllowed = false
  const restricted = lines(sparklineCommandsForCell(local, local.cells.G1, bounds, context))[0]
  near(restricted.y2, 24); assert.notEqual(shared.y2, restricted.y2)
})

test('gap breaks a line, connect bridges it and zero contributes a real point at the original index', () => {
  assert.equal(lines(createSparklineCommands(spec(), [1, 'gap', 3], bounds)).length, 0)
  const connected = lines(createSparklineCommands(spec(), [1, 'connect', 3], bounds))
  assert.equal(connected.length, 1); near(connected[0].x1, 14); near(connected[0].x2, 126)
  const zeroed = lines(createSparklineCommands(spec(), [1, 0, 3], bounds))
  assert.equal(zeroed.length, 2); near(zeroed[0].x2, 70)
})

test('a connect blank cannot bridge an independent nonnumeric gap', () => {
  assert.equal(lines(createSparklineCommands(spec(), [1, 'connect', 'gap', 4], bounds)).length, 0)
  assert.equal(lines(createSparklineCommands(spec(), [1, 'connect', 'connect', 4], bounds)).length, 1)
})

test('six marker policies draw all ties and keep endpoint identities under right-to-left rendering', () => {
  const colors = { highPoint: '#110001', firstPoint: '#110002', lowPoint: '#110003', lastPoint: '#110004', negativePoint: '#110005', markersPoint: '#110006' }
  const points = Object.fromEntries(Object.entries(colors).map(([name, color]) => [name, { visible: true, color }]))
  const commands = createSparklineCommands(spec({ points }), [-2, 8, 8, -2], bounds)
  const expected = { highPoint: 2, firstPoint: 1, lowPoint: 2, lastPoint: 1, negativePoint: 2, markersPoint: 4 }
  for (const [name, count] of Object.entries(expected)) assert.equal(commands.filter(command => command.kind === 'polygon' && command.fill === colors[name]).length, count)
  const reversed = createSparklineCommands(spec({ points, axisOptions: { reverse: true } }), [-2, 8, 8, -2], bounds)
  const first = center(commands.find(command => command.fill === colors.firstPoint)), mirrored = center(reversed.find(command => command.fill === colors.firstPoint))
  near(first.x + mirrored.x, 140); near(first.y, mirrored.y); finiteGeometry(commands); finiteGeometry(reversed)
})

test('independent custom minimum, group maximum and visible zero axis use one projection', () => {
  const commands = createSparklineCommands(spec({ extremumMin: { type: 'custom', value: -10 }, extremumMax: { type: 'group' },
    axisOptions: { visible: true, color: '#654321' }, lineWidth: 4 }), [-2, 8], bounds, { groupExtrema: { min: -100, max: 100 } })
  const series = lines(commands)[0], axis = commands.find(command => command.kind === 'line' && command.color === '#654321')
  near(series.y1, 56 - 8 / 110 * 32); near(series.y2, 56 - 18 / 110 * 32); near(axis.y1, 56 - 10 / 110 * 32)
  assert.equal(series.width, 4); assert.equal(axis.width, 1)
})

test('positive and negative columns share zero, omit zero/missing bars and honor conditional marker colors', () => {
  const columns = createSparklineCommands(spec({ type: 'column', points: { negativePoint: { visible: true, color: '#ff0000' }, highPoint: { visible: true, color: '#00aa00' } } }), [5, -5, 0, 'gap'], bounds)
  assert.equal(columns.length, 2); assert.deepEqual(columns.map(command => command.fill), ['#00aa00', '#ff0000'])
  near(columns[0].y + columns[0].height, columns[1].y); near(columns[0].height, columns[1].height)
})

test('winloss normalizes magnitudes to equal positive/negative bars and reversal retains source colors', () => {
  const settings = { type: 'winloss', points: { negativePoint: { visible: true, color: '#ff0000' } } }
  const commands = createSparklineCommands(spec(settings), [1000, -1, 0, 2], bounds)
  assert.equal(commands.length, 3); near(commands[0].height, commands[1].height); near(commands[0].height, commands[2].height)
  const reversed = createSparklineCommands(spec({ ...settings, axisOptions: { reverse: true } }), [1000, -1, 0, 2], bounds)
  assert.equal(reversed[1].fill, '#ff0000'); assert.ok(reversed[0].x > reversed[2].x)
})

test('rounded columns emit finite polygons with the same bar extents as square columns', () => {
  const input = [5, -3], square = createSparklineCommands(spec({ type: 'column' }), input, bounds)
  const rounded = createSparklineCommands(spec({ type: 'column', showRadius: true }), input, bounds)
  assert.equal(rounded.length, 2); assert.ok(rounded.every(command => command.kind === 'polygon' && command.points.length > 8))
  rounded.forEach((command, index) => {
    near(Math.min(...command.points.map(point => point.x)), square[index].x)
    near(Math.max(...command.points.map(point => point.y)), square[index].y + square[index].height)
  }); finiteGeometry(rounded)
})

test('gradient fills use theme alpha and obey the same gap/connect segmentation as strokes', () => {
  const settings = { showGradient: true, color: '#246fce80' }
  assert.equal(createSparklineCommands(spec(settings), [1, 'gap', 3], bounds).length, 0)
  const commands = createSparklineCommands(spec(settings), [1, 'connect', 3], bounds), fills = commands.filter(command => command.kind === 'polygon')
  assert.ok(fills.length > 2); assert.equal(commands.filter(command => command.kind === 'line').length, 1)
  assert.ok(fills.every(command => /^rgba\(36,111,206,/.test(command.fill)))
  assert.ok(fills.every(command => Number(command.fill.slice(command.fill.lastIndexOf(',') + 1, -1)) <= .11))
  finiteGeometry(commands)
})

test('pie uses magnitude for negative slices, skips zero/missing values and produces a closed colored circle', () => {
  const commands = createSparklineCommands(spec({ type: 'pie' }), [2, -2, 0, 'gap'], bounds), slices = commands.filter(command => command.kind === 'polygon')
  assert.equal(slices.length, 2); assert.deepEqual(slices.map(command => command.fill), ['#FF6384', '#36A2EB'])
  assert.equal(commands.filter(command => command.kind === 'line').length, 2)
  near(slices[0].points.at(-1).x, slices[1].points[1].x); near(slices[0].points.at(-1).y, slices[1].points[1].y)
  near(slices[1].points.at(-1).x, slices[0].points[1].x); near(slices[1].points.at(-1).y, slices[0].points[1].y)
  finiteGeometry(commands)
})

for (const type of ['line', 'column', 'winloss', 'pie']) test(`finite extreme values never emit nonfinite geometry for ${type}`, () => {
  const commands = createSparklineCommands(spec({ type, points: { highPoint: { visible: true }, lowPoint: { visible: true } }, axisOptions: { visible: true } }),
    [-Number.MAX_VALUE, Number.MAX_VALUE, 0], bounds)
  finiteGeometry(commands)
  if (type !== 'pie') assert.ok(commands.length)
  assert.deepEqual(createSparklineCommands(spec({ type }), [], bounds), [])
  assert.deepEqual(createSparklineCommands(spec({ type }), [1, 2], { x: 0, y: 0, width: 1, height: 1 }), [])
})
