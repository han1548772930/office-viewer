import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartCandlestickMode, chartCandlestickOverlayOptions, chartCandlestickState, chartCandlestickStyle,
  chartCandlestickWick, chartCandlestickWidth } from '../src/features/charts/chart-candlestick.ts'
import { drawChartCandlesticks } from '../src/features/charts/chart-candlestick-draw.ts'
import { chartCandlestickEChartsStyle } from '../src/features/charts/echarts-candlestick-option.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { chartCandlestickBodyControls, chartCandlestickBorderControls, chartCandlestickCloseMarkerControls,
  chartCandlestickHighLowControls, chartCandlestickWidthControls, clearChartCandlestickSettings } from '../src/features/charts/chart-candlestick-settings-controls.ts'
import { chartDataMappingControls } from '../src/features/charts/chart-type-settings-controls.ts'
import { createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { updateChart } from '../src/features/charts/chart-engine.ts'
import { validateChartOptions } from '../src/features/charts/chart-options.ts'
import { CANDLE_HLC_MAPPING, candlestickFixture } from './fixtures/chart-candlestick-fixtures.mjs'

// P25 drafts. No runner, type check, Canvas/browser, Office file or benchmark ran.
const json = value => JSON.parse(JSON.stringify(value))
const control = (controls, key) => { const value = controls.find(item => item.key === key); assert.ok(value, key); return value }
const layout = { plot: { x: 0, y: 0, width: 180, height: 100 }, min: 0, max: 20, categoryBandwidth: 45 }
function draw(data, chart) { const commands = []; drawChartCandlesticks(commands, data, chart, layout, [], '#111111'); return commands }

test('source roles retain three strict states, the Close identity and missing sample gaps', () => {
  const { chart, data } = candlestickFixture()
  assert.equal(chart.legend.visible, false)
  assert.equal(data.series[0].field, 4); assert.equal(data.series[0].sourceName, 'Close'); assert.equal(data.series[0].name, 'Close')
  assert.deepEqual(data.series[0].values, [12, 9, 11, null]); assert.equal(data.candles[3], null)
  assert.deepEqual(data.candles.slice(0, 3).map(chartCandlestickState), ['rising', 'falling', 'doji'])
  assert.deepEqual(data.candleFields, { openIndex: 1, highIndex: 2, lowIndex: 3, closeIndex: 4 })
  assert.equal(chartCandlestickMode(chart, data), 'ohlc')
})

test('explicit HLC and empty sources keep their roles without inferring Open from old valueFields or mutating the draft', () => {
  const options = { mapping: { candlestick: CANDLE_HLC_MAPPING } }
  for (const rows of [undefined, [['Blank', null, null, null, null]]]) {
    const { chart, data } = candlestickFixture(options, rows), before = structuredClone(chart)
    assert.equal(chartCandlestickMode(chart, data), 'hlc'); assert.deepEqual(data.candleFields, CANDLE_HLC_MAPPING)
    const mapping = chartDataMappingControls(chart, 5, data)
    assert.equal(control(mapping, 'mapping.candlestick.openIndex').read(chart), undefined)
    assert.equal(control(mapping, 'mapping.candlestick.highIndex').read(chart), 2)
    assert.deepEqual(chart, before)
    if (rows) assert.deepEqual(data.candles, [null])
    else assert.ok(data.candles.every(candle => candle.open === undefined))
  }
})

test('three state defaults and authored layers retain separate body and border alpha with sample overrides last', () => {
  const { chart, data } = candlestickFixture({ candlestick: {
    rising: { color: 'red', fillOpacity: 0.25, border: { color: 'navy', opacity: 0.8, width: 3 } },
    falling: { color: 'blue', hollow: true }, doji: { border: { color: 'black', width: 2, opacity: 0.4 } },
  }, series: { 4: { name: 'Prices', dataPoints: { 1: { color: 'orange', fillOpacity: 0.5, border: { color: 'gold', width: 4 } } } } } })
  const before = structuredClone(chart), source = data.series[0]
  const rising = chartCandlestickStyle(chart, 4, 0, 'rising', [], source.sourceName)
  assert.equal(rising.name, 'Prices'); assert.equal(rising.color, 'red'); assert.equal(rising.fillOpacity, 0.25)
  const paint = chartCandlestickEChartsStyle(rising)
  assert.equal(paint.color, 'rgba(255, 0, 0, 0.25)'); assert.equal(paint.color0, paint.color)
  assert.equal(paint.borderColor, 'rgba(0, 0, 128, 0.8)'); assert.equal(paint.borderColorDoji, paint.borderColor); assert.equal(paint.borderWidth, 3)
  const falling = chartCandlestickStyle(chart, 4, 1, 'falling', [], source.sourceName)
  assert.equal(falling.color, 'orange'); assert.equal(falling.fillOpacity, 0.5); assert.equal(falling.hollow, true)
  assert.equal(falling.border.color, 'gold'); assert.equal(falling.border.width, 4)
  assert.equal(chartCandlestickEChartsStyle(falling).color, 'rgba(255, 165, 0, 0)')
  assert.equal(chartCandlestickEChartsStyle(chartCandlestickStyle(chart, 4, 2, 'doji')).borderColorDoji, 'rgba(0, 0, 0, 0.4)')
  assert.deepEqual(chart, before)
  const defaults = candlestickFixture().chart
  assert.deepEqual(['rising', 'falling', 'doji'].map(state => chartCandlestickStyle(defaults, 4, undefined, state).color),
    ['#0DA471', '#F05252', '#4E5565'])
})

test('automatic width follows 8..32px while author widths and outlines accept their finite positive/nonnegative contracts', () => {
  assert.deepEqual([4, 40, 200].map(band => chartCandlestickWidth(band)), [8, 20, 32])
  assert.equal(chartCandlestickWidth(40, { candleWidth: 0.125 }), 0.125)
  assert.equal(chartCandlestickWidth(40, { candleWidth: 512 }), 512)
  assert.doesNotThrow(() => validateChartOptions({ candlestick: { candleWidth: 0.125, rising: { border: { width: 512 } }, highLowLine: { border: { width: 512 } } } }))
  for (const candleWidth of [0, -1, NaN, Infinity]) assert.throws(() => validateChartOptions({ candlestick: { candleWidth } }))
  for (const candlestick of [{ rising: { fillOpacity: 2 } }, { doji: { border: { width: -1 } } },
    { falling: { hollow: 'yes' } }, { highLowLine: { visible: 1 } }, { closeMarker: { point: { color: {} } } },
    { closeMarker: { point: { size: -1 } } }, { closeMarker: { point: { shape: 'invalid' } } }]) assert.throws(() => validateChartOptions({ candlestick }))
})

test('HLC defaults and explicit OHLC overlay objects preserve visibility, wick shortening and a real clear transition', () => {
  assert.deepEqual(chartCandlestickOverlayOptions(undefined, 'hlc'), {
    highLowLine: { visible: true, border: { color: '#4E5565', width: 1, opacity: 1, dashType: 'solid' } },
    closeMarker: { visible: true, point: { color: '#4E5565', size: 8, shape: 'dot' }, lineWidth: 1 },
  })
  assert.equal(chartCandlestickOverlayOptions(undefined, 'ohlc').highLowLine.visible, false)
  assert.equal(chartCandlestickOverlayOptions({ closeMarker: {} }, 'ohlc').closeMarker.visible, true)
  const candle = { open: 10, close: 12, high: 15, low: 8 }, chart = { options: { candlestick: { highLowLine: { visible: false } } } }
  assert.deepEqual(chartCandlestickWick(candle, undefined), { low: 8, high: 15 })
  assert.deepEqual(chartCandlestickWick(candle, chart.options.candlestick), { low: 10, high: 12 })
  clearChartCandlestickSettings(chart, 'highLowLine')
  assert.deepEqual(chartCandlestickWick(candle, chart.options.candlestick), { low: 8, high: 15 })
  assert.equal(chartCandlestickOverlayOptions(chart.options.candlestick, 'ohlc').highLowLine.visible, false)
})

test('ordinary and reversed Y wicks join the higher/lower value ends of the body and Doji has a horizontal outline', () => {
  const { chart, data } = candlestickFixture(), normal = draw(data, chart).filter(command => command.kind === 'line')
  assert.deepEqual(normal.slice(0, 2).map(command => [command.y1, command.y2]), [[25, 40], [50, 60]])
  chart.yAxis.reverse = true
  const reverse = draw(data, chart).filter(command => command.kind === 'line')
  assert.deepEqual(reverse.slice(0, 2).map(command => [command.y1, command.y2]), [[75, 60], [50, 40]])
  const doji = reverse.filter(command => command.color === 'rgba(78, 85, 101, 1)' && command.y1 === command.y2)
  assert.equal(doji.length, 1); assert.equal(doji[0].x2 - doji[0].x1, 22.5)
  assert.equal(draw(data, chart).filter(command => command.kind === 'rect').length, 2)
})

test('HLC draws high/low and Dot close lines without a synthetic body, while None and zero-size markers really disappear', () => {
  const { chart, data } = candlestickFixture({ mapping: { candlestick: CANDLE_HLC_MAPPING } })
  const commands = draw(data, chart)
  assert.equal(commands.some(command => command.kind === 'rect'), false)
  assert.equal(commands.length, 8)
  assert.deepEqual(commands.filter(command => command.x1 !== command.x2).map(command => [command.x2 - command.x1, command.width]),
    [[4, 1], [4, 1], [4, 1], [4, 1]])
  for (const point of [{ shape: 'none' }, { size: 0 }]) {
    chart.options.candlestick = { closeMarker: { point } }
    const hidden = draw(data, chart)
    assert.equal(hidden.length, 4); assert.ok(hidden.every(command => command.x1 === command.x2))
  }
  const invalid = { ...data, candles: [{ open: 10, close: 12, low: 0, high: 15, categoryIndex: 0 }] }
  chart.yAxis = { scaleType: 'log', min: 1, max: 20 }
  assert.deepEqual(draw(invalid, chart), [])
})

test('native OHLC supplies explicit Doji paint and disables large aggregation; HLC and overlays share the graphic scene', () => {
  const { workbook, chart } = candlestickFixture({ candlestick: { candleWidth: 48, doji: { border: { color: 'black', width: 2 } } } })
  const native = chartEChartsOption(workbook, chart)
  assert.equal(native.series[0].type, 'candlestick'); assert.equal(native.series[0].large, false); assert.equal(native.series[0].barWidth, 48)
  assert.deepEqual(native.series[0].data[3], []); assert.equal(native.series[0].data[2].itemStyle.borderColorDoji, 'rgba(0, 0, 0, 1)')
  chart.options.candlestick.highLowLine = { visible: true }
  assert.equal(chartEChartsOption(workbook, chart).series, undefined)
  const hlc = candlestickFixture({ mapping: { candlestick: CANDLE_HLC_MAPPING } })
  assert.equal(chartEChartsOption(hlc.workbook, hlc.chart).series, undefined)
})

test('state, width and overlay resets survive the actual settings patch while legacy inheritance clears only with the full reset', () => {
  const { workbook, chart, data } = candlestickFixture({ candlestick: { candleWidth: 48, risingColor: 'purple', hollow: true,
    rising: { color: 'red', fillOpacity: 0.4, border: { color: 'navy', width: 3 } },
    falling: { color: 'blue' }, doji: { border: { width: 2 } }, highLowLine: { visible: true, border: { width: 2 } }, closeMarker: { visible: true, point: { size: 12 } } } })
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart)
  control(chartCandlestickWidthControls(), 'candlestick.candleWidth').reset(draft.config)
  control(chartCandlestickBodyControls('rising', data.series[0]), 'candlestick.rising.fillOpacity').reset(draft.config)
  control(chartCandlestickBorderControls('rising', data.series[0]), 'candlestick.rising.border.color').reset(draft.config)
  control(chartCandlestickHighLowControls('ohlc'), 'candlestick.highLowLine.border.width').reset(draft.config)
  control(chartCandlestickCloseMarkerControls('ohlc'), 'candlestick.closeMarker.point.size').reset(draft.config)
  assert.deepEqual(chart, before)
  updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.deepEqual(json(chart.options.candlestick.rising), { color: 'red', border: { width: 3 } })
  assert.equal(chart.options.candlestick.candleWidth, undefined)
  assert.deepEqual(json(chart.options.candlestick.highLowLine), { visible: true, border: {} })
  assert.deepEqual(json(chart.options.candlestick.closeMarker), { visible: true, point: {} })
  const groups = createChartSettingsDraft(chart)
  for (const target of ['rising', 'falling', 'doji', 'highLowLine', 'closeMarker']) clearChartCandlestickSettings(groups.config, target)
  updateChart(workbook, chart, chartSettingsPatch(groups))
  assert.deepEqual(json(chart.options.candlestick), { risingColor: 'purple', hollow: true })
  assert.equal(chartCandlestickStyle(chart, 4, undefined, 'rising').color, 'purple')
  const all = createChartSettingsDraft(chart)
  clearChartCandlestickSettings(all.config, 'all'); updateChart(workbook, chart, chartSettingsPatch(all))
  assert.equal(chart.options.candlestick, undefined); assert.equal(chartCandlestickStyle(chart, 4, undefined, 'rising').color, '#0DA471')
  assert.equal(chart.revision, before.revision + 3)
})
