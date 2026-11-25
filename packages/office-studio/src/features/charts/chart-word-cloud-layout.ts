import type { ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartStatisticsParams, ChartStatisticsTask, ChartWordCloudGlyph } from '../../compute/chart-protocol'
import { COMPUTE_LIMITS } from '../../compute/protocol'
import { readChartStatisticsOutput } from '../../compute/chart-results'
import { estimateChartTaskMemory, validateChartTask } from '../../compute/chart-tasks'
import { ComputeError } from '../../compute/errors'
import type { ChartData } from './chart-engine'
import { chartPlotFrame } from './chart-layout'
import { chartPointStyle } from './chart-derived'
import { chartSeriesStyle } from './chart-options'
import { chartFont } from './chart-geometry'
import { CHART_WORD_CLOUD_DEFAULTS, chartWordCloudData, chartWordCloudEffectiveMask, chartWordCloudFontSize, type ChartWordCloudDatum } from './chart-word-cloud'
import { createChartWordCloudCanvas, drawChartWordCloudMaskImage, drawChartWordCloudShape, waitChartWordCloudFont } from './chart-word-cloud-mask'
import { assertChartPreparationCurrent, assertChartPreparationMemory, yieldChartPreparation, CHART_PREPARATION_CHUNK, type ChartComputePreparationContext } from './chart-compute-lifecycle'

export interface ChartCloudWord extends RenderRect, ChartWordCloudDatum { fontSize: number; rotation: number; font: string }
export interface ChartWordCloudLayoutResult { width: number; height: number; words: ChartCloudWord[]; omittedRepeats: number }
export interface ChartWordCloudPlan {
  params: ChartStatisticsParams
  words: Array<ChartWordCloudDatum & { fontSize: number; rotation: number; font: string }>
  width: number; height: number; gridSize: number
}
interface GlyphRaster { cells: Uint32Array; width: number; height: number }

function nextRandom(state: { seed: number }): number {
  state.seed ^= state.seed << 13; state.seed ^= state.seed >>> 17; state.seed ^= state.seed << 5
  return (state.seed >>> 0) / 0x1_0000_0000
}

async function rasterGlyph(word: ChartWordCloudPlan['words'][number], grid: number, retainedBytes: number,
  context: ChartComputePreparationContext): Promise<GlyphRaster> {
  await waitChartWordCloudFont(word.font, word.text, context)
  assertChartPreparationCurrent(context)
  const owner = createChartWordCloudCanvas(1, 1)
  try {
    const drawing = owner.context
    drawing.font = word.font; drawing.textAlign = 'center'; drawing.textBaseline = 'middle'
    const metrics = drawing.measureText(word.text)
    const halfWidth = Math.max(Math.abs(metrics.actualBoundingBoxLeft), Math.abs(metrics.actualBoundingBoxRight))
    const halfHeight = Math.max(Math.abs(metrics.actualBoundingBoxAscent), Math.abs(metrics.actualBoundingBoxDescent))
    if (!Number.isFinite(halfWidth) || !Number.isFinite(halfHeight)) throw new ComputeError('COMPUTE_UNSUPPORTED_TYPE', '浏览器未提供词云实际字形度量')
    const radians = word.rotation * Math.PI / 180, padding = CHART_WORD_CLOUD_DEFAULTS.padding + 1
    const width = Math.max(1, Math.ceil((Math.abs(Math.cos(radians)) * halfWidth * 2 + Math.abs(Math.sin(radians)) * halfHeight * 2 + padding * 2) / grid))
    const height = Math.max(1, Math.ceil((Math.abs(Math.sin(radians)) * halfWidth * 2 + Math.abs(Math.cos(radians)) * halfHeight * 2 + padding * 2) / grid))
    if (width > 4096 || height > 4096) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '词云字形尺寸超出栅格预算')
    assertChartPreparationMemory(retainedBytes + width * height * (grid * grid * 8 + 4), context)
    owner.canvas.width = width * grid; owner.canvas.height = height * grid
    drawing.font = word.font; drawing.textAlign = 'center'; drawing.textBaseline = 'middle'
    drawing.translate(owner.canvas.width / 2, owner.canvas.height / 2); drawing.rotate(radians)
    drawing.fillStyle = '#000'; drawing.strokeStyle = '#000'; drawing.lineWidth = CHART_WORD_CLOUD_DEFAULTS.padding * 2
    drawing.strokeText(word.text, 0, 0); drawing.fillText(word.text, 0, 0)
    const pixels = drawing.getImageData(0, 0, owner.canvas.width, owner.canvas.height).data, cells = new Uint32Array(width * height)
    let count = 0, work = 0
    for (let cell = 0; cell < cells.length; cell++) {
      const left = cell % width * grid, top = Math.floor(cell / width) * grid
      let occupied = false
      for (let y = 0; y < grid; y++) for (let x = 0; x < grid; x++) {
        occupied ||= pixels[((top + y) * owner.canvas.width + left + x) * 4 + 3]! > 8
      }
      if (occupied) cells[count++] = cell
      if (++work % 512 === 0) await yieldChartPreparation(context)
    }
    return { cells: cells.subarray(0, count), width, height }
  } finally { owner.release() }
}

async function rasterMask(chart: ChartModel, width: number, height: number, grid: number,
  context: ChartComputePreparationContext): Promise<Uint8Array> {
  assertChartPreparationMemory(width * height * (grid * grid * 8 + 1), context)
  const owner = createChartWordCloudCanvas(width * grid, height * grid)
  try {
    const effective = chartWordCloudEffectiveMask(chart.options?.wordCloud)
    if (effective.type === 'image') await drawChartWordCloudMaskImage(owner, effective.value, context)
    else drawChartWordCloudShape(owner.context, effective.value, owner.canvas.width, owner.canvas.height)
    const pixels = owner.context.getImageData(0, 0, owner.canvas.width, owner.canvas.height).data
    let transparent = false
    for (let at = 3; at < pixels.length; at += 4) {
      if (pixels[at]! < 255) transparent = true
      if (at % (CHART_PREPARATION_CHUNK * 4) === 3) await yieldChartPreparation(context)
    }
    const mask = new Uint8Array(width * height)
    for (let cell = 0; cell < mask.length; cell++) {
      const left = cell % width * grid, top = Math.floor(cell / width) * grid
      let inside = true
      for (let y = 0; y < grid; y++) for (let x = 0; x < grid; x++) {
        const at = ((top + y) * owner.canvas.width + left + x) * 4
        const dark = (pixels[at]! + pixels[at + 1]! + pixels[at + 2]!) < 255 * 3 * 0.9
        inside &&= pixels[at + 3]! > 127 && (transparent || dark)
      }
      mask[cell] = Number(inside)
      if (cell % 512 === 0) await yieldChartPreparation(context)
    }
    if (!mask.some(value => value === 1)) throw new ComputeError('COMPUTE_INVALID_INPUT', '词云遮罩没有可排布的有效区域')
    return mask
  } finally { owner.release() }
}

export async function prepareChartWordCloudLayout(data: ChartData, chart: ChartModel,
  context: ChartComputePreparationContext): Promise<ChartData> {
  if (chart.type !== 'wordCloud') return data
  assertChartPreparationCurrent(context)
  const plot = chartPlotFrame(chart, data.series).plot, grid = CHART_WORD_CLOUD_DEFAULTS.gridSize
  const width = Math.max(1, Math.floor(plot.width / grid)), height = Math.max(1, Math.floor(plot.height / grid))
  if (width > 4096 || height > 4096) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '词云画布尺寸超出排布预算')
  const words = chartWordCloudData(data, chart.options?.wordCloud).filter(word => word.text.trim())
  if (words.length > 10_000) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '词云词条数量超出排布预算')
  if (!words.length) {
    // Authored image references must resolve/decode even before words exist;
    // empty data cannot turn a broken resource into a publishable candidate.
    if (chartWordCloudEffectiveMask(chart.options?.wordCloud).type === 'image') await rasterMask(chart, width, height, grid, context)
    return { ...data, wordCloud: { width: width * grid, height: height * grid, words: [], omittedRepeats: 0 } }
  }
  if (!context.runChartTask) throw new ComputeError('COMPUTE_WASM_API_MISSING', '词云排布需要图表 Worker/Rust 计算服务')
  let minimum = Infinity, maximum = 0
  for (const word of words) { minimum = Math.min(minimum, word.value); maximum = Math.max(maximum, word.value) }
  const mask = await rasterMask(chart, width, height, grid, context), random = { seed: CHART_WORD_CLOUD_DEFAULTS.randomSeed }
  const range: readonly [number, number] = [chart.options?.wordCloud?.minFontSize ?? 16, chart.options?.wordCloud?.maxFontSize ?? 56]
  const prepared: ChartWordCloudPlan['words'] = [], glyphs: ChartWordCloudGlyph[] = [], rasters: GlyphRaster[] = []
  let offset = 0, retainedBytes = mask.byteLength + words.length * 1024
  for (const word of words) {
    const fontSize = chartWordCloudFontSize(word.value, [minimum, maximum], range, words.length)
    const rotation = chart.options?.wordCloud?.rotation ?? CHART_WORD_CLOUD_DEFAULTS.rotations[Math.floor(nextRandom(random) * CHART_WORD_CLOUD_DEFAULTS.rotations.length)]!
    const source = data.series[0]!, style = chartPointStyle(chartSeriesStyle(chart, source.field, word.categoryIndex, undefined, source.sourceName ?? source.name), word.categoryIndex, chart.type)
    const font = chartFont({ ...chart.options?.appearance?.font, ...style.label, fontSize })
    const item = { ...word, fontSize, rotation, font }
    const raster = await rasterGlyph(item, grid, retainedBytes, context)
    glyphs.push({ width: raster.width, height: raster.height, offset, length: raster.cells.length, fontSize, weight: word.value })
    rasters.push(raster); prepared.push(item); offset += raster.cells.length; retainedBytes += raster.cells.buffer.byteLength
    assertChartPreparationMemory(retainedBytes, context)
    if (offset > COMPUTE_LIMITS.maxRows || Math.max(offset, width * height) * 2 > COMPUTE_LIMITS.maxElements) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '词云字形批量数据超出 Worker 容量')
  }
  const rowCount = Math.max(width * height, offset, glyphs.length)
  const params: ChartStatisticsParams = { rowCount, columnCount: 2, categoryCount: 1,
    series: [{ field: data.series[0]!.field, seriesPosition: 0, column: 0 }],
    jobs: [{ kind: 'wordCloud', seriesIndex: 0, slot: 0, glyphColumn: 0, maskColumn: 1, width, height, glyphs, maxAttempts: 2400, minFontSize: range[0] }] }
  assertChartPreparationMemory(retainedBytes + estimateChartTaskMemory(params).estimatedPeakBytes, context)
  const values = new Float64Array(rowCount * 2), states = new Uint8Array(rowCount * 2), categoryIds = new Uint32Array(rowCount)
  let at = 0
  for (const raster of rasters) for (const cell of raster.cells) {
    values[at] = cell; states[at++] = 1
    if (at % CHART_PREPARATION_CHUNK === 0) await yieldChartPreparation(context)
  }
  for (let row = 0; row < mask.length; row++) {
    values[rowCount + row] = mask[row]!; states[rowCount + row] = 1
    if (row % CHART_PREPARATION_CHUNK === 0) await yieldChartPreparation(context)
  }
  rasters.length = 0
  const task: ChartStatisticsTask = { kind: 'chartStatistics', values, states, categoryIds, params }
  validateChartTask(task)
  const plan: ChartWordCloudPlan = { params, words: prepared, width: width * grid, height: height * grid, gridSize: grid }
  const output = await context.runChartTask(task)
  assertChartPreparationCurrent(context)
  return { ...data, wordCloud: await decodeChartWordCloudLayout(plan, output, context) }
}

export async function decodeChartWordCloudLayout(plan: ChartWordCloudPlan, candidate: unknown,
  context: ChartComputePreparationContext = {}): Promise<ChartWordCloudLayoutResult> {
  const output = readChartStatisticsOutput(candidate, plan.params), job = plan.params.jobs[0]
  if (!output || output.segments.length !== 1 || job?.kind !== 'wordCloud' || job.glyphs.length !== plan.words.length) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '词云结果与字形任务不匹配')
  const segment = output.segments[0]!, words: ChartCloudWord[] = [], seen = new Set<number>()
  assertChartPreparationMemory(plan.words.length * 2048 + output.values.byteLength + output.indices.byteLength, context)
  for (let index = 0; index < segment.indexLength; index++) {
    const id = output.indices[segment.indexOffset + index]!, source = plan.words[id], glyph = job.glyphs[id], at = segment.valueOffset + index * 3
    const x = output.values[at]!, y = output.values[at + 1]!, scale = output.values[at + 2]!
    if (!source || !glyph || seen.has(id) || ![x, y, scale].every(Number.isFinite) || scale <= 0 || scale > 1
      || scale + 1e-12 < Math.min(1, job.minFontSize / glyph.fontSize)
      || x < glyph.width * scale / 2 || x + glyph.width * scale / 2 > job.width + 1e-9
      || y < glyph.height * scale / 2 || y + glyph.height * scale / 2 > job.height + 1e-9) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '词云字形身份、边界或缩放无效')
    seen.add(id)
    const fontSize = source.fontSize * scale
    words.push({ ...source, x: (x - glyph.width * scale / 2) * plan.gridSize, y: (y - glyph.height * scale / 2) * plan.gridSize,
      width: glyph.width * scale * plan.gridSize, height: glyph.height * scale * plan.gridSize, fontSize,
      font: source.font.replace(/([\d.]+)px /, `${fontSize}px `) })
    if (index % 128 === 0) await yieldChartPreparation(context)
  }
  let omittedRepeats = 0
  for (let index = 0; index < plan.words.length; index++) {
    if (!seen.has(index)) {
      if (plan.words[index]!.repeated) omittedRepeats++
      else throw new ComputeError('COMPUTE_FAILED', '词云词条无法完整排布，请增大图表或减小字号', plan.words[index]!.text)
    }
  }
  assertChartPreparationCurrent(context)
  return { width: plan.width, height: plan.height, words, omittedRepeats }
}

/** Position changes translate the same result; a size change requests a new
 * Worker task through the owning chart signature before it is published. */
export function chartWordCloudLayout(data: ChartData, plot: RenderRect): ChartCloudWord[] {
  const retained = data.wordCloud
  if (!retained) {
    if (!data.categories.length || !data.series.length) return []
    throw new ComputeError('COMPUTE_NOT_READY', '词云尚未获得同一版本的字形排布结果')
  }
  if (Math.abs(retained.width - Math.max(4, Math.floor(plot.width / 4) * 4)) > 1e-9
    || Math.abs(retained.height - Math.max(4, Math.floor(plot.height / 4) * 4)) > 1e-9) throw new ComputeError('COMPUTE_STALE_RESULT', '词云画布尺寸已改变，需要重新排布')
  return retained.words.map(word => ({ ...word, x: plot.x + word.x, y: plot.y + word.y }))
}
