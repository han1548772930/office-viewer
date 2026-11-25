import type {
  ChartChordEmphasisOptions, ChartChordLabelOptions, ChartChordLayoutOptions, ChartChordMapping,
  ChartChordNodeOptions, ChartChordOptions, ChartChordRibbonOptions, ChartNullablePatch, ChartType,
} from '../../domain/model'
import { chartNullableValues } from './chart-cartesian-api'
import { validateChartChordColor } from './chart-chord-color'
import { chartChordAuthoredLayout, chartChordAuthoredRibbon } from './chart-chord'

function input(value: unknown, fields: readonly string[], name: string): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype
    || Object.getOwnPropertySymbols(value).length || Object.keys(value).some(key => !fields.includes(key))) {
    throw new TypeError(`${name}包含无效字段。`)
  }
}
function number(value: unknown, name: string, minimum = -Number.MAX_VALUE, maximum = Number.MAX_VALUE): void {
  if (value !== undefined && (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > maximum)) {
    throw new RangeError(`${name}须为 ${minimum >= 0 ? '非负' : ''}有限数值${maximum === 1 ? '，且不能大于 1' : ''}。`)
  }
}
function boolean(value: unknown, name: string): void {
  if (value !== undefined && typeof value !== 'boolean') throw new TypeError(`${name}须为布尔值。`)
}
function choice(value: unknown, values: readonly string[], name: string): void {
  if (value !== undefined && (typeof value !== 'string' || !values.includes(value))) throw new TypeError(`${name}包含无效选项。`)
}
function color(value: unknown, name: string): void {
  if (value === undefined) return
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${name}须为颜色字符串。`)
  validateChartChordColor(value, name)
}
function index(value: unknown, name: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value > 1_048_575) {
    throw new RangeError(`${name}须为数据区域内的非负字段索引。`)
  }
  return value
}

export function assertChartChordType(type: ChartType): void {
  if (type !== 'chord') throw new TypeError('和弦图专属设置仅适用于和弦图。')
}

export function chartChordMappingValues(value: ChartChordMapping): ChartChordMapping {
  input(value, ['sourceIndex', 'targetIndex', 'valueIndex'], '和弦图映射')
  const sourceIndex = index(value.sourceIndex, '和弦来源字段'), targetIndex = index(value.targetIndex, '和弦目标字段')
  const valueIndex = value.valueIndex === undefined ? undefined : index(value.valueIndex, '和弦权重字段')
  if (sourceIndex === targetIndex || valueIndex === sourceIndex || valueIndex === targetIndex) {
    throw new RangeError('和弦来源、目标和权重须使用互不重复的字段。')
  }
  return { sourceIndex, targetIndex, ...(valueIndex === undefined ? {} : { valueIndex }) }
}

/** Validate the authored leaves only. The final model merge validates the two
 * radii together; a partial inner-radius patch must not inherit a default outer
 * radius before it is merged with the current chart. Null resets stay explicit. */
export function chartChordLayoutValues(value: ChartNullablePatch<ChartChordLayoutOptions>): ChartChordLayoutOptions {
  input(value, ['startAngle', 'clockwise', 'padAngle', 'minAngle', 'innerRadiusRatio', 'outerRadiusRatio'], '和弦图布局')
  const patch = chartNullableValues<ChartChordLayoutOptions>(value)
  number(patch.startAngle, '和弦起始角'); boolean(patch.clockwise, '和弦方向')
  number(patch.padAngle, '和弦节点间隔角', 0); number(patch.minAngle, '和弦最小节点角', 0)
  number(patch.innerRadiusRatio, '和弦内半径比例', 0, 1); number(patch.outerRadiusRatio, '和弦外半径比例', 0, 1)
  return patch
}

export function chartChordLayoutPatch(options: ChartChordOptions | undefined,
  value: ChartNullablePatch<ChartChordLayoutOptions>): ChartChordOptions {
  return { layout: { ...chartChordAuthoredLayout(options), ...chartChordLayoutValues(value) }, startAngle: undefined,
    clockwise: undefined, padAngle: undefined, minAngle: undefined, innerRadiusRatio: undefined, outerRadiusRatio: undefined }
}

export function chartChordNodeValues(value: ChartNullablePatch<ChartChordNodeOptions>): ChartChordNodeOptions {
  input(value, ['color', 'opacity', 'borderColor', 'borderWidth', 'borderType'], '和弦图节点')
  const patch = chartNullableValues<ChartChordNodeOptions>(value)
  color(patch.color, '和弦节点颜色'); color(patch.borderColor, '和弦节点边框颜色')
  number(patch.opacity, '和弦节点不透明度', 0, 1); number(patch.borderWidth, '和弦节点边框宽度', 0)
  choice(patch.borderType, ['solid', 'dashed', 'dotted'], '和弦节点边框类型')
  return patch
}

export function chartChordLabelValues(value: ChartNullablePatch<ChartChordLabelOptions>): ChartChordLabelOptions {
  input(value, ['visible', 'position', 'rotation', 'distance', 'color', 'fontSize', 'bold', 'italic', 'family'], '和弦图标签')
  const patch = chartNullableValues<ChartChordLabelOptions>(value)
  boolean(patch.visible, '和弦标签显示'); boolean(patch.bold, '和弦标签粗体'); boolean(patch.italic, '和弦标签斜体')
  choice(patch.position, ['inside', 'outside'], '和弦标签位置'); number(patch.rotation, '和弦标签旋转角')
  number(patch.distance, '和弦标签距离', 0); number(patch.fontSize, '和弦标签字号', 0)
  color(patch.color, '和弦标签颜色')
  if (patch.family !== undefined && (typeof patch.family !== 'string' || !patch.family.trim())) throw new TypeError('和弦标签字体须为非空字符串。')
  return patch
}

export function chartChordRibbonValues(value: ChartNullablePatch<ChartChordRibbonOptions>): ChartChordRibbonOptions {
  input(value, ['color', 'opacity', 'width', 'type'], '和弦图弦带')
  const patch = chartNullableValues<ChartChordRibbonOptions>(value)
  choice(patch.color, ['source', 'target', 'gradient'], '和弦弦带颜色模式')
  number(patch.opacity, '和弦弦带不透明度', 0, 1); number(patch.width, '和弦弦带边框宽度', 0)
  choice(patch.type, ['solid', 'dashed', 'dotted'], '和弦弦带边框类型')
  return patch
}

export function chartChordRibbonPatch(options: ChartChordOptions | undefined,
  value: ChartNullablePatch<ChartChordRibbonOptions>): ChartChordOptions {
  return { ribbon: { ...chartChordAuthoredRibbon(options), ...chartChordRibbonValues(value) }, ribbonOpacity: undefined }
}

export function chartChordEmphasisValues(value: ChartNullablePatch<ChartChordEmphasisOptions>): ChartChordEmphasisOptions {
  input(value, ['focus'], '和弦图强调')
  const patch = chartNullableValues<ChartChordEmphasisOptions>(value)
  choice(patch.focus, ['adjacency', 'self', 'none'], '和弦悬停聚焦')
  return patch
}
