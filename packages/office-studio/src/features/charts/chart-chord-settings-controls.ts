import type {
  ChartChordEmphasisOptions, ChartChordLabelOptions, ChartChordLayoutOptions, ChartChordMapping,
  ChartChordNodeOptions, ChartChordRibbonOptions, ChartNullablePatch,
} from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { chartChordEmphasisValues, chartChordLabelValues, chartChordLayoutPatch, chartChordMappingValues,
  chartChordNodeValues, chartChordRibbonPatch } from './chart-chord-api'
import { CHART_CHORD_LAYOUT_KEYS, resolveChartChordLayout, resolveChartChordOptions } from './chart-chord'
import { colorControl, fieldControl, numberControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'

const layoutKeys = CHART_CHORD_LAYOUT_KEYS
const lineTypes = [['solid', '实线'], ['dashed', '虚线'], ['dotted', '点线']] as const
const resolved = (config: ChartConfig) => resolveChartChordOptions(config.options?.chord, config.options?.appearance?.font)

function editLayout(config: ChartConfig, patch: ChartNullablePatch<ChartChordLayoutOptions>): void {
  Object.assign(optionGroup(config, 'chord'), chartChordLayoutPatch(config.options?.chord, patch))
}
function editNode(config: ChartConfig, patch: ChartNullablePatch<ChartChordNodeOptions>): void {
  const values = chartChordNodeValues(patch)
  Object.assign(optionGroup(config, 'chord').node ??= {}, values)
}
function editLabel(config: ChartConfig, patch: ChartNullablePatch<ChartChordLabelOptions>): void {
  const values = chartChordLabelValues(patch)
  Object.assign(optionGroup(config, 'chord').label ??= {}, values)
}
function editRibbon(config: ChartConfig, patch: ChartNullablePatch<ChartChordRibbonOptions>): void {
  Object.assign(optionGroup(config, 'chord'), chartChordRibbonPatch(config.options?.chord, patch))
}
function editEmphasis(config: ChartConfig, patch: ChartNullablePatch<ChartChordEmphasisOptions>): void {
  const values = chartChordEmphasisValues(patch)
  Object.assign(optionGroup(config, 'chord').emphasis ??= {}, values)
}

export function chartChordLayoutControls(): ChartSettingControl[] {
  return [
    numberControl('chord.layout.innerRadiusRatio', '内半径比例', config => resolved(config).layout.innerRadiusRatio, (config, value) => editLayout(config, { innerRadiusRatio: value }), 0, 1),
    numberControl('chord.layout.outerRadiusRatio', '外半径比例', config => resolved(config).layout.outerRadiusRatio, (config, value) => editLayout(config, { outerRadiusRatio: value }), 0, 1),
    numberControl('chord.layout.padAngle', '节点间隔角', config => resolved(config).layout.padAngle, (config, value) => editLayout(config, { padAngle: value }), 0),
    numberControl('chord.layout.startAngle', '和弦起始角', config => resolved(config).layout.startAngle, (config, value) => editLayout(config, { startAngle: value })),
    switchControl('chord.layout.clockwise', '顺时针', config => resolved(config).layout.clockwise, (config, value) => editLayout(config, { clockwise: value }), true),
    numberControl('chord.layout.minAngle', '最小节点角', config => resolved(config).layout.minAngle, (config, value) => editLayout(config, { minAngle: value }), 0),
  ]
}
export function chartChordNodeControls(): ChartSettingControl[] {
  return [
    colorControl('chord.node.color', '节点颜色', config => config.options?.chord?.node?.color, (config, value) => editNode(config, { color: value })),
    numberControl('chord.node.opacity', '节点不透明度', config => resolved(config).node.opacity, (config, value) => editNode(config, { opacity: value }), 0, 1),
    colorControl('chord.node.borderColor', '节点边框颜色', config => resolved(config).node.borderColor, (config, value) => editNode(config, { borderColor: value })),
    numberControl('chord.node.borderWidth', '节点边框宽度', config => resolved(config).node.borderWidth, (config, value) => editNode(config, { borderWidth: value }), 0),
    selectControl('chord.node.borderType', '节点边框类型', lineTypes, config => resolved(config).node.borderType, (config, value) => editNode(config, { borderType: value }), 'solid'),
  ]
}
export function chartChordLabelControls(): ChartSettingControl[] {
  const visible = (config: ChartConfig) => resolved(config).label.visible
  const controls: ChartSettingControl[] = [
    switchControl('chord.label.visible', '显示节点标签', config => resolved(config).label.visible, (config, value) => editLabel(config, { visible: value }), true),
    selectControl('chord.label.position', '节点标签位置', [['inside', '内部'], ['outside', '外部']], config => resolved(config).label.position, (config, value) => editLabel(config, { position: value }), 'outside'),
    colorControl('chord.label.color', '节点标签颜色', config => config.options?.chord?.label?.color ?? config.options?.appearance?.font?.color ?? 'gray.700', (config, value) => editLabel(config, { color: value })),
    numberControl('chord.label.fontSize', '节点标签字号', config => resolved(config).label.fontSize, (config, value) => editLabel(config, { fontSize: value }), 0),
    switchControl('chord.label.bold', '节点标签粗体', config => resolved(config).label.bold, (config, value) => editLabel(config, { bold: value })),
    switchControl('chord.label.italic', '节点标签斜体', config => resolved(config).label.italic, (config, value) => editLabel(config, { italic: value })),
    textControl('chord.label.family', '节点标签字体', config => resolved(config).label.family, (config, value) => editLabel(config, { family: value })),
    numberControl('chord.label.rotation', '节点标签旋转角', config => resolved(config).label.rotation, (config, value) => editLabel(config, { rotation: value })),
    numberControl('chord.label.distance', '节点标签距离', config => resolved(config).label.distance, (config, value) => editLabel(config, { distance: value }), 0),
  ]
  return controls.map((control, index) => index === 0 ? control : { ...control, visible })
}
export function chartChordRibbonControls(): ChartSettingControl[] {
  return [
    selectControl('chord.ribbon.color', '弦带颜色模式', [['source', '来源'], ['target', '目标'], ['gradient', '渐变']], config => resolved(config).ribbon.color, (config, value) => editRibbon(config, { color: value }), 'gradient'),
    numberControl('chord.ribbon.opacity', '弦带不透明度', config => resolved(config).ribbon.opacity, (config, value) => editRibbon(config, { opacity: value }), 0, 1),
    numberControl('chord.ribbon.width', '弦带边框宽度', config => resolved(config).ribbon.width, (config, value) => editRibbon(config, { width: value }), 0),
    selectControl('chord.ribbon.type', '弦带边框类型', lineTypes, config => resolved(config).ribbon.type, (config, value) => editRibbon(config, { type: value }), 'solid'),
  ]
}
export function chartChordEmphasisControls(): ChartSettingControl[] {
  return [selectControl('chord.emphasis.focus', '悬停聚焦', [['adjacency', '相邻节点与弦带'], ['self', '自身'], ['none', '无强调']],
    config => resolved(config).emphasis.focus, (config, value) => editEmphasis(config, { focus: value }), 'adjacency')]
}

export function chartChordSettingsRadiiError(config: ChartConfig): string | undefined {
  const { innerRadiusRatio: inner, outerRadiusRatio: outer } = resolveChartChordLayout(config.options?.chord)
  return !Number.isFinite(inner) || !Number.isFinite(outer) || inner < 0 || outer > 1 || inner >= outer
    ? '内半径须小于外半径，且两者须位于 0 到 1 之间。修正后可应用设置。' : undefined
}
export type ChartChordSettingsGroup = 'layout' | 'node' | 'label' | 'ribbon' | 'emphasis' | 'all'
export function canResetChartChordSettings(config: ChartConfig, target: ChartChordSettingsGroup): boolean {
  const chord = config.options?.chord
  if (!chord) return false
  return target === 'all' || chord[target] !== undefined || target === 'layout' && layoutKeys.some(key => chord[key] !== undefined)
    || target === 'ribbon' && chord.ribbonOpacity !== undefined
}
export function clearChartChordSettings(config: ChartConfig, target: ChartChordSettingsGroup): void {
  const chord = config.options?.chord
  if (!chord) return
  if (target === 'all') { config.options!.chord = undefined; return }
  chord[target] = undefined
  if (target === 'layout') for (const key of layoutKeys) chord[key] = undefined
  if (target === 'ribbon') chord.ribbonOpacity = undefined
}
export function resetChartChordLabelFont(config: ChartConfig, target: 'color' | 'font'): void {
  const label = config.options?.chord?.label
  if (!label) return
  if (target === 'color') label.color = undefined
  else { label.fontSize = undefined; label.bold = undefined; label.italic = undefined; label.family = undefined }
}

export function chartChordSettingsMapping(config: ChartConfig, fieldCount: number): ChartChordMapping {
  const authored = config.options?.mapping?.chord ?? config.options?.mapping?.edge
  return authored ? { ...authored } : { sourceIndex: config.categoryField ?? 0, targetIndex: config.valueFields?.[0] ?? 1,
    valueIndex: config.valueFields?.[1] ?? (fieldCount > 2 ? 2 : undefined) }
}
export function chartChordSettingsMappingError(config: ChartConfig, fieldCount: number): string | undefined {
  try {
    const mapping = chartChordMappingValues(chartChordSettingsMapping(config, fieldCount))
    if (!Number.isSafeInteger(fieldCount) || fieldCount < 2 || [mapping.sourceIndex, mapping.targetIndex, ...(mapping.valueIndex === undefined ? [] : [mapping.valueIndex])].some(index => index >= fieldCount)) {
      throw new RangeError('来源、目标和权重字段须位于当前数据区域内。')
    }
    return undefined
  } catch (cause) { return cause instanceof Error ? cause.message : '和弦字段映射无效。' }
}
export function chartChordMappingControls(fieldCount: number): ChartSettingControl[] {
  return [
    fieldControl('mapping.chord.sourceIndex', '和弦来源字段', config => chartChordSettingsMapping(config, fieldCount).sourceIndex, (config, value) => setChartChordSettingsField(config, 'sourceIndex', value, fieldCount)),
    fieldControl('mapping.chord.targetIndex', '和弦目标字段', config => chartChordSettingsMapping(config, fieldCount).targetIndex, (config, value) => setChartChordSettingsField(config, 'targetIndex', value, fieldCount)),
    fieldControl('mapping.chord.valueIndex', '和弦权重字段', config => chartChordSettingsMapping(config, fieldCount).valueIndex, (config, value) => setChartChordSettingsField(config, 'valueIndex', value, fieldCount)),
  ]
}
/** UI draft edits may temporarily duplicate roles while the user swaps two
 * fields. The panel diagnoses the complete mapping and Apply validates it
 * atomically. Required roles are never silently cleared or replaced. */
export function setChartChordSettingsField(config: ChartConfig, role: keyof ChartChordMapping, value: number | undefined, fieldCount: number): void {
  if (!['sourceIndex', 'targetIndex', 'valueIndex'].includes(role)) throw new TypeError('和弦字段角色无效。')
  if (value === undefined && role !== 'valueIndex') return
  if (value !== undefined && (!Number.isSafeInteger(value) || value < 0 || value >= fieldCount)) throw new RangeError('字段须位于当前数据区域内。')
  const current = chartChordSettingsMapping(config, fieldCount), next = { ...current, [role]: value }
  if (value === undefined) delete next.valueIndex
  const mapping = optionGroup(config, 'mapping')
  mapping.chord = next
  mapping.edge = undefined
  config.categoryField = next.sourceIndex
  config.valueFields = [next.targetIndex, ...(next.valueIndex === undefined ? [] : [next.valueIndex])]
}
