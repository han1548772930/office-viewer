import type { ChartConfig } from './chart-engine'
import type { ChartRelationOptions } from '../../domain/model'
import { CHART_RELATION_DEFAULTS } from './chart-relation'
import { CHART_POINT_SHAPE_CHOICES } from './chart-markers'
import { fieldControl, numberControl, optionGroup, selectControl, switchControl, type ChartSettingControl } from './chart-settings-controls'

function reset(control: ChartSettingControl, key: keyof ChartRelationOptions): ChartSettingControl {
  return { ...control, canReset: config => config.options?.relation?.[key] !== undefined,
    reset: config => { if (config.options?.relation) config.options.relation[key] = undefined } }
}
export function chartRelationMappingControls(fieldCount: number): ChartSettingControl[] {
  const edgeOnly = (item: ChartConfig) => item.options?.mapping?.relationFormat !== 'coWordMatrix'
  const edge = (item: ChartConfig) => (optionGroup(item, 'mapping').edge ??= { sourceIndex: 0, targetIndex: 1,
    valueIndex: fieldCount > 2 ? 2 : undefined })
  return [
    { ...selectControl('mapping.relationFormat', '关系数据格式', [['edgeList', '来源、目标、权重'], ['coWordMatrix', '共词矩阵']],
      item => item.options?.mapping?.relationFormat, (item, value) => { optionGroup(item, 'mapping').relationFormat = value }, 'edgeList'),
      extra: '共词矩阵使用名称、分类两列，后续列标题为节点名称、数值为关系权重。',
      canReset: item => item.options?.mapping?.relationFormat !== undefined,
      reset: item => { if (item.options?.mapping) item.options.mapping.relationFormat = undefined } },
    { ...fieldControl('mapping.edge.sourceIndex', '来源字段', item => item.options?.mapping?.edge?.sourceIndex ?? 0,
      (item, value) => { edge(item).sourceIndex = value ?? 0 }), visible: edgeOnly },
    { ...fieldControl('mapping.edge.targetIndex', '目标字段', item => item.options?.mapping?.edge?.targetIndex ?? 1,
      (item, value) => { edge(item).targetIndex = value ?? 1 }), visible: edgeOnly },
    { ...fieldControl('mapping.edge.valueIndex', '权重字段（可选）', item => item.options?.mapping?.edge
      ? item.options.mapping.edge.valueIndex : fieldCount > 2 ? 2 : undefined,
      (item, value) => { edge(item).valueIndex = value }), visible: edgeOnly },
  ]
}
export function chartRelationSettingsControls(): ChartSettingControl[] {
  const forceOnly = (config: ChartConfig) => config.options?.relation?.layout === 'force'
  const circularOnly = (config: ChartConfig) => config.options?.relation?.layout === 'circular'
  return [
    reset(selectControl('relation.layout', '关系布局', [['none', '固定布局'], ['circular', '圆形'], ['force', '力导向']], item => item.options?.relation?.layout,
      (item, value) => { optionGroup(item, 'relation').layout = value }, CHART_RELATION_DEFAULTS.layout), 'layout'),
    reset(switchControl('relation.useValueAsSymbolSize', '权重决定节点大小', item => item.options?.relation?.useValueAsSymbolSize,
      (item, value) => { optionGroup(item, 'relation').useValueAsSymbolSize = value }, true), 'useValueAsSymbolSize'),
    reset(switchControl('relation.emphasisEnabled', '悬浮强调邻接节点和连线', item => item.options?.relation?.emphasisEnabled,
      (item, value) => { optionGroup(item, 'relation').emphasisEnabled = value }, true), 'emphasisEnabled'),
    { ...reset(switchControl('relation.circularLabelRotation', '旋转圆形布局标签', item => item.options?.relation?.circularLabelRotation,
      (item, value) => { optionGroup(item, 'relation').circularLabelRotation = value }, false), 'circularLabelRotation'), visible: circularOnly },
    reset(selectControl('relation.nodeShape', '节点形状', CHART_POINT_SHAPE_CHOICES, item => item.options?.relation?.nodeShape,
      (item, value) => { optionGroup(item, 'relation').nodeShape = value }, 'circle'), 'nodeShape'),
    { ...numberControl('relation.force.repulsion', '排斥强度', item => item.options?.relation?.force?.repulsion,
      (item, value) => { (optionGroup(item, 'relation').force ??= {}).repulsion = value }, 0, 100_000, 5000), visible: forceOnly },
    { ...numberControl('relation.force.gravity', '中心引力', item => item.options?.relation?.force?.gravity,
      (item, value) => { (optionGroup(item, 'relation').force ??= {}).gravity = value }, 0, 10, 0.1), visible: forceOnly },
    { ...numberControl('relation.force.edgeLength', '连线长度', item => item.options?.relation?.force?.edgeLength,
      (item, value) => { (optionGroup(item, 'relation').force ??= {}).edgeLength = value }, 0, 10_000),
      visible: forceOnly, extra: '默认使用图表短边的五分之一。', reset: config => { if (config.options?.relation) config.options.relation.force = undefined },
      resetLabel: '恢复力导向默认', canReset: config => config.options?.relation?.force !== undefined },
  ]
}
