import type { ChartErrorBarOptions, ChartTrendlineOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { chartLineControls, fieldControl, numberControl, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'

function trendAt(config: ChartConfig, index: number): ChartTrendlineOptions { const line = config.options?.trendlines?.[index]; if (!line) throw new RangeError('所选趋势线已移除。'); return line }
function errorAt(config: ChartConfig, index: number): ChartErrorBarOptions { const error = config.options?.errorBars?.[index]; if (!error) throw new RangeError('所选误差线已移除。'); return error }

export function chartTrendlineControls(config: ChartConfig, index: number): ChartSettingControl[] {
  const key = `trendlines.${index}`, read = (item: ChartConfig) => item.options?.trendlines?.[index]
  return [
    fieldControl(`${key}.seriesIndex`, '趋势线系列', item => read(item)?.seriesIndex, (item, value) => { if (value !== undefined) trendAt(item, index).seriesIndex = value }),
    selectControl(`${key}.type`, '趋势类型', [['linear', '线性'], ['exp', '指数'], ['log', '对数'], ['poly', '多项式'], ['pow', '幂'], ['movingAvg', '移动平均']], item => read(item)?.type, (item, value) => { trendAt(item, index).type = value ?? 'linear' }, 'linear'),
    textControl(`${key}.name`, '趋势线名称', item => read(item)?.name, (item, value) => { trendAt(item, index).name = value }),
    ...(config.options?.trendlines?.[index]?.type === 'poly' ? [numberControl(`${key}.order`, '多项式阶数', item => read(item)?.order, (item, value) => { trendAt(item, index).order = value }, 2, 6, 2)] : []),
    ...(config.options?.trendlines?.[index]?.type === 'movingAvg' ? [numberControl(`${key}.period`, '移动平均周期', item => read(item)?.period, (item, value) => { trendAt(item, index).period = value }, 2, 100_000, 2)] : [
      switchControl(`${key}.showFormula`, '显示趋势公式', item => read(item)?.showFormula, (item, value) => { trendAt(item, index).showFormula = value }),
      switchControl(`${key}.showR2`, '显示 R²', item => read(item)?.showR2, (item, value) => { trendAt(item, index).showR2 = value }),
    ]),
    ...chartLineControls(key, read, item => trendAt(item, index)),
  ]
}

export function chartErrorBarControls(config: ChartConfig, index: number, numericX: boolean): ChartSettingControl[] {
  const key = `errorBars.${index}`, read = (item: ChartConfig) => item.options?.errorBars?.[index]
  return [
    fieldControl(`${key}.seriesIndex`, '误差线系列', item => read(item)?.seriesIndex, (item, value) => { if (value !== undefined) errorAt(item, index).seriesIndex = value }),
    selectControl(`${key}.direction`, '方向', numericX ? [['x', 'X 方向'], ['y', 'Y 方向']] : [['y', 'Y 方向']], item => read(item)?.direction, (item, value) => { errorAt(item, index).direction = value ?? 'y' }, 'y'),
    selectControl(`${key}.type`, '正负方向', [['both', '正负'], ['plus', '正向'], ['minus', '负向']], item => read(item)?.type, (item, value) => { errorAt(item, index).type = value ?? 'both' }, 'both'),
    selectControl(`${key}.valueType`, '误差计算', [['fixedVal', '固定值'], ['percentage', '百分比'], ['stdDev', '标准差'], ['stdErr', '标准误差'], ['cust', '自定义字段']], item => read(item)?.valueType, (item, value) => { errorAt(item, index).valueType = value ?? 'fixedVal' }, 'fixedVal'),
    ...(config.options?.errorBars?.[index]?.valueType === 'cust' ? [
      fieldControl(`${key}.plusIndex`, '正误差数据字段', item => read(item)?.plusIndex, (item, value) => { errorAt(item, index).plusIndex = value }),
      fieldControl(`${key}.minusIndex`, '负误差数据字段', item => read(item)?.minusIndex, (item, value) => { errorAt(item, index).minusIndex = value }),
    ] : [numberControl(`${key}.value`, '误差值 / 倍数', item => read(item)?.value, (item, value) => { errorAt(item, index).value = value }, 0, undefined, 1)]),
    switchControl(`${key}.noEndCap`, '隐藏端帽', item => read(item)?.noEndCap, (item, value) => { errorAt(item, index).noEndCap = value }),
    ...chartLineControls(`${key}.line`, item => read(item)?.line, item => errorAt(item, index).line ??= {}),
  ]
}
