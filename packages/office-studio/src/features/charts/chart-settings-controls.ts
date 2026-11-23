import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartLegendOptions, ChartLineOptions, ChartSemanticOptions, ChartSeriesOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'

interface ControlBase { key: string; label: string; extra?: string }
export type ChartSettingControl = ControlBase & (
  | { kind: 'number'; read: (config: ChartConfig) => number | undefined; write: (config: ChartConfig, value: number | undefined) => void; min?: number; max?: number; step?: number; fallback?: number }
  | { kind: 'text' | 'color'; read: (config: ChartConfig) => string | undefined; write: (config: ChartConfig, value: string | undefined) => void }
  | { kind: 'switch'; read: (config: ChartConfig) => boolean | undefined; write: (config: ChartConfig, value: boolean) => void; fallback?: boolean }
  | { kind: 'select'; read: (config: ChartConfig) => string | undefined; write: (config: ChartConfig, value: string | undefined) => void; choices: Array<{ label: string; value: string }>; fallback?: string }
  | { kind: 'field'; read: (config: ChartConfig) => number | undefined; write: (config: ChartConfig, value: number | undefined) => void }
  | { kind: 'fields' | 'points'; read: (config: ChartConfig) => number[] | undefined; write: (config: ChartConfig, value: number[]) => void }
)

export function numberControl(key: string, label: string, read: (config: ChartConfig) => number | undefined, write: (config: ChartConfig, value: number | undefined) => void, min?: number, max?: number, fallback?: number): ChartSettingControl {
  return { key, label, kind: 'number', read, write, min, max, fallback }
}
export function textControl(key: string, label: string, read: (config: ChartConfig) => string | undefined, write: (config: ChartConfig, value: string | undefined) => void): ChartSettingControl { return { key, label, kind: 'text', read, write } }
export function colorControl(key: string, label: string, read: (config: ChartConfig) => string | undefined, write: (config: ChartConfig, value: string | undefined) => void): ChartSettingControl { return { key, label, kind: 'color', read, write } }
export function switchControl(key: string, label: string, read: (config: ChartConfig) => boolean | undefined, write: (config: ChartConfig, value: boolean) => void, fallback = false): ChartSettingControl { return { key, label, kind: 'switch', read, write, fallback } }
export function fieldControl(key: string, label: string, read: (config: ChartConfig) => number | undefined, write: (config: ChartConfig, value: number | undefined) => void): ChartSettingControl { return { key, label, kind: 'field', read, write } }
export function fieldsControl(key: string, label: string, read: (config: ChartConfig) => number[] | undefined, write: (config: ChartConfig, value: number[]) => void, points = false): ChartSettingControl { return { key, label, kind: points ? 'points' : 'fields', read, write } }
export function selectControl<T extends string>(key: string, label: string, values: ReadonlyArray<readonly [T, string]>, read: (config: ChartConfig) => T | undefined, write: (config: ChartConfig, value: T | undefined) => void, fallback?: T): ChartSettingControl {
  return { key, label, kind: 'select', read, fallback, choices: values.map(([value, title]) => ({ value, label: title })),
    write: (config, value) => write(config, values.find(([choice]) => choice === value)?.[0]) }
}

type ChartObjectOptionGroup = Exclude<keyof ChartSemanticOptions, 'invalidValueStrategy' | 'plotVisibleOnly' | 'series' | 'trendlines' | 'errorBars'>
export function optionGroup<K extends ChartObjectOptionGroup>(config: ChartConfig, key: K): NonNullable<ChartSemanticOptions[K]> {
  const options = config.options ??= {}
  if (options[key] === undefined) Object.assign(options, { [key]: {} })
  // K is restricted to optional object groups, initialized above; callers never
  // create a required mapping role through this helper.
  return options[key] as NonNullable<ChartSemanticOptions[K]>
}
export function seriesStyleGroup(config: ChartConfig, field?: number): ChartSeriesOptions {
  const options = config.options ??= {}
  if (field === undefined) return options.allSeriesStyle ??= {}
  return (options.series ??= {})[field] ??= {}
}

export function chartFontControls(key: string, read: (config: ChartConfig) => ChartFontOptions | undefined, write: (config: ChartConfig) => ChartFontOptions): ChartSettingControl[] {
  return [
    textControl(`${key}.family`, '字体', config => read(config)?.family, (config, value) => { write(config).family = value }),
    numberControl(`${key}.fontSize`, '字号', config => read(config)?.fontSize, (config, value) => { write(config).fontSize = value }, 1, 256),
    colorControl(`${key}.color`, '文字颜色', config => read(config)?.color, (config, value) => { write(config).color = value }),
    switchControl(`${key}.bold`, '粗体', config => read(config)?.bold, (config, value) => { write(config).bold = value }),
    switchControl(`${key}.italic`, '斜体', config => read(config)?.italic, (config, value) => { write(config).italic = value }),
  ]
}

export function chartLineControls(key: string, read: (config: ChartConfig) => ChartLineOptions | undefined, write: (config: ChartConfig) => ChartLineOptions, defaultVisible = true): ChartSettingControl[] {
  return [
    switchControl(`${key}.visible`, '显示线条', config => read(config)?.visible, (config, value) => { write(config).visible = value }, defaultVisible),
    colorControl(`${key}.color`, '线条颜色', config => read(config)?.color, (config, value) => { write(config).color = value }),
    numberControl(`${key}.width`, '线宽', config => read(config)?.width, (config, value) => { write(config).width = value }, 0, 30),
    numberControl(`${key}.opacity`, '线条不透明度', config => read(config)?.opacity, (config, value) => { write(config).opacity = value }, 0, 1, 1),
    selectControl(`${key}.dashType`, '线型', [['solid', '实线'], ['dashed', '虚线'], ['dotted', '点线']], config => read(config)?.dashType, (config, value) => { write(config).dashType = value }, 'solid'),
  ]
}

export function chartLabelControls(key: string, read: (config: ChartConfig) => ChartLabelOptions | undefined, write: (config: ChartConfig) => ChartLabelOptions, defaultVisible = false): ChartSettingControl[] {
  return [
    switchControl(`${key}.visible`, '显示数据标签', config => read(config)?.visible, (config, value) => { write(config).visible = value }, defaultVisible),
    selectControl(`${key}.contentType`, '标签内容', [['value', '数值'], ['category', '分类'], ['series', '系列'], ['percentage', '百分比']], config => read(config)?.contentType, (config, value) => { write(config).contentType = value }, 'value'),
    selectControl(`${key}.position`, '标签位置', [['inside', '内部'], ['outside', '外部'], ['top', '上方'], ['bottom', '下方'], ['left', '左侧'], ['right', '右侧'], ['center', '中心']], config => read(config)?.position, (config, value) => { write(config).position = value }),
    textControl(`${key}.format`, '标签数字格式', config => read(config)?.format, (config, value) => { write(config).format = value }),
    ...chartFontControls(key, read, write),
  ]
}

export function chartCommonControls(): ChartSettingControl[] {
  return [
    textControl('title', '图表标题', config => config.title, (config, value) => { config.title = value }),
    textControl('subtitle', '副标题', config => config.subtitle, (config, value) => { config.subtitle = value }),
    switchControl('legend.visible', '显示图例', config => config.legend?.visible, (config, value) => { (config.legend ??= {}).visible = value }, true),
    selectControl('legend.position', '图例位置', [['top', '上方'], ['right', '右侧'], ['bottom', '下方'], ['left', '左侧']], config => config.legend?.position, (config, value) => { (config.legend ??= {}).position = value }, 'bottom'),
    colorControl('appearance.background', '图表背景', config => config.options?.appearance?.background, (config, value) => { optionGroup(config, 'appearance').background = value }),
    switchControl('plotVisibleOnly', '仅绘制可见单元格', config => config.options?.plotVisibleOnly, (config, value) => { (config.options ??= {}).plotVisibleOnly = value }),
    selectControl('invalidValueStrategy', '空值处理', [['break', '断开'], ['link', '连接'], ['zero', '按零值']], config => config.options?.invalidValueStrategy, (config, value) => { (config.options ??= {}).invalidValueStrategy = value }, 'break'),
    numberControl('size.width', '宽度', config => config.size.width, (config, value) => { if (value !== undefined) config.size.width = value }, 1, 100_000),
    numberControl('size.height', '高度', config => config.size.height, (config, value) => { if (value !== undefined) config.size.height = value }, 1, 100_000),
  ]
}

export function chartSeriesControls(field?: number, combination = false, secondaryAxis = false): ChartSettingControl[] {
  const read = (config: ChartConfig): ChartSeriesOptions | undefined => field === undefined ? config.options?.allSeriesStyle : config.options?.series?.[field]
  const write = (config: ChartConfig): ChartSeriesOptions => seriesStyleGroup(config, field)
  const key = field === undefined ? 'allSeriesStyle' : `series.${field}`
  return [
    ...(field === undefined ? [] : [textControl(`${key}.name`, '系列名称', config => read(config)?.name, (config, value) => { write(config).name = value })]),
    ...(combination ? [selectControl(`${key}.type`, '系列图型', [['column', '柱形'], ['columnStacked', '堆积柱形'], ['line', '折线'], ['area', '面积'], ['areaStacked', '堆积面积']], config => read(config)?.type, (config, value) => { write(config).type = value })] : []),
    ...(secondaryAxis ? [switchControl(`${key}.rightAxis`, '使用右侧数值轴', config => read(config)?.rightAxis, (config, value) => { write(config).rightAxis = value })] : []),
    colorControl(`${key}.color`, '系列颜色', config => read(config)?.color, (config, value) => { write(config).color = value }),
    numberControl(`${key}.fillOpacity`, '填充不透明度', config => read(config)?.fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1),
    selectControl(`${key}.point.shape`, '标记形状', [['circle', '圆形'], ['rect', '方形'], ['triangle', '三角形'], ['diamond', '菱形'], ['none', '无标记']], config => read(config)?.point?.shape, (config, value) => { (write(config).point ??= {}).shape = value }),
    numberControl(`${key}.point.size`, '标记大小', config => read(config)?.point?.size, (config, value) => { (write(config).point ??= {}).size = value }, 0, 256),
    colorControl(`${key}.point.color`, '标记颜色', config => read(config)?.point?.color, (config, value) => { (write(config).point ??= {}).color = value }),
    ...chartLineControls(`${key}.border`, config => read(config)?.border, config => write(config).border ??= {}),
    ...chartLabelControls(`${key}.label`, config => read(config)?.label, config => write(config).label ??= {}),
  ]
}

function axisGroup(config: ChartConfig, axis: 'xAxis' | 'yAxis' | 'rightYAxis'): ChartAxisOptions {
  if (axis === 'rightYAxis') return (config.options ??= {}).rightYAxis ??= { scaleType: 'linear' }
  const current = config[axis]
  const next: ChartAxisOptions = { scaleType: 'linear', ...current }
  config[axis] = next
  return next
}

export function chartAxisControls(axis: 'xAxis' | 'yAxis' | 'rightYAxis', allowScale = true): ChartSettingControl[] {
  const read = (config: ChartConfig): Partial<ChartAxisOptions> | undefined => axis === 'rightYAxis' ? config.options?.rightYAxis : config[axis]
  const write = (config: ChartConfig): ChartAxisOptions => axisGroup(config, axis)
  return [
    textControl(`${axis}.title`, '轴标题', config => read(config)?.title, (config, value) => { write(config).title = value }),
    switchControl(`${axis}.reverse`, '反向轴', config => read(config)?.reverse, (config, value) => { write(config).reverse = value }),
    ...(allowScale ? [
      selectControl(`${axis}.scaleType`, '刻度', [['linear', '线性'], ['log', '对数']], config => read(config)?.scaleType, (config, value) => { write(config).scaleType = value ?? 'linear' }, 'linear'),
      numberControl(`${axis}.logBase`, '对数底数', config => read(config)?.logBase, (config, value) => { write(config).logBase = value }, 2, 1000, 10),
      numberControl(`${axis}.min`, '最小刻度', config => read(config)?.min, (config, value) => { write(config).min = value }),
      numberControl(`${axis}.max`, '最大刻度', config => read(config)?.max, (config, value) => { write(config).max = value }),
      textControl(`${axis}.format`, '轴数字格式', config => read(config)?.format, (config, value) => { write(config).format = value }),
    ] : []),
    switchControl(`${axis}.label.visible`, '显示轴标签', config => read(config)?.label?.visible, (config, value) => { (write(config).label ??= {}).visible = value }, true),
    numberControl(`${axis}.label.rotate`, '标签旋转', config => read(config)?.label?.rotate, (config, value) => { (write(config).label ??= {}).rotate = value }, -90, 90),
    ...chartFontControls(`${axis}.label`, config => read(config)?.label, config => write(config).label ??= {}),
    ...chartLineControls(`${axis}.line`, config => read(config)?.line, config => write(config).line ??= {}),
    ...chartLineControls(`${axis}.gridLine`, config => read(config)?.gridLine, config => write(config).gridLine ??= {}),
    switchControl(`${axis}.tick.visible`, '显示刻度线', config => read(config)?.tick?.visible, (config, value) => { (write(config).tick ??= {}).visible = value }),
    numberControl(`${axis}.tick.length`, '刻度线长度', config => read(config)?.tick?.length, (config, value) => { (write(config).tick ??= {}).length = value }, 0, 50, 4),
    selectControl(`${axis}.tick.position`, '刻度线位置', [['inside', '内部'], ['outside', '外部']], config => read(config)?.tick?.position, (config, value) => { (write(config).tick ??= {}).position = value }, 'outside'),
  ]
}

export function chartLegendFont(config: ChartConfig): ChartFontOptions { const legend: Partial<ChartLegendOptions> = config.legend ??= {}; return legend.font ??= {} }
