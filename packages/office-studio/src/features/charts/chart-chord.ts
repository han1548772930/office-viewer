import type {
  ChartChordEmphasisOptions, ChartChordLabelOptions, ChartChordLayoutOptions, ChartChordNodeOptions,
  ChartChordOptions, ChartChordRibbonOptions, ChartFontOptions, ChartLabelOptions,
} from '../../domain/model'

export const CHART_CHORD_DEFAULT_LAYOUT: Readonly<Required<ChartChordLayoutOptions>> = Object.freeze({
  startAngle: 90, clockwise: true, padAngle: 2, minAngle: 0, innerRadiusRatio: 0.58, outerRadiusRatio: 0.78,
})
export const CHART_CHORD_DEFAULT_NODE = Object.freeze({ opacity: 1, borderColor: '#FFFFFF', borderWidth: 1, borderType: 'solid' as const })
export const CHART_CHORD_DEFAULT_LABEL = Object.freeze({
  visible: true, position: 'outside' as const, rotation: 0, distance: 8, fontSize: 12, bold: false, italic: false, family: 'Arial',
})
export const CHART_CHORD_DEFAULT_RIBBON: Readonly<Required<ChartChordRibbonOptions>> = Object.freeze({ color: 'gradient', opacity: 0.65, width: 0, type: 'solid' })
export const CHART_CHORD_DEFAULT_EMPHASIS: Readonly<Required<ChartChordEmphasisOptions>> = Object.freeze({ focus: 'adjacency' })
export const CHART_CHORD_LAYOUT_KEYS = ['startAngle', 'clockwise', 'padAngle', 'minAngle', 'innerRadiusRatio', 'outerRadiusRatio'] as const

export interface ChartResolvedChordOptions {
  layout: Required<ChartChordLayoutOptions>
  node: ChartChordNodeOptions & Required<Pick<ChartChordNodeOptions, 'opacity' | 'borderColor' | 'borderWidth' | 'borderType'>>
  label: ChartChordLabelOptions & Required<Pick<ChartChordLabelOptions, 'visible' | 'position' | 'rotation' | 'distance' | 'fontSize' | 'color' | 'bold' | 'italic' | 'family'>>
  ribbon: Required<ChartChordRibbonOptions>
  emphasis: Required<ChartChordEmphasisOptions>
}

/** Fixed Mee defaults are resolved without authoring renderer values. The old
 * flat startAngle used the Canvas angle direction, so its sign is translated while
 * the new public layout uses the fixed positive-90 convention. */
export function resolveChartChordLayout(options?: ChartChordOptions): Required<ChartChordLayoutOptions> {
  const layout = options?.layout
  return {
    startAngle: layout?.startAngle ?? (options?.startAngle === undefined ? CHART_CHORD_DEFAULT_LAYOUT.startAngle : -options.startAngle),
    clockwise: layout?.clockwise ?? options?.clockwise ?? CHART_CHORD_DEFAULT_LAYOUT.clockwise,
    padAngle: layout?.padAngle ?? options?.padAngle ?? CHART_CHORD_DEFAULT_LAYOUT.padAngle,
    minAngle: layout?.minAngle ?? options?.minAngle ?? CHART_CHORD_DEFAULT_LAYOUT.minAngle,
    innerRadiusRatio: layout?.innerRadiusRatio ?? options?.innerRadiusRatio ?? CHART_CHORD_DEFAULT_LAYOUT.innerRadiusRatio,
    outerRadiusRatio: layout?.outerRadiusRatio ?? options?.outerRadiusRatio ?? CHART_CHORD_DEFAULT_LAYOUT.outerRadiusRatio,
  }
}

/** Preserve only authored leaves when migrating a partial edit. Nested values
 * win over legacy values; the old Canvas-angle sign changes exactly once.
 * Effective defaults remain renderer state and never become authored options. */
export function chartChordAuthoredLayout(options?: ChartChordOptions): ChartChordLayoutOptions {
  const layout: ChartChordLayoutOptions = { ...options?.layout }, effective = resolveChartChordLayout(options)
  for (const key of CHART_CHORD_LAYOUT_KEYS) {
    if (options?.[key] !== undefined && layout[key] === undefined) Object.assign(layout, { [key]: effective[key] })
  }
  return layout
}

export function chartChordAuthoredRibbon(options?: ChartChordOptions): ChartChordRibbonOptions {
  return { ...options?.ribbon, ...(options?.ribbon?.opacity === undefined && options?.ribbonOpacity !== undefined
    ? { opacity: options.ribbonOpacity } : {}) }
}

export function resolveChartChordOptions(options?: ChartChordOptions, font: ChartFontOptions = {}, textColor = '#334155'): ChartResolvedChordOptions {
  const node = options?.node, label = options?.label, ribbon = options?.ribbon
  return {
    layout: resolveChartChordLayout(options),
    node: { color: node?.color, opacity: node?.opacity ?? CHART_CHORD_DEFAULT_NODE.opacity,
      borderColor: node?.borderColor ?? CHART_CHORD_DEFAULT_NODE.borderColor,
      borderWidth: node?.borderWidth ?? CHART_CHORD_DEFAULT_NODE.borderWidth,
      borderType: node?.borderType ?? CHART_CHORD_DEFAULT_NODE.borderType },
    label: { visible: label?.visible ?? true, position: label?.position ?? 'outside', rotation: label?.rotation ?? 0, distance: label?.distance ?? 8,
      color: label?.color ?? font.color ?? textColor, fontSize: label?.fontSize ?? font.fontSize ?? 12,
      bold: label?.bold ?? font.bold ?? false, italic: label?.italic ?? font.italic ?? false, family: label?.family ?? font.family ?? 'Arial' },
    ribbon: { color: ribbon?.color ?? 'gradient', opacity: ribbon?.opacity ?? options?.ribbonOpacity ?? 0.65,
      width: ribbon?.width ?? 0, type: ribbon?.type ?? 'solid' },
    emphasis: { focus: options?.emphasis?.focus ?? 'adjacency' },
  }
}

/** Legacy generic label masks/formats remain readable. New Chord methods only
 * author the fixed name-label contract; this does not discard source metadata. */
export function chartChordLegacyLabel(options?: ChartChordOptions): Pick<ChartLabelOptions, 'contentType' | 'format'> {
  const label: ChartLabelOptions | undefined = options?.label
  return { ...(label?.contentType === undefined ? {} : { contentType: label.contentType }),
    ...(label?.format === undefined ? {} : { format: label.format }) }
}
