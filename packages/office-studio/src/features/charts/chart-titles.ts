import type { ChartFontOptions, ChartLegendOptions, ChartSemanticOptions, ChartSubtitleStyleOptions, ChartTextAlignment, ChartTitleManualLayout, ChartTitlePosition, ChartTitleStyleOptions } from '../../domain/model'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts'
import type { ChartUpdatePatch } from './chart-engine'
import { chartFont } from './chart-geometry'
import { cloneChartOptions, definedChartStyleFields } from './chart-options'

export const ChartTitlePositions = Object.freeze({ Top: 'top', Bottom: 'bottom', Left: 'left', Right: 'right', Hide: 'hide' } as const)
export const ChartTextAlignments = Object.freeze({ Left: 'left', Right: 'right', Center: 'center' } as const)
export type ChartTitleSpec = string | (ChartTitleStyleOptions & { text?: string })
export type ChartSubtitleSpec = string | (ChartSubtitleStyleOptions & { text?: string })
export type ChartTitleKind = 'title' | 'subtitle'
export type ChartTitleObjectSpec = ChartTitleStyleOptions & { text?: string }
export interface ChartTitlePaintLayout {
  kind: ChartTitleKind
  lineIndex: number
  anchor: { x: number; y: number }
  clip: RenderRect
  manual: boolean
  vertical: boolean
  alignment: ChartTextAlignment
}
/** Chart-local metadata stays outside the shared rendering protocol. */
export interface ChartTitleTextCommand extends TextCommand { chartTitle: ChartTitlePaintLayout }
export function isChartTitleTextCommand(command: DrawCommand): command is ChartTitleTextCommand {
  return command.kind === 'text' && 'chartTitle' in command
}
export interface ChartTitleGeometrySource {
  title?: string
  subtitle?: string
  position?: { x: number; y: number }
  size: { width: number; height: number }
  legend?: Partial<ChartLegendOptions>
  options?: ChartSemanticOptions
}

/** Strings replace text only; objects patch authored fields. Undefined clears
 * the selected title group, without changing the other title or appearance. */
export function chartTitlePatch(value: ChartTitleSpec | ChartSubtitleSpec | undefined, kind: ChartTitleKind = 'title'): Pick<ChartUpdatePatch, 'title' | 'subtitle' | 'options'> {
  const group = kind === 'title' ? 'titleStyle' : 'subtitleStyle'
  if (value === undefined) return { [kind]: undefined, options: { [group]: undefined } }
  if (typeof value === 'string') { validateTitleText(value); return { [kind]: value } }
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('标题须为文字或配置对象。')
  const { text, ...style } = value
  if (text !== undefined) validateTitleText(text)
  if (kind === 'subtitle' && ('visible' in style || 'position' in style)) throw new TypeError('副标题位置由主标题控制。')
  const options = cloneChartOptions({ [group]: style })!
  return { ...(Object.hasOwn(value, 'text') ? { [kind]: text } : {}), ...(Object.keys(style).length ? { options } : {}) }
}

function validateTitleText(value: string): void {
  if (typeof value !== 'string' || value.length > 100_000) throw new TypeError('图表标题须为不超过 100000 字符的文本。')
}

/** Existing local font defaults remain compatible with old snapshots. */
export function chartTitleFont(chart: Pick<ChartTitleGeometrySource, 'options'>, kind: ChartTitleKind): ChartFontOptions {
  return { fontSize: kind === 'title' ? 14 : 11, ...(kind === 'title' ? { bold: true } : {}),
    ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.options?.[kind === 'title' ? 'titleStyle' : 'subtitleStyle']) }
}

interface TitleTextLayout { kind: ChartTitleKind; font: ChartFontOptions; alignment: ChartTextAlignment; lines: string[]; lineHeight: number; height: number; width: number; offset: number; manualLayout?: ChartTitleManualLayout; rotate: number }
export interface ChartTitleGeometry {
  position: ChartTitlePosition
  vertical: boolean
  frame: { x: number; y: number; width: number; height: number }
  entries: TitleTextLayout[]
  bands: { top: number; bottom: number; left: number; right: number }
  startY: number
  availableHeight: number
}

/** Pure, bounded text slots keep Worker plot reservation and both Canvas paths
 * together. The model retains full text when a small viewport needs truncation. */
export function chartTitleGeometry(chart: ChartTitleGeometrySource, legendBand?: number): ChartTitleGeometry {
  const frame = { x: chart.position?.x ?? 0, y: chart.position?.y ?? 0, width: chart.size.width, height: chart.size.height }
  const position = chart.options?.titleStyle?.visible === false ? 'hide' : chart.options?.titleStyle?.position ?? 'top'
  const vertical = position === 'left' || position === 'right', bands = { top: 0, bottom: 0, left: 0, right: 0 }
  const band = legendBand ?? (chart.legend?.font?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 11) + 16
  const topLegend = vertical && chart.legend?.visible !== false && chart.legend?.position === 'top' ? band : 0
  const bottomLegend = vertical && chart.legend?.visible !== false && chart.legend?.position === 'bottom' ? band : 0
  const availableHeight = Math.max(1, frame.height - 24 - topLegend - bottomLegend), startY = frame.y + 12 + topLegend
  const entries: TitleTextLayout[] = []
  if (position === 'hide') return { frame, position, vertical, bands, entries, startY, availableHeight }
  let offset = 0
  for (const kind of ['title', 'subtitle'] as const) {
    const text = chart[kind]
    if (!text) continue
    const style = chart.options?.[kind === 'title' ? 'titleStyle' : 'subtitleStyle']
    const manualLayout = style?.manualLayout, font = chartTitleFont(chart, kind), size = font.fontSize!, lineHeight = size * 1.25
    const remaining = vertical ? availableHeight : Math.max(lineHeight, availableHeight - offset - (kind === 'title' && chart.subtitle ? chartTitleFont(chart, 'subtitle').fontSize! * 1.25 + 6 : 0))
    const lines = titleTextLines(text, vertical ? size : manualLayout ? Infinity : Math.max(1, frame.width - 24), size,
      manualLayout && !vertical ? 100_001 : Math.max(1, Math.floor(remaining / lineHeight)), vertical)
    const height = lines.length * lineHeight, width = vertical ? size : Math.min(manualLayout ? Infinity : Math.max(1, frame.width - 24), lines.reduce((maximum, line) => Math.max(maximum, titleTextWidth(line, size)), 0))
    const alignment = manualLayout && !vertical ? 'left' : style?.alignment ?? 'center'
    entries.push({ kind, font, alignment, lines, lineHeight, height, width, offset, manualLayout, rotate: style?.rotate ?? 0 })
    if (!manualLayout) offset += (vertical ? size : height) + 6
  }
  if (offset) bands[position] = Math.max(0, offset - 6) + 12
  return { frame, position, vertical, bands, entries, startY, availableHeight }
}

function titleTextWidth(text: string, size: number): number {
  let width = 0
  for (const character of text) width += size * ((character.codePointAt(0) ?? 0) > 255 ? 1 : 0.6)
  return width
}

function titleTextLines(text: string, width: number, size: number, maximum: number, vertical: boolean): string[] {
  const lines: string[] = []
  const normalized = text.replace(/\r\n?/g, '\n')
  let line = '', used = 0, consumed = 0, truncated = false
  for (const character of normalized) {
    consumed += character.length
    const advance = titleTextWidth(character, size)
    if (character === '\n' || vertical || line && used + advance > width) {
      if (vertical && character !== '\n') line = character
      lines.push(line); line = ''; used = 0
      if (lines.length >= maximum) { truncated = consumed < normalized.length || !vertical && character !== '\n'; break }
      if (character === '\n' || vertical) continue
    }
    line += character; used += advance
  }
  if (line || !lines.length) lines.push(line)
  if (truncated) {
    const last = lines.at(-1)!, end = last.charCodeAt(last.length - 1), before = last.charCodeAt(last.length - 2)
    const remove = end >= 0xdc00 && end <= 0xdfff && before >= 0xd800 && before <= 0xdbff ? 2 : 1
    lines[lines.length - 1] = vertical ? '…' : `${last.slice(0, -remove)}…`
  }
  return lines
}

export function chartTitleCommands(chart: ChartTitleGeometrySource, color: string, legendBand?: number): ChartTitleTextCommand[] {
  const geometry = chartTitleGeometry(chart, legendBand), { frame, vertical, position } = geometry
  return geometry.entries.flatMap(entry => {
    const size = entry.font.fontSize!, x = entry.manualLayout ? frame.x + entry.manualLayout.left * frame.width : vertical ? position === 'left' ? frame.x + 8 + entry.offset + size / 2
      : frame.x + frame.width - 8 - entry.offset - size / 2 : entry.alignment === 'left' ? frame.x + 12
      : entry.alignment === 'right' ? frame.x + frame.width - 12 : frame.x + frame.width / 2
    const y = entry.manualLayout ? frame.y + entry.manualLayout.top * frame.height : vertical ? geometry.startY + (entry.alignment === 'left' ? 0 : entry.alignment === 'right' ? geometry.availableHeight - entry.height : (geometry.availableHeight - entry.height) / 2)
      : position === 'bottom' ? frame.y + frame.height - 8 - entry.offset - entry.height : frame.y + 8 + entry.offset
    const radians = entry.rotate * Math.PI / 180
    return entry.lines.flatMap((text, index) => text ? [{ kind: 'text' as const, x: x - Math.sin(radians) * (index + 0.5) * entry.lineHeight,
      y: y + Math.cos(radians) * (index + 0.5) * entry.lineHeight, text,
      font: chartFont(entry.font, size), color: entry.font.color ?? color, align: vertical ? 'center' as const : entry.alignment,
      baseline: 'middle' as const, ...(entry.manualLayout && !vertical ? {} : { maxWidth: Math.max(1, vertical ? size : frame.width - 24) }),
      ...(entry.rotate ? { rotation: entry.rotate } : {}), clip: frame,
      chartTitle: { kind: entry.kind, lineIndex: index, anchor: { x, y }, clip: frame, manual: !!entry.manualLayout, vertical, alignment: entry.alignment } }] : [])
  })
}
