import type { SlicerType } from '../../domain/model.ts'
import { SlicerStyleType, TimelineStyleType, type ISlicerItemStyleState, type SlicerStyleBorderSide,
  type SlicerStyleColor, type SlicerStyleFormatting, type SlicerStyleTextDecoration } from '../../domain/slicer-styles.ts'
import { getSlicerStyleType, projectSlicerStyleFormatting, projectTimelineStyleFormatting,
  type SlicerStyleDescriptor, type SlicerStyleOwner } from './slicer-style-catalog.ts'

export interface SlicerCanvasTheme {
  background: string
  text: string
  secondaryFill: string
  mutedText: string
  /** DARK1, LIGHT1, DARK2, LIGHT2, ACCENT1–6, HYPERLINK, FOLLOWED_HYPERLINK. */
  themeColors?: readonly string[]
}
export interface SlicerStylePalette {
  background: string; border: string; header: string; headerBorder: string; headerText: string
  item: string; itemBorder: string; selected: string; selectedText: string; text: string
  selectionText: string; scrollTrack: string; scrollThumb: string; periodAlternate: string; periodText: string
}
export interface SlicerRenderBorder { style: string; color: string }
export interface SlicerRenderRegion {
  fill: string
  color: string
  bold: boolean
  horizontal: number
  vertical: number
  borders: Partial<Record<SlicerStyleBorderSide, SlicerRenderBorder | null>>
  overline?: SlicerStyleTextDecoration & { color: string }
}
interface SlicerRenderProfileBase { reference: SlicerStyleDescriptor; palette: SlicerStylePalette; whole: SlicerRenderRegion; header: SlicerRenderRegion }
export type SlicerRenderProfile = SlicerRenderProfileBase & ({ type: 'slicer'; items: Record<SlicerStyleType, SlicerRenderRegion> }
  | { type: 'timeline'; elements: Record<TimelineStyleType, SlicerRenderRegion> })

export const SLICER_BORDER_STYLE_NAMES = ['none', 'thin', 'hair', 'dotted', 'dashed', 'dash-dot', 'dash-dot-dot', 'double',
  'medium', 'medium-dashed', 'medium-dash-dot', 'medium-dash-dot-dot', 'slant-dash-dot', 'thick'] as const
const accents = ['#3F83F8', '#FF5A1F', '#4E5565', '#F1B312', '#0694A2', '#0DA471']
export function mixSlicerColor(foreground: string, background: string, ratio: number): string {
  if (!/^#[a-f\d]{6}$/i.test(foreground) || !/^#[a-f\d]{6}$/i.test(background)) return ratio < .5 ? foreground : background
  const first = foreground.slice(1).match(/.{2}/g)!.map(part => parseInt(part, 16)), second = background.slice(1).match(/.{2}/g)!.map(part => parseInt(part, 16))
  return `#${first.map((value, index) => Math.round(value * (1 - ratio) + second[index]! * ratio).toString(16).padStart(2, '0')).join('')}`
}
function contrast(color: string, against: string): number {
  const luminance = (hex: string) => hex.slice(1).match(/.{2}/g)!.map(part => parseInt(part, 16) / 255)
    .map(value => value <= .04045 ? value / 12.92 : ((value + .055) / 1.055) ** 2.4)
    .reduce((sum, value, index) => sum + value * [0.2126, 0.7152, 0.0722][index]!, 0)
  const values = [luminance(color), luminance(against)].sort((first, second) => second - first)
  return (values[0]! + .05) / (values[1]! + .05)
}
function basePalette(name: string, timeline: boolean, theme?: SlicerCanvasTheme): SlicerStylePalette {
  const match = /Style(Light|Dark)([1-6])$/.exec(name), accent = name.endsWith('Other1') ? '#4E5565' : name.endsWith('Other2') ? '#1F2124'
    : accents[Number(match?.[2] ?? 1) - 1] ?? accents[0]!, dark = match?.[1] === 'Dark' || name.endsWith('Other2')
  const item = mixSlicerColor(accent, '#FFFFFF', dark ? .76 : .9), header = dark ? mixSlicerColor(accent, '#1B1C1F', .24) : theme?.background ?? '#FFFFFF'
  return { background: theme?.background ?? '#FFFFFF', text: theme?.text ?? '#4E5565',
    border: dark ? mixSlicerColor(accent, '#1B1C1F', .5) : accent, header,
    headerBorder: dark ? header : accent, headerText: dark ? contrast(header, '#FFFFFF') >= contrast(header, '#1B1C1F') ? '#FFFFFF' : '#1B1C1F' : theme?.text ?? '#1B1C1F',
    item: theme?.secondaryFill ?? (timeline ? mixSlicerColor(item, accent, .12) : item), itemBorder: mixSlicerColor(accent, '#1B1C1F', .5), selected: accent,
    selectedText: contrast(accent, '#FFFFFF') >= contrast(accent, '#1B1C1F') ? '#FFFFFF' : '#1B1C1F',
    selectionText: theme?.text ?? mixSlicerColor(accent, '#1B1C1F', .42), scrollTrack: theme?.background ?? '#FFFFFF', scrollThumb: theme?.mutedText ?? '#A8B0BD',
    periodAlternate: theme?.secondaryFill ?? mixSlicerColor(accent, '#FFFFFF', .94), periodText: theme?.text ?? '#4E5565' }
}
export function resolveSlicerStyleColor(value: SlicerStyleColor | null | undefined, fallback: string, theme?: SlicerCanvasTheme): string {
  if (!value) return fallback
  if (typeof value.rgb === 'string' && value.rgb.trim()) {
    const rgb = value.rgb.trim()
    if (/^#?[a-f\d]{8}$/i.test(rgb)) { const hex = rgb.replace(/^#/, ''); return `#${hex.slice(2)}${hex.slice(0, 2)}` }
    return /^[a-f\d]{6}$/i.test(rgb) ? `#${rgb}` : rgb
  }
  if (value.th !== undefined) {
    const colors = theme?.themeColors ?? [theme?.text ?? '#1B1C1F', theme?.background ?? '#FFFFFF', '#4E5565', theme?.secondaryFill ?? '#E8EBF1',
      ...accents, '#0563C1', '#954F72']
    return colors[value.th] ?? fallback
  }
  return fallback
}
const edges = (color: string): SlicerRenderRegion['borders'] => Object.fromEntries(['t', 'r', 'b', 'l'].map(side => [side, { style: 'thin', color }]))
function region(fill: string, color: string, border?: string, bold = false, horizontal = 1): SlicerRenderRegion {
  return { fill, color, bold, horizontal, vertical: 2, borders: border ? edges(border) : {} }
}
/** Only typography inherits from the whole control; each region owns its own fill and border. */
function inheritTypography(whole: SlicerStyleFormatting, local: SlicerStyleFormatting): SlicerStyleFormatting {
  return { ...Object.fromEntries(['cl', 'ol', 'ht', 'vt', 'bl'].filter(key => Object.hasOwn(whole, key)).map(key => [key, whole[key as keyof SlicerStyleFormatting]])), ...local }
}
function applyFormatting(base: SlicerRenderRegion, formatting: SlicerStyleFormatting, theme?: SlicerCanvasTheme): SlicerRenderRegion {
  const next = { ...base, borders: { ...base.borders },
    fill: resolveSlicerStyleColor(formatting.bg, base.fill, theme), color: resolveSlicerStyleColor(formatting.cl, base.color, theme),
    bold: formatting.bl === undefined ? base.bold : formatting.bl === 1,
    horizontal: formatting.ht ?? base.horizontal, vertical: formatting.vt ?? base.vertical }
  if (formatting.bd === null) next.borders = {}
  else if (formatting.bd) for (const [side, border] of Object.entries(formatting.bd)) next.borders[side as SlicerStyleBorderSide] = border === null ? null
    : { style: SLICER_BORDER_STYLE_NAMES[border.s] ?? 'none', color: resolveSlicerStyleColor(border.cl, base.color, theme) }
  if (formatting.ol !== undefined) next.overline = { ...formatting.ol,
    color: formatting.ol.c === 0 ? resolveSlicerStyleColor(formatting.ol.cl, next.color, theme) : next.color }
  return next
}
export function resolveSlicerRenderProfile(workbook: SlicerStyleOwner, type: SlicerType, name?: string, theme?: SlicerCanvasTheme): SlicerRenderProfile {
  const projection = type === 'slicer' ? projectSlicerStyleFormatting(workbook, name) : projectTimelineStyleFormatting(workbook, name)
  const palette = basePalette(projection.reference.name, type === 'timeline', theme), whole = applyFormatting(region(palette.background, palette.text, palette.border), projection.wholeStyle, theme)
  const header = applyFormatting(region(palette.header, palette.headerText, palette.headerBorder, true), inheritTypography(projection.wholeStyle, projection.headerRowStyle), theme)
  Object.assign(palette, { background: whole.fill, text: whole.color, border: whole.borders.l?.color ?? 'transparent', header: header.fill,
    headerText: header.color, headerBorder: header.borders.b?.color ?? 'transparent' })
  if ('items' in projection) {
    const items = Object.fromEntries(Object.values(SlicerStyleType).map(style => {
      const selected = style.includes('Selected') || style.startsWith('selected'), hasData = !style.endsWith('NoData') && !style.endsWith('noData'), hovered = style.startsWith('hovered')
      const fill = selected ? palette.selected : palette.item, text = selected ? palette.selectedText : palette.text
      const base = region(hovered ? mixSlicerColor(fill, selected ? '#FFFFFF' : palette.selected, .16) : fill,
        hasData ? text : theme?.mutedText ?? mixSlicerColor(text, palette.background, .45), palette.itemBorder)
      return [style, applyFormatting(base, inheritTypography(projection.wholeStyle, projection.items[style]), theme)]
    })) as Record<SlicerStyleType, SlicerRenderRegion>
    Object.assign(palette, { item: items.unselectedItemWithData.fill, text: items.unselectedItemWithData.color,
      itemBorder: items.unselectedItemWithData.borders.l?.color ?? 'transparent', selected: items.selectedItemWithData.fill, selectedText: items.selectedItemWithData.color })
    return { type: 'slicer', reference: projection.reference, palette, whole, header, items }
  }
  const defaults: Record<TimelineStyleType, SlicerRenderRegion> = {
    selectionLabel: region(palette.background, palette.selectionText, undefined, true), timeLevel: region(palette.background, palette.text),
    periodLabel1: region(palette.background, palette.periodText, undefined, false, 2), periodLabel2: region(palette.periodAlternate, palette.periodText, undefined, false, 2),
    selectedTimeBlock: region(palette.selected, palette.selectedText), unselectedTimeBlock: region(palette.item, palette.text), selectedTimeBlockSpace: region(palette.background, palette.text),
  }
  const elements = Object.fromEntries(Object.values(TimelineStyleType).map(style => [style,
    applyFormatting(defaults[style], inheritTypography(projection.wholeStyle, projection.elements[style]), theme)])) as Record<TimelineStyleType, SlicerRenderRegion>
  Object.assign(palette, { selectionText: elements.selectionLabel.color, periodText: elements.periodLabel1.color, periodAlternate: elements.periodLabel2.fill,
    selected: elements.selectedTimeBlock.fill, selectedText: elements.selectedTimeBlock.color, item: elements.unselectedTimeBlock.fill })
  return { type: 'timeline', reference: projection.reference, palette, whole, header, elements }
}
export function slicerPalette(style: string | undefined, timeline: boolean, theme?: SlicerCanvasTheme, workbook: SlicerStyleOwner = {}): SlicerStylePalette {
  return resolveSlicerRenderProfile(workbook, timeline ? 'timeline' : 'slicer', style, theme).palette
}
export function slicerItemRenderRegion(profile: SlicerRenderProfile, state: ISlicerItemStyleState): SlicerRenderRegion {
  return profile.type === 'slicer' ? profile.items[getSlicerStyleType(state)] : profile.elements[state.selected ? TimelineStyleType.SelectedTimeBlock : TimelineStyleType.UnselectedTimeBlock]
}
export function slicerRenderRegionCss(style: SlicerRenderRegion): Record<string, string | number> {
  const css: Record<string, string | number> = { background: style.fill, color: style.color, fontWeight: style.bold ? 700 : 400,
    textAlign: style.horizontal === 3 ? 'right' : style.horizontal === 2 ? 'center' : style.horizontal >= 4 ? 'justify' : 'left',
    justifyContent: style.horizontal === 3 ? 'flex-end' : style.horizontal === 2 ? 'center' : style.horizontal >= 4 ? 'space-between' : 'flex-start',
    alignItems: style.vertical === 1 ? 'flex-start' : style.vertical === 3 ? 'flex-end' : 'center' }
  for (const [side, name] of [['t', 'Top'], ['r', 'Right'], ['b', 'Bottom'], ['l', 'Left']] as const) {
    const border = style.borders[side]
    css[`border${name}`] = !border || border.style === 'none' ? '0 solid transparent'
      : `${border.style === 'thick' ? 3 : border.style.startsWith('medium') ? 2 : border.style === 'hair' ? .5 : 1}px ${border.style === 'double' ? 'double' : border.style.includes('dotted') ? 'dotted' : border.style.includes('dash') ? 'dashed' : 'solid'} ${border.color}`
  }
  const overline = style.overline
  css.textDecorationLine = overline?.s === 1 && overline.t !== 11 ? 'overline' : 'none'
  if (overline?.s === 1) {
    css.textDecorationColor = overline.color
    css.textDecorationStyle = overline.t === 10 || overline.t === 19 ? 'double' : [14, 15, 16].includes(overline.t ?? 12) ? 'wavy'
      : [8, 9].includes(overline.t ?? 12) ? 'dotted' : (overline.t ?? 12) <= 7 ? 'dashed' : 'solid'
    css.textDecorationThickness = [1, 2, 3, 5, 9, 13, 16].includes(overline.t ?? 12) ? '2px' : '1px'
  }
  return css
}
