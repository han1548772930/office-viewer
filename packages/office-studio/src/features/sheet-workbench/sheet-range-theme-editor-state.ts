import { SHEET_RANGE_THEME_SECTIONS, type SheetCellStylePatch, type SheetCellStyleProperty, type SheetRangeThemeSection,
  type SheetRangeThemeStyle } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { normalizeSheetBorderColor, normalizeSheetBorderStyle } from './sheet-cell-border.ts'
import { SHEET_BUILTIN_RANGE_THEME_NAMES } from './sheet-range-theme-presets.ts'
import { normalizeSheetRangeThemeStyle } from './sheet-range-theme-style.ts'

export const SHEET_RANGE_THEME_SECTION_LABELS: Readonly<Record<SheetRangeThemeSection, string>> = Object.freeze({
  wholeStyle: '整个区域', firstColumnStyle: '第一列条纹', secondColumnStyle: '第二列条纹',
  firstRowStyle: '第一行条纹', secondRowStyle: '第二行条纹', headerColumnStyle: '首列', lastColumnStyle: '末列',
  headerRowStyle: '表头行', lastRowStyle: '末行',
})
interface SheetRangeThemeEditorField {
  property: SheetCellStyleProperty
  label: string
  kind: 'text' | 'number' | 'toggle' | 'color' | 'enum'
  min?: number
  max?: number
  options?: readonly { value: string; label: string }[]
}
/** The nine sections use the same sparse cell-style schema and inheritance rules. */
export const SHEET_RANGE_THEME_EDITOR_FIELDS: readonly SheetRangeThemeEditorField[] = [
  { property: 'fill', label: '填充颜色', kind: 'color' }, { property: 'font.color', label: '字体颜色', kind: 'color' },
  { property: 'font.family', label: '字体', kind: 'text' }, { property: 'font.size', label: '字号', kind: 'number', min: 1, max: 409 },
  { property: 'font.bold', label: '加粗', kind: 'toggle' }, { property: 'font.italic', label: '斜体', kind: 'toggle' },
  { property: 'font.underline', label: '下划线', kind: 'toggle' }, { property: 'font.strike', label: '删除线', kind: 'toggle' },
  { property: 'horizontal', label: '水平对齐', kind: 'enum', options: [
    { value: 'general', label: '常规' }, { value: 'left', label: '左对齐' }, { value: 'center', label: '居中' },
    { value: 'right', label: '右对齐' }, { value: 'justify', label: '两端对齐' },
  ] },
  { property: 'vertical', label: '垂直对齐', kind: 'enum', options: [
    { value: 'top', label: '顶端' }, { value: 'middle', label: '居中' }, { value: 'bottom', label: '底端' },
  ] },
  { property: 'wrapStrategy', label: '文本换行', kind: 'enum', options: [
    { value: 'overflow', label: '溢出' }, { value: 'wrap', label: '换行' }, { value: 'clip', label: '裁剪' },
  ] },
  { property: 'numberFormat', label: '数字格式', kind: 'text' },
  { property: 'rotation', label: '文字角度', kind: 'number', min: -90, max: 90 },
  { property: 'rotationVertical', label: '竖排文字', kind: 'toggle' },
  { property: 'indent', label: '缩进', kind: 'number', min: 0, max: 250 },
  { property: 'shrinkToFit', label: '缩小字体适应', kind: 'toggle' },
]
export const SHEET_RANGE_THEME_BORDER_EDGES = [
  { value: 'top', label: '上框线', source: 't' }, { value: 'right', label: '右框线', source: 'r' },
  { value: 'bottom', label: '下框线', source: 'b' }, { value: 'left', label: '左框线', source: 'l' },
  { value: 'diagonalDown', label: '右下斜线', source: 'tl_br' }, { value: 'diagonalUp', label: '右上斜线', source: 'bl_tr' },
  { value: 'topLeftBottomCenter', label: '左上至下中', source: 'tl_bc' }, { value: 'topLeftMiddleRight', label: '左上至右中', source: 'tl_mr' },
  { value: 'middleLeftTopRight', label: '左中至右上', source: 'ml_tr' }, { value: 'bottomCenterTopRight', label: '下中至右上', source: 'bc_tr' },
] as const
export type SheetRangeThemeBorderEdge = typeof SHEET_RANGE_THEME_BORDER_EDGES[number]['value']
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
function sectionStyle(theme: SheetRangeThemeStyle, section: SheetRangeThemeSection): SheetCellStylePatch {
  if (!SHEET_RANGE_THEME_SECTIONS.includes(section)) throw new TypeError('范围主题区段无效。')
  return cloneModel(theme[section] ?? {})
}
export function sheetRangeThemeEditorValue(theme: SheetRangeThemeStyle, section: SheetRangeThemeSection, property: SheetCellStyleProperty): unknown {
  const style = theme[section]
  if (!style) return undefined
  if (property.startsWith('font.')) return style.font?.[property.slice(5) as keyof NonNullable<SheetCellStylePatch['font']>]
  if (property === 'wrapStrategy') return style.wrapStrategy ?? (style.wrap === undefined ? undefined : style.wrap ? 'wrap' : 'overflow')
  if (property === 'rotation') return style.rotation === 255 ? 0 : style.rotation
  if (property === 'rotationVertical') return style.rotation === 255 ? true : style.rotationVertical
  return style[property as Exclude<keyof SheetCellStylePatch, 'font'>]
}
export function setSheetRangeThemeEditorValue(theme: SheetRangeThemeStyle, section: SheetRangeThemeSection,
  property: SheetCellStyleProperty, value: unknown): SheetRangeThemeStyle {
  if (!SHEET_RANGE_THEME_EDITOR_FIELDS.some(field => field.property === property)) throw new TypeError('范围主题编辑属性无效。')
  const style = sectionStyle(theme, section), field = SHEET_RANGE_THEME_EDITOR_FIELDS.find(item => item.property === property)!
  if (value === null && field.kind !== 'color' || value === '' && field.kind === 'text') value = undefined
  if (property.startsWith('font.')) {
    const key = property.slice(5), font = { ...style.font }
    if (value === undefined) delete font[key as keyof typeof font]; else Object.assign(font, { [key]: value })
    if (Object.keys(font).length) style.font = font; else delete style.font
  } else if (value === undefined) delete style[property as Exclude<keyof SheetCellStylePatch, 'font'>]
  else Object.assign(style, { [property]: value })
  if (property === 'wrapStrategy') {
    if (value === undefined) delete style.wrap; else style.wrap = value === 'wrap'
    if (value === 'wrap') style.shrinkToFit = false
  }
  if (property === 'shrinkToFit' && value === true) { style.wrap = false; style.wrapStrategy = 'overflow' }
  if (property === 'rotation' && value !== undefined) style.rotationVertical = false
  if (property === 'rotationVertical' && value === true) style.rotation = 0
  if (property === 'rotationVertical' && value === false && style.rotation === 255) style.rotation = 0
  // Choosing inheritance or a concrete color releases the old theme/tint source.
  const sourceColor = property === 'fill' ? 'bg' : property === 'font.color' ? 'cl' : undefined
  const sourceStyle = style.metadata?.sourceStyle
  if (sourceColor && object(sourceStyle)) delete sourceStyle[sourceColor]
  return normalizeSheetRangeThemeStyle({ ...cloneModel(theme), [section]: style })
}
export function resetSheetRangeThemeEditorSection(theme: SheetRangeThemeStyle, section: SheetRangeThemeSection): SheetRangeThemeStyle {
  sectionStyle(theme, section)
  return { ...cloneModel(theme), [section]: null }
}
export function sheetRangeThemeEditorBorder(theme: SheetRangeThemeStyle, section: SheetRangeThemeSection, edge: SheetRangeThemeBorderEdge): unknown {
  return theme[section]?.border?.[edge]
}
export function setSheetRangeThemeEditorBorder(theme: SheetRangeThemeStyle, section: SheetRangeThemeSection,
  edge: SheetRangeThemeBorderEdge, property: 'style' | 'color', value: unknown): SheetRangeThemeStyle {
  const descriptor = SHEET_RANGE_THEME_BORDER_EDGES.find(item => item.value === edge)
  if (!descriptor) throw new TypeError('范围主题边框位置无效。')
  const style = sectionStyle(theme, section), border = { ...style.border }, previous = border[edge]
  if (property === 'style' && value === undefined) delete border[edge]
  else if (property === 'style' && (value === null || value === 'none')) border[edge] = null
  else {
    const current = object(previous) ? cloneModel(previous) : {}
    if (property === 'style') current.style = normalizeSheetBorderStyle(value)
    else {
      delete current.cl
      if (value === undefined || value === null) delete current.color
      else current.color = normalizeSheetBorderColor(value)
      const source = style.metadata?.sourceStyle
      const sourceEdge = object(source) && object(source.bd) ? source.bd[descriptor.source] : undefined
      if (object(sourceEdge)) delete sourceEdge.cl
    }
    border[edge] = current
  }
  if (Object.keys(border).length) style.border = border; else delete style.border
  return normalizeSheetRangeThemeStyle({ ...cloneModel(theme), [section]: style })
}
export function uniqueSheetRangeThemeDraftName(preferred: string, names: readonly string[]): string {
  const used = new Set([...SHEET_BUILTIN_RANGE_THEME_NAMES, ...names]), base = preferred.trim() || 'Custom'
  let name = base, index = 2
  while (used.has(name)) name = `${base} ${index++}`
  return name
}
export function prepareSheetRangeThemeEditorDraft(theme: SheetRangeThemeStyle, name: string,
  names: readonly string[], editingName?: string): SheetRangeThemeStyle {
  const next = name.trim()
  if (SHEET_BUILTIN_RANGE_THEME_NAMES.includes(next)) throw new Error('内置主题名称已保留，请使用自定义名称。')
  if (names.includes(next) && next !== editingName) throw new Error('该主题名称已存在，请选择其他名称。')
  return normalizeSheetRangeThemeStyle({ ...cloneModel(theme), name: next })
}
