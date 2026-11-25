import { defaultCellStyle, type CellStyle, type FontStyle, type SheetCellStylePatch, type SheetCellStyleProperty, type SheetStyleSource,
  type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { formatSheetValue } from './sheet-number-format.ts'
import { assignSheetAxisSpan, canonicalSheetAxisSpans, sheetAxisSpanAt } from '../../domain/sheet-axis.ts'
import { mapSheetAxisIndex, sheetAxisMaximum, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { transformSheetAxisSpans } from '../../domain/sheet-axis-transform.ts'
import { normalizeSheetBorderColor, normalizeSheetBorderStyle } from './sheet-cell-border.ts'
import { resolveSheetRangeThemeStyle, type SheetRangeThemeSheet } from './sheet-range-theme-resolver.ts'

export type SheetStyleContext = Pick<WorkbookModel, 'defaultStyle' | 'styles' | 'rangeThemes' | 'tables'>
export type SheetStyleSheet = Pick<SpreadsheetSheet, 'defaultStyle' | 'rowStyles' | 'columnStyles' | 'rowStyleSpans' | 'columnStyleSpans' | 'rowStylePriority'> & SheetRangeThemeSheet
export type SheetStyleProperty = SheetCellStyleProperty
export type SheetStyleInput = SheetStyleSource | Record<string, unknown>

const fontKeys = ['family', 'size', 'bold', 'italic', 'underline', 'strike', 'color'] as const
const styleKeys = ['fill', 'horizontal', 'vertical', 'wrap', 'wrapStrategy', 'numberFormat', 'border', 'rotation',
  'rotationVertical', 'indent', 'shrinkToFit', 'metadata'] as const
const own = (value: object, key: PropertyKey) => Object.hasOwn(value, key)
const same = (first: unknown, second: unknown) => Object.is(first, second)
  || typeof first === 'object' && typeof second === 'object' && JSON.stringify(first) === JSON.stringify(second)
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)

function fixedSheetBorderData(input: unknown): Record<string, unknown> {
  if (!object(input)) throw new TypeError('边框必须是对象。')
  const result: Record<string, unknown> = {}
  const aliases: Record<string, string> = { t: 'top', r: 'right', b: 'bottom', l: 'left', tl_br: 'diagonalDown',
    tl_bc: 'topLeftBottomCenter', tl_mr: 'topLeftMiddleRight', bl_tr: 'diagonalUp', ml_tr: 'middleLeftTopRight', bc_tr: 'bottomCenterTopRight' }
  for (const [key, raw] of Object.entries(input)) {
    if (raw === undefined) continue
    const name = aliases[key] ?? key
    if (raw === null || !own(aliases, key)) { result[name] = cloneModel(raw); continue }
    if (!object(raw)) throw new TypeError('固定版边框边必须是对象或 null。')
    const style = normalizeSheetBorderStyle(raw.s)
    const color = object(raw.cl) ? raw.cl.rgb : raw.cl
    result[name] = { ...cloneModel(raw), style,
      ...(typeof color === 'string' ? { color: normalizeSheetBorderColor(color) } : {}) }
  }
  return result
}

function coherentWrapPatch(source: SheetCellStylePatch): SheetCellStylePatch {
  if (source.wrapStrategy !== undefined) return { ...source, wrap: source.wrapStrategy === 'wrap' }
  if (source.wrap !== undefined) return { ...source, wrapStrategy: source.wrap ? 'wrap' : 'overflow' }
  return source
}

/** The fixed source composes each defined property; the local font object is one group of those properties. */
export function composeSheetCellStyles(...sources: (SheetCellStylePatch | null | undefined)[]): CellStyle {
  const result = defaultCellStyle()
  for (const source of sources) {
    if (!source) continue
    const patch = coherentWrapPatch(source)
    for (const key of styleKeys) if (patch[key] !== undefined) Object.assign(result, { [key]: cloneModel(patch[key]) })
    for (const key of fontKeys) if (source.font?.[key] !== undefined) Object.assign(result.font, { [key]: source.font[key] })
  }
  return result
}

export function resolveSheetStyleSource(source: SheetStyleSource | undefined, context: SheetStyleContext = {}): SheetCellStylePatch | undefined {
  if (source == null) return undefined
  if (typeof source !== 'string') return source
  const style = context.styles && own(context.styles, source) ? context.styles[source] : undefined
  if (!style) throw new Error(`找不到单元格样式：${source}`)
  return style
}

export function sheetInheritedCellStyle(sheet: SheetStyleSheet, row: number, col: number, context: SheetStyleContext = {},
  themeStyle?: SheetStyleSource): CellStyle {
  const rowStyle = resolveSheetStyleSource(sheetAxisStyleSource(sheet, 'row', row), context)
  const columnStyle = resolveSheetStyleSource(sheetAxisStyleSource(sheet, 'column', col), context)
  return composeSheetCellStyles(resolveSheetStyleSource(context.defaultStyle, context), resolveSheetStyleSource(sheet.defaultStyle, context),
    ...(sheet.rowStylePriority === false ? [rowStyle, columnStyle] : [columnStyle, rowStyle]),
    resolveSheetRangeThemeStyle(sheet, row, col, context), resolveSheetStyleSource(themeStyle, context))
}

export function sheetAxisStyleSource(sheet: SheetStyleSheet, axis: 'row' | 'column', index: number): SheetStyleSource | undefined {
  const points = axis === 'row' ? sheet.rowStyles : sheet.columnStyles
  return points && own(points, index) ? points[index] : sheetAxisSpanAt(axis === 'row' ? sheet.rowStyleSpans : sheet.columnStyleSpans, index)?.value
}

export function hasInheritedSheetStyle(sheet: SheetStyleSheet, row: number, col: number, context: SheetStyleContext = {}, themeStyle?: SheetStyleSource): boolean {
  return context.defaultStyle != null || sheet.defaultStyle != null || sheetAxisStyleSource(sheet, 'row', row) != null
    || sheetAxisStyleSource(sheet, 'column', col) != null || themeStyle != null || resolveSheetRangeThemeStyle(sheet, row, col, context) != null
}

type SheetCellStyleData = Pick<SpreadsheetCell, 'style' | 'authoredStyle' | 'authoredStyleReference' | 'clearedStyleProperties' | 'themeStyle'>

export function sheetCellDirectStyle(cell: SheetCellStyleData | undefined, context: SheetStyleContext = {}): SheetCellStylePatch | undefined {
  if (!cell) return undefined
  if (cell.authoredStyle === undefined && !cell.authoredStyleReference) return cell.style
  const referenced = resolveSheetStyleSource(cell.authoredStyleReference, context)
  const direct: SheetCellStylePatch = referenced ? cloneModel(referenced) : {}
  for (const property of cell.clearedStyleProperties ?? []) {
    if (property.startsWith('font.')) {
      if (direct.font) delete direct.font[property.slice(5) as keyof FontStyle]
    } else delete direct[property as Exclude<keyof CellStyle, 'font'>]
  }
  if (direct.font && !Object.keys(direct.font).length) delete direct.font
  return mergeSheetCellStylePatch(direct, cell.authoredStyle ?? {})
}

export function resolveSheetCellStyle(sheet: SheetStyleSheet, row: number, col: number, cell?: SheetCellStyleData,
  context: SheetStyleContext = {}): CellStyle {
  return composeSheetCellStyles(sheetInheritedCellStyle(sheet, row, col, context, cell?.themeStyle), sheetCellDirectStyle(cell, context))
}

/** Preserve legacy full direct styles when the first property is explicitly cleared. */
export function sheetCellAuthoredStyle(cell: Pick<SpreadsheetCell, 'style' | 'authoredStyle' | 'authoredStyleReference'>): SheetCellStylePatch {
  return cloneModel(cell.authoredStyle ?? (cell.authoredStyleReference ? {} : cell.style))
}

export function setSheetCellStyleProperties(cell: SpreadsheetCell, patch: SheetCellStylePatch): void {
  patch = coherentWrapPatch(patch)
  cell.authoredStyle = mergeSheetCellStylePatch(sheetCellAuthoredStyle(cell), patch)
  cell.clearedStyleProperties = sheetStyleResetsAfterPatch(cell.clearedStyleProperties, patch)
  cell.style = composeSheetCellStyles(cell.style, patch)
}

export function sheetStyleResetsAfterPatch(properties: readonly SheetStyleProperty[] | undefined, patch: SheetCellStylePatch): SheetStyleProperty[] | undefined {
  patch = coherentWrapPatch(patch)
  const result = properties?.filter(property => property.startsWith('font.')
    ? patch.font?.[property.slice(5) as keyof FontStyle] === undefined : patch[property as Exclude<keyof CellStyle, 'font'>] === undefined)
  return result?.length ? result : undefined
}

export function mergeSheetCellStylePatch(source: SheetCellStylePatch, patch: SheetCellStylePatch): SheetCellStylePatch {
  patch = coherentWrapPatch(patch)
  const authored = cloneModel(source)
  for (const key of styleKeys) if (patch[key] !== undefined) Object.assign(authored, { [key]: cloneModel(patch[key]) })
  if (patch.font) authored.font = { ...authored.font, ...patch.font }
  return authored
}

export function clearSheetCellStyleProperties(cell: SpreadsheetCell, properties: readonly SheetStyleProperty[]): boolean {
  const authored = sheetCellAuthoredStyle(cell)
  let changed = false
  for (const property of properties) {
    if (cell.authoredStyleReference && !cell.clearedStyleProperties?.includes(property)) {
      cell.clearedStyleProperties = [...(cell.clearedStyleProperties ?? []), property]; changed = true
    }
    if (property.startsWith('font.')) {
      const key = property.slice(5) as keyof FontStyle
      if (authored.font && own(authored.font, key)) { delete authored.font[key]; changed = true }
    } else if (own(authored, property)) { delete authored[property as Exclude<keyof CellStyle, 'font'>]; changed = true }
  }
  if (!changed) return false
  if (authored.font && !Object.keys(authored.font).length) delete authored.font
  cell.authoredStyle = authored
  return true
}

/** Existing mutation services still edit the complete materialized style. Capture only their actual changes. */
export function captureSheetCellStyleEdits(cell: SpreadsheetCell, previous?: SpreadsheetCell): void {
  if (cell.authoredStyle === undefined) return
  if (cell.authoredStyleReference !== previous?.authoredStyleReference) return
  if (!previous && Object.keys(cell.authoredStyle).length) return
  const baseline = previous?.style ?? defaultCellStyle(), authored = sheetCellAuthoredStyle(cell)
  const previousAuthored: SheetCellStylePatch = previous?.authoredStyle ?? (previous ? previous.style : {})
  for (const key of styleKeys) {
    if (!!cell.clearedStyleProperties?.includes(key) !== !!previous?.clearedStyleProperties?.includes(key)) continue
    if (!same(authored[key], previousAuthored[key])) continue
    if (!same(cell.style[key], baseline[key])) {
      if (cell.style[key] === undefined) delete authored[key]
      else Object.assign(authored, { [key]: cloneModel(cell.style[key]) })
    }
  }
  for (const key of fontKeys) {
    if (!!cell.clearedStyleProperties?.includes(`font.${key}`) !== !!previous?.clearedStyleProperties?.includes(`font.${key}`)) continue
    if (!same(authored.font?.[key], previousAuthored.font?.[key])) continue
    if (!same(cell.style.font[key], baseline.font[key])) {
      authored.font ??= {}; Object.assign(authored.font, { [key]: cell.style.font[key] })
    }
  }
  cell.authoredStyle = authored
}

/** Adopt imperative nondefault edits on freshly constructed cells while keeping saved inherited caches sparse. */
export function adoptSheetCellStyleEdits(sheet: SheetStyleSheet, cell: SpreadsheetCell, context: SheetStyleContext = {}): void {
  if (cell.authoredStyle === undefined || cell.styleMaterialized || cell.authoredStyleReference || cell.clearedStyleProperties?.length) return
  const composed = resolveSheetCellStyle(sheet, cell.row, cell.col, cell, context), defaults = defaultCellStyle()
  const authored = sheetCellAuthoredStyle(cell)
  for (const key of styleKeys) if (authored[key] === undefined && cell.style[key] !== undefined
    && !same(cell.style[key], composed[key]) && !same(cell.style[key], defaults[key])) Object.assign(authored, { [key]: cloneModel(cell.style[key]) })
  for (const key of fontKeys) if (authored.font?.[key] === undefined && !same(cell.style.font[key], composed.font[key])
    && !same(cell.style.font[key], defaults.font[key])) {
    authored.font ??= {}; Object.assign(authored.font, { [key]: cell.style.font[key] })
  }
  cell.authoredStyle = authored
}

export function initializeWorkbookCellStyles(workbook: WorkbookModel): void {
  if (workbook.defaultStyle !== undefined) workbook.defaultStyle = normalizeSheetStyleInput(workbook.defaultStyle)
  if (workbook.styles) for (const [id, input] of Object.entries(workbook.styles)) {
    const normalized = normalizeSheetStyleInput(input)
    if (typeof normalized === 'string' || normalized === null) throw new TypeError('样式目录必须包含样式对象。')
    workbook.styles[id] = normalized
  }
  resolveSheetStyleSource(workbook.defaultStyle, workbook)
  for (const sheet of workbook.sheets) {
    if (sheet.defaultStyle !== undefined) sheet.defaultStyle = normalizeSheetStyleInput(sheet.defaultStyle)
    resolveSheetStyleSource(sheet.defaultStyle, workbook)
    if (sheet.rowStylePriority !== undefined && typeof sheet.rowStylePriority !== 'boolean') throw new TypeError('行样式优先级必须是布尔值。')
    for (const key of ['rowStyles', 'columnStyles'] as const) if (sheet[key]) {
      for (const [index, input] of Object.entries(sheet[key]!)) {
        const position = Number(index), limit = key === 'rowStyles' ? 1_048_576 : 16_384
        if (!Number.isSafeInteger(position) || position < 0 || position >= limit) throw new RangeError('默认样式的行列索引越界。')
        const normalized = normalizeSheetStyleInput(input)
        resolveSheetStyleSource(normalized, workbook); sheet[key]![position] = normalized
      }
    }
    for (const key of ['rowStyleSpans', 'columnStyleSpans'] as const) if (sheet[key]) {
      const limit = key === 'rowStyleSpans' ? 1_048_576 : 16_384
      sheet[key] = canonicalSheetAxisSpans(sheet[key]!.map(span => {
        if (span.end >= limit) throw new RangeError('默认样式的行列范围越界。')
        const value = normalizeSheetStyleInput(span.value); resolveSheetStyleSource(value, workbook)
        return { ...span, value }
      }))
    }
    for (const cell of Object.values(sheet.cells)) {
      if (cell.styleMaterialized !== undefined && cell.styleMaterialized !== true) throw new TypeError('单元格样式缓存标记无效。')
      if (cell.authoredStyleReference !== undefined && (typeof cell.authoredStyleReference !== 'string' || !cell.authoredStyleReference.trim())) throw new TypeError('单元格样式引用无效。')
      resolveSheetStyleSource(cell.authoredStyleReference, workbook)
      if (cell.clearedStyleProperties !== undefined && (!Array.isArray(cell.clearedStyleProperties)
        || cell.clearedStyleProperties.some(property => !styleKeys.includes(property as typeof styleKeys[number])
          && !fontKeys.some(key => property === `font.${key}`)))) throw new TypeError('单元格样式清除属性无效。')
      adoptSheetCellStyleEdits(sheet, cell, workbook)
    }
  }
  refreshWorkbookCellStyles(workbook, undefined, false)
}

/** Axis styles follow insert/delete/move/copy just like dimensions; inserted repeats use one compact span. */
export function transformSheetDefaultStyles(sheet: SpreadsheetSheet, edit: SheetAxisEdit, inheritedIndex?: number): void {
  const key = edit.axis === 'row' ? 'rowStyles' : 'columnStyles'
  const spanKey = edit.axis === 'row' ? 'rowStyleSpans' : 'columnStyleSpans'
  const sources = sheet[key], spans = sheet[spanKey]
  const inherited = inheritedIndex === undefined ? undefined : sheetAxisStyleSource(sheet, edit.axis, inheritedIndex)
  const result: Record<number, SheetStyleSource> = {}
  for (const [key, source] of Object.entries(sources ?? {})) {
    const index = Number(key), moved = mapSheetAxisIndex(index, edit)
    if (moved !== undefined && moved < sheetAxisMaximum(edit.axis)) result[moved] = cloneModel(source)
    if (edit.kind === 'copy' && index >= edit.start && index < edit.start + edit.count) result[edit.destination + index - edit.start] = cloneModel(source)
  }
  sheet[key] = Object.keys(result).length ? result : undefined
  const transformed = transformSheetAxisSpans(spans, edit)
  sheet[spanKey] = edit.kind === 'insert' && inherited !== undefined
    ? assignSheetAxisSpan(transformed, edit.start, edit.start + edit.count - 1, cloneModel(inherited)) : transformed
}

/** One materialization path serves commands, snapshots, formula/layout readers and Office writer DTOs. */
export function refreshWorkbookCellStyles(workbook: WorkbookModel, previous?: WorkbookModel, captureEdits = true, numberFormatLocale = 'zh-CN'): void {
  for (const sheet of workbook.sheets) {
    const original = previous?.sheets.find(item => item.id === sheet.id)
    for (const cell of Object.values(sheet.cells)) {
      if (captureEdits) captureSheetCellStyleEdits(cell, original?.cells[cell.id])
      const oldPattern = cell.style.numberFormat
      cell.style = resolveSheetCellStyle(sheet, cell.row, cell.col, cell, workbook)
      cell.styleMaterialized = true
      if (oldPattern !== cell.style.numberFormat && (!cell.formula || cell.value !== null)) {
        cell.display = formatSheetValue(cell.value, cell.style.numberFormat, workbook.date1904, numberFormatLocale) ?? cell.display
      }
    }
  }
}

/** Accept the project's typed style and the fixed Univer short-property public DTO. */
export function normalizeSheetStyleInput(input: SheetStyleInput): SheetStyleSource {
  if (input === null) return null
  if (typeof input === 'string') {
    if (!input.trim()) throw new TypeError('样式 ID 不能为空。')
    return input
  }
  if (!object(input)) throw new TypeError('默认样式必须是样式对象、样式 ID 或 null。')
  const result: SheetCellStylePatch = {}, font: Partial<FontStyle> = {}
  if (input.font !== undefined && !object(input.font)) throw new TypeError('字体样式必须是对象。')
  const rawFont = object(input.font) ? input.font : {}
  for (const key of fontKeys) if (rawFont[key] !== undefined) Object.assign(font, { [key]: rawFont[key] })
  const aliases: Record<string, keyof FontStyle> = { ff: 'family', fs: 'size', bl: 'bold', it: 'italic', ul: 'underline', st: 'strike' }
  for (const [key, field] of Object.entries(aliases)) if (input[key] !== undefined && input[key] !== null) {
    const raw = input[key]
    const value = (field === 'underline' || field === 'strike') && object(raw) ? raw.s : raw
    Object.assign(font, { [field]: ['bold', 'italic', 'underline', 'strike'].includes(field) ? value === 1 ? true : value === 0 ? false : value : value })
  }
  const color = (value: unknown) => object(value) ? value.rgb : value
  if (input.cl !== undefined) font.color = color(input.cl) as string | null
  if (font.family !== undefined && (typeof font.family !== 'string' || !font.family.trim())) throw new TypeError('字体名称不能为空。')
  if (font.size !== undefined && (typeof font.size !== 'number' || !Number.isFinite(font.size) || font.size <= 0)) throw new RangeError('字号必须是有限正数。')
  for (const key of ['bold', 'italic', 'underline', 'strike'] as const) if (font[key] !== undefined && typeof font[key] !== 'boolean') throw new TypeError('字体开关必须是布尔值或 0/1。')
  if (font.color !== undefined && font.color !== null && typeof font.color !== 'string') throw new TypeError('字体颜色必须是字符串或 null。')
  if (Object.keys(font).length) result.font = font
  for (const key of styleKeys) if (input[key] !== undefined) Object.assign(result, { [key]: cloneModel(input[key]) })
  if (input.bg !== undefined) result.fill = color(input.bg) as string | null
  const enumIndex = (value: unknown, maximum: number) => {
    if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value > maximum) throw new TypeError('样式枚举值无效。')
    return value
  }
  if (input.ht !== undefined && input.ht !== null) result.horizontal = (['general', 'left', 'center', 'right', 'justify', 'justify', 'justify'] as const)[enumIndex(input.ht, 6)]
  if (input.vt !== undefined && input.vt !== null) result.vertical = (['middle', 'top', 'middle', 'bottom'] as const)[enumIndex(input.vt, 3)]
  if (input.tb !== undefined && input.tb !== null) {
    result.wrapStrategy = (['overflow', 'overflow', 'clip', 'wrap'] as const)[enumIndex(input.tb, 3)]
    result.wrap = result.wrapStrategy === 'wrap'
  }
  if (object(input.n) && input.n.pattern !== undefined) result.numberFormat = input.n.pattern as string
  if (input.bd !== undefined && input.bd !== null) result.border = fixedSheetBorderData(input.bd)
  if (input.stf !== undefined && input.stf !== null) {
    if (input.stf !== 0 && input.stf !== 1 && typeof input.stf !== 'boolean') throw new TypeError('缩小适应开关无效。')
    result.shrinkToFit = input.stf === 1 || input.stf === true
  }
  if (object(input.tr)) {
    if (input.tr.a !== undefined) result.rotation = input.tr.a as number
    if (input.tr.v !== undefined) result.rotationVertical = input.tr.v === 1 || input.tr.v === true
  }
  if (result.fill !== undefined && result.fill !== null && typeof result.fill !== 'string') throw new TypeError('填充颜色必须是字符串或 null。')
  if (result.horizontal !== undefined && !['general', 'left', 'center', 'right', 'justify'].includes(result.horizontal)) throw new TypeError('水平对齐无效。')
  if (result.vertical !== undefined && !['top', 'middle', 'bottom'].includes(result.vertical)) throw new TypeError('垂直对齐无效。')
  if (result.wrapStrategy !== undefined && !['overflow', 'wrap', 'clip'].includes(result.wrapStrategy)) throw new TypeError('换行策略无效。')
  for (const key of ['wrap', 'rotationVertical', 'shrinkToFit'] as const) if (result[key] !== undefined && typeof result[key] !== 'boolean') throw new TypeError('布局开关必须是布尔值。')
  for (const key of ['rotation', 'indent'] as const) if (result[key] !== undefined && (typeof result[key] !== 'number' || !Number.isFinite(result[key]))) throw new TypeError('布局数值必须有限。')
  if (result.numberFormat !== undefined && typeof result.numberFormat !== 'string') throw new TypeError('数字格式必须是字符串。')
  if (result.border !== undefined && !object(result.border)) throw new TypeError('边框必须是对象。')
  if (result.metadata !== undefined && !object(result.metadata)) throw new TypeError('样式扩展必须是对象。')
  const known = new Set<string>(['font', ...styleKeys, ...Object.keys(aliases), 'cl', 'bg', 'ht', 'vt', 'tb', 'n', 'bd', 'tr', 'stf'])
  const extensions = Object.fromEntries(Object.entries(input).filter(([key]) => !known.has(key)))
  const fontExtensions = Object.fromEntries(Object.entries(rawFont).filter(([key]) => !fontKeys.includes(key as typeof fontKeys[number])))
  for (const key of ['ht', 'vt', 'tb']) if (input[key] !== undefined && input[key] !== null) extensions[key] = cloneModel(input[key])
  if (Object.keys(fontExtensions).length) extensions.font = cloneModel(fontExtensions)
  for (const key of ['ul', 'st', 'cl', 'bg', 'bd', 'tr', 'n']) if (object(input[key])) extensions[key] = cloneModel(input[key])
  if (Object.keys(extensions).length) result.metadata = { ...result.metadata, sourceStyle: cloneModel(extensions) }
  return result
}
