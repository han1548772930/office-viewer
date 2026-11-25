import { SHEET_RANGE_THEME_SECTIONS, type SheetRangeThemeSection, type SheetRangeThemeStyle } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { normalizeSheetStyleInput, type SheetStyleInput } from './sheet-style-inheritance.ts'
import { sheetStyleToPublicData, type SheetPublicStyleData } from './sheet-public-style.ts'
import { sheetRangeThemeStyleAt } from './sheet-range-theme-resolver.ts'

export type SheetRangeThemeStyleJson = { name: string; metadata?: Record<string, unknown> } & Partial<Record<SheetRangeThemeSection, SheetStyleInput>>
export function assertSheetRangeThemeName(name: unknown): asserts name is string {
  if (typeof name !== 'string' || !name.trim()) throw new TypeError('范围主题名称不能为空。')
}
export function normalizeSheetRangeThemeStyle(input: unknown): SheetRangeThemeStyle {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('范围主题配置必须是对象。')
  const source = input as Record<string, unknown>
  assertSheetRangeThemeName(source.name)
  const result: SheetRangeThemeStyle = { name: source.name }
  for (const section of SHEET_RANGE_THEME_SECTIONS) if (source[section] !== undefined) {
    const style = normalizeSheetStyleInput(source[section] as SheetStyleInput)
    if (typeof style === 'string') throw new TypeError('范围主题区段必须是样式对象或 null。')
    result[section] = style
  }
  if (source.metadata !== undefined && (!source.metadata || typeof source.metadata !== 'object' || Array.isArray(source.metadata))) throw new TypeError('范围主题扩展必须是对象。')
  const extensions = Object.fromEntries(Object.entries(source).filter(([key]) => key !== 'name' && key !== 'metadata' && !SHEET_RANGE_THEME_SECTIONS.some(section => section === key)))
  if (source.metadata || Object.keys(extensions).length) result.metadata = {
    ...(source.metadata ? cloneModel(source.metadata as Record<string, unknown>) : {}), ...cloneModel(extensions),
  }
  return result
}
export function sheetRangeThemeToJson(style: SheetRangeThemeStyle): SheetRangeThemeStyleJson {
  const result: SheetRangeThemeStyleJson = { name: style.name, ...(style.metadata ? { metadata: cloneModel(style.metadata) } : {}) }
  for (const section of SHEET_RANGE_THEME_SECTIONS) if (style[section] !== undefined) result[section] = style[section] === null ? null : sheetStyleToPublicData(style[section]!)
  return result
}

/** Public builder is detached; editing it affects a workbook only after registering its serialized author configuration. */
export class RangeThemeStyle {
  private data: SheetRangeThemeStyle
  constructor(name: string, options: Omit<SheetRangeThemeStyleJson, 'name'> = {}) {
    if (!options || typeof options !== 'object' || Array.isArray(options)) throw new TypeError('范围主题配置必须是对象。')
    this.data = normalizeSheetRangeThemeStyle({ ...options, name })
  }
  getName(): string { return this.data.name }
  toJson(): SheetRangeThemeStyleJson { return sheetRangeThemeToJson(this.data) }
  fromJson(json: SheetRangeThemeStyleJson): void {
    if (!json || typeof json !== 'object' || Array.isArray(json)) throw new TypeError('范围主题配置必须是对象。')
    this.data = normalizeSheetRangeThemeStyle({ ...this.data, ...json })
  }
  getStyle(offsetRow: number, offsetCol: number, isLastRow: boolean, isLastCol: boolean, isToggled = false): SheetPublicStyleData | null {
    if (!Number.isSafeInteger(offsetRow) || !Number.isSafeInteger(offsetCol)) throw new RangeError('范围主题偏移必须是整数。')
    if (![isLastRow, isLastCol, isToggled].every(value => typeof value === 'boolean')) throw new TypeError('范围主题首尾和条纹状态必须是布尔值。')
    const hasSection = isLastRow || isLastCol || offsetRow >= 0 && offsetCol >= 0
      || offsetRow % 2 === 0 || offsetRow % 2 === 1 || offsetCol % 2 === 0 || offsetCol % 2 === 1
    if (!hasSection) return null
    const style = sheetRangeThemeStyleAt(this.data, offsetRow, offsetCol, isLastRow, isLastCol, isToggled)
    return sheetStyleToPublicData(style)
  }
  // This builder owns no browser handles, subscriptions or retained buffers.
  dispose(): void { }
  private get(section: SheetRangeThemeSection): SheetPublicStyleData | null {
    return this.data[section] ? sheetStyleToPublicData(this.data[section]!) : null
  }
  private set(section: SheetRangeThemeSection, style: SheetStyleInput): void {
    const patch = normalizeSheetStyleInput(style)
    if (typeof patch === 'string') throw new TypeError('范围主题区段必须是样式对象或 null。')
    this.data[section] = patch
  }
  get wholeStyle(): SheetPublicStyleData | null { return this.get('wholeStyle') }
  set wholeStyle(value: SheetStyleInput) { this.set('wholeStyle', value) }
  get firstRowStyle(): SheetPublicStyleData | null { return this.get('firstRowStyle') }
  set firstRowStyle(value: SheetStyleInput) { this.set('firstRowStyle', value) }
  get secondRowStyle(): SheetPublicStyleData | null { return this.get('secondRowStyle') }
  set secondRowStyle(value: SheetStyleInput) { this.set('secondRowStyle', value) }
  get lastRowStyle(): SheetPublicStyleData | null { return this.get('lastRowStyle') }
  set lastRowStyle(value: SheetStyleInput) { this.set('lastRowStyle', value) }
  get firstColumnStyle(): SheetPublicStyleData | null { return this.get('firstColumnStyle') }
  set firstColumnStyle(value: SheetStyleInput) { this.set('firstColumnStyle', value) }
  get secondColumnStyle(): SheetPublicStyleData | null { return this.get('secondColumnStyle') }
  set secondColumnStyle(value: SheetStyleInput) { this.set('secondColumnStyle', value) }
  get lastColumnStyle(): SheetPublicStyleData | null { return this.get('lastColumnStyle') }
  set lastColumnStyle(value: SheetStyleInput) { this.set('lastColumnStyle', value) }
  get headerRowStyle(): SheetPublicStyleData | null { return this.get('headerRowStyle') }
  set headerRowStyle(value: SheetStyleInput) { this.set('headerRowStyle', value) }
  get headerColumnStyle(): SheetPublicStyleData | null { return this.get('headerColumnStyle') }
  set headerColumnStyle(value: SheetStyleInput) { this.set('headerColumnStyle', value) }
  getWholeStyle(): SheetPublicStyleData | null { return this.wholeStyle }
  setWholeStyle(value: SheetStyleInput): void { this.wholeStyle = value }
  getFirstRowStyle(): SheetPublicStyleData | null { return this.firstRowStyle }
  setFirstRowStyle(value: SheetStyleInput): void { this.firstRowStyle = value }
  getSecondRowStyle(): SheetPublicStyleData | null { return this.secondRowStyle }
  setSecondRowStyle(value: SheetStyleInput): void { this.secondRowStyle = value }
  getLastRowStyle(): SheetPublicStyleData | null { return this.lastRowStyle }
  setLastRowStyle(value: SheetStyleInput): void { this.lastRowStyle = value }
  getFirstColumnStyle(): SheetPublicStyleData | null { return this.firstColumnStyle }
  setFirstColumnStyle(value: SheetStyleInput): void { this.firstColumnStyle = value }
  getSecondColumnStyle(): SheetPublicStyleData | null { return this.secondColumnStyle }
  setSecondColumnStyle(value: SheetStyleInput): void { this.secondColumnStyle = value }
  getLastColumnStyle(): SheetPublicStyleData | null { return this.lastColumnStyle }
  setLastColumnStyle(value: SheetStyleInput): void { this.lastColumnStyle = value }
  getHeaderRowStyle(): SheetPublicStyleData | null { return this.headerRowStyle }
  setHeaderRowStyle(value: SheetStyleInput): void { this.headerRowStyle = value }
  getHeaderColumnStyle(): SheetPublicStyleData | null { return this.headerColumnStyle }
  setHeaderColumnStyle(value: SheetStyleInput): void { this.headerColumnStyle = value }
}
