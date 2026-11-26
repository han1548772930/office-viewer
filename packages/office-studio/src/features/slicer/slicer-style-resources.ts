import type { SlicerModel, SlicerType, WorkbookModel } from '../../domain/model.ts'
import type { ISlicerStyleDefinition, ISlicerStylesDefinition, ITimelineStyleDefinition,
  ITimelineStylesDefinition, SlicerStyleBorder, SlicerStyleBorders, SlicerStyleColor,
  SlicerStyleFormatting, SlicerStyleTextDecoration } from '../../domain/slicer-styles.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { getDefaultSlicerStyleName, getSlicerStyle, hasSlicerStyle, SLICER_ITEM_STYLE_TYPES, SLICER_STYLE_NAMES,
  TIMELINE_ELEMENT_STYLE_TYPES, type SlicerStyleOwner } from './slicer-style-catalog.ts'

const record = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
const integer = (value: unknown, min: number, max: number): value is number => typeof value === 'number' && Number.isInteger(value) && value >= min && value <= max
const flag = (value: unknown): value is 0 | 1 => value === 0 || value === 1
const borderSides = ['t', 'r', 'b', 'l', 'tl_br', 'tl_bc', 'tl_mr', 'bl_tr', 'ml_tr', 'bc_tr'] as const
const resourceKey = (type: SlicerType) => type === 'slicer' ? 'slicerStyles' : 'timelineStyles'
export const SLICER_STYLE_SOURCE_METADATA_KEY = 'slicerStyleResourcesSource'
export const SLICER_STYLE_DIAGNOSTICS_METADATA_KEY = 'slicerStyleResourceDiagnostics'

export class SlicerStyleResourceError extends Error {
  readonly code = 'SLICER_INVALID_STYLE_RESOURCE'
  constructor(readonly path: string, message: string) { super(`${message}（${path}）`); this.name = 'SlicerStyleResourceError' }
}
const invalid = (path: string, message = '切片器或时间线样式资源无效。'): never => { throw new SlicerStyleResourceError(path, message) }
function styleName(value: unknown, path: string): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 255) return invalid(path, '样式名称须为 1–255 个字符。')
  return value
}
function color(value: unknown, path: string): SlicerStyleColor {
  if (!record(value)) return invalid(path)
  const result: SlicerStyleColor = {}
  if (value.rgb !== undefined) {
    if (value.rgb !== null && typeof value.rgb !== 'string') return invalid(`${path}.rgb`)
    result.rgb = value.rgb
  }
  if (value.th !== undefined) {
    if (!integer(value.th, 0, 11)) return invalid(`${path}.th`)
    result.th = value.th as SlicerStyleColor['th']
  }
  return result
}
function overline(value: unknown, path: string): SlicerStyleTextDecoration {
  if (!record(value) || !flag(value.s)) return invalid(path)
  const result: SlicerStyleTextDecoration = { s: value.s }
  if (value.c !== undefined) { if (!flag(value.c)) return invalid(`${path}.c`); result.c = value.c }
  if (value.t !== undefined) { if (!integer(value.t, 0, 19)) return invalid(`${path}.t`); result.t = value.t }
  if (value.cl !== undefined) result.cl = color(value.cl, `${path}.cl`)
  return result
}
function borders(value: unknown, path: string): SlicerStyleBorders {
  if (!record(value)) return invalid(path)
  const result: SlicerStyleBorders = {}
  for (const side of borderSides) if (value[side] !== undefined) {
    const border = value[side]
    if (border === null) { result[side] = null; continue }
    if (!record(border) || !integer(border.s, 0, 13)) return invalid(`${path}.${side}`)
    const normalized: SlicerStyleBorder = { s: border.s, cl: color(border.cl, `${path}.${side}.cl`) }
    result[side] = normalized
  }
  return result
}
/** Mirrors the published projection, including null color/border/alignment resets. Unknown source fields are retained by import. */
export function normalizeSlicerStyleFormatting(value: unknown, path = 'formatting'): SlicerStyleFormatting {
  if (!record(value)) return invalid(path)
  const result: SlicerStyleFormatting = {}
  if (value.bg !== undefined) result.bg = value.bg === null ? null : color(value.bg, `${path}.bg`)
  if (value.cl !== undefined) result.cl = value.cl === null ? null : color(value.cl, `${path}.cl`)
  if (value.ol !== undefined) result.ol = overline(value.ol, `${path}.ol`)
  if (value.bd !== undefined) result.bd = value.bd === null ? null : borders(value.bd, `${path}.bd`)
  if (value.bl !== undefined) { if (!flag(value.bl)) return invalid(`${path}.bl`); result.bl = value.bl }
  if (value.ht !== undefined) {
    if (value.ht !== null && !integer(value.ht, 0, 6)) return invalid(`${path}.ht`)
    result.ht = value.ht as SlicerStyleFormatting['ht']
  }
  if (value.vt !== undefined) {
    if (value.vt !== null && !integer(value.vt, 0, 3)) return invalid(`${path}.vt`)
    result.vt = value.vt as SlicerStyleFormatting['vt']
  }
  return result
}
function definitionHeader(value: Record<string, unknown>, path: string): Pick<ISlicerStyleDefinition, 'name' | 'wholeStyle' | 'headerRowStyle'> {
  const result: Pick<ISlicerStyleDefinition, 'name' | 'wholeStyle' | 'headerRowStyle'> = { name: styleName(value.name, `${path}.name`) }
  for (const key of ['wholeStyle', 'headerRowStyle'] as const) if (value[key] !== undefined) result[key] = normalizeSlicerStyleFormatting(value[key], `${path}.${key}`)
  return result
}
function elements<T extends string>(value: unknown, allowed: readonly T[], path: string): Array<{ type: T; formatting?: SlicerStyleFormatting }> {
  if (!Array.isArray(value)) return invalid(path)
  const seen = new Set<string>(), result: Array<{ type: T; formatting?: SlicerStyleFormatting }> = []
  for (const [index, element] of value.entries()) {
    if (!record(element) || typeof element.type !== 'string' || !(allowed as readonly string[]).includes(element.type) || seen.has(element.type)) return invalid(`${path}[${index}]`)
    seen.add(element.type)
    result.push({ type: element.type as T, ...(element.formatting !== undefined ? { formatting: normalizeSlicerStyleFormatting(element.formatting, `${path}[${index}].formatting`) } : {}) })
  }
  return result
}
export function normalizeSlicerStyleDefinition(value: unknown, path = 'slicerStyle'): ISlicerStyleDefinition {
  if (!record(value)) return invalid(path)
  return { ...definitionHeader(value, path), slicerStyleElements: elements(value.slicerStyleElements, SLICER_ITEM_STYLE_TYPES, `${path}.slicerStyleElements`) }
}
export function normalizeTimelineStyleDefinition(value: unknown, path = 'timelineStyle'): ITimelineStyleDefinition {
  if (!record(value)) return invalid(path)
  return { ...definitionHeader(value, path), timelineStyleElements: elements(value.timelineStyleElements, TIMELINE_ELEMENT_STYLE_TYPES, `${path}.timelineStyleElements`) }
}
function uniqueDefinitions<T extends { name: string }>(value: unknown, normalize: (input: unknown, path: string) => T, type: SlicerType): T[] {
  const path = resourceKey(type)
  if (!Array.isArray(value)) return invalid(`${path}.styles`)
  const seen = new Set<string>(), result: T[] = []
  for (const [index, input] of value.entries()) {
    const definition = normalize(input, `${path}.styles[${index}]`), canonical = definition.name.toLowerCase()
    if (seen.has(canonical) || SLICER_STYLE_NAMES[type].includes(definition.name)) return invalid(`${path}.styles[${index}].name`, '自定义样式名称重复或覆盖内置样式。')
    seen.add(canonical); result.push(definition)
  }
  return result
}
export function normalizeSlicerStylesDefinition(value: unknown): ISlicerStylesDefinition {
  if (!record(value)) return invalid('slicerStyles')
  const result = { defaultSlicerStyle: styleName(value.defaultSlicerStyle, 'slicerStyles.defaultSlicerStyle'),
    styles: uniqueDefinitions(value.styles, normalizeSlicerStyleDefinition, 'slicer') }
  if (!hasSlicerStyle({ slicerStyles: result }, 'slicer', result.defaultSlicerStyle)) return invalid('slicerStyles.defaultSlicerStyle', '默认切片器样式不存在。')
  return result
}
export function normalizeTimelineStylesDefinition(value: unknown): ITimelineStylesDefinition {
  if (!record(value)) return invalid('timelineStyles')
  const result = { defaultTimelineStyle: styleName(value.defaultTimelineStyle, 'timelineStyles.defaultTimelineStyle'),
    styles: uniqueDefinitions(value.styles, normalizeTimelineStyleDefinition, 'timeline') }
  if (!hasSlicerStyle({ timelineStyles: result }, 'timeline', result.defaultTimelineStyle)) return invalid('timelineStyles.defaultTimelineStyle', '默认时间线样式不存在。')
  return result
}
export function projectSlicerStyles(value: unknown): ISlicerStylesDefinition | undefined {
  try { return normalizeSlicerStylesDefinition(value) } catch (error) { if (!(error instanceof SlicerStyleResourceError)) throw error; return undefined }
}
export function projectTimelineStyles(value: unknown): ITimelineStylesDefinition | undefined {
  try { return normalizeTimelineStylesDefinition(value) } catch (error) { if (!(error instanceof SlicerStyleResourceError)) throw error; return undefined }
}

export interface SlicerStyleResourceDiagnostic {
  code: 'SLICER_INVALID_STYLE_RESOURCE' | 'SLICER_STYLE_SOURCE_EXTENSIONS' | 'SLICER_UNRESOLVED_STYLE_REFERENCE'
  kind: SlicerType
  path: string
  message: string
}
export interface ImportedSlicerStyleResources extends SlicerStyleOwner { metadata?: Record<string, unknown> }
/** Valid families are projected independently; rejected or unrecognized source data is retained in metadata, never discarded. */
export function importSlicerStyleResources(input: unknown): ImportedSlicerStyleResources {
  if (!record(input)) return {}
  const result: ImportedSlicerStyleResources = {}, metadata: Record<string, unknown> = record(input.metadata) ? structuredClone(input.metadata) : {}
  const origin = metadata[SLICER_STYLE_SOURCE_METADATA_KEY]
  const retained: Record<string, unknown> = record(origin) ? structuredClone(origin) : {}
  const diagnostics: SlicerStyleResourceDiagnostic[] = []
  const inspect = (source: unknown, type: SlicerType, path: string) => {
    try {
      const projected = type === 'slicer' ? normalizeSlicerStylesDefinition(source) : normalizeTimelineStylesDefinition(source)
      if (!equivalentModelData(source, projected)) diagnostics.push({ code: 'SLICER_STYLE_SOURCE_EXTENSIONS', kind: type, path,
        message: '样式源中的扩展字段已保留，当前格式投影只使用固定公开字段。' })
      return projected
    } catch (error) {
      if (!(error instanceof SlicerStyleResourceError)) throw error
      diagnostics.push({ code: error.code, kind: type, path, message: error.message }); return undefined
    }
  }
  for (const type of ['slicer', 'timeline'] as const) {
    const key = resourceKey(type), source = input[key], originPath = `metadata.${SLICER_STYLE_SOURCE_METADATA_KEY}.${key}`
    if (source !== undefined) {
      if (!Object.hasOwn(retained, key)) retained[key] = structuredClone(source)
      const projected = inspect(source, type, equivalentModelData(source, retained[key]) ? originPath : key)
      if (type === 'slicer' && projected) result.slicerStyles = projected as ISlicerStylesDefinition
      if (type === 'timeline' && projected) result.timelineStyles = projected as ITimelineStylesDefinition
    }
    if (retained[key] !== undefined && !equivalentModelData(source, retained[key])) inspect(retained[key], type, originPath)
  }
  for (const [key, fallback] of [['slicers', 'slicer'], ['timelines', 'timeline']] as const) {
    const controls = input[key]
    if (!Array.isArray(controls)) continue
    for (const [index, control] of controls.entries()) {
      if (!record(control) || typeof control.style !== 'string') continue
      const type = control.type === undefined ? fallback : control.type
      if (type !== 'slicer' && type !== 'timeline') continue
      if (!hasSlicerStyle(result, type, control.style)) diagnostics.push({ code: 'SLICER_UNRESOLVED_STYLE_REFERENCE', kind: type,
        path: `${key}[${index}].style`, message: `作者样式 ${control.style} 尚未解析；显示使用文档默认，原引用保留。` })
    }
  }
  if (Object.keys(retained).length) metadata[SLICER_STYLE_SOURCE_METADATA_KEY] = retained
  if (diagnostics.length) metadata[SLICER_STYLE_DIAGNOSTICS_METADATA_KEY] = diagnostics
  else delete metadata[SLICER_STYLE_DIAGNOSTICS_METADATA_KEY]
  if (Object.keys(metadata).length) result.metadata = metadata
  return result
}

export function readSlicerStyleResources(workbook: SlicerStyleOwner, type: 'slicer'): ISlicerStylesDefinition | undefined
export function readSlicerStyleResources(workbook: SlicerStyleOwner, type: 'timeline'): ITimelineStylesDefinition | undefined
export function readSlicerStyleResources(workbook: SlicerStyleOwner, type: SlicerType): ISlicerStylesDefinition | ITimelineStylesDefinition | undefined
export function readSlicerStyleResources(workbook: SlicerStyleOwner, type: SlicerType): ISlicerStylesDefinition | ITimelineStylesDefinition | undefined {
  const value = type === 'slicer' ? workbook.slicerStyles : workbook.timelineStyles
  return value === undefined ? undefined : structuredClone(value)
}
/** Pure-domain author mutation; the host command owns permissions, cancellation, history and events. */
export function setSlicerStyleResources(workbook: WorkbookModel, type: SlicerType, input: unknown): boolean {
  const nextSlicer = type === 'slicer' && input !== undefined ? normalizeSlicerStylesDefinition(input) : undefined
  const nextTimeline = type === 'timeline' && input !== undefined ? normalizeTimelineStylesDefinition(input) : undefined
  const next = type === 'slicer' ? nextSlicer : nextTimeline
  if (next !== undefined && !equivalentModelData(input, next)) return invalid(resourceKey(type), '编辑样式含尚未支持的扩展字段，请通过导入保留源数据。')
  const previous = readSlicerStyleResources(workbook, type)
  if (equivalentModelData(previous, next)) return false
  const candidate: SlicerStyleOwner = { slicerStyles: workbook.slicerStyles, timelineStyles: workbook.timelineStyles }
  if (type === 'slicer') candidate.slicerStyles = nextSlicer
  else candidate.timelineStyles = nextTimeline
  const affected: SlicerModel[] = []
  for (const control of workbook.slicers ?? []) if (control.type === type) {
    if (control.style !== undefined && hasSlicerStyle(workbook, type, control.style) && !hasSlicerStyle(candidate, type, control.style)) return invalid(`slicers.${control.id}.style`, '不能删除仍被控件引用的自定义样式。')
    if (!equivalentModelData(getSlicerStyle(workbook, type, control.style), getSlicerStyle(candidate, type, control.style))) affected.push(control)
  }
  if (type === 'slicer') { if (nextSlicer === undefined) delete workbook.slicerStyles; else workbook.slicerStyles = nextSlicer }
  else { if (nextTimeline === undefined) delete workbook.timelineStyles; else workbook.timelineStyles = nextTimeline }
  for (const control of affected) control.revision += 1
  return true
}
export function upsertSlicerStyleDefinition(workbook: WorkbookModel, type: SlicerType, input: unknown): boolean {
  if (type === 'slicer') {
    const definition = normalizeSlicerStyleDefinition(input), source = readSlicerStyleResources(workbook, 'slicer') ?? { defaultSlicerStyle: getDefaultSlicerStyleName(workbook), styles: [] }
    if (!equivalentModelData(input, definition)) return invalid('slicerStyle', '编辑样式含尚未支持的扩展字段。')
    const index = source.styles.findIndex(style => style.name.toLowerCase() === definition.name.toLowerCase())
    if (index < 0) source.styles.push(definition); else source.styles[index] = definition
    return setSlicerStyleResources(workbook, type, source)
  }
  const definition = normalizeTimelineStyleDefinition(input), source = readSlicerStyleResources(workbook, 'timeline') ?? { defaultTimelineStyle: getDefaultSlicerStyleName(workbook, 'timeline'), styles: [] }
  if (!equivalentModelData(input, definition)) return invalid('timelineStyle', '编辑样式含尚未支持的扩展字段。')
  const index = source.styles.findIndex(style => style.name.toLowerCase() === definition.name.toLowerCase())
  if (index < 0) source.styles.push(definition); else source.styles[index] = definition
  return setSlicerStyleResources(workbook, type, source)
}
export function removeSlicerStyleDefinition(workbook: WorkbookModel, type: SlicerType, name: string): boolean {
  styleName(name, `${resourceKey(type)}.name`)
  if (SLICER_STYLE_NAMES[type].includes(name)) return invalid(`${resourceKey(type)}.name`, '内置样式不能删除。')
  if (type === 'slicer') {
    const source = readSlicerStyleResources(workbook, 'slicer')
    if (!source || !source.styles.some(style => style.name === name)) return false
    source.styles = source.styles.filter(style => style.name !== name)
    return setSlicerStyleResources(workbook, type, source)
  }
  const source = readSlicerStyleResources(workbook, 'timeline')
  if (!source || !source.styles.some(style => style.name === name)) return false
  source.styles = source.styles.filter(style => style.name !== name)
  return setSlicerStyleResources(workbook, type, source)
}
export function setDefaultSlicerStyle(workbook: WorkbookModel, type: SlicerType, name: string): boolean {
  styleName(name, `${resourceKey(type)}.default`)
  if (!hasSlicerStyle(workbook, type, name)) return invalid(`${resourceKey(type)}.default`, '默认样式不存在。')
  if (type === 'slicer') {
    const source = readSlicerStyleResources(workbook, 'slicer') ?? { defaultSlicerStyle: getDefaultSlicerStyleName(workbook), styles: [] }
    source.defaultSlicerStyle = name; return setSlicerStyleResources(workbook, type, source)
  }
  const source = readSlicerStyleResources(workbook, 'timeline') ?? { defaultTimelineStyle: getDefaultSlicerStyleName(workbook, 'timeline'), styles: [] }
  source.defaultTimelineStyle = name; return setSlicerStyleResources(workbook, type, source)
}
