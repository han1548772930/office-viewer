import type { DocumentContentControl } from '../../../domain/model'
import type { ContentControlPatch } from './content-controls'
import { validateContentControlCheckboxState, validateContentControlDateProperties, validateContentControlListItems } from './content-control-presentation'
import { assertDocumentContentControlTableRowRange } from '../model/content-control-table-rows'
import { assertDocumentContentControlBlockRange } from '../model/content-control-block-range'

const propertyKeys = [
  'type', 'placement', 'alias', 'title', 'tag', 'appearance', 'color', 'placeholder', 'placeholderDocPart', 'showingPlaceholder',
  'temporary', 'lock', 'checkbox', 'multiLine', 'dateFormat', 'dateLocale', 'dateCalendar', 'dateStorage', 'items', 'parentId',
  'repeatingSection', 'tableRowRange', 'blockRange', 'docPart', 'dataBinding', 'rawSdtPr', 'rawSdtEndPr', 'metadata',
] as const satisfies ReadonlyArray<keyof ContentControlPatch>

const stringKeys = ['alias', 'title', 'tag', 'color', 'placeholder', 'dateFormat', 'dateLocale', 'dateCalendar', 'dateStorage', 'parentId', 'rawSdtPr', 'rawSdtEndPr'] as const
const booleanKeys = ['showingPlaceholder', 'temporary', 'multiLine'] as const
const types = ['richText', 'text', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj']

/** Mutations validate supplied properties; imported unknown XML remains in raw fields. */
export function validateContentControlPropertyPatch(patch: ContentControlPatch, current?: DocumentContentControl): void {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new TypeError('内容控件属性必须是对象。')
  for (const key of Object.keys(patch)) if (!propertyKeys.some(property => property === key)) throw new TypeError(`不能通过属性命令修改内容控件字段：${key}`)
  assertJsonProperties(patch)
  for (const key of stringKeys) if (patch[key] !== undefined && typeof patch[key] !== 'string') throw new TypeError(`内容控件 ${key} 必须是文本。`)
  for (const key of booleanKeys) if (patch[key] !== undefined && typeof patch[key] !== 'boolean') throw new TypeError(`内容控件 ${key} 必须是布尔值。`)
  if (patch.type !== undefined && !types.includes(patch.type)) throw new TypeError('Unsupported content control type')
  if (patch.placement !== undefined && !['inline', 'block', 'table', 'row', 'cell', 'header', 'footer'].includes(patch.placement)) throw new TypeError('内容控件位置无效。')
  if (patch.appearance !== undefined && !['boundingBox', 'tags', 'hidden'].includes(patch.appearance)) throw new TypeError('内容控件外观无效。')
  if (patch.lock !== undefined && !['unlocked', 'controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(patch.lock)) throw new TypeError('内容控件锁定状态无效。')
  if (patch.color !== undefined && patch.color !== 'auto' && !/^#?[\da-f]{6}$/i.test(patch.color)) throw new RangeError('颜色必须是六位十六进制色值。')
  if (patch.checkbox !== undefined) {
    assertObject(patch.checkbox, 'checkbox')
    if (typeof patch.checkbox.checked !== 'boolean') throw new TypeError('复选框状态必须是布尔值。')
    for (const key of ['checkedState', 'uncheckedState'] as const) {
      const state = patch.checkbox[key]
      if (state !== undefined) { assertObject(state, key); validateContentControlCheckboxState(state) }
    }
  }
  if (patch.items !== undefined) validateContentControlListItems(patch.items)
  if (patch.tableRowRange !== undefined) assertDocumentContentControlTableRowRange(patch.tableRowRange)
  if (patch.blockRange !== undefined) assertDocumentContentControlBlockRange(patch.blockRange)
  const rowRange = Object.hasOwn(patch, 'tableRowRange') ? patch.tableRowRange : current?.tableRowRange
  const blockRange = Object.hasOwn(patch, 'blockRange') ? patch.blockRange : current?.blockRange
  const placement = Object.hasOwn(patch, 'placement') ? patch.placement : current?.placement
  if (rowRange && blockRange) throw new RangeError('内容控件不能同时包含完整块范围和表格行范围。')
  if (blockRange && placement !== 'block') throw new RangeError('完整块范围必须使用块级内容控件位置。')
  const dates = ['dateFormat', 'dateLocale', 'dateCalendar'] as const
  if (dates.some(key => Object.hasOwn(patch, key))) validateContentControlDateProperties(
    Object.hasOwn(patch, 'dateFormat') ? patch.dateFormat : current?.dateFormat,
    Object.hasOwn(patch, 'dateLocale') ? patch.dateLocale : current?.dateLocale,
    Object.hasOwn(patch, 'dateCalendar') ? patch.dateCalendar : current?.dateCalendar,
  )
  if (patch.dateStorage !== undefined && !['text', 'date', 'dateTime'].includes(patch.dateStorage)) throw new RangeError('日期存储格式无效。')
  for (const key of ['placeholderDocPart', 'repeatingSection', 'docPart', 'dataBinding', 'metadata'] as const) if (patch[key] !== undefined) assertObject(patch[key], key)
  if (patch.placeholderDocPart) validateStringFields(patch.placeholderDocPart, ['name', 'text'])
  if (patch.repeatingSection) {
    validateStringFields(patch.repeatingSection, ['sectionTitle'])
    validateBooleanField(patch.repeatingSection, 'doNotAllowInsertDelete')
  }
  if (patch.docPart) { validateStringFields(patch.docPart, ['gallery', 'category']); validateBooleanField(patch.docPart, 'unique') }
  if (patch.dataBinding) {
    validateStringFields(patch.dataBinding, ['prefixMappings', 'xpath', 'storeItemID', 'cachedResult'])
    validateBooleanField(patch.dataBinding, 'dirty')
  }
}

function assertObject(value: unknown, label: string): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`内容控件 ${label} 必须是对象。`)
}

function validateStringFields(value: object, fields: readonly string[]): void {
  for (const key of fields) if (Reflect.get(value, key) !== undefined && typeof Reflect.get(value, key) !== 'string') throw new TypeError(`内容控件 ${key} 必须是文本。`)
}

function validateBooleanField(value: object, key: string): void {
  if (Reflect.get(value, key) !== undefined && typeof Reflect.get(value, key) !== 'boolean') throw new TypeError(`内容控件 ${key} 必须是布尔值。`)
}

function assertJsonProperties(value: unknown, ancestors = new Set<object>()): void {
  if (value === undefined || value === null || typeof value === 'string' || typeof value === 'boolean') return
  if (typeof value === 'number' && Number.isFinite(value)) return
  if (typeof value !== 'object') throw new TypeError('内容控件属性只接受可保存的 JSON 数据。')
  if (ancestors.has(value)) throw new TypeError('内容控件属性不能包含循环引用。')
  if (!Array.isArray(value) && Object.getPrototypeOf(value) !== Object.prototype && Object.getPrototypeOf(value) !== null) throw new TypeError('内容控件属性只接受普通数据对象。')
  ancestors.add(value)
  for (const child of Object.values(value)) assertJsonProperties(child, ancestors)
  ancestors.delete(value)
}

/** Only edited properties are submitted; a properties dialog must not overwrite later value edits. */
export function contentControlPropertyDraftPatch(baseline: DocumentContentControl, draft: ContentControlPatch): ContentControlPatch {
  const result: ContentControlPatch = {}
  for (const key of propertyKeys) {
    if (!Object.hasOwn(draft, key) || JSON.stringify(baseline[key]) === JSON.stringify(draft[key])) continue
    Object.assign(result, { [key]: draft[key] })
  }
  return result
}

export function assertContentControlPropertyDraftCurrent(baseline: DocumentContentControl, current: DocumentContentControl, patch: ContentControlPatch): void {
  if (current.id !== baseline.id || current.type !== baseline.type || current.blockId !== baseline.blockId
    || current.startOffset !== baseline.startOffset || current.endOffset !== baseline.endOffset
    || current.startNoteReferenceCount !== baseline.startNoteReferenceCount || current.endNoteReferenceCount !== baseline.endNoteReferenceCount
    || JSON.stringify(current.tableRowRange) !== JSON.stringify(baseline.tableRowRange)
    || JSON.stringify(current.blockRange) !== JSON.stringify(baseline.blockRange)) throw new Error('内容控件位置或类型已变化，请重新打开属性。')
  for (const key of propertyKeys) {
    if (Object.hasOwn(patch, key) && JSON.stringify(current[key]) !== JSON.stringify(baseline[key])) throw new Error('内容控件属性已变化，请重新打开属性。')
  }
}
