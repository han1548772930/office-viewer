import type { DocumentTableBorder, DocumentTableCellMargin, DocumentTableColumnTypeConfig } from '../../features/document/model/document-model'
import type { DocumentTableCellRange } from '../../features/document/model/document-table-grid'
import { assertDocumentTableBorder, assertDocumentTableColor, assertDocumentTableColumnType, assertDocumentTableMargin,
  assertDocumentTableBorderPatch, assertDocumentTableMarginPatch, DOCUMENT_TABLE_BORDER_PRESETS, type DocumentTableBorderPreset } from '../../features/document/model/document-table-style'
import type { DocumentTableCommand, DocumentTableCommandTarget, DocumentTableStylePatch } from '../../features/document/commands/table-command-types'
import { assertDocumentTableStylePatch } from '../../features/document/commands/table-mutations'
import { validateDocumentRunPatch } from '../../features/document/model/document-run-range'
import type { RunPatch } from '../../features/document/model/document-model'

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
const optionalNumber = (value: unknown) => value === undefined || typeof value === 'number' && Number.isFinite(value)

function range(value: unknown): DocumentTableCellRange | undefined {
  if (!record(value)) return undefined
  const { startRow, endRow, startColumn, endColumn } = value
  return typeof startRow === 'number' && typeof endRow === 'number' && typeof startColumn === 'number' && typeof endColumn === 'number'
    && [startRow, endRow, startColumn, endColumn].every(item => Number.isSafeInteger(item) && item >= 0) && startRow <= endRow && startColumn <= endColumn
    ? { startRow, endRow, startColumn, endColumn } : undefined
}
function border(value: unknown): DocumentTableBorder | undefined {
  if (!record(value) || typeof value.width !== 'number' || typeof value.color !== 'string' || value.style !== 'solid' && value.style !== 'dot' && value.style !== 'dash') return undefined
  const result: DocumentTableBorder = { width: value.width, color: value.color, style: value.style }
  try { assertDocumentTableBorder(result); return result } catch { return undefined }
}
function margin(value: unknown): DocumentTableCellMargin | undefined {
  if (!record(value) || typeof value.top !== 'number' || typeof value.right !== 'number' || typeof value.bottom !== 'number' || typeof value.left !== 'number') return undefined
  const result = { top: value.top, right: value.right, bottom: value.bottom, left: value.left }
  try { assertDocumentTableMargin(result); return result } catch { return undefined }
}
function borderPatch(value: unknown): Partial<DocumentTableBorder> | undefined {
  if (!record(value) || Object.keys(value).some(key => !['width', 'color', 'style'].includes(key))
    || value.width !== undefined && typeof value.width !== 'number' || value.color !== undefined && typeof value.color !== 'string'
    || value.style !== undefined && value.style !== 'solid' && value.style !== 'dot' && value.style !== 'dash') return undefined
  const result: Partial<DocumentTableBorder> = { ...(typeof value.width === 'number' ? { width: value.width } : {}),
    ...(typeof value.color === 'string' ? { color: value.color } : {}), ...(value.style === 'solid' || value.style === 'dot' || value.style === 'dash' ? { style: value.style } : {}) }
  try { assertDocumentTableBorderPatch(result); return result } catch { return undefined }
}
function marginPatch(value: unknown): Partial<DocumentTableCellMargin> | undefined {
  if (!record(value) || Object.keys(value).some(key => !['top', 'right', 'bottom', 'left'].includes(key))) return undefined
  const result: Partial<DocumentTableCellMargin> = {}
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const amount = value[side]
    if (amount !== undefined && typeof amount !== 'number') return undefined
    if (typeof amount === 'number') result[side] = amount
  }
  try { assertDocumentTableMarginPatch(result); return result } catch { return undefined }
}
function columnType(value: unknown): DocumentTableColumnTypeConfig | undefined {
  if (!record(value) || value.type !== 'text' && value.type !== 'number' && value.type !== 'date' && value.type !== 'checkbox' && value.type !== 'dropdown') return undefined
  const result: DocumentTableColumnTypeConfig = { type: value.type }
  if (value.dropdownOptions !== undefined) {
    if (!Array.isArray(value.dropdownOptions)) return undefined
    const options: NonNullable<DocumentTableColumnTypeConfig['dropdownOptions']> = []
    for (const option of value.dropdownOptions) {
      if (!record(option) || typeof option.id !== 'string' || typeof option.label !== 'string' || option.color !== undefined && typeof option.color !== 'string') return undefined
      options.push({ id: option.id, label: option.label, ...(typeof option.color === 'string' ? { color: option.color } : {}) })
    }
    result.dropdownOptions = options
  }
  try { assertDocumentTableColumnType(result); return result } catch { return undefined }
}

function stylePatch(value: unknown): DocumentTableStylePatch | undefined {
  if (!record(value)) return undefined
  const patch: DocumentTableStylePatch = {}
  if (value.backgroundColor !== undefined) {
    if (typeof value.backgroundColor !== 'string' && value.backgroundColor !== null) return undefined
    patch.backgroundColor = value.backgroundColor
  }
  if (value.align !== undefined) { if (value.align !== 'left' && value.align !== 'center' && value.align !== 'right') return undefined; patch.align = value.align }
  if (value.layout !== undefined) { if (value.layout !== 'auto' && value.layout !== 'fixed') return undefined; patch.layout = value.layout }
  for (const property of ['width', 'indent', 'offsetY'] as const) {
    if (!optionalNumber(value[property])) return undefined
    if (typeof value[property] === 'number') patch[property] = value[property]
  }
  if (value.columnWidths !== undefined) {
    if (!Array.isArray(value.columnWidths) || value.columnWidths.some(width => typeof width !== 'number' || !Number.isFinite(width) || width <= 0)) return undefined
    patch.columnWidths = value.columnWidths.map(Number)
  }
  if (value.cellMargin !== undefined) {
    const parsed = value.cellMargin === null ? null : margin(value.cellMargin)
    if (parsed === undefined) return undefined
    patch.cellMargin = parsed
  }
  if (value.border !== undefined) {
    if (!record(value.border)) return undefined
    const input = value.border, parsed: NonNullable<DocumentTableStylePatch['border']> = {}
    if (input.preset !== undefined) {
      if (!isPreset(input.preset)) return undefined
      parsed.preset = input.preset
    }
    if (input.border !== undefined) { const item = border(input.border); if (!item) return undefined; parsed.border = item }
    if (input.patch !== undefined) { const item = borderPatch(input.patch); if (!item) return undefined; parsed.patch = item }
    if (input.color !== undefined) { if (typeof input.color !== 'string' && input.color !== null) return undefined; parsed.color = input.color }
    if (!optionalNumber(input.width)) return undefined
    if (typeof input.width === 'number') parsed.width = input.width
    if (input.style !== undefined) { if (input.style !== 'solid' && input.style !== 'dot' && input.style !== 'dash') return undefined; parsed.style = input.style }
    patch.border = parsed
  }
  try { assertDocumentTableStylePatch(patch); return patch } catch { return undefined }
}

function isPreset(value: unknown): value is DocumentTableBorderPreset {
  return typeof value === 'string' && DOCUMENT_TABLE_BORDER_PRESETS.some(preset => preset === value)
}

/** Preserve explicit table IDs/ranges; invalid parameters never degrade to selection defaults. */
export function documentTableCommandFromRibbon(id: string, input?: unknown): DocumentTableCommand | undefined {
  if (input !== undefined && !record(input)) return undefined
  const params = record(input) ? input : {}, target: DocumentTableCommandTarget = {}
  if (params.tableId !== undefined) { if (typeof params.tableId !== 'string' || !params.tableId.trim()) return undefined; target.tableId = params.tableId }
  if (params.range !== undefined) { const parsed = range(params.range); if (!parsed) return undefined; target.range = parsed }
  if (id === 'doc.command.table-delete-rows' || id === 'doc.command.table-delete-columns' || id === 'doc.command.table-delete'
    || id === 'doc.command.table-merge-cells' || id === 'doc.command.table-unmerge-cells' || id === 'doc.command.table-distribute-columns') return { ...target, id }
  if (id === 'doc.command.table-insert-row' || id === 'doc.command.table-insert-column') {
    if (params.position !== 'before' && params.position !== 'after' || !optionalNumber(params.count) || !optionalNumber(params.row) || !optionalNumber(params.column)) return undefined
    const count = typeof params.count === 'number' ? params.count : undefined
    return id === 'doc.command.table-insert-row' ? { ...target, id, position: params.position, row: typeof params.row === 'number' ? params.row : undefined, count }
      : { ...target, id, position: params.position, column: typeof params.column === 'number' ? params.column : undefined, count }
  }
  if (id === 'doc.command.table-move-rows' || id === 'doc.command.table-move-columns') {
    if (params.position !== 'before' && params.position !== 'after' || !optionalNumber(params.count) || !optionalNumber(params.fromRow) || !optionalNumber(params.fromColumn)) return undefined
    const count = typeof params.count === 'number' ? params.count : undefined
    return id === 'doc.command.table-move-rows' && typeof params.toRow === 'number' ? { ...target, id, toRow: params.toRow, position: params.position, fromRow: typeof params.fromRow === 'number' ? params.fromRow : undefined, count }
      : id === 'doc.command.table-move-columns' && typeof params.toColumn === 'number' ? { ...target, id, toColumn: params.toColumn, position: params.position, fromColumn: typeof params.fromColumn === 'number' ? params.fromColumn : undefined, count } : undefined
  }
  if (id === 'doc.command.table-move') {
    if (params.position !== undefined && params.position !== 'before' && params.position !== 'after' || params.targetBlockId !== undefined && typeof params.targetBlockId !== 'string'
      || ![params.targetOffset, params.deltaX, params.deltaY].every(optionalNumber)) return undefined
    return { ...target, id, ...(params.position === 'before' || params.position === 'after' ? { position: params.position } : {}),
      ...(typeof params.targetBlockId === 'string' ? { targetBlockId: params.targetBlockId } : {}), ...(typeof params.targetOffset === 'number' ? { targetOffset: params.targetOffset } : {}),
      ...(typeof params.deltaX === 'number' ? { deltaX: params.deltaX } : {}), ...(typeof params.deltaY === 'number' ? { deltaY: params.deltaY } : {}) }
  }
  if (id === 'doc.command.table-resize-row') {
    if (!optionalNumber(params.height) || !optionalNumber(params.row) || params.heightRule !== undefined && params.heightRule !== 'auto' && params.heightRule !== 'atLeast' && params.heightRule !== 'exact'
      || params.allowSplit !== undefined && typeof params.allowSplit !== 'boolean' || params.preserveHeightRule !== undefined && typeof params.preserveHeightRule !== 'boolean'
      || params.height === undefined && params.heightRule === undefined && params.allowSplit === undefined) return undefined
    return { ...target, id, ...(typeof params.height === 'number' ? { height: params.height } : {}), ...(typeof params.row === 'number' ? { row: params.row } : {}),
      ...(params.heightRule === 'auto' || params.heightRule === 'atLeast' || params.heightRule === 'exact' ? { heightRule: params.heightRule } : {}), ...(typeof params.allowSplit === 'boolean' ? { allowSplit: params.allowSplit } : {}),
      ...(typeof params.preserveHeightRule === 'boolean' ? { preserveHeightRule: params.preserveHeightRule } : {}) }
  }
  if (id === 'doc.command.table-resize-column' && typeof params.width === 'number' && optionalNumber(params.column)) {
    if (params.widths !== undefined && (!Array.isArray(params.widths) || params.widths.some(value => typeof value !== 'number' || !Number.isFinite(value) || value <= 0))) return undefined
    return { ...target, id, width: params.width, column: typeof params.column === 'number' ? params.column : undefined,
      ...(Array.isArray(params.widths) ? { widths: params.widths.map(Number) } : {}) }
  }
  if (id === 'doc.command.table-distribute-rows') {
    if (params.heights !== undefined && (!Array.isArray(params.heights) || !params.heights.every(value => typeof value === 'number' && Number.isFinite(value)))) return undefined
    return { ...target, id, ...(Array.isArray(params.heights) ? { heights: params.heights.filter((value): value is number => typeof value === 'number') } : {}) }
  }
  if (id === 'doc.command.table-set-background' || id === 'doc.command.table-set-border-color') {
    if (typeof params.color !== 'string' && params.color !== null || params.scope !== undefined && params.scope !== 'table' && params.scope !== 'selection') return undefined
    try { assertDocumentTableColor(params.color) } catch { return undefined }
    return id === 'doc.command.table-set-background' ? { ...target, id, color: params.color, ...(params.scope === 'table' || params.scope === 'selection' ? { scope: params.scope } : {}) } : { ...target, id, color: params.color }
  }
  if (id === 'doc.command.table-set-border') {
    if (!isPreset(params.preset)) return undefined
    if (params.border !== undefined && params.borderPatch !== undefined) return undefined
    const parsed = params.border === undefined ? undefined : border(params.border), patch = params.borderPatch === undefined ? undefined : borderPatch(params.borderPatch)
    return params.border !== undefined && !parsed || params.borderPatch !== undefined && !patch ? undefined
      : { ...target, id, preset: params.preset, ...(parsed ? { border: parsed } : {}), ...(patch ? { borderPatch: patch } : {}) }
  }
  if (id === 'doc.command.table-set-border-width' && typeof params.width === 'number') return { ...target, id, width: params.width }
  if (id === 'doc.command.table-set-border-style' && (params.style === 'solid' || params.style === 'dot' || params.style === 'dash')) return { ...target, id, style: params.style }
  if (id === 'doc.command.table-set-vertical-align' && (params.value === 'top' || params.value === 'center' || params.value === 'bottom')) return { ...target, id, value: params.value }
  if (id === 'doc.command.table-set-text-align' && (params.value === 'left' || params.value === 'center' || params.value === 'right' || params.value === 'justify')) return { ...target, id, value: params.value }
  if (id === 'doc.command.table-set-align' && (params.value === 'left' || params.value === 'center' || params.value === 'right')) return { ...target, id, value: params.value }
  if (id === 'doc.command.table-set-cell-margin') {
    if (params.margin !== undefined && params.marginPatch !== undefined || params.scope !== undefined && params.scope !== 'table' && params.scope !== 'selection') return undefined
    const parsed = params.margin === null ? null : margin(params.margin), patch = params.marginPatch === undefined ? undefined : marginPatch(params.marginPatch)
    if (params.marginPatch !== undefined ? !patch : parsed === undefined) return undefined
    return { ...target, id, ...(patch ? { marginPatch: patch } : { margin: parsed }), ...(params.scope === 'table' || params.scope === 'selection' ? { scope: params.scope } : {}) }
  }
  if (id === 'doc.command.table-set-cell-text' && typeof params.row === 'number' && typeof params.column === 'number' && typeof params.text === 'string' && (params.optionId === undefined || typeof params.optionId === 'string')) return { ...target, id, row: params.row, column: params.column, text: params.text,
    ...(typeof params.optionId === 'string' ? { optionId: params.optionId } : {}) }
  if (id === 'doc.command.table-set-range-text') {
    if (params.mode !== 'clear' && params.mode !== 'replace' && params.mode !== 'matrix'
      || params.text !== undefined && typeof params.text !== 'string'
      || params.mode === 'replace' && typeof params.text !== 'string') return undefined
    let values: string[][] | undefined
    if (params.values !== undefined) {
      if (!Array.isArray(params.values) || params.values.some(row => !Array.isArray(row) || row.some(value => typeof value !== 'string'))) return undefined
      values = params.values.map(row => [...row] as string[])
    }
    if (params.mode === 'matrix' && !values?.length) return undefined
    let marks: RunPatch | undefined
    if (params.marks !== undefined) {
      if (!record(params.marks)) return undefined
      try { validateDocumentRunPatch(params.marks as RunPatch); marks = { ...params.marks } as RunPatch } catch { return undefined }
    }
    return { ...target, id, mode: params.mode, ...(typeof params.text === 'string' ? { text: params.text } : {}), ...(values ? { values } : {}), ...(marks ? { marks } : {}) }
  }
  if (id === 'doc.command.table-set-header-row-count' && typeof params.count === 'number') return { ...target, id, count: params.count }
  if (id === 'doc.command.table-insert-title-row' && (params.text === undefined || typeof params.text === 'string')) return { ...target, id, ...(typeof params.text === 'string' ? { text: params.text } : {}) }
  if (id === 'doc.command.table-set-column-type' && optionalNumber(params.column)) {
    const parsed = columnType(params.config)
    return parsed ? { ...target, id, column: typeof params.column === 'number' ? params.column : undefined, config: parsed } : undefined
  }
  if (id === 'doc.command.table-sort' && (params.direction === 'ascending' || params.direction === 'descending') && optionalNumber(params.column) && (params.locale === undefined || typeof params.locale === 'string')) return { ...target, id, direction: params.direction,
    ...(typeof params.column === 'number' ? { column: params.column } : {}), ...(typeof params.locale === 'string' ? { locale: params.locale } : {}) }
  if (id === 'doc.command.table-set-style') { const parsed = stylePatch(params.patch); return parsed ? { ...target, id, patch: parsed } : undefined }
  return undefined
}
