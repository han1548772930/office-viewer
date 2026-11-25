import type { PivotTableModel, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { mapConditionalFormatFormulas } from '../conditional-formatting/conditional-format-formulas.ts'
import { mapFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { remapSheetClipboardTableFormula } from './sheet-clipboard-workbook-resources.ts'
import { moveSheetClipboardHyperlink } from './sheet-clipboard-references.ts'
import { validateSheetName } from './sheet-workbench.ts'
import { copySheetCellImage } from '../sheet-images/sheet-image-model.ts'
import type { SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { drawingRefKey } from '../drawing-groups/index.ts'
import { createSheetDrawingScene } from '../sheet-drawings/sheet-drawing-host.ts'

export const COPY_WORKSHEET_COMMAND = 'sheet.command.copy-sheet'
const fullSheet = { startRow: 0, startCol: 0, endRow: 1_048_575, endCol: 16_383 }
const sameName = (first: string | undefined, second: string): boolean => first?.toLowerCase() === second.toLowerCase()
const newId = (kind: string): string => `copy-${kind}-${crypto.randomUUID()}`
const quoteSheet = (name: string): string => `'${name.replace(/'/g, "''")}'!`

export interface CopyWorksheetCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  /** Local extension: identifies the new sheet without reading mutable active-sheet state. */
  targetSheetId?: string
  targetSheetName?: string
  expectedRevision?: number
  expectedHistoryGeneration?: number
  host?: { unitId: string; sheetId: string }
}

export interface WorksheetCopyPlan {
  workbook: WorkbookModel
  sourceSheetId: string
  targetSheetId: string
  targetSheetName: string
  index: number
  resources: { tables: number; pivotTables: number; charts: number; slicers: number }
}

/** Keep the Office 31 UTF-16 limit, including long names and surrogate pairs. */
export function createWorksheetCopyName(sheets: readonly Pick<SpreadsheetSheet, 'id' | 'name'>[], sourceName: string): string {
  for (let index = 2; ; index++) {
    const suffix = ` (${index})`
    let base = sourceName.slice(0, 31 - suffix.length)
    const last = base.charCodeAt(base.length - 1)
    if (last >= 0xd800 && last <= 0xdbff) base = base.slice(0, -1)
    const name = `${base}${suffix}`
    if (validateSheetName(name, sheets).valid) return name
    if (!validateSheetName(sourceName, []).valid) throw new TypeError('源工作表名称无效，请先重命名。')
  }
}

/** Rewrite explicit self-sheet references; coordinates, literals, external and 3D references retain their meaning. */
export function remapWorksheetCopyFormula(formula: string, sourceName: string, targetName: string, tables: ReadonlyMap<string, string> = new Map()): string {
  return remapSheetClipboardTableFormula(mapFormulaReferenceAreas(formula, area => {
    if (area.external || area.threeDimensional) return { ...area }
    return { ...area,
      ...(sameName(area.sheetName, sourceName) ? { qualifier: quoteSheet(targetName), sheetName: targetName } : {}),
      ...(sameName(area.endSheetName, sourceName) ? { endQualifier: quoteSheet(targetName), endSheetName: targetName } : {}),
    }
  }), tables)
}

function uniqueResourceName(name: string, used: Set<string>): string {
  let candidate = name, suffix = 2
  while (used.has(candidate.toLowerCase())) candidate = `${name.slice(0, 240)}_${suffix++}`
  used.add(candidate.toLowerCase())
  return candidate
}

function copyLocalNames(metadata: Record<string, unknown> | undefined, source: SpreadsheetSheet, target: SpreadsheetSheet,
  sourceIndex: number, remap: (formula: string) => string, workbookScope: boolean): Record<string, unknown> | undefined {
  if (!metadata) return undefined
  const next = cloneModel(metadata)
  for (const key of ['definedNames', 'localDefinedNames']) {
    const entries = next[key]
    if (!Array.isArray(entries)) continue
    const copies: unknown[] = []
    const shifted = entries.map((entry: unknown) => {
      if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return entry
      const record = entry as Record<string, unknown>
      const owns = record.sheetId === source.id || record.subUnitId === source.id || record.sheet === source.name
        || record.localSheetId === sourceIndex || !workbookScope && record.sheetId === undefined && record.sheet === undefined
          && record.localSheetId === undefined && record.subUnitId === undefined
      if (owns) {
        const copied = { ...record,
          ...(record.id === undefined ? {} : { id: newId('defined-name') }),
          ...(record.sheetId === undefined ? {} : { sheetId: target.id }),
          ...(record.subUnitId === undefined ? {} : { subUnitId: target.id }),
          ...(record.sheet === undefined ? {} : { sheet: target.name }),
          ...(record.localSheetId === undefined ? {} : { localSheetId: sourceIndex + 1 }),
          ...(typeof record.refersTo === 'string' ? { refersTo: remap(record.refersTo) } : {}),
        }
        if (!workbookScope) return copied
        copies.push(copied)
      }
      return workbookScope && Number.isSafeInteger(record.localSheetId) && Number(record.localSheetId) > sourceIndex
        ? { ...record, localSheetId: Number(record.localSheetId) + 1 } : record
    })
    next[key] = [...shifted, ...copies]
  }
  return next
}

/** Plans a complete detached transaction; hidden rows and sparse sheet dimensions never pass through range clipboard projection. */
export function planWorksheetCopy(original: WorkbookModel, sourceSheetId: string,
  options: { userId: string; targetSheetId?: string; targetSheetName?: string }): WorksheetCopyPlan {
  const sourceIndex = original.sheets.findIndex(sheet => sheet.id === sourceSheetId), source = original.sheets[sourceIndex]
  if (!source) throw new Error('待复制的工作表不存在。')
  if (!canViewWorksheetProtection(source, options.userId) || !canViewProtectedRanges(source, [fullSheet], options.userId)) {
    throw new Error('当前用户没有查看或复制整个工作表的权限。')
  }
  const targetSheetId = options.targetSheetId ?? newId('sheet')
  if (!targetSheetId.trim() || original.sheets.some(sheet => sheet.id === targetSheetId)) throw new Error('副本工作表 ID 必须唯一。')
  const targetSheetName = options.targetSheetName ?? createWorksheetCopyName(original.sheets, source.name)
  const name = validateSheetName(targetSheetName, original.sheets)
  if (!name.valid) throw new Error(name.message)
  const workbook = cloneModel(original), target = cloneModel(source)
  const drawingIds = new Map<string, SheetDrawingRef>()
  target.id = targetSheetId
  target.name = targetSheetName
  target.visible = true
  if (target.images) target.images = target.images.map(image => {
    const id = newId('image'); drawingIds.set(drawingRefKey({ kind: 'image', id: image.id }), { kind: 'image', id })
    return { ...cloneModel(image), id }
  })
  if (target.shapes) target.shapes = target.shapes.map(shape => {
    const id = newId('shape'); drawingIds.set(drawingRefKey({ kind: 'shape', id: shape.id }), { kind: 'shape', id })
    return { ...cloneModel(shape), id, sheetId: targetSheetId }
  })
  if (target.metadata && Object.hasOwn(target.metadata, 'hiddenState')) target.metadata.hiddenState = 'visible'
  workbook.sheets.splice(sourceIndex + 1, 0, target)
  workbook.activeSheetId = target.id
  workbook.modified = true

  const tableNames = new Map<string, string>(), tableIds = new Map<string, { id: string; columns: Map<string, string> }>()
  const names = new Set((original.tables ?? []).map(table => table.name.toLowerCase()))
  const copiedTables = (original.tables ?? []).filter(table => table.sheetId === source.id).map(table => {
    const next = cloneModel(table)
    next.id = newId('table'); next.sheetId = target.id; next.name = uniqueResourceName(table.name, names)
    next.columns = next.columns.map(column => ({ ...column, id: newId('table-column') }))
    tableNames.set(table.name.toLowerCase(), next.name)
    tableIds.set(table.id, { id: next.id, columns: new Map(table.columns.map((column, index) => [column.id, next.columns[index]!.id])) })
    return next
  })
  if (copiedTables.length) workbook.tables = [...(workbook.tables ?? []), ...copiedTables]

  const pivotIds = new Map<string, { id: string; fields: Map<string, string> }>()
  const copiedPivots = (original.pivotTables ?? []).filter(pivot => pivot.anchor.subUnitId === source.id).map(pivot => {
    const next = cloneModel(pivot), fields = new Map(pivot.fields.map(field => [field.id, newId('pivot-field')]))
    next.id = newId('pivot'); next.revision = 0
    next.anchor = { ...next.anchor, subUnitId: target.id }
    if (pivot.source.subUnitId === source.id && (!original.id || pivot.source.unitId === original.id)) {
      next.source = { ...next.source, subUnitId: target.id, sheetName: target.name }
    }
    next.fields = next.fields.map(field => ({ ...field, id: fields.get(field.id)!,
      ...(field.valueFilter ? { valueFilter: { ...field.valueFilter, valueFieldId: fields.get(field.valueFilter.valueFieldId) ?? field.valueFilter.valueFieldId } } : {}) }))
    if (next.drillDownMap) next.drillDownMap = Object.fromEntries(Object.entries(next.drillDownMap).map(([key, entry]) => [key, { ...entry,
      valueFieldId: fields.get(entry.valueFieldId) ?? entry.valueFieldId,
      fieldFilters: Object.fromEntries(Object.entries(entry.fieldFilters).map(([id, value]) => [fields.get(id) ?? id, value])),
      ...(entry.fieldMembers ? { fieldMembers: Object.fromEntries(Object.entries(entry.fieldMembers).map(([id, value]) => [fields.get(id) ?? id, value])) } : {}),
    }]))
    pivotIds.set(pivot.id, { id: next.id, fields }); return next
  })
  if (copiedPivots.length) workbook.pivotTables = [...(workbook.pivotTables ?? []), ...copiedPivots]
  const pivotRange = (pivot: PivotTableModel) => pivot.outputRange ?? { startRow: pivot.anchor.row, endRow: pivot.anchor.row, startCol: pivot.anchor.col, endCol: pivot.anchor.col }
  const copiedCharts = (original.charts ?? []).filter(chart => chart.sheetId === source.id).map(chart => {
    const next = cloneModel(chart); next.id = newId('chart'); next.sheetId = target.id; next.revision = 0
    drawingIds.set(drawingRefKey({ kind: 'chart', id: chart.id }), { kind: 'chart', id: next.id })
    if ((chart.source.sheetId ?? chart.sheetId) === source.id) next.source.sheetId = target.id
    if (chart.pivotTableId) {
      next.pivotTableId = pivotIds.get(chart.pivotTableId)?.id ?? chart.pivotTableId
      const pivot = workbook.pivotTables?.find(item => item.id === next.pivotTableId)
      if (!pivot) throw new Error('透视图的数据源不存在，无法完整复制。')
      next.source = { ...next.source, sheetId: pivot.anchor.subUnitId, range: { ...pivotRange(pivot) } }
    }
    return next
  })
  if (copiedCharts.length) workbook.charts = [...(workbook.charts ?? []), ...copiedCharts]
  const slicerNames = new Set((original.slicers ?? []).flatMap(slicer => slicer.name ? [slicer.name.toLowerCase()] : []))
  const copiedSlicers = (original.slicers ?? []).filter(slicer => slicer.hostSheetId === source.id).map(slicer => {
    const next = cloneModel(slicer); next.id = newId('slicer'); next.hostSheetId = target.id; next.revision = 0
    drawingIds.set(drawingRefKey({ kind: slicer.type, id: slicer.id }), { kind: slicer.type, id: next.id })
    if (next.name) next.name = uniqueResourceName(next.name, slicerNames)
    if (slicer.tableId) {
      const table = tableIds.get(slicer.tableId)
      next.tableId = table?.id ?? slicer.tableId
      next.sourceFieldId = table?.columns.get(slicer.sourceFieldId) ?? slicer.sourceFieldId
      if (!workbook.tables?.some(table => table.id === next.tableId)) throw new Error('切片器的表格数据源不存在。')
    } else if (slicer.pivotTableId) {
      const pivot = pivotIds.get(slicer.pivotTableId)
      next.pivotTableId = pivot?.id ?? slicer.pivotTableId
      next.sourceFieldId = pivot?.fields.get(slicer.sourceFieldId) ?? slicer.sourceFieldId
      if (!workbook.pivotTables?.some(pivot => pivot.id === next.pivotTableId)) throw new Error('切片器的透视数据源不存在。')
    } else throw new Error('切片器缺少数据源，无法完整复制。')
    return next
  })
  if (copiedSlicers.length) workbook.slicers = [...(workbook.slicers ?? []), ...copiedSlicers]

  const remap = (formula: string) => remapWorksheetCopyFormula(formula, source.name, target.name, tableNames)
  const move = { sourceSheetId: source.id, targetSheetId: target.id, sourceSheetName: source.name, targetSheetName: target.name,
    sourceRange: fullSheet, targetRange: fullSheet }
  for (const cell of Object.values(target.cells)) {
    // Cell IDs are sheet-local coordinates; only workbook-level resource identities are regenerated.
    cell.id = `${cell.row}:${cell.col}`
    if (cell.image) cell.image = copySheetCellImage(cell.image)
    if (cell.formula) cell.formula = remap(cell.formula)
    if (cell.sparkline) cell.sparkline.sourceRange = remap(cell.sparkline.sourceRange)
    if (cell.hyperlink) cell.hyperlink = moveSheetClipboardHyperlink(cell.hyperlink, original, move)
    if (cell.richText) cell.richText = cell.richText.map(run => typeof run.hyperlink === 'string'
      ? { ...run, hyperlink: moveSheetClipboardHyperlink(run.hyperlink, original, move) } : run)
    if (typeof cell.metadata?.pivotTableId === 'string') {
      const pivot = pivotIds.get(cell.metadata.pivotTableId)
      if (pivot) cell.metadata.pivotTableId = pivot.id
    }
  }
  target.dataValidations = target.dataValidations?.map(rule => ({ ...rule, id: newId('validation'),
    ...(rule.formula1 && (!(rule.type === 'list' || rule.type === 'listMultiple') || rule.formula1.startsWith('=')) ? { formula1: remap(rule.formula1) } : {}),
    ...(rule.formula2 ? { formula2: remap(rule.formula2) } : {}) }))
  target.conditionalFormats = target.conditionalFormats?.map(rule => ({ ...mapConditionalFormatFormulas(rule, remap), id: newId('conditional-format') }))
  target.rangeThemeRules = target.rangeThemeRules?.map(rule => ({ ...rule, id: newId('range-theme') }))
  if (target.outline) target.outline.groups = target.outline.groups.map(group => ({ ...group, id: newId('outline') }))
  target.rangeProtectionRules = target.rangeProtectionRules?.map(rule => ({ ...rule,
    id: newId('range-protection'), permissionId: newId('permission'), ownerUserId: options.userId }))
  if (target.worksheetProtection) target.worksheetProtection = { ...target.worksheetProtection,
    id: newId('worksheet-protection'), permissionId: newId('permission'), ownerUserId: options.userId }
  target.metadata = copyLocalNames(target.metadata, source, target, sourceIndex, remap, false)
  workbook.metadata = copyLocalNames(workbook.metadata, source, target, sourceIndex, remap, true)
  if (drawingIds.size || source.drawingState) {
    const state = cloneModel(source.drawingState ?? createSheetDrawingScene(original, source.id).state)
    for (const group of state.groups) drawingIds.set(drawingRefKey({ kind: 'group', id: group.id }), { kind: 'group', id: newId('drawing-group') })
    const mapped = (ref: SheetDrawingRef): SheetDrawingRef => {
      const target = drawingIds.get(drawingRefKey(ref))
      if (!target) throw new Error('绘图组合的完整对象映射缺失，无法复制工作表。')
      return { ...target }
    }
    target.drawingState = { version: 1, revision: 0, rootOrder: state.rootOrder.map(mapped), groups: state.groups.map(group => ({
      ...cloneModel(group), id: mapped({ kind: 'group', id: group.id }).id, members: group.members.map(mapped),
    })) }
    createSheetDrawingScene(workbook, target.id)
  }
  return { workbook, sourceSheetId: source.id, targetSheetId: target.id, targetSheetName: target.name, index: sourceIndex + 1,
    resources: { tables: copiedTables.length, pivotTables: copiedPivots.length, charts: copiedCharts.length, slicers: copiedSlicers.length } }
}
