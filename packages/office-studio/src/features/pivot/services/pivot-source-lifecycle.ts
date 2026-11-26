import type { CellRange, PivotCacheModel, PivotFieldModel, PivotTableModel, PivotTableSource, WorkbookModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import { equivalentModelData } from '../../../platform/commands/model-equivalence.ts'
import { retainPivotCollapseInfo } from './pivot-axis-state.ts'
import { renderPivotTable } from './pivot-engine.ts'
import { getPivotShowDataAsBaseFieldIds, getPivotShowDataAsRule } from './pivot-show-data-as.ts'
import { retainPivotValuesIndexAfterRemoval } from './pivot-values-position.ts'

export interface PivotSourceCacheTransform {
  sheetId: string
  range: (range: CellRange) => CellRange | undefined
  column?: (column: number, sourceRange: CellRange) => number | undefined
  /** Range cuts may relocate a complete source to another worksheet in the same transaction. */
  sourceWorksheet?: (range: CellRange) => Pick<PivotTableSource, 'subUnitId' | 'sheetName'> | undefined
  /** Data can change even when an internal permutation leaves the range unchanged. */
  sourceDirty?: (range: CellRange) => boolean
}

export interface PivotSourceStructureTransform extends PivotSourceCacheTransform {
  point: (row: number, col: number) => { row: number; col: number } | undefined
  outputRange?: (range: CellRange) => CellRange | undefined
  anchorWorksheet?: (row: number, col: number) => string | undefined
  /** Cell-band shifts cannot silently discard part of a generated output. */
  preserveGeneratedCells?: boolean
}

const sourceKey = (source: PivotTableSource): unknown[] => [source.unitId, source.subUnitId,
  source.range.startRow, source.range.startCol, source.range.endRow, source.range.endCol]
const sameSource = (first: PivotTableSource, second: PivotTableSource): boolean => equivalentModelData(sourceKey(first), sourceKey(second))
const localSource = (workbook: WorkbookModel, source: PivotTableSource, sheetId: string): boolean =>
  source.subUnitId === sheetId && (!workbook.id || source.unitId === workbook.id)
const cacheId = (report: PivotTableModel): string | undefined => report.pivotCacheId ?? report.source.pivotCacheId
const reports = (workbook: WorkbookModel): PivotTableModel[] => [...(workbook.pivotTables ?? []),
  ...(workbook.charts ?? []).flatMap(chart => chart.ownedPivot ? [chart.ownedPivot] : [])]
const inColumns = (column: number, range: CellRange): boolean => column >= range.startCol && column <= range.endCol

/** Freeze provenance before changing physical columns; report-local IDs are never source IDs. */
function fieldIdentity(report: PivotTableModel, field: PivotFieldModel, cache?: PivotCacheModel): string {
  return field.sourceFieldId ?? cache?.fields?.find(source => source.sourceColumn === field.sourceColumn)?.id
    ?? String(field.sourceColumn - report.source.range.startCol)
}

/** Output hit metadata belongs to one generated matrix, never to a source/cache snapshot. */
export function invalidatePivotSourceReportOutput(report: PivotTableModel, owned = false): void {
  delete report.axisOutput
  if (owned) {
    delete report.outputRange; delete report.generatedCellKeys; delete report.drillDownMap
  } else report.drillDownMap = {}
}

/** Structural source deletion removes every instance and its dangling local dependencies atomically. */
function reconcileFieldDependencies(report: PivotTableModel): void {
  const values = new Set(report.fields.filter(field => field.area === 'value').map(field => field.id))
  for (const area of ['row', 'column', 'value', 'filter'] as const) {
    report.fields.filter(field => field.area === area).sort((a, b) => a.position - b.position)
      .forEach((field, position) => { field.position = position })
  }
  for (const field of report.fields) {
    if (field.valueFilter && !values.has(field.valueFilter.valueFieldId)) delete field.valueFilter
    if (field.showDataAs) {
      const rule = getPivotShowDataAsRule(field.showDataAs.type)
      if (rule.requiresParameters && !getPivotShowDataAsBaseFieldIds(report, rule.type).includes(field.showDataAs.baseFieldId)) delete field.showDataAs
    }
  }
  const collapseInfo = retainPivotCollapseInfo(report, report.collapseInfo)
  if (Object.keys(collapseInfo).length) report.collapseInfo = collapseInfo
  else delete report.collapseInfo
}

function assertSourceRange(range: CellRange | undefined): asserts range is CellRange {
  if (!range || range.endRow <= range.startRow || range.endCol < range.startCol) throw new RangeError('透视数据源需要保留表头、字段和记录行。')
}

/** The caller supplies one coordinate mapping for reports, retained caches and Slicer resources. */
export function transformPivotSourceReport(before: WorkbookModel, report: PivotTableModel,
  transform: PivotSourceStructureTransform, owned = false): PivotTableModel {
  const next = cloneModel(report)
  let sourceChanged = false
  if (localSource(before, report.source, transform.sheetId)) {
    const range = transform.range(report.source.range)
    assertSourceRange(range)
    const retained = before.pivotCaches?.find(cache => cache.id === cacheId(report))
    const fields = report.fields.flatMap(field => {
      const column = transform.column ? transform.column(field.sourceColumn, report.source.range) : field.sourceColumn
      return column === undefined || !inColumns(column, range) ? [] : [{ ...cloneModel(field), sourceColumn: column }]
    })
    const worksheet = transform.sourceWorksheet?.(report.source.range)
    sourceChanged = transform.sourceDirty?.(report.source.range) === true || !equivalentModelData(range, report.source.range)
      || !equivalentModelData(fields, report.fields) || !!worksheet
    if (sourceChanged) {
      next.source = { ...next.source, range, ...worksheet }
      next.fields = fields.map(field => ({ ...field, sourceFieldId: fieldIdentity(report, report.fields.find(previous => previous.id === field.id)!, retained) }))
      reconcileFieldDependencies(next)
      retainPivotValuesIndexAfterRemoval(next, report.fields)
    }
  }
  if (!owned && report.anchor.subUnitId === transform.sheetId && (!before.id || report.anchor.unitId === before.id)) {
    const anchor = transform.point(report.anchor.row, report.anchor.col)
    if (!anchor) throw new RangeError('结构编辑会移除透视输出起点，请先移动或删除透视对象。')
    Object.assign(next.anchor, anchor)
    const worksheet = transform.anchorWorksheet?.(report.anchor.row, report.anchor.col)
    if (worksheet) next.anchor.subUnitId = worksheet
    if (report.outputRange) {
      const range = (transform.outputRange ?? transform.range)(report.outputRange)
      if (!range) throw new RangeError('结构编辑不能删除整个透视输出。')
      next.outputRange = range
    }
    if (report.generatedCellKeys) next.generatedCellKeys = report.generatedCellKeys.flatMap(key => {
      const match = /^(\d+):(\d+)$/.exec(key)
      if (!match) throw new TypeError('透视输出坐标无效。')
      const point = transform.point(Number(match[1]), Number(match[2]))
      if (!point && transform.preserveGeneratedCells) throw new RangeError('局部移位不能删除透视生成格。')
      return point ? [`${point.row}:${point.col}`] : []
    })
  }
  if (sourceChanged || !equivalentModelData(next, report)) {
    invalidatePivotSourceReportOutput(next, owned)
    next.revision = report.revision + 1
  }
  return next
}

export function transformWorkbookPivotSourceReports(before: WorkbookModel, candidate: WorkbookModel,
  transform: PivotSourceStructureTransform): void {
  if (before.pivotTables) candidate.pivotTables = before.pivotTables.map(report => transformPivotSourceReport(before, report, transform))
  const previousCharts = new Map((before.charts ?? []).map(chart => [chart.id, chart]))
  for (const chart of candidate.charts ?? []) {
    const previous = previousCharts.get(chart.id)
    if (previous?.ownedPivot) chart.ownedPivot = transformPivotSourceReport(before, previous.ownedPivot, transform, true)
  }
}

/** Cache-scoped canonical fields are transformed once, including explicitly disconnected sources. */
export function reconcilePivotSourceCaches(before: WorkbookModel, candidate: WorkbookModel,
  transform?: PivotSourceCacheTransform): Map<string, Set<string>> {
  candidate.pivotCaches ??= cloneModel(before.pivotCaches ?? [])
  const currentReports = reports(candidate), previousReports = reports(before), removedSourceFields = new Map<string, Set<string>>()
  for (const cache of candidate.pivotCaches) {
    const original = before.pivotCaches?.find(previous => previous.id === cache.id), baseline = cloneModel(cache)
    let dirty = false
    if (original && transform && localSource(before, original.source, transform.sheetId)) {
      const range = transform.range(original.source.range)
      assertSourceRange(range)
      const removed = new Set<string>()
      cache.source = { ...cloneModel(original.source), range, ...transform.sourceWorksheet?.(original.source.range) }
      cache.fields = original.fields?.flatMap(field => {
        const column = transform.column ? transform.column(field.sourceColumn, original.source.range) : field.sourceColumn
        if (column === undefined || !inColumns(column, range)) { removed.add(field.id); return [] }
        return [{ ...cloneModel(field), sourceColumn: column }]
      })
      if (cache.fields) {
        const columns = new Set(cache.fields.map(field => field.sourceColumn)), ids = new Set(cache.fields.map(field => field.id))
        let sequence = original.nextFieldId ?? 0
        for (let column = range.startCol; column <= range.endCol; column++) {
          if (columns.has(column)) continue
          let id = `source-field:${++sequence}`
          while (ids.has(id)) id = `source-field:${++sequence}`
          const header = candidate.sheets.find(sheet => sheet.id === cache.source.subUnitId)?.cells[`${range.startRow}:${column}`]?.value
          cache.fields.push({ id, sourceColumn: column, name: header == null ? undefined : String(header) })
          ids.add(id); columns.add(column)
        }
        cache.fields.sort((a, b) => a.sourceColumn - b.sourceColumn)
        cache.nextFieldId = sequence
      }
      if (removed.size) removedSourceFields.set(cache.id, removed)
      dirty = transform.sourceDirty?.(original.source.range) === true
    }
    const owners = currentReports.filter(report => cacheId(report) === cache.id)
    const previousOwners = previousReports.filter(report => cacheId(report) === cache.id)
    const coherentBefore = !original || previousOwners.every(report => sameSource(original.source, report.source))
    if (owners.length && coherentBefore) {
      if (new Set(owners.map(report => JSON.stringify(sourceKey(report.source)))).size > 1) throw new TypeError('结构编辑后同一透视缓存出现不同数据源。')
      if (!sameSource(cache.source, owners[0]!.source)) {
        const delta = owners[0]!.source.range.startCol - cache.source.range.startCol
        cache.fields = cache.fields?.map(field => ({ ...field, sourceColumn: field.sourceColumn + delta }))
      }
      cache.source = cloneModel(owners[0]!.source)
      for (const report of owners) {
        const previous = previousOwners.find(owner => owner.id === report.id)
        if (!previous) continue
        const removed = removedSourceFields.get(cache.id)
        if (removed?.size) {
          const previousFields = report.fields
          report.fields = report.fields.filter(field => !removed.has(fieldIdentity(report, field, cache)))
          reconcileFieldDependencies(report)
          retainPivotValuesIndexAfterRemoval(report, previousFields)
        }
        if (transform && (dirty || !equivalentModelData(report.fields, previous.fields) || !equivalentModelData(report.source, previous.source)
          || !equivalentModelData(report.collapseInfo, previous.collapseInfo))) {
          invalidatePivotSourceReportOutput(report, candidate.charts?.some(chart => chart.ownedPivot === report) === true)
          report.revision = Math.max(report.revision, previous.revision + 1)
        }
      }
    }
    if (dirty || !equivalentModelData(cache, baseline)) cache.revision = Math.max(cache.revision, (original?.revision ?? baseline.revision) + 1)
  }
  return removedSourceFields
}

export function getPivotSourceAvailability(workbook: WorkbookModel, report: PivotTableModel): 'available' | 'unit-missing' | 'worksheet-missing' | 'cache-mismatch' {
  if (report.source.unitId !== workbook.id) return 'unit-missing'
  if (!workbook.sheets.some(sheet => sheet.id === report.source.subUnitId)) return 'worksheet-missing'
  const cache = workbook.pivotCaches?.find(cache => cache.id === cacheId(report))
  return cache && !sameSource(cache.source, report.source) ? 'cache-mismatch' : 'available'
}

/** Publish fresh worksheet output through the existing sole aggregator after resource mappings settle. */
export function refreshPivotSourceWorksheetOutputs(before: WorkbookModel, candidate: WorkbookModel, copiedIds: readonly string[] = []): void {
  const previous = new Map((before.pivotTables ?? []).map(report => [report.id, report])), copied = new Set(copiedIds)
  for (const report of candidate.pivotTables ?? []) {
    const old = previous.get(report.id)
    if (!copied.has(report.id) && (!old || report.revision === old.revision)) continue
    if (getPivotSourceAvailability(candidate, report) !== 'available') continue
    const hostIndex = candidate.sheets.findIndex(sheet => sheet.id === report.anchor.subUnitId)
    if (hostIndex >= 0 && candidate.sheets[hostIndex] === before.sheets.find(sheet => sheet.id === report.anchor.subUnitId)) {
      candidate.sheets[hostIndex] = cloneModel(candidate.sheets[hostIndex]!)
    }
    renderPivotTable(candidate, report)
  }
}

/** Copies own a new report identity; source/cache provenance and local hierarchy keys remain separate. */
export function prepareCopiedPivotSourceReport(report: PivotTableModel, owned = false): void {
  reconcileFieldDependencies(report)
  invalidatePivotSourceReportOutput(report, owned)
}

/** Sheet names are descriptive metadata; the source sheet ID and cache ID remain authoritative. */
export function renamePivotSourceWorksheet(before: WorkbookModel, candidate: WorkbookModel, sheetId: string, name: string): void {
  for (const report of reports(candidate)) if (localSource(candidate, report.source, sheetId) && report.source.sheetName !== name) {
    const previous = reports(before).find(previous => previous.id === report.id)
    report.source.sheetName = name
    report.revision = Math.max(report.revision, (previous?.revision ?? report.revision) + 1)
  }
  for (const cache of candidate.pivotCaches ?? []) if (localSource(candidate, cache.source, sheetId) && cache.source.sheetName !== name) {
    const previous = before.pivotCaches?.find(previous => previous.id === cache.id)
    cache.source.sheetName = name
    cache.revision = Math.max(cache.revision, (previous?.revision ?? cache.revision) + 1)
  }
}

/** A removed source never silently rebinds to another sheet with the same name. Retain its cache reference. */
export function invalidateRemovedPivotSourceWorksheet(before: WorkbookModel, candidate: WorkbookModel, sheetId: string): void {
  if (!before.sheets.some(sheet => sheet.id === sheetId) || candidate.sheets.some(sheet => sheet.id === sheetId)) return
  for (const report of reports(candidate)) if (localSource(before, report.source, sheetId)) {
    const previous = reports(before).find(previous => previous.id === report.id)
    invalidatePivotSourceReportOutput(report, candidate.charts?.some(chart => chart.ownedPivot === report) === true)
    report.revision = Math.max(report.revision, (previous?.revision ?? report.revision) + 1)
  }
  for (const cache of candidate.pivotCaches ?? []) if (localSource(before, cache.source, sheetId)) {
    const previous = before.pivotCaches?.find(previous => previous.id === cache.id)
    cache.revision = Math.max(cache.revision, (previous?.revision ?? cache.revision) + 1)
  }
}

/** Clipboard content changes supply their real source/target scopes instead of inventing coordinate moves. */
export function invalidatePivotSourceContentRanges(before: WorkbookModel, candidate: WorkbookModel,
  scopes: readonly { sheetId: string; range: CellRange }[]): void {
  const changes: Array<{ sheetId: string; row: number; col: number }> = [], seen = new Set<string>()
  for (const scope of scopes) {
    const old = before.sheets.find(sheet => sheet.id === scope.sheetId), current = candidate.sheets.find(sheet => sheet.id === scope.sheetId)
    if (!old || !current || old === current) continue
    for (const key of new Set([...Object.keys(old.cells), ...Object.keys(current.cells)])) {
      const cell = current.cells[key] ?? old.cells[key]!
      if (cell.row < scope.range.startRow || cell.row > scope.range.endRow || cell.col < scope.range.startCol || cell.col > scope.range.endCol
        || Object.is(old.cells[key]?.value, current.cells[key]?.value) && old.cells[key]?.formula === current.cells[key]?.formula) continue
      const id = JSON.stringify([scope.sheetId, key])
      if (!seen.has(id)) { seen.add(id); changes.push({ sheetId: scope.sheetId, row: cell.row, col: cell.col }) }
    }
  }
  const affected = (source: PivotTableSource): boolean => source.unitId === candidate.id && changes.some(change =>
    change.sheetId === source.subUnitId && change.row >= source.range.startRow && change.row <= source.range.endRow
      && change.col >= source.range.startCol && change.col <= source.range.endCol)
  const previousReports = reports(before)
  for (const report of reports(candidate)) if (affected(report.source)) {
    const previous = previousReports.find(previous => previous.id === report.id)
    if (!previous) continue
    invalidatePivotSourceReportOutput(report, candidate.charts?.some(chart => chart.ownedPivot === report) === true)
    report.revision = Math.max(report.revision, previous.revision + 1)
  }
  for (const cache of candidate.pivotCaches ?? []) if (affected(cache.source)) {
    const previous = before.pivotCaches?.find(previous => previous.id === cache.id)
    if (previous) cache.revision = Math.max(cache.revision, previous.revision + 1)
  }
}
