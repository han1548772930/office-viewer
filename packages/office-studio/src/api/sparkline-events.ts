import type { OfficeStudio } from './office-studio.ts'
import { FWorkbook, FWorksheet } from './facade.ts'
import type { WorkbookModel, SpreadsheetSheet } from '../domain/model.ts'
import { equivalentModelData } from '../platform/commands/model-equivalence.ts'
import { readSparklineGroups } from '../features/sparkline/sparkline-model.ts'
import { SparklineRangeMatrix } from '../features/sparkline/sparkline-range-matrix.ts'
import { sparklineGroupToPublic } from './sparklines.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../features/sheet-protection/range-protection.ts'
import { parseSparklineReference } from '../features/sparkline/sparkline-reference.ts'
import type { ApiRange } from './events.ts'
import { createSparklineSourceChangeDetector } from '../features/sparkline/sparkline-source-changes.ts'
import { resolveSparklineSource } from '../features/sparkline/sparkline-source.ts'
import type { SparklineGroupInfo } from '../features/sparkline/sparkline-model.ts'

/** Publish after the content history settles, for commands and Undo/Redo alike. */
export function publishWorkbookSparklineChanges(studio: OfficeStudio, before: WorkbookModel, after: WorkbookModel): void {
  const unitId = after.id
  if (!unitId || !studio.events.hasListeners('SheetSparklineChanged')) return
  const revision = studio.getWorkbookRevision(unitId), generation = studio.editor.history.captureSavepoint(unitId).generation, userId = studio.currentUserId
  const current = () => !studio.isDisposed && !!studio.getWorkbookModel(unitId) && studio.getWorkbookRevision(unitId) === revision
    && studio.editor.history.captureSavepoint(unitId).generation === generation && studio.currentUserId === userId
  const sourceChanged = createSparklineSourceChangeDetector(before, after)
  for (const sheet of after.sheets) {
    if (!current()) break
    try {
    const oldSheet = before.sheets.find(previous => previous.id === sheet.id), beforeGroups = oldSheet ? readSparklineGroups(oldSheet) : [], afterGroups = readSparklineGroups(sheet)
    const previous = new Map(beforeGroups.map(group => [group.groupId, group])), next = new Map(afterGroups.map(group => [group.groupId, group]))
    const changedIds = [...new Set([...previous.keys(), ...next.keys()])].filter(id => !equivalentModelData(previous.get(id), next.get(id))
      || next.has(id) && sourceChanged(sheet.id, next.get(id)!))
    if (!changedIds.length) continue
    const visible = (target: SpreadsheetSheet, addresses: string[]) => canViewWorksheetProtection(target, userId)
      && canViewProtectedRanges(target, addresses.map(address => parseSparklineReference(address)!.range), userId)
    const sourcesVisible = (workbook: WorkbookModel, host: SpreadsheetSheet, group: SparklineGroupInfo) => group.members.every(member => {
      const target = parseSparklineReference(member.address)!, spec = host.cells[`${target.range.startRow}:${target.range.startCol}`]?.sparkline
      const source = resolveSparklineSource({ sourceRange: member.sourceRange, sourceReference: spec?.sourceReference }, host, workbook)
      return source.status !== 'resolved' || canViewWorksheetProtection(source.sheet, userId) && canViewProtectedRanges(source.sheet, [source.range], userId)
    })
    const readableIds = changedIds.filter(id => (!previous.get(id) || oldSheet && visible(oldSheet, previous.get(id)!.targetRanges)
      && sourcesVisible(before, oldSheet, previous.get(id)!))
      && visible(sheet, next.get(id)?.targetRanges ?? previous.get(id)!.targetRanges)
      && (!next.get(id) || sourcesVisible(after, sheet, next.get(id)!)))
    if (!readableIds.length || !current()) continue
    studio.publishObserverEvent('SheetSparklineChanged', () => ({
      workbook: new FWorkbook(studio, unitId), worksheet: new FWorksheet(studio, unitId, sheet.id),
      sparklines: readableIds.map(id => {
        const group = next.get(id) ?? previous.get(id)!, data = sparklineGroupToPublic(group,
          name => after.sheets.find(sheet => sheet.id === name || sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())?.id,
          { host: sheet, workbook: after })
        return next.has(id) ? data : { config: data.config, sparklines: new SparklineRangeMatrix<ApiRange>() }
      }),
    }), current)
    } catch (error) { studio.reportObserverError('SheetSparklineChanged', error) }
  }
}
