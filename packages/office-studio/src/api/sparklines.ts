import type { ApiRange } from './events.ts'
import type { FWorksheet } from './facade.ts'
import type { OfficeStudio } from './office-studio.ts'
import type { SparklineGroupInfo } from '../features/sparkline/sparkline-model.ts'
import { readSparklineGroups } from '../features/sparkline/sparkline-model.ts'
import { sparklineConfigFromPublic, sparklineConfigToPublic, type ISparklineGroup, type ISparklineGroupConfig } from '../features/sparkline/sparkline-api-types.ts'
import { SparklineRangeMatrix } from '../features/sparkline/sparkline-range-matrix.ts'
import { parseSparklineReference, sparklineCellAddress, sparklineRangeString } from '../features/sparkline/sparkline-reference.ts'

export function sparklineApiRangeReference(studio: OfficeStudio, unitId: string, sheetId: string, range: ApiRange, target = false): string {
  if (!range || typeof range !== 'object' || Array.isArray(range)
    || ![range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isSafeInteger)
    || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn
    || range.endRow > 1_048_575 || range.endColumn > 16_383 || range.unitId !== undefined && range.unitId !== unitId) {
    throw new RangeError('迷你图 API 区域无效或不属于当前工作簿。')
  }
  const workbook = studio.getWorkbookModel(unitId), sourceId = range.sheetId ?? sheetId
  const source = workbook?.sheets.find(sheet => sheet.id === sourceId)
  if (!source || target && sourceId !== sheetId) throw new Error('迷你图目标必须属于当前工作表，来源必须属于当前工作簿。')
  return sparklineRangeString({ startRow: range.startRow, startCol: range.startColumn, endRow: range.endRow, endCol: range.endColumn }, sourceId === sheetId ? undefined : source.name)
}

/** Group snapshots are detached; editing the matrix never bypasses the command transaction. */
export function sparklineGroupToPublic(group: SparklineGroupInfo, resolveSheetId?: (name: string) => string | undefined): ISparklineGroup {
  const sparklines = new SparklineRangeMatrix<ApiRange>()
  for (const member of group.members) {
    const target = parseSparklineReference(member.address), source = parseSparklineReference(member.sourceRange)
    if (!target || !source) throw new TypeError('迷你图组包含失效的区域引用。')
    sparklines.setValue(target.range.startRow, target.range.startCol, {
      startRow: source.range.startRow, startColumn: source.range.startCol, endRow: source.range.endRow, endColumn: source.range.endCol,
      ...(source.sheetName === undefined ? {} : { sheetId: resolveSheetId?.(source.sheetName) ?? source.sheetName }),
    })
  }
  return { config: sparklineConfigToPublic(group.settings), sparklines }
}

class SparklineHandleOwner {
  private readonly generation: number
  constructor(protected readonly studio: OfficeStudio, protected readonly worksheet: FWorksheet, protected readonly groupId: string) {
    this.generation = studio.editor.history.captureSavepoint(worksheet.unitId).generation
  }
  protected currentGroup(): SparklineGroupInfo {
    if (this.studio.isDisposed || !this.studio.getWorkbookModel(this.worksheet.unitId)
      || this.studio.editor.history.captureSavepoint(this.worksheet.unitId).generation !== this.generation) throw new Error('迷你图 API 所属工作簿已关闭或被替换。')
    const sheet = this.studio.getWorkbookModel(this.worksheet.unitId)?.sheets.find(sheet => sheet.id === this.worksheet.id)
    const found = sheet && readSparklineGroups(sheet).find(group => group.groupId === this.groupId)
    const group = found && this.worksheet.getSparklineGroups(found.targetRanges).find(group => group.groupId === this.groupId)
    if (!group) throw new Error('迷你图组已删除或重新组合，请重新获取。')
    return group
  }
  protected sourceReferences(ranges: readonly ApiRange[], target = false): string[] {
    if (!Array.isArray(ranges) || !ranges.length) throw new TypeError('请指定迷你图源数据和放置区域。')
    return ranges.map(range => sparklineApiRangeReference(this.studio, this.worksheet.unitId, this.worksheet.id, range, target))
  }
}

export class FSparkline extends SparklineHandleOwner {
  constructor(studio: OfficeStudio, worksheet: FWorksheet, groupId: string, private row: number, private column: number) { super(studio, worksheet, groupId) }
  changeDataSource(sourceRange: ApiRange, targetRange: ApiRange): this | undefined {
    const group = this.currentGroup(), address = sparklineCellAddress(this.row, this.column)
    if (!group.members.some(member => member.address === address)) throw new Error('原单元格的迷你图已移动或重新组合。')
    const sources = this.sourceReferences([sourceRange]), targets = this.sourceReferences([targetRange], true)
    const target = parseSparklineReference(targets[0]!)!
    if (target.range.startRow !== target.range.endRow || target.range.startCol !== target.range.endCol) throw new RangeError('单个迷你图必须放置在一个单元格中。')
    this.worksheet.applySparklineMutation({ kind: 'change-source', resetType: 'item', ranges: [address], primary: address, sourceRanges: sources, targetRanges: targets })
    this.row = target.range.startRow; this.column = target.range.startCol
    return this.currentGroup().members.some(member => member.address === targets[0]) ? this : undefined
  }
  removeSparkline(): void {
    const group = this.currentGroup(), address = sparklineCellAddress(this.row, this.column)
    if (!group.members.some(member => member.address === address)) throw new Error('原单元格的迷你图已移动或重新组合。')
    this.worksheet.applySparklineMutation({ kind: 'clear', scope: 'item', ranges: [address] })
  }
}

export class FSparklineGroup extends SparklineHandleOwner {
  changeDataSource(sourceRanges: ApiRange[], targetRanges: ApiRange[]): this | undefined {
    const group = this.currentGroup()
    this.worksheet.applySparklineMutation({ kind: 'change-source', resetType: 'group', ranges: group.targetRanges,
      sourceRanges: this.sourceReferences(sourceRanges), targetRanges: this.sourceReferences(targetRanges, true) })
    return this.currentGroup().members.length ? this : undefined
  }
  removeSparklineGroup(): void { this.worksheet.applySparklineMutation({ kind: 'clear', scope: 'group', ranges: this.currentGroup().targetRanges }) }
  setConfig(config: ISparklineGroupConfig): this {
    const group = this.currentGroup()
    this.worksheet.applySparklineMutation({ kind: 'settings', ranges: group.targetRanges, settings: sparklineConfigFromPublic(config) })
    return this
  }
}
