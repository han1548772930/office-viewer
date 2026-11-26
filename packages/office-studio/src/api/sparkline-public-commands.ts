import type { ApiRange } from './events.ts'
import type { OfficeStudio } from './office-studio.ts'
import { sparklineApiRangeReference } from './sparklines.ts'
import { SET_SPARKLINES_COMMAND, type SparklineMutationContext } from '../features/sparkline/sparkline-command.ts'
import { readSparklineGroups, type SparklineMutationRequest } from '../features/sparkline/sparkline-model.ts'
import { sparklineConfigFromPublic, type ISparklineGroupConfig } from '../features/sparkline/sparkline-api-types.ts'
import { sparklineCellAddress } from '../features/sparkline/sparkline-reference.ts'

export const SPARKLINE_PUBLIC_COMMANDS = Object.freeze({ add: 'sheet.command.add-sparkline', set: 'sheet.command.set-sparkline', remove: 'sheet.command.remove-sparkline' } as const)
type TargetInfo = { unitId: string; subUnitId: string }
interface PublicSparklineBase extends SparklineMutationContext { targetInfo?: TargetInfo; ranges?: ApiRange[] }
export interface AddSheetSparklineCommandInput extends PublicSparklineBase { sourceRanges: ApiRange[]; targetRanges: ApiRange[]; config?: ISparklineGroupConfig }
export interface RemoveSheetSparklineCommandInput extends PublicSparklineBase { isSingle?: boolean }
export interface SetSheetSparklineCommandInput extends PublicSparklineBase {
  config: { config: ISparklineGroupConfig }
  isChangeDataSource: boolean
  combine?: boolean
  unCombine?: boolean
  changeDataSourceInfo?: { sourceRanges: ApiRange[]; targetRanges: ApiRange[]; groupId: string; resetType: 'item' | 'group'; primary: { actualRow?: number; actualColumn?: number; startRow?: number; startColumn?: number } }
}
type PublicInput = AddSheetSparklineCommandInput | RemoveSheetSparklineCommandInput | SetSheetSparklineCommandInput

export function registerPublicSparklineCommands(studio: OfficeStudio): void {
  for (const id of Object.values(SPARKLINE_PUBLIC_COMMANDS)) studio.registerCommand<PublicInput, boolean>(id, raw => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new TypeError('迷你图公开命令参数无效。')
    const unitId = raw.targetInfo?.unitId ?? studio.activeWorkbookId, workbook = studio.getWorkbookModel(unitId)
    const sheetId = raw.targetInfo?.subUnitId ?? workbook?.activeSheetId, sheet = workbook?.sheets.find(sheet => sheet.id === sheetId)
    if (!workbook || !sheet || raw.targetInfo !== undefined && (!raw.targetInfo || typeof raw.targetInfo !== 'object'
      || typeof raw.targetInfo.unitId !== 'string' || typeof raw.targetInfo.subUnitId !== 'string')) throw new Error('迷你图公开命令目标已关闭或无效。')
    const ranges = raw.ranges ?? studio.getSelectionState(unitId, sheet.id)?.ranges ?? []
    const references = (values: readonly ApiRange[], target = false): string[] => {
      if (!Array.isArray(values) || !values.length) throw new TypeError('迷你图公开命令须提供有效区域数组。')
      return values.map(range => sparklineApiRangeReference(studio, unitId, sheet.id, range, target))
    }
    let request: SparklineMutationRequest
    if (id === SPARKLINE_PUBLIC_COMMANDS.add) {
      if (!('sourceRanges' in raw) || !('targetRanges' in raw)) throw new TypeError('插入迷你图缺少源数据或放置区域。')
      request = { kind: 'insert', sourceRanges: references(raw.sourceRanges), targetRanges: references(raw.targetRanges, true), settings: sparklineConfigFromPublic(raw.config ?? {}) }
    } else if (id === SPARKLINE_PUBLIC_COMMANDS.remove) {
      if ('isSingle' in raw && raw.isSingle !== undefined && typeof raw.isSingle !== 'boolean') throw new TypeError('迷你图删除模式无效。')
      request = { kind: 'clear', ranges: references(ranges, true), scope: 'isSingle' in raw && raw.isSingle === false ? 'group' : 'item' }
    } else {
      if (!('isChangeDataSource' in raw) || typeof raw.isChangeDataSource !== 'boolean' || !raw.config || typeof raw.config !== 'object') throw new TypeError('迷你图设置命令缺少配置。')
      if (raw.combine !== undefined && typeof raw.combine !== 'boolean' || raw.unCombine !== undefined && typeof raw.unCombine !== 'boolean'
        || [raw.isChangeDataSource, raw.combine === true, raw.unCombine === true].filter(Boolean).length > 1) throw new TypeError('迷你图设置模式必须互斥。')
      const selected = references(ranges, true), settings = sparklineConfigFromPublic(raw.config.config)
      if (raw.isChangeDataSource) {
        const source = raw.changeDataSourceInfo
        if (!source || typeof source !== 'object' || typeof source.groupId !== 'string' || !source.primary || typeof source.primary !== 'object') throw new TypeError('迷你图改源信息无效。')
        const primary = sparklineCellAddress(source.primary.actualRow ?? source.primary.startRow ?? -1, source.primary.actualColumn ?? source.primary.startColumn ?? -1)
        if (!readSparklineGroups(sheet, [primary]).some(group => group.groupId === source.groupId)) throw new Error('迷你图改源信息属于已改变的分组。')
        request = { kind: 'change-source', ranges: selected, resetType: source.resetType, sourceRanges: references(source.sourceRanges), targetRanges: references(source.targetRanges, true), primary, settings }
      } else if (raw.combine) request = { kind: 'combine', ranges: selected, settings }
      else if (raw.unCombine) request = { kind: 'uncombine', ranges: selected }
      else request = { kind: 'settings', ranges: selected, settings }
    }
    return studio.executeCommandSync(SET_SPARKLINES_COMMAND, { unitId, sheetId: sheet.id, request, host: raw.host,
      expectedRevision: raw.expectedRevision ?? studio.getWorkbookRevision(unitId), historyGeneration: raw.historyGeneration ?? studio.editor.history.captureSavepoint(unitId).generation,
      expectedUserId: raw.expectedUserId ?? studio.currentUserId })
  })
}
