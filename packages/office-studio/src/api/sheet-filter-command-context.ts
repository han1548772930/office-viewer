import type { OfficeStudio } from './office-studio.ts'
import type { CellRange } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { workbookCommandGuard } from './workbook-command-guards.ts'

interface FilterCommandTarget { unitId: string; sheetId: string; expectedRevision?: number }

/** Filter configuration owns the source range and complete intersecting merges.
 * The same check runs during preparation and immediately before publication. */
export function sheetFilterCommandContext(api: OfficeStudio, input: FilterCommandTarget, nextRange?: CellRange | null) {
  const guard = workbookCommandGuard(input)
  guard?.()
  if (!input || typeof input.unitId !== 'string' || !input.unitId.trim() || typeof input.sheetId !== 'string' || !input.sheetId.trim()) throw new TypeError('筛选目标工作簿或工作表无效。')
  const unitId = input.unitId, sheetId = input.sheetId, model = api.getWorkbookModel(unitId), sheet = model?.sheets.find(sheet => sheet.id === sheetId)
  if (!model || !sheet) throw new Error('筛选目标工作簿或工作表已关闭。')
  const generation = api.editor.history.captureSavepoint(unitId).generation, revision = api.getWorkbookRevision(unitId), userId = api.currentUserId
  if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0 || input.expectedRevision !== revision)) throw new Error('数据已更改，请重新读取数据区域后再应用筛选。')
  const ranges = [sheet.autoFilter?.range, nextRange].filter((range): range is CellRange => !!range).map(range => cloneModel(range))
  const signature = JSON.stringify(input)
  const assertCurrent = (): void => {
    guard?.()
    if (api.isDisposed || api.getWorkbookModel(unitId) !== model || api.getWorkbookRevision(unitId) !== revision
      || api.editor.history.captureSavepoint(unitId).generation !== generation || api.currentUserId !== userId || JSON.stringify(input) !== signature) throw new Error('筛选目标、用户、参数或内容版本已更改。')
    api.assertCanViewWorksheet(unitId, sheetId); api.assertCanEditWorksheet(unitId, sheetId)
    for (const range of ranges) api.assertCanEditRange(unitId, sheetId, range)
  }
  assertCurrent()
  return { assertCurrent, expectedRevision: revision, expectedHistoryGeneration: generation }
}
