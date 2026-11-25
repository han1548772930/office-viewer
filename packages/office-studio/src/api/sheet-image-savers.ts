import type { OfficeStudio } from './office-studio.ts'
import type { SheetImageSaveConfig, SheetImageSaveRequest } from '../features/sheet-images/sheet-image-save.ts'
import type { SheetImageCommandInput } from '../features/sheet-images/sheet-image-commands.ts'
export type SheetImageSaveAction = (config: SheetImageSaveConfig, signal?: AbortSignal) => Promise<boolean>
export type SheetImageSaver = (request: SheetImageSaveRequest, input: SheetImageCommandInput, save: SheetImageSaveAction) => Promise<boolean>
interface Owner { save: SheetImageSaver }
const owners = new WeakMap<OfficeStudio, Map<string, Owner>>()
export function bindSheetImageSaver(api: OfficeStudio, unitId: string, save: SheetImageSaver) {
  if (api.isDisposed || !api.getWorkbookModel(unitId)) throw new Error('找不到图片保存器的工作簿。')
  const units = owners.get(api) ?? new Map<string, Owner>(), owner = { save }
  owners.set(api, units); units.set(unitId, owner)
  return { dispose: () => { if (units.get(unitId) === owner) units.delete(unitId); if (!units.size && owners.get(api) === units) owners.delete(api) } }
}
export async function requestSheetImageSaver(api: OfficeStudio, unitId: string, request: SheetImageSaveRequest, input: SheetImageCommandInput, save: SheetImageSaveAction) {
  const owner = owners.get(api)?.get(unitId)
  if (!owner || api.isDisposed) return false
  const saved = await owner.save(request, input, save)
  return !api.isDisposed && owners.get(api)?.get(unitId) === owner && saved
}
export function clearSheetImageSavers(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) owners.delete(api)
  else { const units = owners.get(api); units?.delete(unitId); if (!units?.size) owners.delete(api) }
}
