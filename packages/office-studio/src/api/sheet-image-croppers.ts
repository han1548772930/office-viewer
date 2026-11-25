import type { OfficeStudio } from './office-studio.ts'
import type { SheetImageCommandInput } from '../features/sheet-images/sheet-image-commands.ts'
import type { SheetImageCropOperation, SheetImageCropResult } from '../features/sheet-images/sheet-image-crop-session.ts'

interface Owner { operate(id: SheetImageCropOperation, input: SheetImageCommandInput): SheetImageCropResult }
const owners = new WeakMap<OfficeStudio, Map<string, Owner>>()
export function bindSheetImageCropper(api: OfficeStudio, unitId: string, operate: Owner['operate']) {
  if (api.isDisposed || !api.getWorkbookModel(unitId)) throw new Error('找不到图片裁剪器的工作簿。')
  const units = owners.get(api) ?? new Map<string, Owner>(), owner = { operate }
  owners.set(api, units); units.set(unitId, owner)
  return { dispose: () => { if (units.get(unitId) === owner) units.delete(unitId); if (!units.size && owners.get(api) === units) owners.delete(api) } }
}
export function requestSheetImageCropper(api: OfficeStudio, unitId: string, id: SheetImageCropOperation, input: SheetImageCommandInput): SheetImageCropResult {
  return api.isDisposed ? { handled: false } : owners.get(api)?.get(unitId)?.operate(id, input) ?? { handled: false }
}
export function clearSheetImageCroppers(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) owners.delete(api)
  else { const units = owners.get(api); units?.delete(unitId); if (!units?.size) owners.delete(api) }
}
