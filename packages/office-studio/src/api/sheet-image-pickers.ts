import type { OfficeStudio } from './office-studio.ts'
import type { SheetImageResource } from '../domain/sheet-images.ts'
import type { SheetImageCommandId, SheetImageCommandInput } from '../features/sheet-images/sheet-image-commands.ts'

export type SheetImagePicker = (commandId: SheetImageCommandId, input: SheetImageCommandInput) => Promise<readonly SheetImageResource[] | null>
interface PickerOwner { pick: SheetImagePicker }
const owners = new WeakMap<OfficeStudio, Map<string, PickerOwner>>()
/** Runtime callbacks are never saved; replacement owners cannot be removed by an old handle. */
export function bindSheetImagePicker(api: OfficeStudio, unitId: string, pick: SheetImagePicker) {
  if (api.isDisposed || !api.getWorkbookModel(unitId)) throw new Error('找不到图片选择器的工作簿。')
  const units = owners.get(api) ?? new Map<string, PickerOwner>(), owner = { pick }
  owners.set(api, units); units.set(unitId, owner)
  return { isCurrent: () => !api.isDisposed && units.get(unitId) === owner,
    dispose: () => { if (units.get(unitId) === owner) units.delete(unitId); if (!units.size && owners.get(api) === units) owners.delete(api) } }
}
export async function requestSheetImagePicker(api: OfficeStudio, unitId: string, id: SheetImageCommandId, input: SheetImageCommandInput) {
  const owner = owners.get(api)?.get(unitId)
  if (!owner || api.isDisposed) return null
  const resources = await owner.pick(id, input)
  return !api.isDisposed && owners.get(api)?.get(unitId) === owner ? resources : null
}
export function clearSheetImagePickers(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) owners.delete(api)
  else { const units = owners.get(api); units?.delete(unitId); if (!units?.size) owners.delete(api) }
}
