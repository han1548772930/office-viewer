import type { OfficeStudio } from './office-studio.ts'
import type { SheetImageCommandInput } from '../features/sheet-images/sheet-image-commands.ts'
interface Owner { open(input: SheetImageCommandInput): boolean }
const owners = new WeakMap<OfficeStudio, Map<string, Owner>>()
export function bindSheetImageEditor(api: OfficeStudio, unitId: string, open: Owner['open']) {
  if (api.isDisposed || !api.getWorkbookModel(unitId)) throw new Error('找不到图片编辑器的工作簿。')
  const units = owners.get(api) ?? new Map<string, Owner>(), owner = { open }
  owners.set(api, units); units.set(unitId, owner)
  return { dispose: () => { if (units.get(unitId) === owner) units.delete(unitId); if (!units.size && owners.get(api) === units) owners.delete(api) } }
}
export function requestSheetImageEditor(api: OfficeStudio, unitId: string, input: SheetImageCommandInput): boolean {
  return !api.isDisposed && (owners.get(api)?.get(unitId)?.open(input) ?? false)
}
export function clearSheetImageEditors(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) owners.delete(api)
  else { const units = owners.get(api); units?.delete(unitId); if (!units?.size) owners.delete(api) }
}
