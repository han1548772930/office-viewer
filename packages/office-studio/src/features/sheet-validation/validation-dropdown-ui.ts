export interface SheetValidationDropdownUiOptions {
  showEditOnDropdown?: boolean
  showSearchOnDropdown?: boolean
}
export interface SheetValidationDropdownUiState {
  showEditOnDropdown: boolean
  showSearchOnDropdown: boolean
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
/** Workbench-local preferences mirror the fixed public UI config defaults.
 * They affect discoverability, never View/Edit capability or document content.
 */
export function resolveSheetValidationDropdownUi(value: unknown): SheetValidationDropdownUiState {
  if (value === undefined) return { showEditOnDropdown: true, showSearchOnDropdown: true }
  if (!record(value)) throw new TypeError('下拉列表界面配置必须是对象。')
  const edit = value.showEditOnDropdown, search = value.showSearchOnDropdown
  if ((edit !== undefined && typeof edit !== 'boolean') || (search !== undefined && typeof search !== 'boolean')) {
    throw new TypeError('下拉列表搜索和编辑入口配置必须是布尔值。')
  }
  return { showEditOnDropdown: edit ?? true, showSearchOnDropdown: search ?? true }
}
