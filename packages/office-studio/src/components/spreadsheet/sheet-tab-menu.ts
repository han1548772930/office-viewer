export interface SheetTabMenuActions {
  rename(sheetId: string): void
  copy(sheetId: string): void
  protect(sheetId: string): void
  unprotect(sheetId: string): void
  setColor(sheetId: string, color: string): void
  hide(sheetId: string): void
  delete(sheetId: string): void
}

interface SheetTabMenuItem {
  key: string
  label: string
  disabled?: boolean
  danger?: boolean
  children?: SheetTabMenuItem[]
}

export function createSheetTabMenu(
  sheetId: string,
  visibleSheetCount: number,
  isProtected: boolean,
  translate: (key: string, fallback: string) => string,
  actions: SheetTabMenuActions,
) {
  return {
    items: [
      { key: 'rename', label: '重命名' },
      { key: 'copy', label: translate('sheet.footer.copy-sheet', '复制工作表') },
      { key: 'tab-color', label: translate('sheet.footer.tab-color', '标签颜色'), children: [
        { key: 'color-reset', label: translate('sheet.footer.no-color', '无颜色') },
        { key: 'color:#c00000', label: translate('sheet.footer.red', '红色') },
        { key: 'color:#ed7d31', label: translate('sheet.footer.orange', '橙色') },
        { key: 'color:#ffc000', label: translate('sheet.footer.yellow', '黄色') },
        { key: 'color:#70ad47', label: translate('sheet.footer.green', '绿色') },
        { key: 'color:#00b0f0', label: translate('sheet.footer.cyan', '青色') },
        { key: 'color:#4472c4', label: translate('sheet.footer.blue', '蓝色') },
        { key: 'color:#7030a0', label: translate('sheet.footer.purple', '紫色') },
      ] },
      { key: 'protect-sheet', label: translate('sheet.footer.protect-sheet', isProtected ? '工作表保护设置' : '保护工作表') },
      ...(isProtected ? [{ key: 'unprotect-sheet', label: translate('sheet.footer.unprotect-sheet', '移除工作表保护') }] : []),
      { key: 'hide', label: '隐藏工作表', disabled: visibleSheetCount <= 1 },
      { key: 'delete', label: '删除工作表', danger: true, disabled: visibleSheetCount <= 1 },
    ] satisfies SheetTabMenuItem[],
    onClick: ({ key }: { key: string | number }) => {
      if (key === 'rename') actions.rename(sheetId)
      if (key === 'copy') actions.copy(sheetId)
      if (key === 'protect-sheet') actions.protect(sheetId)
      if (key === 'unprotect-sheet' && isProtected) actions.unprotect(sheetId)
      if (key === 'color-reset') actions.setColor(sheetId, '')
      if (typeof key === 'string' && key.startsWith('color:#')) actions.setColor(sheetId, key.slice('color:'.length))
      if (key === 'hide' && visibleSheetCount > 1) actions.hide(sheetId)
      if (key === 'delete' && visibleSheetCount > 1) actions.delete(sheetId)
    },
  }
}
