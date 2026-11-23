import type { SheetPasteMode } from './sheet-paste.ts'

export const SHEET_CLIPBOARD_COMMANDS = Object.freeze({
  copy: 'sheet.command.copy', cut: 'sheet.command.cut', paste: 'sheet.command.paste',
  copyFormulas: 'sheet.command.copy-formulas', values: 'sheet.command.paste-values',
  formulas: 'sheet.command.paste-formulas', formats: 'sheet.command.paste-formats',
  columnWidths: 'sheet.command.paste-column-widths', exceptBorders: 'sheet.command.paste-except-borders',
  optionalPaste: 'sheet.command.optional-paste', cancelCopy: 'sheet.operation.cancel-copy',
})

export interface SheetClipboardCommandAction {
  action: 'copy' | 'cut' | 'paste' | 'optional-paste' | 'cancel-copy'
  mode?: SheetPasteMode
  formulaOnly?: boolean
}

const aliases: Readonly<Record<string, SheetClipboardCommandAction>> = {
  [SHEET_CLIPBOARD_COMMANDS.copy]: { action: 'copy' },
  [SHEET_CLIPBOARD_COMMANDS.cut]: { action: 'cut' },
  [SHEET_CLIPBOARD_COMMANDS.paste]: { action: 'paste' },
  [SHEET_CLIPBOARD_COMMANDS.copyFormulas]: { action: 'copy', formulaOnly: true },
  [SHEET_CLIPBOARD_COMMANDS.values]: { action: 'paste', mode: 'values' },
  [SHEET_CLIPBOARD_COMMANDS.formulas]: { action: 'paste', mode: 'formulas' },
  [SHEET_CLIPBOARD_COMMANDS.formats]: { action: 'paste', mode: 'formats' },
  [SHEET_CLIPBOARD_COMMANDS.columnWidths]: { action: 'paste', mode: 'columnWidths' },
  [SHEET_CLIPBOARD_COMMANDS.exceptBorders]: { action: 'paste', mode: 'exceptBorders' },
  [SHEET_CLIPBOARD_COMMANDS.optionalPaste]: { action: 'optional-paste' },
  [SHEET_CLIPBOARD_COMMANDS.cancelCopy]: { action: 'cancel-copy' },
  'sheet.command.copy-formula-only': { action: 'copy', formulaOnly: true },
  'sheet.command.paste-value': { action: 'paste', mode: 'values' },
  'sheet.command.paste-formula': { action: 'paste', mode: 'formulas' },
  'sheet.command.paste-format': { action: 'paste', mode: 'formats' },
  'sheet.command.paste-col-width': { action: 'paste', mode: 'columnWidths' },
  'sheet.command.paste-besides-border': { action: 'paste', mode: 'exceptBorders' },
  'sheet.command.paste-by-short-key': { action: 'paste' },
}
export const SHEET_CLIPBOARD_COMMAND_IDS: readonly string[] = Object.freeze(Object.keys(aliases))

export function resolveSheetClipboardCommand(commandId: string): SheetClipboardCommandAction | undefined {
  const action = Object.hasOwn(aliases, commandId) ? aliases[commandId] : undefined
  return action ? { ...action } : undefined
}

const pasteModes: Readonly<Record<string, SheetPasteMode>> = {
  all: 'all', values: 'values', formulas: 'formulas', formats: 'formats', columnWidths: 'columnWidths', exceptBorders: 'exceptBorders',
  'default-paste': 'all', 'special-paste-value': 'values', 'special-paste-formula': 'formulas',
  'special-paste-format': 'formats', 'special-paste-col-width': 'columnWidths', 'special-paste-besides-border': 'exceptBorders',
}

export type SheetClipboardPasteType = 'DEFAULT_PASTE' | 'SPECIAL_PASTE_VALUE' | 'SPECIAL_PASTE_FORMAT'
  | 'SPECIAL_PASTE_COL_WIDTH' | 'SPECIAL_PASTE_BESIDES_BORDER' | 'SPECIAL_PASTE_FORMULA'

const pasteTypes: Readonly<Record<SheetClipboardPasteType, SheetPasteMode>> = Object.freeze({
  DEFAULT_PASTE: 'all', SPECIAL_PASTE_VALUE: 'values', SPECIAL_PASTE_FORMAT: 'formats',
  SPECIAL_PASTE_COL_WIDTH: 'columnWidths', SPECIAL_PASTE_BESIDES_BORDER: 'exceptBorders', SPECIAL_PASTE_FORMULA: 'formulas',
})

/** Univer's optional command uses the enum key in `type`; ordinary paste uses its hook value in `value`. */
export function sheetClipboardPasteMode(mode?: string, value?: string, commandMode?: SheetPasteMode, type?: string): SheetPasteMode {
  const supplied: SheetPasteMode[] = []
  for (const input of [mode, value]) if (input !== undefined) {
    if (!Object.hasOwn(pasteModes, input)) throw new TypeError('选择性粘贴模式无效或参数不一致。')
    supplied.push(pasteModes[input]!)
  }
  if (type !== undefined) {
    if (!Object.hasOwn(pasteTypes, type)) throw new TypeError('选择性粘贴模式无效或参数不一致。')
    supplied.push(pasteTypes[type as SheetClipboardPasteType])
  }
  if (commandMode !== undefined) supplied.push(commandMode)
  if (supplied.some(input => input !== supplied[0])) throw new TypeError('选择性粘贴模式无效或参数不一致。')
  return supplied[0] ?? 'all'
}
