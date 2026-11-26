/** Fixed Univer v1.6.0 drawing-ui context menu contract; command registration belongs to the host. */
export const DRAWING_GROUP_MENU = {
  id: 'contextMenu.drawing-group', label: '组合', contextGroup: 'others', order: 1,
  items: [
    { action: 'group', label: '组合', order: 0, commandId: 'drawing.operation.set-drawing-group', contentCommandId: 'sheet.command.group-sheet-image' },
    { action: 'ungroup', label: '取消组合', order: 1, commandId: 'drawing.operation.cancel-drawing-group', contentCommandId: 'sheet.command.ungroup-sheet-image' },
  ],
} as const

export const DRAWING_ARRANGE_MENU = {
  id: 'contextMenu.drawing-arrange', label: '排列', contextGroup: 'others', order: 2,
  items: [
    { action: 'front', label: '置于顶层', order: 0, commandId: 'drawing.operation.set-drawing-arrange-front' },
    { action: 'forward', label: '上移一层', order: 1, commandId: 'drawing.operation.set-drawing-arrange-forward' },
    { action: 'back', label: '置于底层', order: 2, commandId: 'drawing.operation.set-drawing-arrange-back' },
    { action: 'backward', label: '下移一层', order: 3, commandId: 'drawing.operation.set-drawing-arrange-backward' },
  ],
  contentCommandId: 'sheet.command.set-drawing-arrange',
} as const

export const SHEET_DRAWING_GROUP_COMMAND_IDS = {
  group: 'sheet.command.group-sheet-image', ungroup: 'sheet.command.ungroup-sheet-image',
  arrange: 'sheet.command.set-drawing-arrange', transform: 'sheet.command.transform-drawing-group',
  remove: 'sheet.command.remove-sheet-drawings',
  align: 'sheet.command.align-sheet-drawings',
  move: 'sheet.command.move-drawing',
} as const

export const DRAWING_ALIGN_MENU = {
  id: 'contextMenu.drawing-align', label: '对齐', contextGroup: 'others', order: 3,
  items: [
    { alignType: '1', label: '左对齐', order: 0, commandId: 'sheet.operation.set-drawing-align-left' },
    { alignType: '2', label: '水平居中', order: 1, commandId: 'sheet.operation.set-drawing-align-center' },
    { alignType: '3', label: '右对齐', order: 2, commandId: 'sheet.operation.set-drawing-align-right' },
    { alignType: '4', label: '顶部对齐', order: 3, commandId: 'sheet.operation.set-drawing-align-top' },
    { alignType: '5', label: '垂直居中', order: 4, commandId: 'sheet.operation.set-drawing-align-middle' },
    { alignType: '6', label: '底部对齐', order: 5, commandId: 'sheet.operation.set-drawing-align-bottom' },
    { alignType: '7', label: '水平分布', order: 6, commandId: 'sheet.operation.set-drawing-align-horizon' },
    { alignType: '8', label: '垂直分布', order: 7, commandId: 'sheet.operation.set-drawing-align-vertical' },
  ],
  contentCommandId: SHEET_DRAWING_GROUP_COMMAND_IDS.align,
} as const
