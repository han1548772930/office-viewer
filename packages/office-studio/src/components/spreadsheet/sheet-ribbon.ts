import { AlignCenter, AlignLeft, AlignRight, ArrowDownAZ, ArrowUpAZ, Bold, CalendarDays, ClipboardPaste, Copy, FileChartColumn, Grid2X2, Italic, ListFilter, Merge, PaintBucket, Plus, Redo2, RefreshCw, Scissors, Search, Snowflake, Trash2, Type, Underline, Undo2, WrapText, ZoomIn, ZoomOut } from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema.ts'

export interface SheetRibbonState {
  canUndo: boolean
  canRedo: boolean
  hasSheet: boolean
  canMerge: boolean
  canUnmerge: boolean
  canDeleteRow: boolean
  canDeleteColumn: boolean
  chartAvailable: boolean
  canSort: boolean
  fontFamily: string | null
  fontSize: number | null
  numberFormat: string | null
  bold?: boolean
  italic?: boolean
  underline?: boolean
  wrap?: boolean
  color?: string
  background?: string
  align?: string
  gridlines: boolean
  frozen: boolean
  zoom: number
}

const fonts = ['等线', '微软雅黑', '宋体', 'Aptos', 'Arial', 'Times New Roman']
const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72]
const colors = [
  { label: '白色', value: '#ffffff' }, { label: '浅黄色', value: '#fff4cc' },
  { label: '浅蓝色', value: '#dceaff' }, { label: '浅绿色', value: '#dcf4e4' },
  { label: '浅红色', value: '#ffe0e0' }, { label: '黑色', value: '#1f2937' },
]

/** The schema describes editor intents; UI, keyboard and menu callers use the same dispatcher. */
export function createSheetRibbon(state: SheetRibbonState): RibbonTab[] {
  const button = (id: string, title: string, extra: Partial<RibbonItem> = {}): RibbonItem => ({
    id, title, commandId: id, type: 'button', disabled: !state.hasSheet, ...extra,
  })
  return [
    {
      id: 'home', title: '开始', order: 0, groups: [
        { id: 'history', title: '历史', order: 0, items: [
          button('sheet.command.undo', '撤销', { icon: Undo2, disabled: !state.canUndo, shortcut: 'Ctrl+Z' }),
          button('sheet.command.redo', '重做', { icon: Redo2, disabled: !state.canRedo, shortcut: 'Ctrl+Y' }),
        ] },
        { id: 'clipboard', title: '剪贴板', order: 10, items: [
          button('sheet.command.paste', '粘贴', { icon: ClipboardPaste, shortcut: 'Ctrl+V' }),
          button('sheet.command.cut', '剪切', { icon: Scissors, shortcut: 'Ctrl+X' }),
          button('sheet.command.copy', '复制', { icon: Copy, shortcut: 'Ctrl+C' }),
        ] },
        { id: 'font', title: '字体', order: 20, items: [
          { id: 'sheet.command.set-font-family', commandId: 'sheet.command.set-font-family', type: 'select', title: '字体', width: 112, disabled: !state.hasSheet, value: state.fontFamily, placeholder: '混合', options: [...new Set([...fonts, ...(state.fontFamily ? [state.fontFamily] : [])])].map(value => ({ label: value, value })) },
          { id: 'sheet.command.set-font-size', commandId: 'sheet.command.set-font-size', type: 'select', title: '字号', width: 66, disabled: !state.hasSheet, value: state.fontSize, placeholder: '混合', options: [...new Set([...fontSizes, ...(state.fontSize ? [state.fontSize] : [])])].sort((a, b) => a - b).map(value => ({ label: String(value), value })) },
          button('sheet.command.toggle-bold', '加粗', { icon: Bold, activated: state.bold === true }),
          button('sheet.command.toggle-italic', '斜体', { icon: Italic, activated: state.italic === true }),
          button('sheet.command.toggle-underline', '下划线', { icon: Underline, activated: state.underline === true }),
          { id: 'sheet.menu.text-color', type: 'menu', title: '文字颜色', icon: Type, disabled: !state.hasSheet, children: colors.map(color => button(`sheet.menu.text-color.${color.value}`, color.label, { commandId: 'sheet.command.set-color', params: color.value, activated: state.color === color.value })) },
          { id: 'sheet.menu.fill', type: 'menu', title: '填充颜色', icon: PaintBucket, disabled: !state.hasSheet, children: colors.map(color => button(`sheet.menu.fill.${color.value}`, color.label, { commandId: 'sheet.command.set-background', params: color.value, activated: state.background === color.value })) },
        ] },
        { id: 'alignment', title: '对齐', order: 30, items: [
          button('sheet.command.align-left', '左对齐', { commandId: 'sheet.command.set-align', params: 'left', icon: AlignLeft, activated: state.align === 'left' }),
          button('sheet.command.align-center', '居中', { commandId: 'sheet.command.set-align', params: 'center', icon: AlignCenter, activated: state.align === 'center' }),
          button('sheet.command.align-right', '右对齐', { commandId: 'sheet.command.set-align', params: 'right', icon: AlignRight, activated: state.align === 'right' }),
          button('sheet.command.toggle-wrap', '自动换行', { icon: WrapText, activated: state.wrap === true }),
          { ...button('sheet.command.merge', '合并单元格', { icon: Merge, disabled: !state.canMerge }), type: 'split', children: [
            button('sheet.command.merge-all', '合并单元格', { commandId: 'sheet.command.merge', disabled: !state.canMerge }),
            button('sheet.command.unmerge', '取消合并', { disabled: !state.canUnmerge }),
          ] },
        ] },
        { id: 'number', title: '数字', order: 40, items: [
          { id: 'sheet.command.set-number-format', commandId: 'sheet.command.set-number-format', type: 'select', title: '数字格式', width: 90, disabled: !state.hasSheet, value: state.numberFormat, placeholder: '混合', options: [{ label: '常规', value: 'general' }, { label: '数字', value: 'number' }, { label: '百分比', value: 'percent' }] },
        ] },
        { id: 'editing', title: '编辑', order: 50, items: [
          button('sheet.operation.find', '查找', { icon: Search, showLabel: true, shortcut: 'Ctrl+F' }),
          button('sheet.command.clear-content', '清除内容', { icon: Trash2, showLabel: true }),
        ] },
      ],
    },
    {
      id: 'insert', title: '插入', order: 10, groups: [
        { id: 'cells', title: '行和列', items: [
          button('sheet.command.insert-row', '插入行', { icon: Plus, showLabel: true }),
          button('sheet.command.insert-column', '插入列', { icon: Plus, showLabel: true }),
          button('sheet.command.delete-row', '删除行', { icon: Trash2, showLabel: true, disabled: !state.canDeleteRow }),
          button('sheet.command.delete-column', '删除列', { icon: Trash2, showLabel: true, disabled: !state.canDeleteColumn }),
        ] },
        { id: 'objects', title: '对象', order: 10, items: [
          button('sheet.operation.insert-chart', '图表', { icon: FileChartColumn, showLabel: true, disabled: !state.chartAvailable, disabledReason: '图表需要绑定到已打开的工作簿' }),
          button('sheet.insert.pivot-table', '数据透视表', { commandId: 'sheet.operation.open-pivot', icon: Grid2X2, showLabel: true, disabled: !state.chartAvailable, disabledReason: '数据透视表需要绑定到已打开的工作簿' }),
        ] },
      ],
    },
    {
      id: 'data', title: '数据', order: 30, groups: [
        { id: 'sort', title: '排序', items: [
          button('sheet.command.sort-ascending', '升序', { icon: ArrowDownAZ, showLabel: true, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围' }),
          button('sheet.command.sort-descending', '降序', { icon: ArrowUpAZ, showLabel: true, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围' }),
        ] },
        { id: 'pivot', title: '数据分析', order: 10, items: [
          button('sheet.operation.open-pivot', '数据透视表', { icon: Grid2X2, showLabel: true, disabled: !state.chartAvailable }),
          button('sheet.operation.open-slicer', '切片器', { icon: ListFilter, showLabel: true, disabled: !state.chartAvailable }),
          button('sheet.operation.open-timeline', '时间线', { icon: CalendarDays, showLabel: true, disabled: !state.chartAvailable }),
        ] },
      ],
    },
    {
      id: 'formulas', title: '公式', order: 20, groups: [
        { id: 'formula-edit', title: '公式编辑', items: [
          button('sheet.operation.start-formula', '编辑公式', { icon: Type, showLabel: true }),
          button('sheet.command.recalculate', '重新计算', { icon: RefreshCw, showLabel: true }),
        ] },
      ],
    },
    {
      id: 'view', title: '视图', order: 40, groups: [
        { id: 'display', title: '显示', items: [
          button('sheet.command.toggle-gridlines', '网格线', { icon: Grid2X2, showLabel: true, activated: state.gridlines }),
          { ...button('sheet.command.freeze-selection', '冻结窗格', { icon: Snowflake, showLabel: true, activated: state.frozen }), type: 'split', children: [
            button('sheet.menu.freeze.selection', '冻结到当前单元格', { commandId: 'sheet.command.freeze-selection' }),
            button('sheet.command.freeze-first-row', '冻结首行'),
            button('sheet.command.freeze-first-column', '冻结首列'),
            button('sheet.command.unfreeze', '取消冻结', { disabled: !state.frozen }),
          ] },
        ] },
        { id: 'zoom', title: '缩放', order: 10, items: [
          button('sheet.operation.zoom-out', '缩小', { icon: ZoomOut, disabled: state.zoom <= 10 }),
          button('sheet.operation.reset-zoom', '100%', { showLabel: true }),
          button('sheet.operation.zoom-in', '放大', { icon: ZoomIn, disabled: state.zoom >= 400 }),
        ] },
      ],
    },
  ]
}
