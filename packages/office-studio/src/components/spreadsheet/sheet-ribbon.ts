import { AlignCenter, AlignLeft, AlignRight, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, AlignVerticalJustifyStart, ArrowDownAZ, ArrowUpAZ, AArrowDown, AArrowUp, Bold, CalendarDays, ClipboardPaste, Eraser, FileChartColumn, Grid2X2, Grid3X3, Italic, ListChecks, ListFilter, Merge, PaintBucket, Plus, Redo2, RefreshCw, RotateCw, Search, Shrink, Snowflake, Strikethrough, Trash2, Type, Underline, Undo2, WrapText, ZoomIn, ZoomOut } from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema.ts'
import { SHEET_FUNCTION_CATALOG } from '../../features/sheet-workbench/sheet-function-catalog.ts'

export interface SheetRibbonState {
  canUndo: boolean
  canRedo: boolean
  hasSheet: boolean
  canMerge: boolean
  canUnmerge: boolean
  canDeleteRow: boolean
  canDeleteColumn: boolean
  chartAvailable: boolean
  dataValidationAvailable: boolean
  canSort: boolean
  fontFamily: string | null
  fontSize: number | null
  numberFormat: string | null
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  wrap?: boolean
  color?: string
  background?: string
  align?: string
  vertical?: string
  rotation?: number | null
  rotationVertical?: boolean
  shrinkToFit?: boolean
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
const functionCategories = [...new Set(SHEET_FUNCTION_CATALOG.map(entry => entry.category))]
const textRotationOptions: Array<[number | 'v', string]> = [
  [0, '水平'], [-45, '左斜 45°'], [45, '右斜 45°'], ['v', '竖排'], [-90, '左转 90°'], [90, '右转 90°'],
]

/** The schema describes editor intents; UI, keyboard and menu callers use the same dispatcher. */
export function createSheetRibbon(state: SheetRibbonState): RibbonTab[] {
  const button = (id: string, title: string, extra: Partial<RibbonItem> = {}): RibbonItem => ({
    id, title, commandId: id, type: 'button', disabled: !state.hasSheet, ...extra,
  })
  return [
    {
      id: 'home', title: '开始', order: 0, groups: [
        { id: 'clipboard', title: '剪贴板', order: 0, items: [
          button('sheet.command.undo', '撤销', { icon: Undo2, disabled: !state.canUndo, shortcut: 'Ctrl+Z' }),
          button('sheet.command.redo', '重做', { icon: Redo2, disabled: !state.canRedo, shortcut: 'Ctrl+Y' }),
          button('sheet.command.clear-format', '清除格式', { icon: Eraser, showLabel: true }),
          button('sheet.command.paste', '粘贴', { icon: ClipboardPaste, shortcut: 'Ctrl+V' }),
        ] },
        { id: 'font', title: '字体', order: 10, items: [
          { id: 'sheet.command.set-font-family', commandId: 'sheet.command.set-font-family', type: 'select', title: '字体', width: 112, disabled: !state.hasSheet, value: state.fontFamily, placeholder: '混合', options: [...new Set([...fonts, ...(state.fontFamily ? [state.fontFamily] : [])])].map(value => ({ label: value, value })) },
          { id: 'sheet.command.set-font-size', commandId: 'sheet.command.set-font-size', type: 'select', title: '字号', width: 66, disabled: !state.hasSheet, value: state.fontSize, placeholder: '混合', options: [...new Set([...fontSizes, ...(state.fontSize ? [state.fontSize] : [])])].sort((a, b) => a - b).map(value => ({ label: String(value), value })) },
          button('sheet.command.increase-font-size', '增大字号', { icon: AArrowUp }),
          button('sheet.command.decrease-font-size', '减小字号', { icon: AArrowDown }),
          button('sheet.command.toggle-bold', '加粗', { icon: Bold, activated: state.bold === true }),
          button('sheet.command.toggle-italic', '斜体', { icon: Italic, activated: state.italic === true }),
          button('sheet.command.toggle-underline', '下划线', { icon: Underline, activated: state.underline === true }),
          button('sheet.command.toggle-strike', '删除线', { icon: Strikethrough, activated: state.strike === true }),
          { id: 'sheet.menu.text-color', type: 'menu', title: '文字颜色', icon: Type, disabled: !state.hasSheet, children: colors.map(color => button(`sheet.menu.text-color.${color.value}`, color.label, { commandId: 'sheet.command.set-color', params: color.value, activated: state.color === color.value })) },
          { id: 'sheet.menu.fill', type: 'menu', title: '填充颜色', icon: PaintBucket, disabled: !state.hasSheet, children: colors.map(color => button(`sheet.menu.fill.${color.value}`, color.label, { commandId: 'sheet.command.set-background', params: color.value, activated: state.background === color.value })) },
          { id: 'sheet.menu.borders', type: 'menu', title: '边框', icon: Grid3X3, disabled: !state.hasSheet, children: [
            ['all', '所有边框'], ['outside', '外侧框线'], ['inside', '内部框线'], ['top', '上框线'], ['bottom', '下框线'], ['left', '左框线'], ['right', '右框线'], ['none', '无框线'],
          ].map(([value, label]) => button(`sheet.menu.border.${value}`, label!, { commandId: 'sheet.command.set-borders', params: value })) },
        ] },
        { id: 'layout', title: '布局', order: 20, items: [
          button('sheet.command.align-left', '左对齐', { commandId: 'sheet.command.set-align', params: 'left', icon: AlignLeft, activated: state.align === 'left' }),
          button('sheet.command.align-center', '居中', { commandId: 'sheet.command.set-align', params: 'center', icon: AlignCenter, activated: state.align === 'center' }),
          button('sheet.command.align-right', '右对齐', { commandId: 'sheet.command.set-align', params: 'right', icon: AlignRight, activated: state.align === 'right' }),
          { id: 'sheet.menu.vertical-align', type: 'menu', title: '垂直对齐', icon: AlignVerticalJustifyCenter, disabled: !state.hasSheet, children: [
            button('sheet.menu.vertical.top', '顶端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'top', icon: AlignVerticalJustifyStart, activated: state.vertical === 'top' }),
            button('sheet.menu.vertical.middle', '垂直居中', { commandId: 'sheet.command.set-vertical-align', params: 'middle', icon: AlignVerticalJustifyCenter, activated: state.vertical === 'middle' }),
            button('sheet.menu.vertical.bottom', '底端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'bottom', icon: AlignVerticalJustifyEnd, activated: state.vertical === 'bottom' }),
          ] },
          button('sheet.command.toggle-wrap', '自动换行', { icon: WrapText, activated: state.wrap === true }),
          { id: 'sheet.menu.text-rotation', type: 'menu', title: '文字方向', icon: RotateCw, disabled: !state.hasSheet, children: [
            ...textRotationOptions.map(([angle, label]) => button(`sheet.menu.text-rotation.${angle}`, label, {
            commandId: 'sheet.command.set-text-rotation', params: angle,
            activated: angle === 'v' ? state.rotationVertical === true : state.rotationVertical !== true && state.rotation === angle,
          }))] },
          button('sheet.command.set-shrink-to-fit', '缩小字体填充', {
            icon: Shrink, params: !(state.shrinkToFit === true && state.wrap !== true),
            activated: state.shrinkToFit === true && state.wrap !== true, disabled: !state.hasSheet || state.wrap === true,
            disabledReason: state.wrap === true ? '关闭自动换行后可使用缩小字体填充' : undefined,
          }),
          { ...button('sheet.command.merge', '合并单元格', { icon: Merge, disabled: !state.canMerge }), type: 'split', children: [
            button('sheet.command.merge-all', '合并单元格', { commandId: 'sheet.command.merge', disabled: !state.canMerge }),
            button('sheet.command.unmerge', '取消合并', { disabled: !state.canUnmerge }),
          ] },
        ] },
        { id: 'number', title: '数字', order: 30, items: [
          { id: 'sheet.command.set-number-format', commandId: 'sheet.command.set-number-format', type: 'select', title: '数字格式', width: 108, disabled: !state.hasSheet, value: state.numberFormat, placeholder: '混合或自定义', options: [
            { label: '常规', value: 'general' }, { label: '数字', value: 'number' }, { label: '货币', value: 'currency' }, { label: '会计专用', value: 'accounting' },
            { label: '百分比', value: 'percent' }, { label: '日期', value: 'date' }, { label: '时间', value: 'time' }, { label: '日期和时间', value: 'datetime' }, { label: '文本', value: 'text' },
          ] },
          button('sheet.command.increase-decimals', '增加小数位', { icon: Plus, showLabel: true }),
          button('sheet.command.decrease-decimals', '减少小数位', { icon: Type, showLabel: true }),
        ] },
        { id: 'tools', title: '工具', order: 40, items: [
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
          button('sheet.insert.slicer', '切片器', { commandId: 'sheet.operation.open-slicer', icon: ListFilter, showLabel: true, disabled: !state.chartAvailable }),
          button('sheet.insert.timeline', '时间线', { commandId: 'sheet.operation.open-timeline', icon: CalendarDays, showLabel: true, disabled: !state.chartAvailable }),
        ] },
      ],
    },
    {
      id: 'data', title: '数据', order: 30, groups: [
        { id: 'validation', title: '数据验证', order: 0, items: [
          button('sheet.operation.open-data-validation', '数据验证', { icon: ListChecks, showLabel: true, disabled: !state.dataValidationAvailable, disabledReason: '数据验证需要连接工作簿命令服务' }),
        ] },
        { id: 'sort', title: '排序', order: 10, items: [
          button('sheet.command.sort-ascending', '升序', { icon: ArrowDownAZ, showLabel: true, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围' }),
          button('sheet.command.sort-descending', '降序', { icon: ArrowUpAZ, showLabel: true, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围' }),
          button('sheet.operation.open-data-panel', '自定义排序', { params: 'sort', icon: ArrowDownAZ, showLabel: true, disabled: !state.chartAvailable }),
          button('sheet.data.filter', '筛选', { commandId: 'sheet.operation.open-data-panel', params: 'filter', icon: ListFilter, showLabel: true, disabled: !state.chartAvailable }),
        ] },
        { id: 'pivot', title: '数据分析', order: 20, items: [
          button('sheet.operation.open-pivot', '数据透视表', { icon: Grid2X2, showLabel: true, disabled: !state.chartAvailable }),
          button('sheet.operation.open-slicer', '切片器', { icon: ListFilter, showLabel: true, disabled: !state.chartAvailable }),
          button('sheet.operation.open-timeline', '时间线', { icon: CalendarDays, showLabel: true, disabled: !state.chartAvailable }),
        ] },
      ],
    },
    {
      id: 'formulas', title: '公式', order: 20, groups: [
        { id: 'formula-edit', title: '公式编辑', items: [
          { ...button('sheet.operation.open-function', '插入函数', { icon: Type, showLabel: true }), type: 'split', children: [
            button('sheet.operation.open-function.all', '所有已支持函数', { commandId: 'sheet.operation.open-function' }),
            ...functionCategories.map((category, index) => button(`sheet.operation.open-function.category-${index}`, category, { commandId: 'sheet.operation.open-function', params: category })),
          ] },
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
