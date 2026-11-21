import { AlignCenter, AlignLeft, AlignRight, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, AlignVerticalJustifyStart, ArrowDownAZ, ArrowUpAZ, AArrowDown, AArrowUp, Bold, CalendarDays, CheckSquare, ClipboardPaste, Crosshair, Eraser, FileChartColumn, Grid2X2, Grid3X3, Italic, ListChecks, ListFilter, Maximize, Merge, Minimize, PaintBucket, Paintbrush, Plus, Redo2, RefreshCw, RotateCw, Search, ShieldCheck, Shrink, Snowflake, Sparkles, Strikethrough, TextCursorInput, Trash2, Type, Underline, Undo2, WrapText } from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema.ts'
import { SHEET_FUNCTION_CATALOG } from '../../features/sheet-workbench/sheet-function-catalog.ts'
import { localizeRibbon, type OfficeTranslator } from '../../i18n/office-i18n.ts'

export interface SheetRibbonState {
  canUndo: boolean
  canRedo: boolean
  hasSheet: boolean
  canMerge: boolean
  canUnmerge: boolean
  canDeleteRow: boolean
  canDeleteColumn: boolean
  chartAvailable: boolean
  pivotChartAvailable?: boolean
  dataValidationAvailable: boolean
  canSort: boolean
  filterEnabled?: boolean
  filterCanToggle?: boolean
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
  fullscreen: boolean
  calculationMode?: 'automatic' | 'manual'
  outlineSummaryAbove?: boolean
  outlineSummaryLeft?: boolean
  crosshairEnabled?: boolean
  crosshairColor?: string
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
export function createSheetRibbon(state: SheetRibbonState, translate?: OfficeTranslator): RibbonTab[] {
  const button = (id: string, title: string, extra: Partial<RibbonItem> = {}): RibbonItem => ({
    id, title, commandId: id, type: 'button', disabled: !state.hasSheet, ...extra,
  })
  const tabs: RibbonTab[] = [
    {
      id: 'home', title: '开始', order: 0, groups: [
        { id: 'clipboard', title: '剪贴板', order: 0, items: [
          button('sheet.command.undo', '撤销', { icon: Undo2, disabled: !state.canUndo, shortcut: 'Ctrl+Z', gridLayout: { row: 1, column: 1 } }),
          button('sheet.command.redo', '重做', { icon: Redo2, disabled: !state.canRedo, shortcut: 'Ctrl+Y', gridLayout: { row: 2, column: 1 } }),
          button('sheet.command.clear-format', '清除格式', { icon: Eraser, showLabel: true, gridLayout: { row: 1, column: 2 } }),
          button('sheet.operation.format-painter', '格式刷', { icon: Paintbrush, showLabel: true, gridLayout: { row: 2, column: 2 } }),
          { ...button('sheet.command.paste', '粘贴', { icon: ClipboardPaste, shortcut: 'Ctrl+V', gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }), type: 'split', children: [
            button('sheet.menu.paste-all', '保留源格式粘贴', { commandId: 'sheet.command.paste' }),
            button('sheet.menu.paste-values', '仅粘贴值', { commandId: 'sheet.command.paste-values' }),
            button('sheet.menu.paste-formulas', '仅粘贴公式', { commandId: 'sheet.command.paste-formulas' }),
            button('sheet.menu.paste-formats', '仅粘贴格式', { commandId: 'sheet.command.paste-formats' }),
            button('sheet.menu.paste-column-widths', '仅粘贴列宽', { commandId: 'sheet.command.paste-column-widths' }),
            button('sheet.menu.paste-except-borders', '保留目标边框粘贴', { commandId: 'sheet.command.paste-except-borders' }),
          ] },
        ] },
        { id: 'font', title: '字体', order: 10, items: [
          { id: 'sheet.command.set-font-family', commandId: 'sheet.command.set-font-family', type: 'select', title: '字体', width: 156, disabled: !state.hasSheet, value: state.fontFamily, placeholder: '混合', gridLayout: { row: 1, column: 1, columnSpan: 3, width: 156, minWidth: 120 }, options: [...new Set([...fonts, ...(state.fontFamily ? [state.fontFamily] : [])])].map(value => ({ label: value, value })) },
          { id: 'sheet.command.set-font-size', commandId: 'sheet.command.set-font-size', type: 'select', title: '字号', width: 66, disabled: !state.hasSheet, value: state.fontSize, placeholder: '混合', gridLayout: { row: 1, column: 4, width: 68 }, options: [...new Set([...fontSizes, ...(state.fontSize ? [state.fontSize] : [])])].sort((a, b) => a - b).map(value => ({ label: String(value), value })) },
          button('sheet.command.increase-font-size', '增大字号', { icon: AArrowUp, gridLayout: { row: 1, column: 5 } }),
          button('sheet.command.decrease-font-size', '减小字号', { icon: AArrowDown, gridLayout: { row: 1, column: 6 } }),
          button('sheet.command.toggle-bold', '加粗', { icon: Bold, activated: state.bold === true, gridLayout: { row: 2, column: 1 } }),
          button('sheet.command.toggle-italic', '斜体', { icon: Italic, activated: state.italic === true, gridLayout: { row: 2, column: 2 } }),
          button('sheet.command.toggle-underline', '下划线', { icon: Underline, activated: state.underline === true, gridLayout: { row: 2, column: 3 } }),
          button('sheet.command.toggle-strike', '删除线', { icon: Strikethrough, activated: state.strike === true, gridLayout: { row: 2, column: 4 } }),
          { id: 'sheet.menu.text-color', type: 'menu', title: '文字颜色', icon: Type, disabled: !state.hasSheet, gridLayout: { row: 2, column: 5 }, children: colors.map(color => button(`sheet.menu.text-color.${color.value}`, color.label, { commandId: 'sheet.command.set-color', params: color.value, activated: state.color === color.value })) },
          { id: 'sheet.menu.fill', type: 'menu', title: '填充颜色', icon: PaintBucket, disabled: !state.hasSheet, gridLayout: { row: 2, column: 6 }, children: colors.map(color => button(`sheet.menu.fill.${color.value}`, color.label, { commandId: 'sheet.command.set-background', params: color.value, activated: state.background === color.value })) },
          { id: 'sheet.menu.borders', type: 'menu', title: '边框', icon: Grid3X3, disabled: !state.hasSheet, gridLayout: { row: 2, column: 7 }, children: [
            ['all', '所有边框'], ['outside', '外侧框线'], ['inside', '内部框线'], ['top', '上框线'], ['bottom', '下框线'], ['left', '左框线'], ['right', '右框线'], ['none', '无框线'],
          ].map(([value, label]) => button(`sheet.menu.border.${value}`, label!, { commandId: 'sheet.command.set-borders', params: value })) },
        ] },
        { id: 'layout', title: '布局', order: 20, items: [
          button('sheet.command.align-left', '左对齐', { commandId: 'sheet.command.set-align', params: 'left', icon: AlignLeft, activated: state.align === 'left', gridLayout: { row: 1, column: 1 } }),
          button('sheet.command.align-center', '居中', { commandId: 'sheet.command.set-align', params: 'center', icon: AlignCenter, activated: state.align === 'center', gridLayout: { row: 1, column: 2 } }),
          button('sheet.command.align-right', '右对齐', { commandId: 'sheet.command.set-align', params: 'right', icon: AlignRight, activated: state.align === 'right', gridLayout: { row: 1, column: 3 } }),
          { id: 'sheet.menu.vertical-align', type: 'menu', title: '垂直对齐', icon: AlignVerticalJustifyCenter, disabled: !state.hasSheet, gridLayout: { row: 2, column: 1 }, children: [
            button('sheet.menu.vertical.top', '顶端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'top', icon: AlignVerticalJustifyStart, activated: state.vertical === 'top' }),
            button('sheet.menu.vertical.middle', '垂直居中', { commandId: 'sheet.command.set-vertical-align', params: 'middle', icon: AlignVerticalJustifyCenter, activated: state.vertical === 'middle' }),
            button('sheet.menu.vertical.bottom', '底端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'bottom', icon: AlignVerticalJustifyEnd, activated: state.vertical === 'bottom' }),
          ] },
          button('sheet.command.toggle-wrap', '自动换行', { icon: WrapText, activated: state.wrap === true, gridLayout: { row: 1, column: 4 } }),
          { id: 'sheet.menu.text-rotation', type: 'menu', title: '文字方向', icon: RotateCw, disabled: !state.hasSheet, gridLayout: { row: 1, column: 5 }, children: [
            ...textRotationOptions.map(([angle, label]) => button(`sheet.menu.text-rotation.${angle}`, label, {
            commandId: 'sheet.command.set-text-rotation', params: angle,
            activated: angle === 'v' ? state.rotationVertical === true : state.rotationVertical !== true && state.rotation === angle,
          }))] },
          button('sheet.command.set-shrink-to-fit', '缩小字体填充', {
            icon: Shrink, params: !(state.shrinkToFit === true && state.wrap !== true),
            activated: state.shrinkToFit === true && state.wrap !== true, disabled: !state.hasSheet || state.wrap === true,
            disabledReason: state.wrap === true ? '关闭自动换行后可使用缩小字体填充' : undefined,
            gridLayout: { row: 2, column: 2 },
          }),
          { ...button('sheet.command.merge', '合并单元格', { icon: Merge, disabled: !state.canMerge, gridLayout: { row: 2, column: 3 } }), type: 'split', children: [
            button('sheet.command.merge-all', '合并单元格', { commandId: 'sheet.command.merge', disabled: !state.canMerge }),
            button('sheet.command.unmerge', '取消合并', { disabled: !state.canUnmerge }),
          ] },
        ] },
        { id: 'number', title: '数字', order: 30, items: [
          { id: 'sheet.command.set-number-format', commandId: 'sheet.command.set-number-format', type: 'select', title: '数字格式', width: 128, disabled: !state.hasSheet, value: state.numberFormat, placeholder: '混合或自定义', gridLayout: { row: 1, column: 1, columnSpan: 4, width: 128 }, options: [
            { label: '常规', value: 'general' }, { label: '数字', value: 'number' }, { label: '货币', value: 'currency' }, { label: '会计专用', value: 'accounting' },
            { label: '百分比', value: 'percent' }, { label: '日期', value: 'date' }, { label: '时间', value: 'time' }, { label: '日期和时间', value: 'datetime' }, { label: '文本', value: 'text' },
          ] },
          button('sheet.command.increase-decimals', '增加小数位', { icon: Plus, gridLayout: { row: 2, column: 1 } }),
          button('sheet.command.decrease-decimals', '减少小数位', { icon: Type, gridLayout: { row: 2, column: 2 } }),
          button('sheet.operation.open-custom-number-format', '自定义格式', { icon: Type, gridLayout: { row: 2, column: 3 } }),
        ] },
        { id: 'tools', title: '工具', order: 40, items: [
          button('sheet.operation.find', '查找', { icon: Search, showLabel: true, shortcut: 'Ctrl+F', gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.command.clear-content', '清除内容', { icon: Trash2, showLabel: true, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.open-range-protection', '保护范围', { icon: ShieldCheck, showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
        ] },
      ],
    },
    {
      id: 'insert', title: '插入', order: 10, groups: [
        { id: 'objects', title: '对象', order: 10, items: [
          button('sheet.operation.insert-chart', '图表', { icon: FileChartColumn, showLabel: true, disabled: !state.chartAvailable, disabledReason: '图表需要绑定到已打开的工作簿', gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.insert.pivot-table', '数据透视表', { commandId: 'sheet.operation.open-pivot', icon: Grid2X2, showLabel: true, disabled: !state.chartAvailable, disabledReason: '数据透视表需要绑定到已打开的工作簿', gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.insert.slicer', '切片器', { commandId: 'sheet.operation.open-slicer', icon: ListFilter, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          button('sheet.insert.timeline', '时间线', { commandId: 'sheet.operation.open-timeline', icon: CalendarDays, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
        ] },
        { id: 'quick-validation', title: '快速数据验证', order: 20, items: [
          button('sheet.operation.quick-data-validation.checkbox', '复选框', { commandId: 'sheet.operation.quick-data-validation', params: 'checkbox', icon: CheckSquare, showLabel: true, disabled: !state.dataValidationAvailable, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.quick-data-validation.dropdown', '下拉列表', { commandId: 'sheet.operation.quick-data-validation', params: 'dropdown', icon: ListFilter, showLabel: true, disabled: !state.dataValidationAvailable, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.quick-data-validation.date', '日期', { commandId: 'sheet.operation.quick-data-validation', params: 'date', icon: CalendarDays, showLabel: true, disabled: !state.dataValidationAvailable, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.quick-data-validation.date-time', '日期和时间', { commandId: 'sheet.operation.quick-data-validation', params: 'date-time', icon: CalendarDays, showLabel: true, disabled: !state.dataValidationAvailable, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
        ] },
      ],
    },
    {
      id: 'data', title: '数据', order: 30, groups: [
        { id: 'validation', title: '数据验证', order: 0, items: [
          button('sheet.operation.open-data-validation', '数据验证', { icon: ListChecks, showLabel: true, disabled: !state.dataValidationAvailable, disabledReason: '数据验证需要连接工作簿命令服务', gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
        ] },
        { id: 'rules', title: '规则', order: 10, items: [
          button('sheet.operation.open-conditional-formatting', '条件格式', { icon: PaintBucket, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
        ] },
        { id: 'sort', title: '排序和筛选', order: 20, items: [
          button('sheet.command.sort-ascending', '升序', { icon: ArrowDownAZ, showLabel: true, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围', gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.command.sort-descending', '降序', { icon: ArrowUpAZ, showLabel: true, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围', gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.open-data-panel', '自定义排序', { params: 'sort', icon: ArrowDownAZ, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          { ...button('sheet.data.filter', '筛选', { commandId: 'sheet.command.toggle-auto-filter', icon: ListFilter, showLabel: true, activated: state.filterEnabled, disabled: !state.filterCanToggle, disabledReason: translate?.('sheet.filter.select-range', '请选择包含标题行和数据行的区域后再启用筛选'), gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }), type: 'split', children: [
            button('sheet.menu.filter.clear', '清除筛选', { commandId: 'sheet.command.clear-filter', disabled: !state.filterEnabled }),
            button('sheet.menu.filter.reapply', '重新应用', { commandId: 'sheet.command.reapply-filter', disabled: !state.filterEnabled }),
            button('sheet.menu.filter.panel', '筛选条件', { commandId: 'sheet.operation.open-data-panel', params: 'filter', disabled: !state.filterEnabled }),
          ] },
        ] },
        { id: 'pivot', title: '分析', order: 30, items: [
          button('sheet.operation.open-pivot', '数据透视表', { icon: Grid2X2, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.open-slicer', '切片器', { icon: ListFilter, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.open-timeline', '时间线', { icon: CalendarDays, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.open-sparkline', '迷你图', { icon: Sparkles, showLabel: true, disabled: !state.hasSheet, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.open-pivot-chart', '透视图', { icon: FileChartColumn, showLabel: true, disabled: !state.pivotChartAvailable, disabledReason: '请先创建数据透视表', gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true } }),
        ] },
        { id: 'outline', title: '分组显示', order: 40, items: [
          button('sheet.operation.outline-group-rows', '组合行', { icon: TextCursorInput, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-toggle-rows', '折叠/展开行组', { icon: TextCursorInput, showLabel: true, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-remove-rows', '取消行分组', { icon: Eraser, showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-group-columns', '组合列', { icon: TextCursorInput, showLabel: true, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-toggle-columns', '折叠/展开列组', { icon: TextCursorInput, showLabel: true, gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-remove-columns', '取消列分组', { icon: Eraser, showLabel: true, gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true } }),
          { id: 'sheet.operation.outline-summary-rows', commandId: 'sheet.operation.outline-summary-position', type: 'select', title: '汇总行位置', width: 112, value: state.outlineSummaryAbove ? 'above' : 'below', valueParams: value => ({ axis: 'row', atStart: value === 'above' }), gridLayout: { row: 1, column: 7, rowSpan: 2, width: 112, showLabel: true }, options: [
            { label: '汇总行在下方', value: 'below' }, { label: '汇总行在上方', value: 'above' },
          ] },
          { id: 'sheet.operation.outline-summary-columns', commandId: 'sheet.operation.outline-summary-position', type: 'select', title: '汇总列位置', width: 112, value: state.outlineSummaryLeft ? 'left' : 'right', valueParams: value => ({ axis: 'column', atStart: value === 'left' }), gridLayout: { row: 1, column: 8, rowSpan: 2, width: 112, showLabel: true }, options: [
            { label: '汇总列在右侧', value: 'right' }, { label: '汇总列在左侧', value: 'left' },
          ] },
        ] },
      ],
    },
    {
      id: 'formulas', title: '公式', order: 20, groups: [
        { id: 'formula-edit', title: '公式编辑', items: [
          button('sheet.operation.open-function', '插入函数', { icon: Type, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          ...functionCategories.map((category, index) => ({
            id: `sheet.operation.open-function.category-${index}`,
            type: 'menu' as const,
            title: category,
            icon: Type,
            disabled: !state.hasSheet,
            gridLayout: { row: index % 2 + 1, column: Math.floor(index / 2) + 2, showLabel: true },
            children: SHEET_FUNCTION_CATALOG.filter(entry => entry.category === category).map(entry => button(
              `sheet.function.${entry.name}`,
              entry.name,
              { commandId: 'sheet.operation.insert-function', params: entry.name },
            )),
          })),
          button('sheet.operation.start-formula', '编辑公式', { icon: Type, showLabel: true, gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true } }),
          button('sheet.command.recalculate', '重新计算', { icon: RefreshCw, showLabel: true, gridLayout: { row: 1, column: 7, rowSpan: 2, showLabel: true } }),
        ] },
        { id: 'calculation-options', title: '计算选项', order: 10, items: [
          { id: 'sheet.command.set-calculation-mode', commandId: 'sheet.command.set-calculation-mode', type: 'select', title: '计算模式', width: 116, value: state.calculationMode ?? 'automatic', gridLayout: { row: 1, column: 1, rowSpan: 2, width: 116, showLabel: true }, options: [
            { label: '自动', value: 'automatic' }, { label: '手动', value: 'manual' },
          ] },
        ] },
      ],
    },
    {
      id: 'view', title: '视图', order: 40, groups: [
        { id: 'display', title: '显示', order: 0, items: [
          { ...button('sheet.command.freeze-selection', '冻结窗格', { icon: Snowflake, showLabel: true, activated: state.frozen, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }), type: 'split', children: [
            button('sheet.menu.freeze.selection', '冻结到当前单元格', { commandId: 'sheet.command.freeze-selection' }),
            button('sheet.command.freeze-first-row', '冻结首行'),
            button('sheet.command.freeze-first-column', '冻结首列'),
            button('sheet.command.unfreeze', '取消冻结', { disabled: !state.frozen }),
          ] },
          { id: 'sheet.operation.set-zoom', commandId: 'sheet.operation.set-zoom', type: 'select', title: '缩放', width: 88, value: state.zoom, valueParams: value => Number(value), gridLayout: { row: 1, column: 2, rowSpan: 2, width: 88, showLabel: true }, options: [50, 75, 100, 125, 150, 175, 200, 400].map(value => ({ label: `${value}%`, value })) },
        ] },
        { id: 'visibility', title: '可见性', order: 10, items: [
          button('sheet.command.toggle-gridlines', '网格线', { icon: Grid2X2, showLabel: true, activated: state.gridlines, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.toggle-fullscreen', state.fullscreen ? '退出全屏' : '全屏', { icon: state.fullscreen ? Minimize : Maximize, showLabel: true, activated: state.fullscreen, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          { ...button('sheet.operation.toggle-crosshair-highlight', '十字高亮', { icon: Crosshair, activated: state.crosshairEnabled, gridLayout: { row: 2, column: 2 } }), type: 'split', children: [
            ...[
              ['blue', '蓝色', 'rgba(91, 141, 239, 0.18)'], ['teal', '青色', 'rgba(36, 166, 153, 0.18)'],
              ['green', '绿色', 'rgba(65, 160, 96, 0.18)'], ['yellow', '黄色', 'rgba(235, 190, 58, 0.2)'],
              ['orange', '橙色', 'rgba(235, 128, 55, 0.18)'], ['red', '红色', 'rgba(220, 75, 75, 0.16)'],
              ['purple', '紫色', 'rgba(141, 99, 201, 0.16)'], ['gray', '灰色', 'rgba(105, 119, 140, 0.16)'],
            ].map(([color, title, value]) => button(`sheet.operation.crosshair-color.${color}`, title!, {
              commandId: 'sheet.operation.set-crosshair-highlight-color', params: value, swatch: value,
            })),
          ] },
        ] },
      ],
    },
  ]
  return translate ? localizeRibbon(tabs, 'sheet', translate) : tabs
}
