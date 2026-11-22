import { h, type Component } from 'vue'
import {
  AArrowDown, AArrowUp, AlignCenter, AlignLeft, AlignRight, AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd, AlignVerticalJustifyStart, ArrowDownAZ, ArrowDownWideNarrow, ArrowUpAZ,
  BadgeDollarSign, Baseline, Bold, Brush, Calculator, CalendarDays,
  ChartColumnIncreasing, ChartLine, ChartNoAxesColumnIncreasing, ChartNoAxesCombined,
  ClipboardCheck, ClipboardPaste, Clock, Crosshair, DecimalsArrowLeft, DecimalsArrowRight, Expand,
  Database, Filter, FilterX, FunctionSquare, Grid2X2, Group, Info, Italic, Link, ListFilter, Minimize, PaintBucket, Palette,
  Pin, PinOff, Redo2, RemoveFormatting, RotateCcw,
  RotateCw, Search, SearchCode, Settings2, ShieldCheck, Shrink, Sigma, SquareCheck, Strikethrough,
  TableCellsMerge, TableCellsSplit, TableProperties, TextCursor, Underline, Undo2, Ungroup, Workflow,
  WrapText, Percent,
} from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema.ts'
import { fontSizeOptions, ribbonFontOptions, type OfficeFontList } from '../../platform/ribbon/ribbon-fonts.ts'
import { SHEET_FUNCTION_CATALOG } from '../../features/sheet-workbench/sheet-function-catalog.ts'
import { quickDropdownPresets } from '../../features/sheet-validation/quick-validation.ts'
import QuickDropdownPicker from './QuickDropdownPicker.vue'
import SheetBorderSelector from './SheetBorderSelector.vue'
import SheetNumberFormatControl from './SheetNumberFormatControl.vue'
import SheetColorControl from './SheetColorControl.vue'
import { SHEET_STANDARD_COLORS, sheetColorHex } from '../../features/sheet-workbench/sheet-colors.ts'
import { SHEET_NUMBER_FORMATS, sheetNumberFormatMenu } from '../../features/sheet-workbench/sheet-number-format.ts'
import { numberFormatCurrency, sheetCurrencySymbol } from '../../features/sheet-workbench/sheet-currency.ts'
import { sheetBorderLabels, sheetBorderStyleLabels, type SheetBorderOptions, type SheetBorderPreset } from '../../features/sheet-workbench/sheet-toolbar-format.ts'
import { localizeRibbon, type OfficeTranslator } from '../../i18n/office-i18n.ts'

export interface SheetRibbonState {
  fontList?: OfficeFontList
  fontEditable?: boolean
  painterActive?: boolean
  painterUnavailable?: boolean
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
  quickValidationEditable?: boolean
  validationPanelOpen?: boolean
  hyperlinkAvailable?: boolean
  canSort: boolean
  filterEnabled?: boolean
  filterCanToggle?: boolean
  fontFamily: string | null
  fontSize: number | null
  numberFormat: string | null
  numberFormatPattern?: string | null
  region?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  wrap?: boolean
  wrapStrategy?: 'overflow' | 'wrap' | 'clip' | null
  color?: string | null
  background?: string | null
  lastTextColor?: string
  lastBackgroundColor?: string
  defaultTextColor?: string
  defaultBackgroundColor?: string
  borderColor?: string | null
  borderOptions?: SheetBorderOptions
  borderPreset?: SheetBorderPreset
  onBorderOptions?: (value: SheetBorderOptions) => void
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

const functionCategoryOrder = [
  '常用与统计', '财务', '逻辑', '文本', '日期与时间', '查找与引用', '数学', '统计', '工程', '信息', '数据库',
]
const supportedFunctionCategories = new Set(SHEET_FUNCTION_CATALOG.map(entry => entry.category))
const functionCategories = functionCategoryOrder.filter(category => supportedFunctionCategories.has(category))
const functionCategoryIcons: Record<string, Component> = {
  '常用与统计': FunctionSquare, 数学: Sigma, 逻辑: Workflow, 统计: ChartNoAxesColumnIncreasing,
  '查找与引用': SearchCode, 财务: BadgeDollarSign, 工程: Settings2, 信息: Info, 数据库: Database,
}
const textRotationOptions: Array<[number | 'v', string]> = [
  [0, '水平'], [-45, '左斜 45°'], [45, '右斜 45°'], ['v', '竖排'], [-90, '左转 90°'], [90, '右转 90°'],
]

/** The schema describes editor intents; UI, keyboard and menu callers use the same dispatcher. */
export function createSheetRibbon(state: SheetRibbonState, translate?: OfficeTranslator): RibbonTab[] {
  const button = (id: string, title: string, extra: Partial<RibbonItem> = {}): RibbonItem => ({
    id, title, commandId: id, type: 'button', disabled: !state.hasSheet, ...extra,
  })
  const quickDisabled = !state.dataValidationAvailable || state.quickValidationEditable === false
  const quick = (id: string, title: string, params: string, icon?: Component) => button(id, title, { commandId: 'sheet.operation.quick-data-validation', params, icon, disabled: quickDisabled })
  const tr: OfficeTranslator = translate ?? ((_key, fallback) => fallback)
  const numberFormatDisabled = !state.hasSheet || state.fontEditable === false
  const fontDisabled = !state.hasSheet || state.fontEditable === false
  const layoutDisabled = fontDisabled
  const selectedWrapStrategy = state.wrapStrategy !== undefined ? state.wrapStrategy : state.wrap === undefined ? null : state.wrap ? 'wrap' : 'overflow'
  const mixedLabel = tr('common.mixed', '混合')
  const mark = (id: string, title: string, icon: Component, value: boolean | undefined, column: number, shortcut: string) => button(id, title, {
    icon, activated: value === true, disabled: fontDisabled, shortcut,
    tooltip: value === undefined ? `${tr(`sheet.items.${id}`, title)} (${tr('sheet.color.mixed', '混合')})` : tr(`sheet.items.${id}`, title), gridLayout: { row: 2, column },
  })
  const palette = (id: string, title: string, field: 'color' | 'background', icon: Component, column: number): RibbonItem => {
    const commandId = field === 'color' ? 'sheet.command.set-color' : 'sheet.command.set-background'
    const resetCommandId = field === 'color' ? 'sheet.command.reset-text-color' : 'sheet.command.reset-background-color'
    const resetTitle = field === 'color' ? tr('sheet.color.reset-text', '重置文字颜色') : tr('sheet.color.no-fill', '无填充')
    const defaultColor = field === 'color' ? state.defaultTextColor : state.defaultBackgroundColor
    const lastColor = (field === 'color' ? state.lastTextColor : state.lastBackgroundColor) ?? sheetColorHex(defaultColor) ?? (field === 'color' ? '#000000' : '#3f83f8')
    const selectedColor = state[field]
    return {
      id, title, type: 'split', commandId, params: lastColor, value: lastColor, swatch: lastColor, icon, disabled: fontDisabled,
      gridLayout: { row: 2, column }, inlineComponent: SheetColorControl, popupCommandId: commandId,
      popupProps: { title: tr(`sheet.items.${id}`, title), field, icon, commandId, resetCommandId, resetTitle, defaultColor, lastColor, selectedColor, mixed: selectedColor === undefined },
      children: [button(resetCommandId, resetTitle, { icon: RemoveFormatting, activated: selectedColor === null, disabled: fontDisabled }),
        ...SHEET_STANDARD_COLORS.map(color => button(`${id}.${color}`, color.toUpperCase(), { commandId, params: color, swatch: color, activated: sheetColorHex(selectedColor) === color, disabled: fontDisabled }))],
    }
  }
  const currencySymbol = sheetCurrencySymbol(state.region)
  const currentPattern = state.numberFormatPattern !== undefined ? state.numberFormatPattern
    : state.numberFormat && Object.hasOwn(SHEET_NUMBER_FORMATS, state.numberFormat) ? SHEET_NUMBER_FORMATS[state.numberFormat as keyof typeof SHEET_NUMBER_FORMATS] : null
  const currencyIcon: Component = (props: { size?: number }) => h('span', { 'aria-hidden': 'true', style: { display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: `${props.size ?? 16}px`, height: `${props.size ?? 16}px`, fontSize: `${Math.min(props.size ?? 16, currencySymbol.length > 1 ? 12 : 18)}px`, fontFamily: 'Arial, sans-serif', lineHeight: 1 } }, currencySymbol)
  const presetValues = quickDropdownPresets(tr)
  const presets = (['dropdown-yes-no', 'dropdown-status', 'dropdown-abc'] as const).map(id => ({ id, values: presetValues[id] }))
  const editDropdown = tr('sheet.validation.quick.edit-dropdown', '编辑下拉列表')
  const clearDropdown = tr('sheet.validation.quick.clear-dropdown', '清除下拉列表')
  const tabs: RibbonTab[] = [
    {
      id: 'home', title: '开始', order: 0, groups: [
        { id: 'clipboard', title: '剪贴板', order: 0, items: [
          button('sheet.command.undo', '撤销', { icon: Undo2, disabled: !state.canUndo, shortcut: 'Ctrl+Z', gridLayout: { row: 1, column: 1, iconSize: 18 } }),
          button('sheet.command.redo', '重做', { icon: Redo2, disabled: !state.canRedo, shortcut: 'Ctrl+Y', gridLayout: { row: 2, column: 1, iconSize: 18 } }),
          button('sheet.operation.format-painter', '格式刷', { icon: Brush, activated: Boolean(state.painterActive), disabled: !state.painterActive && (!state.hasSheet || state.painterUnavailable === true), subId: 'sheet.operation.continuous-format-painter', gridLayout: { row: 1, column: 2, showLabel: false } }),
          button('sheet.command.clear-format', '清除格式', { icon: RemoveFormatting, gridLayout: { row: 2, column: 2, showLabel: false } }),
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
          { id: 'sheet.command.set-font-family', commandId: 'sheet.command.set-font-family', type: 'select', edit: { kind: 'font' }, title: '字体', width: 180, disabled: !state.hasSheet || state.fontEditable === false, value: state.fontFamily, placeholder: '混合', gridLayout: { row: 1, column: 1, columnSpan: 4, width: 180, minWidth: 120 }, options: ribbonFontOptions(state.fontFamily, state.fontList) },
          { id: 'sheet.command.set-font-size', commandId: 'sheet.command.set-font-size', type: 'select', edit: { kind: 'number', min: 6, max: 400 }, title: '字号', width: 72, disabled: !state.hasSheet || state.fontEditable === false, value: state.fontSize, placeholder: '混合', gridLayout: { row: 1, column: 5, width: 72 }, options: fontSizeOptions },
          button('sheet.command.increase-font-size', '增大字号', { icon: AArrowUp, disabled: !state.hasSheet || state.fontEditable === false, gridLayout: { row: 1, column: 6 } }),
          button('sheet.command.decrease-font-size', '减小字号', { icon: AArrowDown, disabled: !state.hasSheet || state.fontEditable === false, gridLayout: { row: 1, column: 7 } }),
          mark('sheet.command.toggle-bold', '加粗', Bold, state.bold, 1, 'Ctrl+B'),
          mark('sheet.command.toggle-italic', '斜体', Italic, state.italic, 2, 'Ctrl+I'),
          mark('sheet.command.toggle-underline', '下划线', Underline, state.underline, 3, 'Ctrl+U'),
          mark('sheet.command.toggle-strike', '删除线', Strikethrough, state.strike, 4, 'Ctrl+Shift+X'),
          palette('sheet.menu.text-color', '文字颜色', 'color', Palette, 7),
          palette('sheet.menu.fill', '填充颜色', 'background', PaintBucket, 6),
          { id: 'sheet.menu.borders', type: 'split', commandId: 'sheet.command.set-borders', title: '边框', icon: Grid2X2, params: state.borderPreset ?? 'all', disabled: fontDisabled, gridLayout: { row: 2, column: 5 },
            popupComponent: SheetBorderSelector, popupCommandId: 'sheet.command.set-borders', popupProps: {
              value: state.borderOptions ?? { style: 'thin', color: '#1f2937' }, preset: state.borderPreset ?? 'all', selectedColor: state.borderColor, mixed: state.borderColor === undefined, translate: tr, 'onUpdate:value': state.onBorderOptions,
            }, children: [
              ...Object.entries(sheetBorderLabels).map(([value, label]) => button(`sheet.menu.border.${value}`, label, { commandId: 'sheet.command.set-borders', params: value, disabled: fontDisabled })),
              { id: 'sheet.menu.border-style', type: 'menu', title: '边框样式', disabled: fontDisabled, children: Object.entries(sheetBorderStyleLabels).map(([value, label]) => button(`sheet.menu.border-style.${value}`, tr(`sheet.border.style.${value}`, label), { commandId: 'sheet.command.set-border-style', params: value, disabled: fontDisabled })) },
              { id: 'sheet.menu.border-color', type: 'menu', title: '边框颜色', disabled: fontDisabled, children: SHEET_STANDARD_COLORS.map(value => button(`sheet.menu.border-color.${value}`, value.toUpperCase(), { commandId: 'sheet.command.set-border-color', params: value, swatch: value, disabled: fontDisabled })) },
            ] },
        ] },
        { id: 'layout', title: '布局', order: 20, items: [
          { id: 'sheet.menu.horizontal-align', type: 'menu', title: '水平对齐', tooltip: state.align === undefined ? `水平对齐（${mixedLabel}）` : '水平对齐', icon: state.align === 'center' ? AlignCenter : state.align === 'right' ? AlignRight : AlignLeft, disabled: layoutDisabled, gridLayout: { row: 1, column: 1 }, children: [
            button('sheet.command.align-left', '左对齐', { commandId: 'sheet.command.set-align', params: 'left', icon: AlignLeft, activated: state.align === 'left', disabled: layoutDisabled }),
            button('sheet.command.align-center', '居中', { commandId: 'sheet.command.set-align', params: 'center', icon: AlignCenter, activated: state.align === 'center', disabled: layoutDisabled }),
            button('sheet.command.align-right', '右对齐', { commandId: 'sheet.command.set-align', params: 'right', icon: AlignRight, activated: state.align === 'right', disabled: layoutDisabled }),
          ] },
          { id: 'sheet.menu.vertical-align', type: 'menu', title: '垂直对齐', tooltip: state.vertical === undefined ? `垂直对齐（${mixedLabel}）` : '垂直对齐', icon: state.vertical === 'top' ? AlignVerticalJustifyStart : state.vertical === 'bottom' ? AlignVerticalJustifyEnd : AlignVerticalJustifyCenter, disabled: layoutDisabled, gridLayout: { row: 2, column: 1 }, children: [
            button('sheet.menu.vertical.top', '顶端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'top', icon: AlignVerticalJustifyStart, activated: state.vertical === 'top', disabled: layoutDisabled }),
            button('sheet.menu.vertical.middle', '垂直居中', { commandId: 'sheet.command.set-vertical-align', params: 'middle', icon: AlignVerticalJustifyCenter, activated: state.vertical === 'middle', disabled: layoutDisabled }),
            button('sheet.menu.vertical.bottom', '底端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'bottom', icon: AlignVerticalJustifyEnd, activated: state.vertical === 'bottom', disabled: layoutDisabled }),
          ] },
          { id: 'sheet.menu.text-wrap', type: 'menu', title: '文本换行', tooltip: selectedWrapStrategy === null ? `文本换行（${mixedLabel}）` : '文本换行', icon: WrapText, disabled: layoutDisabled, gridLayout: { row: 1, column: 2 }, children: [
            button('sheet.menu.text-wrap.overflow', '溢出', { commandId: 'sheet.command.set-wrap-strategy', params: 'overflow', activated: selectedWrapStrategy === 'overflow', disabled: layoutDisabled }),
            button('sheet.menu.text-wrap.wrap', '自动换行', { commandId: 'sheet.command.set-wrap-strategy', params: 'wrap', icon: WrapText, activated: selectedWrapStrategy === 'wrap', disabled: layoutDisabled }),
            button('sheet.menu.text-wrap.clip', '截断', { commandId: 'sheet.command.set-wrap-strategy', params: 'clip', activated: selectedWrapStrategy === 'clip', disabled: layoutDisabled }),
          ] },
          button('sheet.command.set-shrink-to-fit', '缩小字体填充', {
            icon: Shrink, params: !(state.shrinkToFit === true && state.wrap !== true),
            activated: state.shrinkToFit === true && state.wrap !== true, disabled: layoutDisabled || state.wrap === true,
            tooltip: state.shrinkToFit === undefined ? `缩小字体填充（${mixedLabel}）` : '缩小字体填充',
            disabledReason: state.wrap === true ? '关闭自动换行后可使用缩小字体填充' : undefined,
            gridLayout: { row: 2, column: 2 },
          }),
          { id: 'sheet.menu.text-rotation', type: 'menu', title: '文字方向', tooltip: state.rotation === null ? `文字方向（${mixedLabel}）` : '文字方向', icon: state.rotationVertical ? TextCursor : state.rotation === -45 ? RotateCcw : state.rotation === 45 ? RotateCw : state.rotation === -90 ? RotateCcw : state.rotation === 90 ? RotateCw : Baseline, disabled: layoutDisabled, gridLayout: { row: 1, column: 3 }, children: textRotationOptions.map(([angle, label]) => button(`sheet.menu.text-rotation.${angle}`, label, {
            commandId: 'sheet.command.set-text-rotation', params: angle,
            disabled: layoutDisabled,
            activated: angle === 'v' ? state.rotationVertical === true : state.rotationVertical !== true && state.rotation === angle,
            icon: angle === 'v' ? TextCursor : angle === 0 ? Baseline : angle === -45 ? RotateCcw : angle === 45 ? RotateCw : angle === -90 ? RotateCcw : RotateCw,
          })) },
          { id: 'sheet.menu.merge', type: 'menu', title: '合并单元格', icon: TableCellsMerge, disabled: layoutDisabled || !state.canMerge && !state.canUnmerge, activated: state.canUnmerge, gridLayout: { row: 2, column: 3 }, children: [
            button('sheet.command.merge-all', '全部合并', { commandId: 'sheet.command.merge', params: 'all', icon: TableCellsMerge, disabled: layoutDisabled || !state.canMerge }),
            button('sheet.command.merge-vertical', '垂直合并', { commandId: 'sheet.command.merge', params: 'vertical', icon: TableCellsMerge, disabled: layoutDisabled || !state.canMerge }),
            button('sheet.command.merge-horizontal', '水平合并', { commandId: 'sheet.command.merge', params: 'horizontal', icon: TableCellsMerge, disabled: layoutDisabled || !state.canMerge }),
            button('sheet.command.unmerge', '取消合并', { icon: TableCellsSplit, disabled: layoutDisabled || !state.canUnmerge }),
          ] },
        ] },
        { id: 'number', title: '数字', order: 30, items: [
          { id: 'sheet.command.set-number-format', commandId: 'sheet.command.set-number-format', type: 'select', title: '数字格式', width: 160, disabled: numberFormatDisabled,
            value: state.numberFormat, placeholder: '混合或自定义', gridLayout: { row: 1, column: 1, columnSpan: 4, width: 160 }, inlineComponent: SheetNumberFormatControl,
            popupCommandId: 'sheet.command.set-number-format', popupProps: { pattern: currentPattern, region: state.region, title: tr('sheet.number-format.title', '数字格式'), translate: tr },
            options: sheetNumberFormatMenu(state.region).flatMap(entry => entry.type === 'divider' ? [] : entry.type === 'more'
              ? [{ label: entry.label, value: entry.id, commandId: 'sheet.operation.open-custom-number-format' }]
              : [{ label: entry.label, value: entry.id, params: { pattern: entry.pattern } }]),
          },
          button('sheet.number-format.percent', '百分比格式', { commandId: 'sheet.command.set-percent-format', icon: Percent, disabled: numberFormatDisabled, activated: state.numberFormat === 'percent', gridLayout: { row: 2, column: 1 } }),
          button('sheet.number-format.currency', '货币格式', { commandId: 'sheet.command.set-currency-format', icon: currencyIcon, disabled: numberFormatDisabled,
            activated: currentPattern !== null && ['currency', 'accounting', 'roundingCurrency'].includes(state.numberFormat ?? '') && numberFormatCurrency(currentPattern) === currencySymbol, gridLayout: { row: 2, column: 2 } }),
          button('sheet.command.increase-decimals', '增加小数位', { icon: DecimalsArrowRight, disabled: numberFormatDisabled, gridLayout: { row: 2, column: 3 } }),
          button('sheet.command.decrease-decimals', '减少小数位', { icon: DecimalsArrowLeft, disabled: numberFormatDisabled, gridLayout: { row: 2, column: 4 } }),
        ] },
        { id: 'tools', title: '工具', order: 40, items: [
          button('sheet.operation.open-range-protection', '保护范围', { icon: ShieldCheck, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.command.clear-content', '清除内容', { icon: RemoveFormatting, showLabel: true, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
        ] },
      ],
    },
    {
      id: 'insert', title: '插入', order: 10, groups: [
        { id: 'media', title: '媒体', order: 10, items: [
          button('sheet.operation.insert-hyperlink', '超链接', { icon: Link, showLabel: true, disabled: !state.hyperlinkAvailable, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.insert-chart', '图表', { icon: ChartNoAxesCombined, order: 4, disabled: !state.chartAvailable, disabledReason: '图表需要绑定到已打开的工作簿' }),
          button('sheet.operation.open-sparkline', '迷你图', { icon: ChartLine, order: 5 }),
        ] },
        { id: 'cells', title: '单元格', order: 20, items: [
          { ...quick('sheet.operation.quick-data-validation.checkbox', '复选框', 'checkbox', SquareCheck), type: 'split', children: [
            quick('sheet.menu.quick-validation.set-checkbox', tr('sheet.validation.quick.set-checkbox', '设置复选框'), 'checkbox'),
            quick('sheet.menu.quick-validation.clear-checkbox', tr('sheet.validation.quick.clear-checkbox', '清除复选框'), 'clear-checkbox'),
          ] },
          { id: 'sheet.operation.quick-data-validation.dropdown', type: 'menu', title: '下拉列表', icon: ListFilter, disabled: quickDisabled,
            popupComponent: QuickDropdownPicker, popupCommandId: 'sheet.operation.quick-data-validation',
            popupProps: { presets, editDropdown, clearDropdown, title: tr('sheet.validation.quick.presets', '下拉预设') }, children: [
              ...presets.map(p => quick(`sheet.menu.quick-validation.${p.id}`, p.values.join(' / '), p.id)),
              quick('sheet.menu.quick-validation.edit-dropdown', editDropdown, 'edit-dropdown'),
              quick('sheet.menu.quick-validation.clear-dropdown', clearDropdown, 'clear-dropdown'),
            ] },
          { ...quick('sheet.operation.quick-data-validation.date', '日期', 'date', CalendarDays), type: 'split', children: [
            quick('sheet.menu.quick-data-validation.date', '日期', 'date', CalendarDays),
            quick('sheet.operation.quick-data-validation.date-time', '日期和时间', 'date-time', CalendarDays),
            quick('sheet.menu.quick-validation.edit-date', tr('sheet.validation.quick.edit-date', '编辑日期规则'), 'edit-date'),
            quick('sheet.menu.quick-validation.clear-date', tr('sheet.validation.quick.clear-date', '移除日期规则'), 'clear-date'),
          ] },
        ] },
      ],
    },
    {
      id: 'data', title: '数据', order: 30, groups: [
        { id: 'organization', title: '数据整理', order: 10, items: [
          { id: 'sheet.menu.sort', type: 'menu', title: '排序', icon: ArrowDownWideNarrow, disabled: !state.hasSheet, gridLayout: { row: 1, column: 1 }, children: [
            button('sheet.command.sort-ascending', '升序', { icon: ArrowUpAZ, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围' }),
            button('sheet.command.sort-ascending-extended', '扩展选定区域升序', { commandId: 'sheet.command.sort-ascending', params: { extendSelection: true }, icon: ArrowUpAZ, disabled: !state.hasSheet }),
            button('sheet.command.sort-descending', '降序', { icon: ArrowDownAZ, disabled: !state.canSort, disabledReason: '请选择至少两行的数据范围' }),
            button('sheet.command.sort-descending-extended', '扩展选定区域降序', { commandId: 'sheet.command.sort-descending', params: { extendSelection: true }, icon: ArrowDownAZ, disabled: !state.hasSheet }),
            button('sheet.operation.open-data-panel', '自定义排序', { params: 'sort', icon: ArrowDownWideNarrow, disabled: !state.hasSheet }),
          ] },
          { ...button('sheet.data.filter', '筛选', { commandId: 'sheet.command.toggle-auto-filter', icon: Filter, showLabel: true, activated: state.filterEnabled, disabled: !state.filterCanToggle, disabledReason: translate?.('sheet.filter.select-range', '请选择包含标题行和数据行的区域后再启用筛选'), gridLayout: { row: 2, column: 1 } }), type: 'split', children: [
            button('sheet.menu.filter.clear', '清除筛选', { commandId: 'sheet.command.clear-filter', icon: FilterX, disabled: !state.filterEnabled }),
            button('sheet.menu.filter.reapply', '重新应用', { commandId: 'sheet.command.reapply-filter', disabled: !state.filterEnabled }),
            button('sheet.menu.filter.panel', '筛选条件', { commandId: 'sheet.operation.open-data-panel', params: 'filter', disabled: !state.filterEnabled }),
          ] },
        ] },
        { id: 'rules', title: '规则', order: 20, items: [
          { id: 'sheet.menu.data-validation', type: 'menu', title: '数据验证', icon: ClipboardCheck, disabled: quickDisabled, gridLayout: { row: 1, column: 1 }, children: [
            button('sheet.operation.open-data-validation', tr('sheet.validation.manage-rules', '管理验证规则'), { disabled: quickDisabled, activated: state.validationPanelOpen }),
            button('sheet.operation.add-data-validation', tr('sheet.validation.add', '添加规则'), { disabled: quickDisabled }),
          ] },
          button('sheet.operation.open-conditional-formatting', '条件格式', { icon: ChartColumnIncreasing, showLabel: true, gridLayout: { row: 2, column: 1, showLabel: true } }),
        ] },
        { id: 'pivot', title: '分析', order: 30, items: [
          button('sheet.operation.open-pivot', '数据透视表', { icon: TableProperties, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 1, showLabel: true } }),
          button('sheet.operation.open-pivot-chart', '透视图', { icon: ChartNoAxesCombined, showLabel: true, disabled: !state.pivotChartAvailable, disabledReason: '请先创建数据透视表', gridLayout: { row: 2, column: 1, showLabel: true } }),
          button('sheet.operation.open-slicer', '切片器', { icon: ListFilter, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 2, showLabel: true } }),
          button('sheet.operation.open-timeline', '时间线', { icon: Clock, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 2, column: 2, showLabel: true } }),
        ] },
        { id: 'outline', title: '分组显示', order: 40, items: [
          button('sheet.operation.outline-group-rows', '组合行', { icon: Group, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-toggle-rows', '折叠/展开行组', { icon: Grid2X2, showLabel: true, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-remove-rows', '取消行分组', { icon: Ungroup, showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-group-columns', '组合列', { icon: Group, showLabel: true, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-toggle-columns', '折叠/展开列组', { icon: Grid2X2, showLabel: true, gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.outline-remove-columns', '取消列分组', { icon: Ungroup, showLabel: true, gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true } }),
          { id: 'sheet.operation.outline-summary-rows', commandId: 'sheet.operation.outline-summary-position', type: 'select', title: '汇总行位置', width: 112, value: state.outlineSummaryAbove ? 'above' : 'below', valueParams: value => ({ axis: 'row', atStart: value === 'above' }), gridLayout: { row: 1, column: 7, rowSpan: 2, width: 112, showLabel: true }, options: [
            { label: '汇总行在下方', value: 'below' }, { label: '汇总行在上方', value: 'above' },
          ] },
          { id: 'sheet.operation.outline-summary-columns', commandId: 'sheet.operation.outline-summary-position', type: 'select', title: '汇总列位置', width: 112, value: state.outlineSummaryLeft ? 'left' : 'right', valueParams: value => ({ axis: 'column', atStart: value === 'left' }), gridLayout: { row: 1, column: 8, rowSpan: 2, width: 112, showLabel: true }, options: [
            { label: '汇总列在右侧', value: 'right' }, { label: '汇总列在左侧', value: 'left' },
          ] },
        ] },
        { id: 'data-tools', title: '工具', order: 50, items: [
          button('sheet.operation.find', '查找', { icon: Search, showLabel: true, shortcut: 'Ctrl+F', gridLayout: { row: 1, column: 1, rowSpan: 2, columnSpan: 2, showLabel: true } }),
        ] },
      ],
    },
    {
      id: 'formulas', title: '公式', order: 20, groups: [
        { id: 'formula-edit', title: '公式编辑', items: [
          ...functionCategories.map((category, index) => ({
            id: `sheet.operation.open-function.category-${index}`,
            type: 'menu' as const,
            title: category,
            icon: functionCategoryIcons[category] ?? FunctionSquare,
            disabled: !state.hasSheet,
            gridLayout: index === 0 ? { row: 1, column: 1, rowSpan: 2, showLabel: false } : undefined,
            children: [...SHEET_FUNCTION_CATALOG.filter(entry => entry.category === category).map(entry => button(
              `sheet.function.${entry.name}`,
              entry.name,
              { commandId: 'sheet.operation.insert-function', params: entry.name },
            )), ...(index === 0 ? [button('sheet.operation.open-function', '插入函数', { icon: FunctionSquare })] : [])],
          })),
        ] },
        { id: 'calculation-options', title: '计算选项', order: 10, items: [
          { id: 'sheet.command.set-calculation-mode', commandId: 'sheet.command.set-calculation-mode', type: 'select', title: '计算模式', width: 116, value: state.calculationMode ?? 'automatic', gridLayout: { row: 1, column: 1, rowSpan: 2, columnSpan: 2, width: 116, showLabel: true }, options: [
            { label: '自动', value: 'automatic' }, { label: '手动', value: 'manual' },
          ] },
          button('sheet.operation.start-formula', '编辑公式', { icon: FunctionSquare, gridLayout: { row: 1, column: 3 } }),
          button('sheet.command.recalculate', '重新计算', { icon: Calculator, gridLayout: { row: 2, column: 3 } }),
        ] },
      ],
    },
    {
      id: 'view', title: '视图', order: 40, groups: [
        { id: 'display', title: '显示', order: 0, items: [
          { ...button('sheet.command.freeze-selection', '冻结窗格', { icon: Pin, showLabel: true, activated: state.frozen, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }), type: 'split', children: [
            button('sheet.menu.freeze.selection', '冻结到当前单元格', { commandId: 'sheet.command.freeze-selection', icon: Pin }),
            button('sheet.command.freeze-first-row', '冻结首行', { icon: Pin }),
            button('sheet.command.freeze-first-column', '冻结首列', { icon: Pin }),
            button('sheet.command.unfreeze', '取消冻结', { icon: PinOff, disabled: !state.frozen }),
          ] },
          { id: 'sheet.operation.set-zoom', commandId: 'sheet.operation.set-zoom', type: 'select', title: '缩放', width: 88, value: state.zoom, valueParams: value => Number(value), gridLayout: { row: 1, column: 2, width: 88, showLabel: true }, options: [50, 75, 100, 125, 150, 175, 200, 400].map(value => ({ label: `${value}%`, value })) },
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
        { id: 'visibility', title: '可见性', order: 10, items: [
          button('sheet.command.toggle-gridlines', '网格线', { icon: Grid2X2, showLabel: true, activated: state.gridlines, gridLayout: { row: 1, column: 1, showLabel: true } }),
          button('sheet.operation.toggle-fullscreen', state.fullscreen ? '退出全屏' : '全屏', { icon: state.fullscreen ? Minimize : Expand, showLabel: true, activated: state.fullscreen, gridLayout: { row: 2, column: 1, showLabel: true } }),
        ] },
      ],
    },
  ]
  return translate ? localizeRibbon(tabs, 'sheet', translate) : tabs
}
