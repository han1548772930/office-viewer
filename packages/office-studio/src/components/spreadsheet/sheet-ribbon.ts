import { h, type Component } from 'vue'
import {
  AArrowDown, AArrowUp, AlignCenter, AlignLeft, AlignRight, AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd, AlignVerticalJustifyStart, ArrowDownAZ, ArrowDownWideNarrow, ArrowUpAZ,
  BadgeDollarSign, Baseline, Bold, Brush, Calculator, CalendarDays,
  ChartColumnIncreasing, ChartLine, ChartNoAxesColumnIncreasing, ChartNoAxesCombined,
  ClipboardCheck, ClipboardPaste, Clock, Crosshair, DecimalsArrowLeft, DecimalsArrowRight, Expand,
  Database, Filter, FilterX, FunctionSquare, Grid2X2, Group, Image, ImagePlus, Info, Italic, Link, ListFilter, Minimize, PaintBucket, Palette,
  Pin, PinOff, Redo2, RemoveFormatting, RotateCcw,
  RotateCw, Search, SearchCode, Settings2, ShieldCheck, Shrink, Sigma, SquareCheck, Strikethrough,
  TableCellsMerge, TableCellsSplit, TableProperties, TextCursor, Underline, Undo2, Ungroup, Workflow,
  WrapText, Percent, ArrowUpRight, CircleDot, Star,
} from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema.ts'
import { fontSizeOptions, ribbonFontOptions, type OfficeFontList } from '../../platform/ribbon/ribbon-fonts.ts'
import { SHEET_COMMON_FUNCTION_CATALOG, SHEET_FUNCTION_CATALOG } from '../../features/sheet-workbench/sheet-function-catalog.ts'
import { quickDropdownPresets } from '../../features/sheet-validation/quick-validation.ts'
import QuickDropdownPicker from './QuickDropdownPicker.vue'
import SheetBorderSelector from './SheetBorderSelector.vue'
import SheetBorderIcon from './SheetBorderIcon.vue'
import SheetNumberFormatControl from './SheetNumberFormatControl.vue'
import SheetColorControl from './SheetColorControl.vue'
import { SHEET_STANDARD_COLORS, sheetColorHex } from '../../features/sheet-workbench/sheet-colors.ts'
import { SHEET_NUMBER_FORMATS, sheetNumberFormatMenu } from '../../features/sheet-workbench/sheet-number-format.ts'
import { numberFormatCurrency, sheetCurrencySymbol } from '../../features/sheet-workbench/sheet-currency.ts'
import { sheetBorderLabels, sheetBorderStyleLabels, type SheetBorderOptions, type SheetBorderPreset } from '../../features/sheet-workbench/sheet-toolbar-format.ts'
import { localizeRibbon, type OfficeTranslator } from '../../i18n/office-i18n.ts'
import { sheetClipboardCommandCapability, type SheetClipboardCapabilities } from '../../features/sheet-workbench/sheet-clipboard-capabilities.ts'
import { OPEN_SHEET_RANGE_THEME_OPERATION } from '../../features/sheet-workbench/sheet-range-theme-command.ts'
import { SHEET_IMAGE_COMMANDS } from '../../features/sheet-images/sheet-image-commands.ts'
import SlicerQuickStylePicker from './SlicerQuickStylePicker.vue'
import { createSheetOutlineContextMenuItem, createSheetOutlineRibbonItem } from '../../features/sheet-outline/sheet-outline-menu.ts'
import type { SheetOutlineCommandContext } from '../../features/sheet-outline/sheet-outline-command.ts'
import { getSheetOutlineUpdate, type SheetOutlineAction } from '../../features/sheet-outline/sheet-outline.ts'
import { sheetOutlineErrorMessage } from '../../features/sheet-outline/sheet-outline-types.ts'
import { OPEN_SHEET_OUTLINE_SETTINGS_OPERATION } from './use-sheet-outline-workbench.ts'
import { QUICK_CONDITIONAL_FORMAT_COMMANDS, QUICK_CONDITIONAL_ICON_SETS, type QuickConditionalFormatSelection } from '../../features/conditional-formatting/quick-conditional-formatting.ts'
import { featureSearchRibbonItem, shortcutPanelRibbonItem } from '../../platform/ribbon/feature-search'

export interface SlicerRibbonControlState {
  id: string
  type: 'slicer' | 'timeline'
  style?: string
  styleOwner?: import('../../features/slicer/slicer-style-catalog.ts').SlicerStyleOwner
  supportsConnections: boolean
  canConfigure: boolean
  revision: number
}

export interface SheetRibbonState {
  clipboardCapabilities?: SheetClipboardCapabilities
  fontList?: OfficeFontList
  fontEditable?: boolean
  clearFormatEditable?: boolean
  clearContentEditable?: boolean
  mergeEditable?: boolean
  numberFormatEditable?: boolean
  borderEditable?: boolean
  layoutEditable?: boolean
  formulaEditable?: boolean
  rangeProtectionEditable?: boolean
  rangeThemeEditable?: boolean
  imageEditable?: boolean
  backgroundImageEditable?: boolean
  hasBackgroundImage?: boolean
  painterActive?: boolean
  painterUnavailable?: boolean
  canUndo: boolean
  canRedo: boolean
  hasSheet: boolean
  canMerge: boolean
  canUnmerge: boolean
  mergeActive?: boolean
  drawingFocused?: boolean
  canDeleteRow: boolean
  canDeleteColumn: boolean
  chartAvailable: boolean
  pivotChartAvailable?: boolean
  dataValidationAvailable: boolean
  quickValidationEditable?: boolean
  quickConditional?: QuickConditionalFormatSelection
  quickConditionalEditable?: boolean
  quickConditionalWorksheetEditable?: boolean
  shortcutPanelOpen?: boolean
  validationPanelOpen?: boolean
  hyperlinkAvailable?: boolean
  canSort: boolean
  filterEnabled?: boolean
  filterHasCriteria?: boolean
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
  outlineContext?: SheetOutlineCommandContext
  outlineSettingsAvailable?: boolean
  crosshairEnabled?: boolean
  crosshairColor?: string
  slicerControl?: SlicerRibbonControlState
}

const functionCategoryOrder = [
  '常用与统计', '财务', '逻辑', '文本', '日期与时间', '查找与引用', '数学', '统计', '工程', '信息', '数据库',
]
const supportedFunctionCategories = new Set(SHEET_FUNCTION_CATALOG.map(entry => entry.category))
const functionCategories = functionCategoryOrder.filter(category => category === '常用与统计' || supportedFunctionCategories.has(category))
const functionCategoryIcons: Record<string, Component> = {
  '常用与统计': FunctionSquare, 数学: Sigma, 逻辑: Workflow, 统计: ChartNoAxesColumnIncreasing,
  文本: TextCursor, '日期与时间': CalendarDays,
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
  const clipboardButton = (id: string, title: string, extra: Partial<RibbonItem> = {}): RibbonItem => {
    const entry = state.clipboardCapabilities && sheetClipboardCommandCapability(state.clipboardCapabilities, extra.commandId ?? id)
    return button(id, title, { ...extra, ...(entry ? { hidden: entry.hidden, disabled: !state.hasSheet || entry.disabled, disabledReason: entry.reason } : {}) })
  }
  const pasteRoot = state.clipboardCapabilities?.ribbonPaste
  // When the default write is unavailable, retain access to permitted special modes through the menu.
  const pasteType = state.clipboardCapabilities?.paste.all.disabled && !pasteRoot?.disabled ? 'menu' : 'split'
  const quickDisabled = !state.dataValidationAvailable || state.quickValidationEditable === false
  const quick = (id: string, title: string, params: string, icon?: Component) => button(id, title, { commandId: 'sheet.operation.quick-data-validation', params, icon, disabled: quickDisabled })
  const tr: OfficeTranslator = translate ?? ((_key, fallback) => fallback)
  const numberFormatDisabled = !state.hasSheet || state.fontEditable === false || state.numberFormatEditable === false
  const fontDisabled = !state.hasSheet || state.fontEditable === false
  const borderDisabled = fontDisabled || state.borderEditable === false
  const formulaDisabled = !state.hasSheet || state.formulaEditable === false || state.drawingFocused === true
  const layoutDisabled = fontDisabled || state.layoutEditable === false
  const mergeDisabled = layoutDisabled || state.mergeEditable === false
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
  const currencyIcon: Component = (props: { size?: number }) => h('svg', {
    'aria-hidden': 'true', focusable: 'false', viewBox: '0 0 24 24', width: props.size ?? 16, height: props.size ?? 16,
  }, [h('text', {
    x: 12, y: 12, fill: 'currentColor', 'text-anchor': 'middle', 'dominant-baseline': 'central',
    'font-size': currencySymbol.length > 1 || currencySymbol === '﷼' ? 16 : 20, 'font-family': 'Arial, sans-serif',
    textLength: currencySymbol.length > 1 || currencySymbol === '﷼' ? 20 : undefined,
    lengthAdjust: 'spacingAndGlyphs',
  }, currencySymbol)])
  const borderIcon: Component = (props: { size?: number }) => h(SheetBorderIcon, { preset: state.borderPreset ?? 'all', size: props.size })
  const presetValues = quickDropdownPresets(tr)
  const presets = (['dropdown-yes-no', 'dropdown-status', 'dropdown-abc'] as const).map(id => ({ id, values: presetValues[id] }))
  const editDropdown = tr('sheet.validation.quick.edit-dropdown', '编辑下拉列表')
  const clearDropdown = tr('sheet.validation.quick.clear-dropdown', '清除下拉列表')
  const { dataBar, iconSet, clear: clearConditional } = QUICK_CONDITIONAL_FORMAT_COMMANDS
  const quickConditionalDisabled = !state.hasSheet || (state.quickConditionalEditable ?? state.fontEditable) === false
  const quickConditionalWorksheetDisabled = !state.hasSheet || (state.quickConditionalWorksheetEditable ?? state.quickConditionalEditable ?? state.fontEditable) === false
  const dataBarLabel = tr('conditional.quick.data-bar', '数据条')
  const iconSetLabel = tr('conditional.quick.icon-set', '图标集')
  const newConditionalLabel = tr('conditional.quick.create', '新建条件格式')
  const clearConditionalLabel = tr('conditional.quick.clear', '清除所选区域的规则')
  const quickIconSetStyles = {
    '3Stars': { label: tr('conditional.quick.rank', '等级'), icon: Star },
    '3TrafficLights1': { label: tr('conditional.quick.shape', '形状'), icon: CircleDot },
    '3Arrows': { label: tr('conditional.quick.direction', '方向'), icon: ArrowUpRight },
  }
  const newConditional = 'sheet.operation.open-conditional-formatting'
  const outlineContext = state.hasSheet ? state.outlineContext : undefined
  const rowOutlineMenu = createSheetOutlineContextMenuItem(outlineContext, 'row')
  const columnOutlineMenu = createSheetOutlineContextMenuItem(outlineContext, 'column')
  const outlineActionState = (axis: 'row' | 'column', kind: 'toggle' | 'collapse' | 'expand' | 'summary-position', atStart = false) => {
    const selection = outlineContext?.selections?.length === 1 ? outlineContext.selections[0] : undefined
    if (!outlineContext || kind !== 'summary-position' && !selection) return { disabled: true }
    const action: SheetOutlineAction = kind === 'summary-position' ? { kind, axis, atStart } : {
      kind, axis, start: axis === 'row' ? selection!.startRow : selection!.startCol,
      end: axis === 'row' ? selection!.endRow : selection!.endCol,
    }
    try {
      return { disabled: !getSheetOutlineUpdate(outlineContext.sheet, action, outlineContext.userId) }
    } catch (error) {
      return { disabled: true, disabledReason: sheetOutlineErrorMessage(error) }
    }
  }
  const outlineItems: RibbonItem[] = [
    createSheetOutlineRibbonItem(outlineContext),
    { id: 'sheet.menu.outline-details', type: 'menu', title: '分组显示', icon: Grid2X2, order: 4, disabled: rowOutlineMenu.disabled,
      gridLayout: { row: 1, column: 2 }, children: [
        { ...rowOutlineMenu.children![0]!, title: '组合行', icon: Group },
        { ...columnOutlineMenu.children![0]!, title: '组合列', icon: Group },
        { ...rowOutlineMenu.children![1]!, title: '取消行分组', icon: Ungroup },
        { ...columnOutlineMenu.children![1]!, title: '取消列分组', icon: Ungroup },
        button('sheet.operation.outline-toggle-rows', '折叠/展开行组', { icon: Grid2X2, ...outlineActionState('row', 'toggle') }),
        button('sheet.operation.outline-toggle-columns', '折叠/展开列组', { icon: Grid2X2, ...outlineActionState('column', 'toggle') }),
        button('sheet.operation.outline-collapse-rows', '隐藏行明细', outlineActionState('row', 'collapse')),
        button('sheet.operation.outline-expand-rows', '显示行明细', outlineActionState('row', 'expand')),
        button('sheet.operation.outline-collapse-columns', '隐藏列明细', outlineActionState('column', 'collapse')),
        button('sheet.operation.outline-expand-columns', '显示列明细', outlineActionState('column', 'expand')),
        { id: 'sheet.menu.outline-summary-rows', type: 'menu', title: '汇总行位置', children: [
          button('sheet.menu.outline-summary-below', '汇总行在下方', { commandId: 'sheet.command.set-outline-summary-position', params: { axis: 'row', atStart: false }, activated: !state.outlineSummaryAbove, ...outlineActionState('row', 'summary-position', false) }),
          button('sheet.menu.outline-summary-above', '汇总行在上方', { commandId: 'sheet.command.set-outline-summary-position', params: { axis: 'row', atStart: true }, activated: state.outlineSummaryAbove, ...outlineActionState('row', 'summary-position', true) }),
        ] },
        { id: 'sheet.menu.outline-summary-columns', type: 'menu', title: '汇总列位置', children: [
          button('sheet.menu.outline-summary-right', '汇总列在右侧', { commandId: 'sheet.command.set-outline-summary-position', params: { axis: 'column', atStart: false }, activated: !state.outlineSummaryLeft, ...outlineActionState('column', 'summary-position', false) }),
          button('sheet.menu.outline-summary-left', '汇总列在左侧', { commandId: 'sheet.command.set-outline-summary-position', params: { axis: 'column', atStart: true }, activated: state.outlineSummaryLeft, ...outlineActionState('column', 'summary-position', true) }),
        ] },
      ] },
    button(OPEN_SHEET_OUTLINE_SETTINGS_OPERATION, '大纲设置', { icon: Settings2, order: 5, disabled: !state.outlineSettingsAvailable,
      showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
  ]
  const tabs: RibbonTab[] = [
    {
      id: 'home', title: '开始', order: 0, groups: [
        { id: 'clipboard', title: '剪贴板', order: 0, items: [
          button('sheet.command.undo', '撤销', { icon: Undo2, disabled: !state.canUndo, shortcut: 'Ctrl+Z', gridLayout: { row: 1, column: 1, iconSize: 18 } }),
          button('sheet.command.redo', '重做', { icon: Redo2, disabled: !state.canRedo, shortcut: 'Ctrl+Y', gridLayout: { row: 2, column: 1, iconSize: 18 } }),
          button('sheet.operation.format-painter', '格式刷', { icon: Brush, activated: Boolean(state.painterActive), disabled: !state.painterActive && (!state.hasSheet || state.painterUnavailable === true), subId: 'sheet.operation.continuous-format-painter', gridLayout: { row: 1, column: 2, showLabel: false } }),
          button('sheet.command.clear-format', '清除格式', { icon: RemoveFormatting, disabled: !state.hasSheet || state.clearFormatEditable === false, gridLayout: { row: 2, column: 2, showLabel: false } }),
          { ...button('sheet.command.paste', '粘贴', { icon: ClipboardPaste, shortcut: 'Ctrl+V',
            ...(pasteRoot ? { hidden: pasteRoot.hidden, disabled: !state.hasSheet || pasteRoot.disabled, disabledReason: pasteRoot.reason } : {}),
            gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }), type: pasteType, children: [
            clipboardButton('sheet.menu.paste-all', '保留源格式粘贴', { commandId: 'sheet.command.paste' }),
            clipboardButton('sheet.menu.paste-values', '仅粘贴值', { commandId: 'sheet.command.paste-values' }),
            clipboardButton('sheet.menu.paste-formulas', '仅粘贴公式', { commandId: 'sheet.command.paste-formulas' }),
            clipboardButton('sheet.menu.paste-formats', '仅粘贴格式', { commandId: 'sheet.command.paste-formats' }),
            clipboardButton('sheet.menu.paste-column-widths', '仅粘贴列宽', { commandId: 'sheet.command.paste-column-widths' }),
            clipboardButton('sheet.menu.paste-except-borders', '保留目标边框粘贴', { commandId: 'sheet.command.paste-except-borders' }),
          ] },
        ] },
        { id: 'font', title: '字体', order: 10, items: [
          { id: 'sheet.command.set-font-family', commandId: 'sheet.command.set-font-family', type: 'select', edit: { kind: 'font' }, title: '字体', width: 180, disabled: !state.hasSheet || state.fontEditable === false, value: state.fontFamily, placeholder: '混合', gridLayout: { row: 1, column: 1, columnSpan: 4, width: 180, minWidth: 120 }, options: ribbonFontOptions(state.fontFamily, state.fontList) },
          { id: 'sheet.command.set-font-size', commandId: 'sheet.command.set-font-size', type: 'select', edit: { kind: 'number', min: 6, max: 400 }, title: '字号', width: 72, disabled: !state.hasSheet || state.fontEditable === false, value: state.fontSize, placeholder: '混合', gridLayout: { row: 1, column: 5, width: 72 }, options: fontSizeOptions },
          button('sheet.command.increase-font-size', '增大字号', { icon: AArrowUp, shortcut: 'Ctrl+Shift+.', disabled: !state.hasSheet || state.fontEditable === false, gridLayout: { row: 1, column: 6 } }),
          button('sheet.command.decrease-font-size', '减小字号', { icon: AArrowDown, shortcut: 'Ctrl+Shift+,', disabled: !state.hasSheet || state.fontEditable === false, gridLayout: { row: 1, column: 7 } }),
          mark('sheet.command.toggle-bold', '加粗', Bold, state.bold, 1, 'Ctrl+B'),
          mark('sheet.command.toggle-italic', '斜体', Italic, state.italic, 2, 'Ctrl+I'),
          mark('sheet.command.toggle-underline', '下划线', Underline, state.underline, 3, 'Ctrl+U'),
          mark('sheet.command.toggle-strike', '删除线', Strikethrough, state.strike, 4, 'Ctrl+Shift+X'),
          palette('sheet.menu.text-color', '文字颜色', 'color', Palette, 7),
          palette('sheet.menu.fill', '填充颜色', 'background', PaintBucket, 6),
          { id: 'sheet.menu.borders', type: 'split', commandId: 'sheet.command.set-borders', title: '边框', icon: borderIcon, params: state.borderPreset ?? 'all', hidden: state.drawingFocused, disabled: borderDisabled, gridLayout: { row: 2, column: 5 },
            popupComponent: SheetBorderSelector, popupCommandId: 'sheet.command.set-borders', popupProps: {
              value: state.borderOptions ?? { style: 'thin', color: '#1f2937' }, preset: state.borderPreset ?? 'all', selectedColor: state.borderColor, mixed: state.borderColor === undefined, translate: tr, 'onUpdate:value': state.onBorderOptions,
            }, children: [
              ...Object.entries(sheetBorderLabels).map(([value, label]) => button(`sheet.menu.border.${value}`, label, { commandId: 'sheet.command.set-borders', params: value, disabled: borderDisabled })),
              { id: 'sheet.menu.border-style', type: 'menu', title: '边框样式', disabled: borderDisabled, children: Object.entries(sheetBorderStyleLabels).map(([value, label]) => button(`sheet.menu.border-style.${value}`, tr(`sheet.border.style.${value}`, label), { commandId: 'sheet.command.set-border-style', params: value, disabled: borderDisabled })) },
              { id: 'sheet.menu.border-color', type: 'menu', title: '边框颜色', disabled: borderDisabled, children: SHEET_STANDARD_COLORS.map(value => button(`sheet.menu.border-color.${value}`, value.toUpperCase(), { commandId: 'sheet.command.set-border-color', params: value, swatch: value, disabled: borderDisabled })) },
            ] },
        ] },
        { id: 'layout', title: '布局', order: 20, items: [
          { id: 'sheet.menu.horizontal-align', type: 'menu', title: '水平对齐', tooltip: state.align === undefined ? `水平对齐（${mixedLabel}）` : '水平对齐', icon: state.align === 'center' ? AlignCenter : state.align === 'right' ? AlignRight : AlignLeft, hidden: state.drawingFocused, disabled: layoutDisabled, gridLayout: { row: 1, column: 1 }, children: [
            button('sheet.command.align-left', '左对齐', { commandId: 'sheet.command.set-align', params: 'left', icon: AlignLeft, activated: state.align === 'left', disabled: layoutDisabled }),
            button('sheet.command.align-center', '居中', { commandId: 'sheet.command.set-align', params: 'center', icon: AlignCenter, activated: state.align === 'center', disabled: layoutDisabled }),
            button('sheet.command.align-right', '右对齐', { commandId: 'sheet.command.set-align', params: 'right', icon: AlignRight, activated: state.align === 'right', disabled: layoutDisabled }),
          ] },
          { id: 'sheet.menu.vertical-align', type: 'menu', title: '垂直对齐', tooltip: state.vertical === undefined ? `垂直对齐（${mixedLabel}）` : '垂直对齐', icon: state.vertical === 'top' ? AlignVerticalJustifyStart : state.vertical === 'bottom' ? AlignVerticalJustifyEnd : AlignVerticalJustifyCenter, hidden: state.drawingFocused, disabled: layoutDisabled, gridLayout: { row: 2, column: 1 }, children: [
            button('sheet.menu.vertical.top', '顶端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'top', icon: AlignVerticalJustifyStart, activated: state.vertical === 'top', disabled: layoutDisabled }),
            button('sheet.menu.vertical.middle', '垂直居中', { commandId: 'sheet.command.set-vertical-align', params: 'middle', icon: AlignVerticalJustifyCenter, activated: state.vertical === 'middle', disabled: layoutDisabled }),
            button('sheet.menu.vertical.bottom', '底端对齐', { commandId: 'sheet.command.set-vertical-align', params: 'bottom', icon: AlignVerticalJustifyEnd, activated: state.vertical === 'bottom', disabled: layoutDisabled }),
          ] },
          { id: 'sheet.menu.text-wrap', type: 'menu', title: '文本换行', tooltip: selectedWrapStrategy === null ? `文本换行（${mixedLabel}）` : '文本换行', icon: WrapText, hidden: state.drawingFocused, disabled: layoutDisabled, gridLayout: { row: 1, column: 2 }, children: [
            button('sheet.menu.text-wrap.overflow', '溢出', { commandId: 'sheet.command.set-wrap-strategy', params: 'overflow', activated: selectedWrapStrategy === 'overflow', disabled: layoutDisabled }),
            button('sheet.menu.text-wrap.wrap', '自动换行', { commandId: 'sheet.command.set-wrap-strategy', params: 'wrap', icon: WrapText, activated: selectedWrapStrategy === 'wrap', disabled: layoutDisabled }),
            button('sheet.menu.text-wrap.clip', '截断', { commandId: 'sheet.command.set-wrap-strategy', params: 'clip', activated: selectedWrapStrategy === 'clip', disabled: layoutDisabled }),
          ] },
          button('sheet.command.set-shrink-to-fit', '缩小字体填充', {
            icon: Shrink,
            activated: state.shrinkToFit === true && state.wrap !== true, disabled: layoutDisabled || state.wrap === true,
            tooltip: state.shrinkToFit === undefined ? `缩小字体填充（${mixedLabel}）` : '缩小字体填充',
            disabledReason: state.wrap === true ? '关闭自动换行后可使用缩小字体填充' : undefined,
            gridLayout: { row: 2, column: 2 },
          }),
          { id: 'sheet.menu.text-rotation', type: 'menu', title: '文字方向', tooltip: state.rotation === null ? `文字方向（${mixedLabel}）` : '文字方向', icon: state.rotationVertical ? TextCursor : state.rotation === -45 ? RotateCcw : state.rotation === 45 ? RotateCw : state.rotation === -90 ? RotateCcw : state.rotation === 90 ? RotateCw : Baseline, hidden: state.drawingFocused, disabled: layoutDisabled, gridLayout: { row: 1, column: 3 }, children: textRotationOptions.map(([angle, label]) => button(`sheet.menu.text-rotation.${angle}`, label, {
            commandId: 'sheet.command.set-text-rotation', params: angle,
            disabled: layoutDisabled,
            activated: angle === 'v' ? state.rotationVertical === true : state.rotationVertical !== true && state.rotation === angle,
            icon: angle === 'v' ? TextCursor : angle === 0 ? Baseline : angle === -45 ? RotateCcw : angle === 45 ? RotateCw : angle === -90 ? RotateCcw : RotateCw,
          })) },
          { id: 'sheet.menu.merge', type: 'menu', title: '合并单元格', icon: TableCellsMerge, hidden: state.drawingFocused, disabled: mergeDisabled || !state.canMerge && !state.canUnmerge, activated: state.mergeActive === true, gridLayout: { row: 2, column: 3 }, children: [
            button('sheet.command.merge-all', '全部合并', { commandId: 'sheet.command.merge', params: 'all', icon: TableCellsMerge, disabled: mergeDisabled || !state.canMerge }),
            button('sheet.command.merge-vertical', '垂直合并', { commandId: 'sheet.command.merge', params: 'vertical', icon: TableCellsMerge, disabled: mergeDisabled || !state.canMerge }),
            button('sheet.command.merge-horizontal', '水平合并', { commandId: 'sheet.command.merge', params: 'horizontal', icon: TableCellsMerge, disabled: mergeDisabled || !state.canMerge }),
            button('sheet.command.unmerge', '取消合并', { icon: TableCellsSplit, disabled: mergeDisabled || !state.canUnmerge }),
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
          button('sheet.operation.open-range-protection', '保护范围', { icon: ShieldCheck, params: { mode: 'add' }, disabled: !state.hasSheet || state.rangeProtectionEditable === false, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
          shortcutPanelRibbonItem(tr('shortcut-panel.title', '快捷键面板'), state.shortcutPanelOpen),
          featureSearchRibbonItem(tr('feature-search.title', '搜索功能')),
          button('sheet.command.clear-content', '清除内容', { icon: RemoveFormatting, order: 3, showLabel: true, disabled: !state.hasSheet || state.clearContentEditable === false, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
          button(OPEN_SHEET_RANGE_THEME_OPERATION, '范围主题', { icon: Palette, order: 4, showLabel: true, disabled: state.rangeThemeEditable === false || !state.hasSheet,
            gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
        ] },
      ],
    },
    {
      id: 'insert', title: '插入', order: 10, groups: [
        { id: 'media', title: '媒体', order: 10, items: [
          { id: 'sheet.menu.image', title: '图片', type: 'menu', icon: Image, order: 0,
            disabled: !state.hasSheet || state.imageEditable === false, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true }, children: [
              button(SHEET_IMAGE_COMMANDS.insertFloat, '浮动图片', { icon: ImagePlus, disabled: !state.hasSheet || state.imageEditable === false }),
              button(SHEET_IMAGE_COMMANDS.insertCell, '单元格图片', { icon: Image, disabled: !state.hasSheet || state.imageEditable === false }),
            ] },
          button('sheet.operation.insert-hyperlink', '超链接', { icon: Link, showLabel: true, disabled: !state.hyperlinkAvailable, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
          button('sheet.operation.insert-chart', '图表', { icon: ChartNoAxesCombined, order: 4, disabled: !state.chartAvailable, disabledReason: '图表需要绑定到已打开的工作簿' }),
          button('sheet.operation.open-sparkline-selector', '迷你图', { icon: ChartLine, order: 5 }),
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
          { ...button(dataBar, dataBarLabel, { icon: ChartNoAxesColumnIncreasing, order: 3, disabled: quickConditionalDisabled,
            params: {}, activated: state.quickConditional?.hasDataBar }), type: 'split', children: [
            button(`${dataBar}.default`, dataBarLabel, { commandId: dataBar, params: {}, disabled: quickConditionalDisabled, activated: state.quickConditional?.hasDataBar }),
            button(`${dataBar}.new`, newConditionalLabel, { commandId: newConditional, params: 'dataBar', disabled: quickConditionalDisabled }),
            button(`${dataBar}.clear`, clearConditionalLabel, { commandId: clearConditional, params: { types: ['dataBar'] }, disabled: quickConditionalDisabled || !state.quickConditional?.canClearDataBar }),
          ] },
          { ...button(iconSet, iconSetLabel, { icon: Star, order: 4, disabled: quickConditionalDisabled,
            params: { iconSet: QUICK_CONDITIONAL_ICON_SETS[0] }, activated: state.quickConditional?.hasIconSet }), type: 'split', children: [
            ...QUICK_CONDITIONAL_ICON_SETS.map(preset => button(`${iconSet}.${preset}`, quickIconSetStyles[preset].label, {
              commandId: iconSet, icon: quickIconSetStyles[preset].icon, disabled: quickConditionalDisabled,
              params: { iconSet: preset }, activated: state.quickConditional?.activeIconSet === preset,
            })),
            button(`${iconSet}.new`, newConditionalLabel, { commandId: newConditional, params: 'iconSet', disabled: quickConditionalDisabled }),
            button(`${iconSet}.clear`, clearConditionalLabel, { commandId: clearConditional, params: { types: ['iconSet'] }, disabled: quickConditionalDisabled || !state.quickConditional?.canClearIconSet }),
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
            button('sheet.menu.filter.clear', '清除筛选', { commandId: 'sheet.command.clear-filter', icon: FilterX, disabled: !state.filterHasCriteria }),
            button('sheet.menu.filter.reapply', '重新应用', { commandId: 'sheet.command.reapply-filter', disabled: !state.filterHasCriteria }),
            button('sheet.menu.filter.panel', '筛选条件', { commandId: 'sheet.operation.open-data-panel', params: 'filter', disabled: !state.filterEnabled }),
          ] },
          ...outlineItems,
        ] },
        { id: 'rules', title: '规则', order: 20, items: [
          { id: 'sheet.menu.data-validation', type: 'menu', title: '数据验证', icon: ClipboardCheck, disabled: quickDisabled, gridLayout: { row: 1, column: 1 }, children: [
            button('sheet.operation.open-data-validation', tr('sheet.validation.manage-rules', '管理验证规则'), { disabled: quickDisabled, activated: state.validationPanelOpen }),
            button('sheet.operation.add-data-validation', tr('sheet.validation.add', '添加规则'), { disabled: quickDisabled }),
          ] },
          { id: newConditional, type: 'menu', title: '条件格式', icon: ChartColumnIncreasing, showLabel: true,
            disabled: quickConditionalDisabled && quickConditionalWorksheetDisabled, order: 1, gridLayout: { row: 2, column: 1, showLabel: true }, children: [
              button(`${newConditional}.highlight`, '突出显示单元格规则', { commandId: newConditional, params: 'highlightCell', disabled: quickConditionalDisabled }),
              button(`${newConditional}.rank`, '排名和平均值规则', { commandId: newConditional, params: 'top10', disabled: quickConditionalDisabled }),
              button(`${newConditional}.formula`, '公式规则', { commandId: newConditional, params: 'expression', disabled: quickConditionalDisabled }),
              button(`${newConditional}.color-scale`, '色阶规则', { commandId: newConditional, params: 'colorScale', disabled: quickConditionalDisabled }),
              button(`${newConditional}.data-bar`, '数据条规则', { commandId: newConditional, params: 'dataBar', disabled: quickConditionalDisabled }),
              button(`${newConditional}.icon-set`, '图标集规则', { commandId: newConditional, params: 'iconSet', disabled: quickConditionalDisabled }),
              button(`${newConditional}.manage`, '管理条件格式', { commandId: newConditional, disabled: quickConditionalWorksheetDisabled }),
              button(`${newConditional}.create`, newConditionalLabel, { commandId: newConditional, params: 'create', disabled: quickConditionalDisabled }),
              button(`${newConditional}.clear-range`, clearConditionalLabel, { commandId: clearConditional, params: {}, disabled: quickConditionalDisabled || !state.quickConditional?.canClearRange }),
              button(`${newConditional}.clear-worksheet`, '清除工作表的规则', { commandId: QUICK_CONDITIONAL_FORMAT_COMMANDS.clearWorksheet, params: {}, disabled: quickConditionalWorksheetDisabled || !state.quickConditional?.canClearWorksheet }),
            ] },
        ] },
        { id: 'pivot', title: '分析', order: 30, items: [
          button('sheet.operation.open-pivot', '数据透视表', { icon: TableProperties, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 1, showLabel: true } }),
          button('sheet.operation.open-pivot-chart', '透视图', { icon: ChartNoAxesCombined, showLabel: true, disabled: !state.pivotChartAvailable, disabledReason: '请打开可用的工作簿和工作表', gridLayout: { row: 2, column: 1, showLabel: true } }),
          button('sheet.operation.open-slicer', '切片器', { icon: ListFilter, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 1, column: 2, showLabel: true } }),
          button('sheet.operation.open-timeline', '时间线', { icon: Clock, showLabel: true, disabled: !state.chartAvailable, gridLayout: { row: 2, column: 2, showLabel: true } }),
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
            disabled: formulaDisabled,
            gridLayout: index === 0 ? { row: 1, column: 1, rowSpan: 2, showLabel: false } : undefined,
            children: [...(category === '常用与统计' ? SHEET_COMMON_FUNCTION_CATALOG : SHEET_FUNCTION_CATALOG.filter(entry => entry.category === category)).map(entry => button(
              `sheet.function.${entry.name}`,
              entry.name,
              { commandId: 'sheet.operation.insert-function', params: entry.name, disabled: formulaDisabled },
            )), ...(index === 0 ? [button('sheet.operation.open-function', '插入函数', { icon: FunctionSquare, disabled: formulaDisabled })] : [])],
          })),
        ] },
        { id: 'calculation-options', title: '计算选项', order: 10, items: [
          { id: 'sheet.command.set-calculation-mode', commandId: 'sheet.command.set-calculation-mode', type: 'select', title: '计算模式', width: 116, disabled: !state.hasSheet, value: state.calculationMode ?? 'automatic', gridLayout: { row: 1, column: 1, rowSpan: 2, columnSpan: 2, width: 116, showLabel: true }, options: [
            { label: '自动', value: 'automatic' }, { label: '手动', value: 'manual' },
          ] },
          button('sheet.operation.start-formula', '编辑公式', { icon: FunctionSquare, disabled: formulaDisabled, gridLayout: { row: 1, column: 3 } }),
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
          { id: 'sheet.menu.worksheet-background-image', title: '背景图', type: 'menu', icon: Image, order: .5,
            disabled: !state.hasSheet || state.backgroundImageEditable === false, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true }, children: [
              button(SHEET_IMAGE_COMMANDS.addBackground, '设置背景图', { icon: ImagePlus, disabled: !state.hasSheet || state.backgroundImageEditable === false }),
              button(SHEET_IMAGE_COMMANDS.deleteBackground, '删除背景图', { icon: RemoveFormatting, disabled: !state.hasBackgroundImage || state.backgroundImageEditable === false }),
            ] },
          { id: 'sheet.operation.set-zoom', commandId: 'sheet.operation.set-zoom', type: 'select', title: '缩放', width: 88, value: state.zoom, valueParams: value => Number(value), gridLayout: { row: 1, column: 3, width: 88, showLabel: true }, options: [50, 75, 100, 125, 150, 175, 200, 400].map(value => ({ label: `${value}%`, value })) },
          { ...button('sheet.operation.toggle-crosshair-highlight', '十字高亮', { icon: Crosshair, activated: state.crosshairEnabled, gridLayout: { row: 2, column: 3 } }), type: 'split', children: [
            ...[
              ['blue', '蓝色', 'rgba(91, 141, 239, 0.18)'], ['teal', '青色', 'rgba(36, 166, 153, 0.18)'],
              ['green', '绿色', 'rgba(65, 160, 96, 0.18)'], ['yellow', '黄色', 'rgba(235, 190, 58, 0.2)'],
              ['orange', '橙色', 'rgba(235, 128, 55, 0.18)'], ['red', '红色', 'rgba(220, 75, 75, 0.16)'],
              ['purple', '紫色', 'rgba(141, 99, 201, 0.16)'], ['gray', '灰色', 'rgba(105, 119, 140, 0.16)'],
            ].map(([color, title, value]) => button(`sheet.operation.crosshair-color.${color}`, title!, {
              commandId: 'sheet.operation.set-crosshair-highlight-color', params: value, swatch: value, activated: state.crosshairColor === value,
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
  const slicer = state.slicerControl
  if (slicer) tabs.push({
    id: 'slicer-format', title: '切片器格式', contextual: true, order: 100,
    groups: [
      { id: 'slicer-format.styles', title: '样式', order: 0, items: [{
        id: 'sheet.menu.slicer-quick-style', commandId: 'sheet.command.update-slicer-presentation', type: 'menu', title: '快速样式', icon: PaintBucket,
        disabled: !slicer.canConfigure, value: slicer.style ?? '', popupComponent: SlicerQuickStylePicker,
        popupCommandId: 'sheet.command.update-slicer-presentation', popupProps: { slicerId: slicer.id, kind: slicer.type, style: slicer.style,
          styleOwner: slicer.styleOwner, expectedRevision: slicer.revision },
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
      }] },
      { id: 'slicer-format.options', title: '选项', order: 1, items: [
        button('sheet.menu.slicer-settings', '切片器设置', { commandId: 'sheet.operation.open-slicer-settings-dialog', icon: Settings2, order: 0,
          hidden: slicer.type !== 'slicer', disabled: !slicer.canConfigure, params: { slicerId: slicer.id, kind: slicer.type }, gridLayout: { row: 1, column: 1, showLabel: true } }),
        button('sheet.menu.timeline-settings', '时间线设置', { commandId: 'sheet.operation.open-slicer-settings-dialog', icon: CalendarDays, order: 1,
          hidden: slicer.type !== 'timeline', disabled: !slicer.canConfigure, params: { slicerId: slicer.id, kind: slicer.type }, gridLayout: { row: 1, column: 1, showLabel: true } }),
        button('sheet.menu.slicer-connections', '报表连接', { commandId: 'sheet.operation.open-slicer-connections-dialog', icon: Link, order: 2,
          hidden: !slicer.supportsConnections, disabled: !slicer.canConfigure, params: { slicerId: slicer.id }, gridLayout: { row: 2, column: 1, showLabel: true } }),
      ] },
    ],
  })
  return translate ? localizeRibbon(tabs, 'sheet', translate) : tabs
}
