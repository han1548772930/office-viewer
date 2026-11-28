import type { Component } from 'vue'
import type { DocumentBlock, DocumentModel } from '../../domain/model'
import {
  AlignCenter, AlignJustify, AlignLeft, AlignRight, ArrowUp, Bold, Brush, CalendarDays, ChevronDown,
  ClipboardPaste, CodeXml, Columns3, Copy, Expand, FileCog, Grid2X2, ImagePlus,
  IndentDecrease, IndentIncrease, Italic, Layers2, Link, List, ListOrdered, ListTodo, ListTree, LockKeyhole, MessageSquarePlus,
  Minimize, Omega, Palette, PanelTop, Printer, Redo2, RemoveFormatting, Scissors, Search, FileText,
  PaintBucket, SeparatorHorizontal, ShieldCheck, Sigma, Smile, SquareCheckBig, Strikethrough, Subscript, Superscript,
  TextCursorInput, Trash2, Underline, Undo2, X, Quote, MessageSquareText,
} from 'lucide-vue-next'
import type { DocumentCommand } from '../../features/document/commands/document-commands'
import { documentStructureCommandFromMenu } from './document-structure-ribbon'
import { DOCUMENT_PASTE_OPTIONS } from '../../features/document/services/document-clipboard'
import { validateDocumentParagraphStylePatch } from '../../features/document/services/document-settings'
import { DOCUMENT_NAMED_PARAGRAPH_STYLES, orderedSelection, sameTarget, textTargets, type DocumentNamedParagraphStyle, type DocumentSelection } from '../../features/document/model/document-model.ts'
import { documentBlockChildContainers } from '../../features/document/model/document-structural-blocks'
import { documentTableSelectedTextTargets } from '../../features/document/model/document-table-selection'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import { featureSearchRibbonItem, shortcutPanelRibbonItem } from '../../platform/ribbon/feature-search'
import { fontSizeOptions, ribbonFontOptions, type OfficeFontList } from '../../platform/ribbon/ribbon-fonts'
import type { DocumentToolbarState } from './document-toolbar-state'
import { localizeRibbon, type OfficeTranslator } from '../../i18n/office-i18n.ts'
import DocumentCharacterPicker from './DocumentCharacterPicker.vue'
import type { DocumentCharacterPreferencesStore } from '../../features/document/stores/document-character-preferences-store'
import DocumentListPicker from './DocumentListPicker.vue'
import DocumentColorControl from './DocumentColorControl.vue'
import { DOCUMENT_DEFAULT_TEXT_COLOR, DOCUMENT_STANDARD_COLORS, documentColorIsAutomatic, documentColorPreview, isDocumentColorField } from '../../features/document/services/document-colors'
import { isListPreset, listKind, type ListKind } from '../../features/document/model/document-lists'
import type { HeaderFooterRibbonState, HeaderFooterRibbonKind } from './DocumentHeaderFooterRibbonControl.vue'
import DocumentHeaderFooterRibbonGroup from './DocumentHeaderFooterRibbonGroup.vue'
import { headerFooterControlDisabled } from './document-header-footer-capabilities'
import type { DocumentTocActionState } from '../../features/document/toc/document-toc-command'
import { createDocumentTableRibbon, type DocumentTableRibbonContext } from './document-table-ribbon'
import { documentTableCommandFromRibbon } from './document-table-command-parser'
import { isDocumentTableCommand } from '../../features/document/commands/table-command-types'
import { assertDocumentTableInsertionSize } from '../../features/document/commands/table-insertion'
import { createDocumentNoteRibbon, type DocumentNoteRibbonContext } from './document-note-ribbon'
import { createDocumentStructureRibbon, type DocumentStructureRibbonContext } from './document-structure-ribbon'
import { createDocumentLatexRibbon, type DocumentLatexRibbonContext } from './document-latex-ribbon'
import type { DocumentHyperlinkSelectionState } from '../../features/document/services/document-hyperlinks'
import { DOCUMENT_ZOOM_SHORTCUTS } from '../../features/document/services/document-zoom'

export interface DocumentTocContextActionStates { update: DocumentTocActionState; delete: DocumentTocActionState }

export interface DocumentRibbonContext {
  latex?: DocumentLatexRibbonContext
  structuredBlock?: DocumentStructureRibbonContext
  notes?: DocumentNoteRibbonContext
  table?: DocumentTableRibbonContext
  fontList?: OfficeFontList
  painterActive?: boolean
  painterUnavailable?: boolean
  clearFormatting?: { available: boolean; reason?: string }
  clipboard?: { copy: boolean; cut: boolean; paste: boolean }
  hyperlink?: DocumentHyperlinkSelectionState
  lastTextColor?: string
  lastHighlightColor?: string
  defaultTextColor?: string
  defaultHighlightColor?: string
  numbering?: { canRestart: boolean; canContinue: boolean }
  indentation?: { canDecrease: boolean; canIncrease: boolean }
  state: DocumentToolbarState
  canUndo: boolean
  canRedo: boolean
  outline: boolean
  focus: boolean
  fullscreen: boolean
  shortcutPanelOpen?: boolean
  paper: string
  orientation: string
  zoom: number
  zoomDisabled?: boolean
  contentControlsAvailable: boolean
  mode?: 'traditional' | 'modern'
  editingHeaderFooter?: boolean
  editingAuxiliaryContent?: boolean
  editingDisabled?: boolean
  selectionHasCode?: boolean
  characterPreferences?: DocumentCharacterPreferencesStore
  pageSettingsDisabled?: boolean
  namedStyleDisabled?: boolean
  namedStyleDisabledReason?: string
  /** Character formatting can target an active story independently of body-only tools. */
  textFormattingDisabled?: boolean
  headerFooter?: HeaderFooterRibbonState
  selectedContentControl?: {
    id: string
    type: string
    lock?: string
    placeholder?: string
    showingPlaceholder?: boolean
    appearance?: string
    multiLine?: boolean
    temporary?: boolean
    checkboxStyle?: string
    dateFormat?: string
    dateValue?: string
    parentId?: string
    canSelectParent?: boolean
    placement?: string
    canResetPlaceholder?: boolean
    resetPlaceholderDisabledReason?: string
    canReplacePicture?: boolean
    replacePictureDisabledReason?: string
    canDuplicateItem?: boolean
    duplicateItemDisabledReason?: string
    canDeleteItem?: boolean
    deleteItemDisabledReason?: string
  }
  commentsPanelOpen?: boolean
  canCreateComment?: boolean
  tableOfContents?: {
    available: boolean
    insertDisabled: boolean
    selectedId?: string
    editDisabled: boolean
    updateDisabled?: boolean
    deleteDisabled?: boolean
    updateReason?: string
    deleteReason?: string
  }
}

const action = (id: string, title: string, icon: Component | undefined, commandId = id, params?: unknown, more: Partial<RibbonItem> = {}): RibbonItem => ({ id, type: 'button', title, icon, commandId, params, ...more })

export function documentRibbonSelectionHasCode(model: DocumentModel, selection: DocumentSelection): boolean {
  const targets = textTargets(model)
  const { start, end } = orderedSelection(model, selection)
  const selected = selection.tableSelection ? documentTableSelectedTextTargets(model, selection.tableSelection)
    : targets.slice(targets.findIndex(target => sameTarget(target, start)), targets.findIndex(target => sameTarget(target, end)) + 1)
  const selectedIds = new Set(selected.map(target => target.cellBlockId ?? target.blockId))
  const intersects = (blocks: readonly DocumentBlock[], inCode: boolean): boolean => blocks.some(block => {
    const code = inCode || block.type === 'code'
    if (code && selectedIds.has(block.id)) return true
    return documentBlockChildContainers(block).some(children => intersects(children, code))
  })
  return intersects(model.blocks, false)
}

export function createDocumentRibbon(context: DocumentRibbonContext, translate?: OfficeTranslator): RibbonTab[] {
  const { state } = context
  const textFormattingDisabled = context.textFormattingDisabled ?? context.editingDisabled
  const tocLabel = (key: string, fallback: string) => translate?.(key, fallback) ?? fallback
  const listControl = (kind: ListKind, icon: Component, title: string, row: number, column: number): RibbonItem => ({
    ...action(`doc-${kind}`, title, icon, 'doc.command.set-block-type', { type: kind }), type: 'split',
    disabled: context.editingDisabled, activated: !state.listPresetMixed && (state.listPreset ? listKind(state.listPreset) === kind : state.blockStyle === kind),
    shortcut: kind === 'numbered' ? 'Ctrl+Shift+7' : 'Ctrl+Shift+8', gridLayout: { row, column },
    popupComponent: DocumentListPicker, popupCommandId: 'doc.command.change-list-type',
    popupProps: { kind, title: tocLabel(`items.doc-${kind}`, title), active: state.listPreset,
      canRestart: kind === 'numbered' && Boolean(context.numbering?.canRestart), canContinue: kind === 'numbered' && Boolean(context.numbering?.canContinue) },
  })
  const format = (id: string, title: string, icon: Component, key: 'bold' | 'italic' | 'underline' | 'strike'): RibbonItem => {
    const column = ['bold', 'italic', 'underline', 'strike'].indexOf(key) + 1
    return action(id, title, icon, 'doc.command.format-text', { patch: { [key]: state.marks[key] !== true } }, { order: column - 1, disabled: textFormattingDisabled, activated: state.marks[key] === true,
      shortcut: ({ bold: 'Ctrl+B', italic: 'Ctrl+I', underline: 'Ctrl+U', strike: 'Ctrl+Shift+X' })[key],
      tooltip: state.mixedMarks.includes(key) ? `${title}（混合格式）` : title, gridLayout: { row: 2, column, width: 32, minWidth: 28 } })
  }
  const verticalFormat = (id: string, title: string, icon: Component, value: 'subscript' | 'superscript'): RibbonItem => action(id, title, icon, 'doc.command.format-text', { patch: { verticalAlign: state.marks.verticalAlign === value ? 'baseline' : value } }, { order: value === 'subscript' ? 4 : 5, disabled: textFormattingDisabled, activated: state.marks.verticalAlign === value,
    shortcut: value === 'subscript' ? 'Ctrl+,' : 'Ctrl+.', tooltip: state.mixedMarks.includes('verticalAlign') ? `${title}（混合格式）` : title, gridLayout: { row: 2, column: value === 'subscript' ? 6 : 5, width: 32, minWidth: 28 } })
  const alignments = [{ value: 'left', title: '左对齐', icon: AlignLeft, shortcut: 'Ctrl+Shift+L' }, { value: 'center', title: '居中', icon: AlignCenter, shortcut: 'Ctrl+Shift+E' }, { value: 'right', title: '右对齐', icon: AlignRight, shortcut: 'Ctrl+Shift+R' }, { value: 'justify', title: '两端对齐', icon: AlignJustify, shortcut: 'Ctrl+Shift+J' }]
  const palette = (id: string, title: string, key: 'color' | 'highlight', icon: Component): RibbonItem => {
    const defaultColor = key === 'color' ? context.defaultTextColor ?? DOCUMENT_DEFAULT_TEXT_COLOR : context.defaultHighlightColor ?? '#315bc2'
    const lastColor = (key === 'color' ? context.lastTextColor : context.lastHighlightColor) ?? defaultColor
    return {
      id, type: 'split', title, icon, order: key === 'color' ? 8 : 9, disabled: textFormattingDisabled,
      commandId: 'doc.command.format-text', params: { patch: { [key]: lastColor } }, value: lastColor, swatch: documentColorPreview(key, lastColor, defaultColor),
      inlineComponent: DocumentColorControl, popupCommandId: 'doc.command.format-text', popupProps: {
        title: tocLabel(`items.${id}`, title), field: key, icon, defaultColor, lastColor,
        selectedColor: state.marks[key], mixed: state.mixedMarks.includes(key),
      },
      gridLayout: { row: 2, column: key === 'color' ? 8 : 7, width: 40, minWidth: 38 },
      children: [
        ...DOCUMENT_STANDARD_COLORS.map((color) => action(`${id}-preset-${color.slice(1)}`, color.toUpperCase(), icon, 'doc.command.format-text', { patch: { [key]: color } }, { activated: !state.mixedMarks.includes(key) && state.marks[key]?.toLowerCase() === color, tooltip: color.toUpperCase(), swatch: color })),
        action(`${id}-reset`, tocLabel('document.color.reset', '重置颜色'), RemoveFormatting, 'doc.command.format-text', { patch: { [key]: defaultColor } }),
        action(`${id}-clear`, tocLabel(key === 'color' ? 'document.color.automatic' : 'document.color.noBackground', key === 'color' ? '自动文字颜色' : '无背景'), RemoveFormatting, 'doc.command.format-text', { patch: {}, clearFields: [key] }, { activated: !state.mixedMarks.includes(key) && documentColorIsAutomatic(key, state.marks[key]) }),
      ],
    }
  }
  const tabs: RibbonTab[] = [
    { id: 'home', title: '开始', order: 0, groups: [
      { id: 'doc-history', title: '历史与剪贴板', order: 0, items: [
        action('doc-undo', '撤销', Undo2, 'core.command.undo', undefined, { disabled: !context.canUndo, disabledReason: '没有可撤销的文档更改', shortcut: 'Ctrl+Z', gridLayout: { row: 1, column: 1, iconSize: 18 } }),
        action('doc-redo', '重做', Redo2, 'core.command.redo', undefined, { disabled: !context.canRedo, disabledReason: '没有可重做的文档更改', shortcut: 'Ctrl+Y', gridLayout: { row: 2, column: 1, iconSize: 18 } }),
        action('doc-format-painter', tocLabel('items.sheet.operation.format-painter', '格式刷'), Brush, 'ui.operation.activate-format-painter', undefined, { subId: 'ui.operation.continuous-format-painter', activated: !!context.painterActive, disabled: !context.painterActive && context.painterUnavailable, gridLayout: { row: 1, column: 2, showLabel: false } }),
        action('doc-clear-format', '清除格式', RemoveFormatting, 'ui.command.clear-formatting', undefined, { disabled: context.clearFormatting ? !context.clearFormatting.available : context.painterUnavailable, disabledReason: context.clearFormatting?.reason, gridLayout: { row: 2, column: 2, showLabel: false } }),
        { ...action('doc-paste', '粘贴', ClipboardPaste, 'doc.command.paste-special', undefined, { order: 4, shortcut: 'Ctrl+V', disabled: context.clipboard ? !context.clipboard.paste : context.editingDisabled, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }), type: 'menu', children:
          DOCUMENT_PASTE_OPTIONS.map(({ value, label }) => action(`doc-paste-mode-${value}`, tocLabel(`doc.context.paste-special.${value}`, label), undefined, 'doc.command.paste-special', { value })) },
      ] },
      { id: 'doc-format', title: '字体与样式', order: 1, items: [
        { id: 'doc-block-style', type: 'select', title: '段落样式', order: 5.5, disabled: context.editingDisabled || context.namedStyleDisabled, disabledReason: context.namedStyleDisabledReason, width: 120, value: state.namedStyle?.startsWith('heading') ? `H${state.namedStyle.slice(-1)}` : state.namedStyle ?? state.blockStyle ?? null, placeholder: '段落样式', gridLayout: { row: 1, column: 1, columnSpan: 2, width: 80, minWidth: 72 }, options: [
          ...([
            ['normal', 'normal', 'Normal'], ['title', 'title', 'Title'], ['subtitle', 'subtitle', 'Subtitle'],
            ['H1', 'heading1', 'Heading 1'], ['H2', 'heading2', 'Heading 2'], ['H3', 'heading3', 'Heading 3'],
            ['H4', 'heading4', 'Heading 4'], ['H5', 'heading5', 'Heading 5'],
          ] as const).map(([value, namedStyle, label]) => {
            const style = DOCUMENT_NAMED_PARAGRAPH_STYLES[namedStyle]
            return { label, value, preview: { fontSize: `${style.fontSize}px`, fontWeight: style.bold ? 700 : 400, color: style.color ? 'var(--ant-color-text-secondary, currentColor)' : undefined }, commandId: 'doc.command.set-paragraph-named-style', params: { namedStyle } }
          }),
        ] },
        { id: 'doc-font-family', type: 'select', edit: { kind: 'font' }, title: '字体', order: 7, disabled: textFormattingDisabled, width: 112, value: state.marks.fontFamily ?? null, placeholder: '混合字体', commandId: 'doc.command.format-text', gridLayout: { row: 1, column: 3, columnSpan: 4, width: 140, minWidth: 124 }, options: ribbonFontOptions(state.marks.fontFamily, context.fontList), valueParams: (value) => ({ patch: { fontFamily: value } }) },
        { id: 'doc-font-size', type: 'select', edit: { kind: 'number', min: 1, max: 400 }, title: '字号', order: 6, disabled: textFormattingDisabled, width: 60, value: state.marks.fontSize ?? null, placeholder: '混合', commandId: 'doc.command.format-text', gridLayout: { row: 1, column: 7, columnSpan: 2, width: 84, minWidth: 80 }, options: fontSizeOptions, valueParams: (value) => ({ patch: { fontSize: Number(value) } }) },
        format('doc-bold', '加粗', Bold, 'bold'), format('doc-italic', '斜体', Italic, 'italic'), format('doc-underline', '下划线', Underline, 'underline'), format('doc-strike', '删除线', Strikethrough, 'strike'),
        verticalFormat('doc-superscript', '上标', Superscript, 'superscript'), verticalFormat('doc-subscript', '下标', Subscript, 'subscript'),
        palette('doc-highlight', '文本高亮', 'highlight', PaintBucket), palette('doc-text-color', '字体颜色', 'color', Palette),
      ] },
      { id: 'doc-layout', title: '段落与列表', order: 2, items: [
        { id: 'doc-align', type: 'menu', title: '段落对齐', disabled: context.editingDisabled, icon: alignments.find((item) => item.value === state.paragraph.align)?.icon ?? AlignLeft, gridLayout: { row: 1, column: 1 }, children: alignments.map((item) => action(`doc-align-${item.value}`, item.title, item.icon, 'doc.command.format-paragraph', { patch: { align: item.value } }, { activated: state.paragraph.align === item.value, shortcut: item.shortcut })) },
        listControl('numbered', ListOrdered, '有序列表', 2, 1),
        listControl('bullet', List, '无序列表', 1, 2),
        action('doc-checklist', '清单', ListTodo, 'doc.command.set-block-type', { type: 'checklist' }, { disabled: context.editingDisabled, activated: state.blockStyle === 'checklist', gridLayout: { row: 2, column: 2 } }),
        action('doc-header-footer', '页眉和页脚', PanelTop, 'doc.operation.open-header-footer', undefined, { hidden: context.mode === 'modern', showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
      ] },
      { id: 'doc-others', title: '其他', order: 3, items: [
        action('doc-permissions', '文档权限', ShieldCheck, 'doc.operation.open-permissions', undefined, { order: 0, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
        shortcutPanelRibbonItem(tocLabel('shortcut-panel.title', '快捷键面板'), context.shortcutPanelOpen),
        featureSearchRibbonItem(tocLabel('feature-search.title', '搜索功能')),
        action('doc-find', '查找', Search, 'doc.operation.toggle-find', undefined, { order: 3, gridLayout: { row: 1, column: 3 } }),
      ] },
      { id: 'doc-layout-extensions', title: '段落工具', order: 4, items: [
        { id: 'doc-extra-styles', type: 'menu', title: '更多段落样式', icon: FileText, order: 10, disabled: context.editingDisabled, gridLayout: { row: 1, column: 3 }, children: [
          action('doc-heading-six', tocLabel('options.doc.doc-block-style.H6', '标题 6'), undefined, 'doc.command.set-block-type', { type: 'heading', level: 6 }, { activated: state.blockStyle === 'H6' }),
          action('doc-quote', tocLabel('options.doc.doc-block-style.quote', '引用'), undefined, 'doc.operation.insert-quote', undefined, { activated: context.structuredBlock?.type === 'quote' || state.blockStyle === 'quote' }),
        ] },
        action('doc-indent-less', '减少缩进', IndentDecrease, 'doc.command.change-paragraph-indent', { direction: -1 }, { disabled: context.editingDisabled || !(context.indentation?.canDecrease ?? ((state.paragraph.indentFirstLine ?? 0) > 0)), disabledReason: '当前选区无法减少缩进', gridLayout: { row: 1, column: 1 } }),
        action('doc-indent-more', '增加缩进', IndentIncrease, 'doc.command.change-paragraph-indent', { direction: 1 }, { disabled: context.editingDisabled || !(context.indentation?.canIncrease ?? ((state.paragraph.indentFirstLine ?? 0) < 100)), disabledReason: '当前选区无法增加缩进', gridLayout: { row: 2, column: 1 } }),
        { id: 'doc-line-height', type: 'select', title: '行距', disabled: context.editingDisabled, width: 76,
          value: !state.paragraph.spacingRule || state.paragraph.spacingRule === 'auto' ? state.paragraph.lineHeight ?? null : null,
          placeholder: state.paragraph.spacingRule && state.paragraph.spacingRule !== 'auto'
            ? `${tocLabel(`doc.paragraph-settings.spacing-rule.${state.paragraph.spacingRule}`, state.paragraph.spacingRule === 'exact' ? '固定值' : '最小值')} ${state.paragraph.lineHeight ?? ''} px` : '混合',
          commandId: 'doc.command.format-paragraph', gridLayout: { row: 1, column: 2 }, options: [1, 1.15, 1.5, 2, 2.5, 3].map((value) => ({ label: `${value} 倍`, value })), valueParams: (value) => ({ patch: { spacingRule: 'auto', lineHeight: Number(value) } }) },
        { id: 'doc-paragraph-spacing', type: 'menu', title: '段落间距', disabled: context.editingDisabled, icon: IndentIncrease, gridLayout: { row: 2, column: 2 }, children: (['spaceBefore', 'spaceAfter'] as const).map((key) => ({ id: `doc-${key}`, type: 'menu', title: key === 'spaceBefore' ? '段前间距' : '段后间距', children: [0, 6, 8, 12, 18, 24, 36].map((value) => action(`doc-${key}-${value}`, `${value} 像素`, IndentIncrease, 'doc.command.format-paragraph', { patch: { [key]: value } }, { activated: state.paragraph[key] === value })) })) },
      ] },
      { id: 'doc-home-extensions', title: '其他工具', order: 5, items: [
        action('doc-copy', '复制', Copy, 'doc.command.copy', undefined, { disabled: context.clipboard ? !context.clipboard.copy : !state.hasSelection, shortcut: 'Ctrl+C', gridLayout: { row: 1, column: 1 } }),
        action('doc-cut', '剪切', Scissors, 'doc.command.cut', undefined, { disabled: context.clipboard ? !context.clipboard.cut : !state.hasSelection || context.editingDisabled, shortcut: 'Ctrl+X', gridLayout: { row: 2, column: 1 } }),
        action('doc-outline', '文档大纲', ListTree, 'doc.operation.toggle-outline', undefined, { activated: context.outline, showLabel: true, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
        action('doc-focus', '专注模式', Expand, 'doc.operation.toggle-focus', undefined, { activated: context.focus, showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
        action('doc-print', '打印文档', Printer, 'doc.command.print', undefined, { showLabel: true, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
      ] },
    ] },
    { id: 'insert', title: '插入', order: 1, groups: [
      { id: 'doc-insert-media', title: '媒体', order: 0, items: [
      { ...action('doc-breaks', '分页符和分节符', Columns3, 'doc.operation.open-breaks', undefined, { order: 4, hidden: context.mode === 'modern',
        disabled: context.editingDisabled || context.editingHeaderFooter || context.editingAuxiliaryContent || context.structuredBlock?.type === 'code',
        showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }), type: 'menu', children: [
        action('doc-page-break', '分页符', undefined, 'doc.command.insert-page-break'),
        action('doc-column-break', '分栏符', undefined, 'doc.command.insert-column-break', undefined, { disabled: !!state.table }),
        { id: 'doc-section-breaks', type: 'menu', title: '分节符', disabled: !!state.table, children: [
          ...([
            ['nextPage', '下一页'],
            ['continuous', '连续'],
            ['nextColumn', '下一栏'],
            ['evenPage', '偶数页'],
            ['oddPage', '奇数页'],
          ] as const).map(([type, title]) => action(`doc-section-break-${type}`, title, undefined, 'doc.command.insert-section-break', { type }, { disabled: !!state.table })),
        ] },
      ] },
      { ...action('doc-table', '表格', Grid2X2, 'doc.operation.create-table', undefined, { order: 2, disabled: context.editingDisabled, gridLayout: { row: 1, column: 2 } }), type: 'menu', children: [
        action('doc-table-insert', '插入表格', Grid2X2, 'doc.operation.create-table'),
      ] },
      { ...action('doc-emoji', '表情', Smile, 'doc.operation.open-characters', 'emoji', { order: 5, disabled: context.editingDisabled, gridLayout: { row: 1, column: 3 } }), type: 'menu', popupComponent: DocumentCharacterPicker, popupProps: { mode: 'emoji', preferences: context.characterPreferences }, popupCommandId: 'doc.command.insert-special-character' },
      { ...action('doc-symbol', '符号', Omega, 'doc.operation.open-characters', 'symbol', { order: 6, disabled: context.editingDisabled, gridLayout: { row: 2, column: 3 } }), type: 'menu', popupComponent: DocumentCharacterPicker, popupProps: { mode: 'symbol', preferences: context.characterPreferences }, popupCommandId: 'doc.command.insert-special-character' },
      action('doc-image', '图片', ImagePlus, 'doc.command.insert-float-image', undefined, { order: 0, disabled: context.editingDisabled, gridLayout: { row: 1, column: 4 } }),
      action('doc-link', '链接', Link, 'doc.operation.show-hyper-link-edit-popup', undefined, { order: 1,
        disabled: context.hyperlink ? !context.hyperlink.available : context.editingDisabled || !state.hasSelection,
        disabledReason: context.hyperlink?.reason, shortcut: 'Ctrl+K', gridLayout: { row: 1, column: 5 } }),
      action('doc-divider', '分隔线', SeparatorHorizontal, 'doc.command.horizontal-line', undefined, { order: 4, disabled: context.editingDisabled || context.selectionHasCode || context.structuredBlock?.type === 'code', gridLayout: { row: 2, column: 5 } }),
      action('doc-code-block', '代码块', CodeXml, 'doc.operation.insert-code-block', undefined, { order: 7, disabled: context.editingDisabled, gridLayout: { row: 1, column: 7 } }),
      action('doc-latex', 'LaTeX 公式', Sigma, 'doc.operation.open-latex', undefined, { order: 8, disabled: context.editingDisabled, gridLayout: { row: 2, column: 7 } }),
      ] },
      { id: 'doc-insert-cell', title: '目录', order: 1, items: [{
        ...action('doc-insert-table-of-contents', tocLabel('document.toc.insert', 'Table of Contents'), ListTree, 'doc.operation.insert-table-of-contents', undefined, {
          order: 7,
          hidden: context.tableOfContents?.available !== true,
          disabled: context.tableOfContents?.insertDisabled ?? true,
          tooltip: tocLabel('document.toc.insert', 'Table of Contents'),
          showLabel: true,
          gridLayout: { row: 1, column: 1, showLabel: true },
        }),
        type: 'menu',
        children: [
          action('doc-insert-table-of-contents-automatic', tocLabel('document.toc.automatic', 'Automatic table'), ListTree, 'doc.operation.insert-table-of-contents', { mode: 'automatic' }),
          action('doc-insert-table-of-contents-custom', tocLabel('document.toc.custom', 'Custom table of contents'), FileCog, 'doc.operation.insert-table-of-contents', { mode: 'custom' }),
        ],
      },
      { ...action('doc-controls', '内容控件', SquareCheckBig, 'doc.operation.insert-content-control', 'richText', { order: 7.1, showLabel: true, disabled: context.editingDisabled || !context.contentControlsAvailable, disabledReason: '内容控件需要连接文档命令服务', gridLayout: { row: 2, column: 1, showLabel: true } }), type: 'split', children: [
        ...[
          { type: 'richText', title: '富文本控件', icon: TextCursorInput },
          { type: 'text', title: '纯文本控件', icon: TextCursorInput },
          { type: 'picture', title: '图片控件', icon: ImagePlus },
          { type: 'checkbox', title: '复选框控件', icon: SquareCheckBig },
          { type: 'comboBox', title: '组合框控件', icon: ChevronDown },
          { type: 'dropDownList', title: '下拉列表控件', icon: ChevronDown },
          { type: 'date', title: '日期控件', icon: CalendarDays },
          { type: 'repeatingSection', title: '重复节控件', icon: Copy },
          { type: 'group', title: '分组控件', icon: Layers2 },
          { type: 'docPartObj', title: '文档部件控件', icon: FileCog },
        ].map((item) => action(`doc-controls-${item.type}`, item.title, item.icon, 'doc.operation.insert-content-control', item.type)),
      ] },
      ] },
      { id: 'doc-insert-structured-blocks', title: '高级块', order: 2, items: [
        action('doc-insert-quote', '引用块', Quote, 'doc.operation.insert-quote', undefined, { order: 0, disabled: context.editingDisabled, gridLayout: { row: 1, column: 1, showLabel: true } }),
        action('doc-insert-callout', '标注块', MessageSquareText, 'doc.operation.insert-callout', undefined, { order: 1, disabled: context.editingDisabled, gridLayout: { row: 2, column: 1, showLabel: true } }),
        { ...action('doc-insert-column-group', '多栏组', Columns3, 'doc.operation.insert-column-group', undefined, { order: 2, disabled: context.editingDisabled,
          gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }), type: 'split', children: [
          ...[2, 3, 4, 5].map(columnCount => action(`doc-insert-column-group-${columnCount}`, `${columnCount} 栏`, undefined, 'doc.operation.insert-column-group', { columnCount })),
          action('doc-insert-column-group-custom', '设置比例和间距', undefined, 'doc.operation.insert-column-group'),
        ] },
      ] },
      { id: 'doc-insert-extensions', title: '更多', order: 3, items: [
      action('doc-comment-toggle', '评论', MessageSquarePlus, 'docs.operation.toggle-comment-panel', undefined, { activated: context.commentsPanelOpen, gridLayout: { row: 1, column: 1, showLabel: true } }),
      action('doc-comment-add', '添加评论', MessageSquarePlus, 'docs.operation.start-add-comment', undefined, { disabled: !context.canCreateComment, disabledReason: '选择文本后可添加评论', gridLayout: { row: 2, column: 1, showLabel: true } }),
      ] },
    ] },
    { id: 'view', title: '视图', order: 2, groups: [
      { id: 'doc-display', title: '显示', order: 0, items: [
      action('doc-page-settings', '页面设置', FileCog, 'doc.operation.open-page-settings', undefined, { disabled: context.pageSettingsDisabled, showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
      { id: 'doc-zoom', type: 'select', title: '缩放', width: 90, value: context.zoom, disabled: context.zoomDisabled, commandId: 'doc.operation.set-zoom-ratio',
        valueParams: value => ({ zoomRatio: Number(value) / 100 }), gridLayout: { row: 1, column: 2 }, options: DOCUMENT_ZOOM_SHORTCUTS.map(value => ({ label: `${value}%`, value })) },
      action('doc-fullscreen', context.fullscreen ? '退出全屏' : '全屏', context.fullscreen ? Minimize : Expand, 'doc.operation.toggle-fullscreen', undefined, { activated: context.fullscreen, gridLayout: { row: 2, column: 2 } }),
      ] },
    ] },
  ]
  if (context.headerFooter?.visible && context.mode !== 'modern') {
    const headerFooter = context.headerFooter
    const headerFooterControl = (kind: HeaderFooterRibbonKind, title: string, icon: Component, order: number): RibbonItem => ({
      id: `doc-header-footer-${kind}`,
      type: 'button',
      title,
      icon,
      order,
      disabled: headerFooterControlDisabled(kind, headerFooter),
      inlineComponent: DocumentHeaderFooterRibbonGroup,
      popupProps: { kind, title, state: headerFooter },
      popupCommandId: 'doc.operation.header-footer-control',
      gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
    })
    tabs.push({
      id: 'header-footer',
      title: '页眉和页脚',
      order: 90,
      contextual: true,
      groups: [
        ...[
          headerFooterControl('insertion', '插入', FileText, 0),
          headerFooterControl('navigation', '导航', PanelTop, 1),
          headerFooterControl('options', '选项', FileCog, 2),
          headerFooterControl('position', '位置', IndentIncrease, 3),
        ].map(item => ({ id: item.id, title: item.title, order: item.order, items: [item] })),
        { id: 'header-footer-close', title: '关闭', order: 5, items: [action('doc-header-footer-close', '关闭页眉和页脚', X, 'doc.operation.close-header-footer', undefined, {
            order: 5,
            showLabel: true,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
          })] },
      ],
    })
  }
  if (context.tableOfContents?.selectedId) {
    const tocId = context.tableOfContents.selectedId
    tabs.push({
      id: 'table-of-contents',
      title: tocLabel('document.toc.insert', 'Table of Contents'),
      order: 100,
      contextual: true,
      groups: [{
        id: 'table-of-contents-manage',
        title: tocLabel('document.toc.manage', 'Manage'),
        order: 0,
        items: [
          action('doc-update-table-of-contents', tocLabel('document.toc.update', 'Update Table of Contents'), ListTree, 'doc.operation.update-table-of-contents', { tocId }, {
            showLabel: true,
            disabled: context.tableOfContents.updateDisabled ?? context.tableOfContents.editDisabled,
            disabledReason: context.tableOfContents.updateReason,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
          }),
          action('doc-delete-table-of-contents', tocLabel('document.toc.delete', 'Remove Table of Contents'), Trash2, 'doc.command.delete-table-of-contents', { tocId }, {
            showLabel: true,
            disabled: context.tableOfContents.deleteDisabled ?? context.tableOfContents.editDisabled,
            disabledReason: context.tableOfContents.deleteReason,
            gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
          }),
        ],
      }],
    })
  }
  const latexRibbon = context.latex && createDocumentLatexRibbon(context.latex, translate)
  if (latexRibbon) tabs.push(latexRibbon)
  if (context.selectedContentControl) {
    const control = context.selectedContentControl
    const values: RibbonItem[] = []
    if (control.type === 'checkbox') {
      values.push({
        ...action('doc-controls-checkbox-style', '复选框样式', SquareCheckBig, 'doc.command.set-selected-sdt-checkbox-style'),
        type: 'select', value: control.checkboxStyle ?? null, placeholder: '自定义',
        hidden: ['contentLocked', 'sdtContentLocked'].includes(control.lock ?? ''), gridLayout: { row: 1, column: 1, showLabel: true },
        options: [
          { value: 'Cross', label: '叉号' }, { value: 'Check', label: '勾选' },
          { value: 'Minimal', label: '简洁' }, { value: 'Round', label: '圆形' },
        ], valueParams: value => ({ value }),
      })
    }
    if (control.type === 'picture') values.push(action('doc-controls-change-picture', '更换图片', ImagePlus, 'doc.command.change-selected-sdt-picture', undefined, {
      disabled: control.canReplacePicture === false, disabledReason: control.replacePictureDisabledReason,
      gridLayout: { row: 1, column: 1, showLabel: true },
    }))
    if (control.type === 'date') values.push({
      ...action('doc-controls-date-format', '日期格式', CalendarDays, 'doc.command.set-selected-sdt-date-format'),
      type: 'select', value: control.dateFormat ?? 'yyyy-MM-dd', gridLayout: { row: 1, column: 1, showLabel: true },
      disabled: context.editingDisabled || (['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '') && !control.showingPlaceholder && !!control.dateValue),
      options: ['yyyy-MM-dd', 'yyyy/MM/dd', 'M/d/yyyy', 'dd MMM yyyy', 'MMMM d, yyyy'].map(value => ({ value, label: value })),
      valueParams: value => ({ value }),
    })
    if (control.placeholder) values.push(action('doc-controls-reset-placeholder', '恢复占位符', TextCursorInput, 'doc.command.reset-selected-sdt-placeholder', undefined, {
      disabled: !control.canResetPlaceholder, disabledReason: control.resetPlaceholderDisabledReason, gridLayout: { row: 1, column: 2, showLabel: true },
    }))
    if (control.type === 'text') values.push(action('doc-controls-multiline', '允许换行', TextCursorInput, 'doc.command.toggle-selected-sdt-multiline', undefined, {
      activated: control.multiLine, disabled: context.editingDisabled, gridLayout: { row: 1, column: 1, showLabel: true },
    }))
    if (control.type === 'text' || control.type === 'richText') values.push(action('doc-controls-temporary', '编辑后移除', RemoveFormatting, 'doc.command.toggle-selected-sdt-temporary', undefined, {
      activated: control.temporary, disabled: context.editingDisabled, gridLayout: { row: 2, column: 1, showLabel: true },
    }))
    const behavior: RibbonItem[] = [
      { id: 'doc-controls-appearance', type: 'select', title: '外观', icon: Layers2,
        value: control.appearance ?? 'boundingBox', commandId: 'doc.command.set-selected-sdt-appearance', disabled: context.editingDisabled,
        gridLayout: { row: 1, column: 1, showLabel: true }, options: [
          { label: '边界框', value: 'boundingBox' }, { label: '开始/结束标记', value: 'tags' }, { label: '不显示标记', value: 'hidden' },
        ].map(option => ({ ...option, commandId: 'doc.command.set-selected-sdt-appearance', params: { value: option.value } })), valueParams: value => ({ value }) },
      { id: 'doc-controls-lock', type: 'select', title: '锁定', icon: LockKeyhole,
        value: control.lock === 'controlLocked' ? 'sdtLocked' : control.lock ?? 'unlocked', commandId: 'doc.command.set-selected-sdt-lock',
        disabled: context.editingDisabled, gridLayout: { row: 2, column: 1, showLabel: true }, options: [
          { label: '不锁定', value: 'unlocked' }, { label: '禁止删除控件', value: 'sdtLocked' },
          { label: '禁止编辑内容', value: 'contentLocked' }, { label: '内容和控件均锁定', value: 'sdtContentLocked' },
        ].map(option => ({ ...option, commandId: 'doc.command.set-selected-sdt-lock', params: { value: option.value } })), valueParams: value => ({ value }) },
      action('doc-controls-properties', '属性', FileCog, 'doc.command.open-selected-sdt-properties', undefined, { gridLayout: { row: 1, column: 2, showLabel: true } }),
      action('doc-controls-remove', '删除内容控件', Trash2, 'doc.command.remove-selected-sdt', undefined, {
        disabled: context.editingDisabled || ['sdtLocked', 'sdtContentLocked', 'controlLocked'].includes(control.lock ?? ''),
        gridLayout: { row: 2, column: 2, showLabel: true },
      }),
    ]
    const actions: RibbonItem[] = [
      action('doc-controls-select-parent', '选择上层控件', ArrowUp, 'doc.operation.select-parent-sdt', undefined, {
        disabled: !control.canSelectParent, gridLayout: { row: 1, column: 2, showLabel: true },
      }),
      ...(control.type === 'repeatingSectionItem' && !['cell', 'table'].includes(control.placement ?? '') ? [
        action('doc-controls-duplicate-item', '复制', Copy, 'doc.command.duplicate-selected-repeating-sdt-item', undefined, {
          disabled: control.canDuplicateItem === false, disabledReason: control.duplicateItemDisabledReason,
          gridLayout: { row: 1, column: 1, showLabel: true },
        }),
      ] : []),
      ...(control.type === 'repeatingSectionItem' && !['cell', 'table'].includes(control.placement ?? '') ? [
        action('doc-controls-delete-item', '删除项目', Trash2, 'doc.command.delete-selected-repeating-sdt-item', undefined, {
          disabled: control.canDeleteItem === false, disabledReason: control.deleteItemDisabledReason,
          gridLayout: { row: 2, column: 1, showLabel: true },
        }),
      ] : []),
    ]
    tabs.push({
      id: 'content-control', title: '内容控件', order: 110, contextual: true,
      groups: [
        { id: 'content-control-values', title: '值', order: 0, items: values },
        { id: 'content-control-behavior', title: '行为', order: 1, items: behavior },
        { id: 'content-control-actions', title: '操作', order: 2, items: actions },
      ],
    })
  }
  const tableTab = context.table && createDocumentTableRibbon(context.table)
  if (tableTab) tabs.push(tableTab)
  const noteTab = context.notes && createDocumentNoteRibbon(context.notes)
  if (noteTab) tabs.push(noteTab)
  if (context.structuredBlock) tabs.push(createDocumentStructureRibbon(context.structuredBlock))
  return translate ? localizeRibbon(tabs, 'doc', translate) : tabs
}

export function documentTableContextMenuItems(translate?: OfficeTranslator): RibbonItem[] {
  const label = (key: string, fallback: string) => translate?.(key, fallback) ?? fallback
  return [
    { id: 'doc-context-row-before', type: 'button', title: label('items.doc-table-row-before', '在上方插入行'), commandId: 'doc.command.table-insert-row', params: { position: 'before' } },
    { id: 'doc-context-row-after', type: 'button', title: label('items.doc-table-row-after', '在下方插入行'), commandId: 'doc.command.table-insert-row', params: { position: 'after' } },
    { id: 'doc-context-column-before', type: 'button', title: label('items.doc-table-column-before', '在左侧插入列'), commandId: 'doc.command.table-insert-column', params: { position: 'before' } },
    { id: 'doc-context-column-after', type: 'button', title: label('items.doc-table-column-after', '在右侧插入列'), commandId: 'doc.command.table-insert-column', params: { position: 'after' } },
    { id: 'doc-context-delete-rows', type: 'button', title: label('items.doc-table-delete-rows', 'Delete Rows'), commandId: 'doc.command.table-delete-rows' },
    { id: 'doc-context-delete-columns', type: 'button', title: label('items.doc-table-delete-columns', 'Delete Columns'), commandId: 'doc.command.table-delete-columns' },
    { id: 'doc-context-delete-table', type: 'button', title: label('items.doc-table-delete', 'Delete Table'), commandId: 'doc.command.table-delete' },
  ]
}

export function documentTocContextMenuItems(tocId: string, states: boolean | DocumentTocContextActionStates, translate?: OfficeTranslator): RibbonItem[] {
  const label = (key: string, fallback: string) => translate?.(key, fallback) ?? fallback
  const update: DocumentTocActionState = typeof states === 'boolean' ? { available: !states } : states.update
  const remove: DocumentTocActionState = typeof states === 'boolean' ? { available: !states } : states.delete
  return [
    { id: 'doc-context-update-table-of-contents', type: 'button', title: label('document.toc.update', 'Update Table of Contents'), commandId: 'doc.operation.update-table-of-contents', params: { tocId }, disabled: !update.available, disabledReason: update.reason, order: -1 },
    { id: 'doc-context-delete-table-of-contents', type: 'button', title: label('document.toc.delete', 'Remove Table of Contents'), commandId: 'doc.command.delete-table-of-contents', params: { tocId }, disabled: !remove.available, disabledReason: remove.reason, order: -0.5 },
  ]
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== null && !Array.isArray(value) }

/** Validate toolbar parameters before they enter the same command path as Canvas input. */
export function documentCommandFromRibbon(id: string, params: unknown): DocumentCommand | undefined {
  const structure = documentStructureCommandFromMenu(id, params)
  if (structure) return structure
  if (isDocumentTableCommand({ id })) return documentTableCommandFromRibbon(id, params)
  if (id === 'doc.command.insert-character') return typeof params === 'string' && params.length > 0 ? { id: 'doc.command.insert-text', text: params } : undefined
  if (id === 'doc.command.insert-special-character') {
    return record(params) && typeof params.value === 'string' && params.value.length > 0 ? { id: 'doc.command.insert-text', text: params.value } : undefined
  }
  if (id === 'ui.command.clear-formatting' || id === 'doc.command.clear-formatting') return { id: 'doc.command.clear-formatting' }
  if (id === 'doc.command.change-paragraph-indent') {
    return record(params) && (params.direction === -1 || params.direction === 1) ? { id, direction: params.direction } : undefined
  }
  if (id === 'doc.command.restart-list-numbering') {
    if (params === undefined) return { id }
    if (!record(params)) return undefined
    const start = params.start
    return start === undefined ? { id } : typeof start === 'number' && Number.isSafeInteger(start) && start >= 0 && start <= 2147483647 ? { id, start } : undefined
  }
  if (id === 'doc.command.continue-list-numbering') return { id }
  if (id === 'doc.command.change-list-type') {
    const listType = record(params) ? params.listType : params
    return isListPreset(listType) ? { id, listType } : undefined
  }
  if (id === 'doc.command.insert-page-break' || id === 'doc.command.insert-column-break' || id === 'doc.command.insert-divider'
    || id === 'doc.command.horizontal-line' || id === 'doc.command.insert-horizontal-line-bellow') return { id }
  if (id === 'doc.command.delete' && record(params) && (params.direction === -1 || params.direction === 1)) return { id, direction: params.direction }
  if ((id === 'doc.command.insert-table' || id === 'doc.command.create-table') && record(params)) {
    const rows = id === 'doc.command.create-table' ? params.rowCount : params.rows
    const columns = id === 'doc.command.create-table' ? params.colCount : params.cols
    if (typeof rows !== 'number' || typeof columns !== 'number') return undefined
    try {
      assertDocumentTableInsertionSize(rows, columns)
      return id === 'doc.command.create-table' ? { id, rowCount: rows, colCount: columns } : { id, rows, cols: columns }
    } catch (error) {
      if (error instanceof RangeError) return undefined
      throw error
    }
  }
  if (id === 'doc.command.insert-section-break' && record(params)) {
    const type = params.type
    if (type === 'nextPage' || type === 'continuous' || type === 'nextColumn' || type === 'evenPage' || type === 'oddPage') return { id, type }
    return undefined
  }
  if (id === 'doc.command.table-delete-rows' || id === 'doc.command.table-delete-columns' || id === 'doc.command.table-delete') return { id }
  if (!record(params)) return undefined
  if (id === 'doc.command.table-insert-row' || id === 'doc.command.table-insert-column') {
    return params.position === 'before' || params.position === 'after' ? { id, position: params.position } : undefined
  }
  if (id === 'doc.command.insert-text') return typeof params.text === 'string' ? { id, text: params.text } : undefined
  if (id === 'doc.command.set-paragraph-named-style') {
    const namedStyle = params.namedStyle
    return typeof namedStyle === 'string' && Object.hasOwn(DOCUMENT_NAMED_PARAGRAPH_STYLES, namedStyle)
      ? { id, namedStyle: namedStyle as DocumentNamedParagraphStyle }
      : undefined
  }
  if (id === 'doc.command.set-block-type') {
    const type = params.type
    if (type !== 'paragraph' && type !== 'heading' && type !== 'bullet' && type !== 'numbered' && type !== 'checklist' && type !== 'quote') return undefined
    const level = params.level
    if (level !== undefined && (typeof level !== 'number' || !Number.isInteger(level) || level < 0 || level > 6)) return undefined
    return { id, type, level }
  }
  if (!record(params.patch)) return undefined
  const patch = params.patch
  if (id === 'doc.command.format-text') {
    const clearFields = params.clearFields
    if (clearFields !== undefined && (!Array.isArray(clearFields) || !clearFields.every(isDocumentColorField))) return undefined
    if (!Object.keys(patch).length && !clearFields?.length) return undefined
    if (!Object.entries(patch).every(([key, value]) => ['bold', 'italic', 'underline', 'strike'].includes(key) ? typeof value === 'boolean' : key === 'verticalAlign' ? value === 'baseline' || value === 'subscript' || value === 'superscript' : key === 'fontSize' ? typeof value === 'number' && Number.isFinite(value) && value > 0 && value <= 400 : ['fontFamily', 'color', 'highlight', 'href'].includes(key) && typeof value === 'string')) return undefined
    return { id, patch, ...(clearFields?.length ? { clearFields: [...new Set(clearFields)] } : {}) }
  }
  if (id === 'doc.command.format-paragraph') {
    if (!Object.keys(patch).length) return undefined
    try {
      validateDocumentParagraphStylePatch(patch)
      return { id, patch }
    } catch { return undefined }
  }
  return undefined
}
