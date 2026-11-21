import type { Component } from 'vue'
import { AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, Clipboard, ClipboardPaste, Eraser, FileText, Highlighter, ImagePlus, IndentDecrease, IndentIncrease, Italic, Link2, List, ListOrdered, ListTodo, ListTree, LockKeyhole, Maximize, MessageSquare, Minimize, Minus, Omega, Plus, Printer, Redo2, Scissors, Search, Settings2, Sigma, Smile, StickyNote, Strikethrough, Subscript, Superscript, Table2, Underline, Undo2 } from 'lucide-vue-next'
import type { DocumentCommand } from '../../features/document/commands/document-commands'
import { DOCUMENT_NAMED_PARAGRAPH_STYLES, type DocumentNamedParagraphStyle } from '../../features/document/model/document-model.ts'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import type { DocumentToolbarState } from './document-toolbar-state'
import { localizeRibbon, type OfficeTranslator } from '../../i18n/office-i18n.ts'

interface DocumentRibbonContext {
  state: DocumentToolbarState
  canUndo: boolean
  canRedo: boolean
  outline: boolean
  focus: boolean
  fullscreen: boolean
  paper: string
  orientation: string
  zoom: number
  contentControlsAvailable: boolean
  commentsPanelOpen?: boolean
  canCreateComment?: boolean
  tableOfContents?: {
    available: boolean
    insertDisabled: boolean
    selectedId?: string
    editDisabled: boolean
  }
}

const action = (id: string, title: string, icon: Component, commandId = id, params?: unknown, more: Partial<RibbonItem> = {}): RibbonItem => ({ id, type: 'button', title, icon, commandId, params, ...more })
const colors = [
  { key: 'black', value: '#1f2937' }, { key: 'gray', value: '#6b7280' }, { key: 'white', value: '#ffffff' },
  { key: 'red', value: '#dc2626' }, { key: 'orange', value: '#ea580c' }, { key: 'yellow', value: '#facc15' },
  { key: 'green', value: '#16a34a' }, { key: 'teal', value: '#0d9488' }, { key: 'blue', value: '#2563eb' },
  { key: 'navy', value: '#1e3a8a' }, { key: 'purple', value: '#9333ea' }, { key: 'pink', value: '#db2777' },
  { key: 'brown', value: '#92400e' }, { key: 'light-gray', value: '#d1d5db' }, { key: 'light-blue', value: '#bfdbfe' },
  { key: 'light-yellow', value: '#fef3c7' },
]

/** Menu ordering follows docs-ui/menu/schema at the fixed Univer baseline. */
export function createDocumentRibbon(context: DocumentRibbonContext, translate?: OfficeTranslator): RibbonTab[] {
  const { state } = context
  const tocLabel = (key: string, fallback: string) => translate?.(key, fallback) ?? fallback
  const format = (id: string, title: string, icon: Component, key: 'bold' | 'italic' | 'underline' | 'strike'): RibbonItem => action(id, title, icon, 'doc.command.format-text', { patch: { [key]: state.marks[key] !== true } }, { activated: state.marks[key] === true, tooltip: state.mixedMarks.includes(key) ? `${title}（混合格式）` : title })
  const verticalFormat = (id: string, title: string, icon: Component, value: 'subscript' | 'superscript'): RibbonItem => action(id, title, icon, 'doc.command.format-text', { patch: { verticalAlign: state.marks.verticalAlign === value ? 'baseline' : value } }, { activated: state.marks.verticalAlign === value, tooltip: state.mixedMarks.includes('verticalAlign') ? `${title}（混合格式）` : title })
  const alignments = [{ value: 'left', title: '左对齐', icon: AlignLeft }, { value: 'center', title: '居中', icon: AlignCenter }, { value: 'right', title: '右对齐', icon: AlignRight }, { value: 'justify', title: '两端对齐', icon: AlignJustify }]
  const palette = (id: string, title: string, key: 'color' | 'highlight', icon: Component): RibbonItem => ({
    id, type: 'menu', title, icon,
    children: [
      ...colors.map((color) => action(`${id}-${color.value}`, translate?.(`colors.${color.key}`, color.key) || color.key, icon, 'doc.command.format-text', { patch: { [key]: color.value } }, { activated: state.marks[key] === color.value, tooltip: color.value, swatch: color.value })),
      action(`${id}-reset`, translate?.(key === 'color' ? 'items.doc-text-color-reset' : 'items.doc-highlight-reset', key === 'color' ? 'Reset Font Color' : 'No Background') || (key === 'color' ? 'Reset Font Color' : 'No Background'), Eraser, 'doc.command.format-text', { patch: { [key]: key === 'color' ? '#1f2937' : '' } }),
    ],
  })
  const tabs: RibbonTab[] = [
    { id: 'home', title: '开始', order: 0, groups: [
      { id: 'doc-history', title: '历史与剪贴板', order: 0, items: [
        action('doc-undo', '撤销', Undo2, 'core.command.undo', undefined, { disabled: !context.canUndo, disabledReason: '没有可撤销的文档更改', shortcut: 'Ctrl+Z', gridLayout: { row: 1, column: 1, iconSize: 18 } }),
        action('doc-redo', '重做', Redo2, 'core.command.redo', undefined, { disabled: !context.canRedo, disabledReason: '没有可重做的文档更改', shortcut: 'Ctrl+Y', gridLayout: { row: 2, column: 1, iconSize: 18 } }),
        action('doc-clear-format', '清除字符格式', Eraser, 'doc.command.format-text', { patch: { bold: false, italic: false, underline: false, strike: false, verticalAlign: 'baseline', fontFamily: 'Arial', fontSize: 11, color: '#1f2937', highlight: '' } }, { gridLayout: { row: 2, column: 2 } }),
        { ...action('doc-paste', '粘贴', ClipboardPaste, 'doc.command.paste', undefined, { shortcut: 'Ctrl+V', gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }), type: 'split', children: [
          action('doc-paste-source', '保留源格式粘贴', ClipboardPaste, 'doc.command.paste'),
          action('doc-paste-text', '仅粘贴文本', FileText, 'doc.command.paste-text', undefined, { shortcut: 'Ctrl+Shift+V' }),
        ] },
      ] },
      { id: 'doc-format', title: '字体与样式', order: 2, items: [
        { id: 'doc-block-style', type: 'select', title: '段落样式', width: 120, value: state.blockStyle ?? null, placeholder: '混合样式', gridLayout: { row: 1, column: 1, columnSpan: 2, width: 80, minWidth: 72 }, options: [
          ...([
            ['normal', 'normal', 'Normal'], ['title', 'title', 'Title'], ['subtitle', 'subtitle', 'Subtitle'],
            ['H1', 'heading1', 'Heading 1'], ['H2', 'heading2', 'Heading 2'], ['H3', 'heading3', 'Heading 3'],
            ['H4', 'heading4', 'Heading 4'], ['H5', 'heading5', 'Heading 5'],
          ] as const).map(([value, namedStyle, label]) => ({ label, value, commandId: 'doc.command.set-paragraph-named-style', params: { namedStyle } })),
          { label: 'Heading 6', value: 'H6', commandId: 'doc.command.set-block-type', params: { type: 'heading', level: 6 } },
          { label: 'Quote', value: 'quote', commandId: 'doc.command.set-block-type', params: { type: 'quote' } },
        ] },
        { id: 'doc-font-family', type: 'select', title: '字体', width: 112, value: state.marks.fontFamily ?? null, placeholder: '混合字体', commandId: 'doc.command.format-text', gridLayout: { row: 1, column: 3, columnSpan: 4, width: 140, minWidth: 124 }, options: ['Arial', 'Aptos', 'Calibri', 'Georgia', 'Times New Roman', 'Microsoft YaHei', 'SimSun'].map((value) => ({ label: value === 'Microsoft YaHei' ? '微软雅黑' : value === 'SimSun' ? '宋体' : value, value })), valueParams: (value) => ({ patch: { fontFamily: value } }) },
        { id: 'doc-font-size', type: 'select', title: '字号', width: 60, value: state.marks.fontSize ?? null, placeholder: '混合', commandId: 'doc.command.format-text', gridLayout: { row: 1, column: 7, columnSpan: 2, width: 84, minWidth: 80 }, options: [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72].map((value) => ({ label: String(value), value })), valueParams: (value) => ({ patch: { fontSize: Number(value) } }) },
        format('doc-bold', '加粗', Bold, 'bold'), format('doc-italic', '斜体', Italic, 'italic'), format('doc-underline', '下划线', Underline, 'underline'), format('doc-strike', '删除线', Strikethrough, 'strike'),
        verticalFormat('doc-superscript', '上标', Superscript, 'superscript'), verticalFormat('doc-subscript', '下标', Subscript, 'subscript'),
        palette('doc-highlight', '文本高亮', 'highlight', Highlighter), palette('doc-text-color', '字体颜色', 'color', FileText),
      ] },
      { id: 'doc-layout', title: '段落与列表', order: 2, items: [
        { id: 'doc-align', type: 'menu', title: '段落对齐', icon: alignments.find((item) => item.value === state.paragraph.align)?.icon ?? AlignLeft, gridLayout: { row: 1, column: 1 }, children: alignments.map((item) => action(`doc-align-${item.value}`, item.title, item.icon, 'doc.command.format-paragraph', { patch: { align: item.value } }, { activated: state.paragraph.align === item.value })) },
        action('doc-bullet', '无序列表', List, 'doc.command.set-block-type', { type: 'bullet' }, { activated: state.blockStyle === 'bullet', gridLayout: { row: 1, column: 2 } }),
        action('doc-numbered', '有序列表', ListOrdered, 'doc.command.set-block-type', { type: 'numbered' }, { activated: state.blockStyle === 'numbered', gridLayout: { row: 2, column: 1 } }),
        action('doc-checklist', '清单', ListTodo, 'doc.command.set-block-type', { type: 'checklist' }, { activated: state.blockStyle === 'checklist', gridLayout: { row: 2, column: 2 } }),
        action('doc-header-footer', '页眉和页脚', FileText, 'doc.operation.open-header-footer', undefined, { showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
      ] },
      { id: 'doc-others', title: '其他', order: 3, items: [
        action('doc-find', '查找', Search, 'doc.operation.toggle-find', undefined, { gridLayout: { row: 1, column: 1 } }),
        action('doc-permissions', '文档权限', LockKeyhole, 'doc.operation.open-permissions', undefined, { gridLayout: { row: 2, column: 1 } }),
      ] },
      { id: 'doc-layout-extensions', title: '段落工具', order: 4, items: [
        action('doc-indent-less', '减少缩进', IndentDecrease, 'doc.command.format-paragraph', { patch: { indent: Math.max(0, (state.paragraph.indent ?? 0) - 1) } }, { disabled: state.paragraph.indent === 0, disabledReason: '当前段落已经位于最左侧', gridLayout: { row: 1, column: 1 } }),
        action('doc-indent-more', '增加缩进', IndentIncrease, 'doc.command.format-paragraph', { patch: { indent: Math.min(20, (state.paragraph.indent ?? 0) + 1) } }, { gridLayout: { row: 2, column: 1 } }),
        { id: 'doc-line-height', type: 'select', title: '行距', width: 76, value: state.paragraph.lineHeight ?? null, placeholder: '混合', commandId: 'doc.command.format-paragraph', gridLayout: { row: 1, column: 2 }, options: [1, 1.15, 1.5, 2, 2.5, 3].map((value) => ({ label: `${value} 倍`, value })), valueParams: (value) => ({ patch: { lineHeight: Number(value) } }) },
        { id: 'doc-paragraph-spacing', type: 'menu', title: '段落间距', icon: IndentIncrease, gridLayout: { row: 2, column: 2 }, children: (['spaceBefore', 'spaceAfter'] as const).map((key) => ({ id: `doc-${key}`, type: 'menu', title: key === 'spaceBefore' ? '段前间距' : '段后间距', children: [0, 6, 8, 12, 18, 24, 36].map((value) => action(`doc-${key}-${value}`, `${value} 像素`, IndentIncrease, 'doc.command.format-paragraph', { patch: { [key]: value } }, { activated: state.paragraph[key] === value })) })) },
      ] },
      { id: 'doc-home-extensions', title: '其他工具', order: 5, items: [
        action('doc-copy', '复制', Clipboard, 'doc.command.copy', undefined, { disabled: !state.hasSelection, disabledReason: '请先选择要复制的文字', gridLayout: { row: 1, column: 1 } }),
        action('doc-cut', '剪切', Scissors, 'doc.command.cut', undefined, { disabled: !state.hasSelection, disabledReason: '请先选择要剪切的文字', shortcut: 'Ctrl+X', gridLayout: { row: 2, column: 1 } }),
        action('doc-outline', '文档大纲', ListTree, 'doc.operation.toggle-outline', undefined, { activated: context.outline, showLabel: true, gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
        action('doc-focus', '专注模式', Maximize, 'doc.operation.toggle-focus', undefined, { activated: context.focus, showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
        action('doc-print', '打印文档', Printer, 'doc.command.print', undefined, { showLabel: true, gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }),
      ] },
    ] },
    { id: 'insert', title: '插入', order: 1, groups: [
      { id: 'doc-insert-media', title: '媒体', order: 0, items: [
      { ...action('doc-breaks', '分页符和分节符', Plus, 'doc.operation.open-breaks', undefined, { showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }), type: 'menu', children: [
        action('doc-page-break', '分页符', Plus, 'doc.command.insert-page-break'),
        action('doc-column-break', '分栏符', Plus, 'doc.command.insert-column-break'),
        { id: 'doc-section-breaks', type: 'menu', title: '分节符', children: [
          ...([
            ['nextPage', '下一页'],
            ['continuous', '连续'],
            ['nextColumn', '下一栏'],
            ['evenPage', '偶数页'],
            ['oddPage', '奇数页'],
          ] as const).map(([type, title]) => action(`doc-section-break-${type}`, title, Plus, 'doc.command.insert-section-break', { type })),
        ] },
      ] },
      action('doc-table', '表格', Table2, 'doc.operation.open-table', undefined, { gridLayout: { row: 1, column: 2 } }),
      action('doc-emoji', '表情', Smile, 'doc.operation.open-characters', 'emoji', { gridLayout: { row: 1, column: 3 } }),
      action('doc-symbol', '符号', Omega, 'doc.operation.open-characters', 'symbol', { gridLayout: { row: 2, column: 3 } }),
      action('doc-image', '图片', ImagePlus, 'doc.operation.open-image', undefined, { gridLayout: { row: 1, column: 4 } }),
      action('doc-link', '链接', Link2, 'doc.operation.open-link', undefined, { gridLayout: { row: 1, column: 5 } }),
      action('doc-divider', '分隔线', Minus, 'doc.command.insert-divider', undefined, { gridLayout: { row: 2, column: 5 } }),
      ] },
      { id: 'doc-insert-cell', title: '目录', order: 1, items: [{
        ...action('doc-insert-table-of-contents', tocLabel('document.toc.insert', 'Table of Contents'), FileText, 'doc.operation.insert-table-of-contents', undefined, {
          order: 7,
          hidden: context.tableOfContents?.available !== true,
          disabled: context.tableOfContents?.insertDisabled ?? true,
          tooltip: tocLabel('document.toc.insert', 'Table of Contents'),
          showLabel: true,
          gridLayout: { row: 1, column: 1, showLabel: true },
        }),
        type: 'menu',
        children: [
          action('doc-insert-table-of-contents-automatic', tocLabel('document.toc.automatic', 'Automatic table'), FileText, 'doc.operation.insert-table-of-contents', { mode: 'automatic' }),
          action('doc-insert-table-of-contents-custom', tocLabel('document.toc.custom', 'Custom table of contents'), Settings2, 'doc.operation.insert-table-of-contents', { mode: 'custom' }),
        ],
      }] },
      { id: 'doc-insert-extensions', title: '更多', order: 3, items: [
      action('doc-comment-toggle', '评论', MessageSquare, 'docs.operation.toggle-comment-panel', undefined, { activated: context.commentsPanelOpen, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
      action('doc-comment-add', '添加评论', Plus, 'docs.operation.start-add-comment', undefined, { disabled: !context.canCreateComment, disabledReason: '选择文本后可添加评论', gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true } }),
      action('doc-code-block', '代码块', FileText, 'doc.operation.insert-code-block', undefined, { showLabel: true, gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true } }),
      { ...action('doc-controls', '内容控件', FileText, 'doc.operation.open-content-controls', undefined, { showLabel: true, disabled: !context.contentControlsAvailable, disabledReason: '内容控件需要连接文档命令服务', gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true } }), type: 'split', children: [
        ...[{ type: 'text', title: '纯文本控件' }, { type: 'richText', title: '富文本控件' }, { type: 'checkbox', title: '复选框控件' }, { type: 'date', title: '日期控件' }, { type: 'dropDownList', title: '下拉列表控件' }, { type: 'comboBox', title: '组合框控件' }].map((item) => action(`doc-controls-${item.type}`, item.title, FileText, 'doc.operation.open-content-controls', item.type)),
      ] },
      { ...action('doc-footnote', '脚注', StickyNote, 'doc.operation.insert-footnote', undefined, { showLabel: true, gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true } }), type: 'split', children: [
        action('doc-footnote-insert', '插入脚注', StickyNote, 'doc.operation.insert-footnote'),
        action('doc-endnote-insert', '插入尾注', StickyNote, 'doc.operation.insert-endnote'),
        action('doc-notes-manage', '管理脚注和尾注', StickyNote, 'doc.operation.open-notes'),
      ] },
      action('doc-latex', 'LaTeX 公式', Sigma, 'doc.operation.open-latex', undefined, { showLabel: true, gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true } }),
      ] },
    ] },
    { id: 'view', title: '视图', order: 2, groups: [
      { id: 'doc-display', title: '显示', order: 0, items: [
      action('doc-page-settings', '页面设置', Settings2, 'doc.operation.open-page-settings', undefined, { showLabel: true, gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true } }),
      { id: 'doc-zoom', type: 'select', title: '缩放', width: 90, value: context.zoom, commandId: 'doc.operation.set-zoom', gridLayout: { row: 1, column: 2 }, options: [25, 50, 75, 100, 125, 150, 200, 300, 400].map((value) => ({ label: `${value}%`, value })) },
      action('doc-fullscreen', context.fullscreen ? '退出全屏' : '全屏', context.fullscreen ? Minimize : Maximize, 'doc.operation.toggle-fullscreen', undefined, { activated: context.fullscreen, gridLayout: { row: 2, column: 2 } }),
      ] },
    ] },
  ]
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
          action('doc-update-table-of-contents', tocLabel('document.toc.update', 'Update Table of Contents'), FileText, 'doc.operation.update-table-of-contents', { tocId }, {
            showLabel: true,
            disabled: context.tableOfContents.editDisabled,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
          }),
          action('doc-delete-table-of-contents', tocLabel('document.toc.delete', 'Remove Table of Contents'), Eraser, 'doc.command.delete-table-of-contents', { tocId }, {
            showLabel: true,
            disabled: context.tableOfContents.editDisabled,
            gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
          }),
        ],
      }],
    })
  }
  return translate ? localizeRibbon(tabs, 'doc', translate) : tabs
}

export function documentTableContextMenuItems(translate?: OfficeTranslator): RibbonItem[] {
  const label = (key: string, fallback: string) => translate?.(key, fallback) ?? fallback
  return [
    { id: 'doc-context-row-before', type: 'button', title: label('items.doc-table-row-before', 'Insert Row Above'), commandId: 'doc.command.table-insert-row', params: { position: 'before' } },
    { id: 'doc-context-row-after', type: 'button', title: label('items.doc-table-row-after', 'Insert Row Below'), commandId: 'doc.command.table-insert-row', params: { position: 'after' } },
    { id: 'doc-context-column-before', type: 'button', title: label('items.doc-table-column-before', 'Insert Column Left'), commandId: 'doc.command.table-insert-column', params: { position: 'before' } },
    { id: 'doc-context-column-after', type: 'button', title: label('items.doc-table-column-after', 'Insert Column Right'), commandId: 'doc.command.table-insert-column', params: { position: 'after' } },
    { id: 'doc-context-delete-rows', type: 'button', title: label('items.doc-table-delete-rows', 'Delete Rows'), commandId: 'doc.command.table-delete-rows' },
    { id: 'doc-context-delete-columns', type: 'button', title: label('items.doc-table-delete-columns', 'Delete Columns'), commandId: 'doc.command.table-delete-columns' },
    { id: 'doc-context-delete-table', type: 'button', title: label('items.doc-table-delete', 'Delete Table'), commandId: 'doc.command.table-delete' },
  ]
}

export function documentTocContextMenuItems(tocId: string, disabled: boolean, translate?: OfficeTranslator): RibbonItem[] {
  const label = (key: string, fallback: string) => translate?.(key, fallback) ?? fallback
  return [
    { id: 'doc-context-update-table-of-contents', type: 'button', title: label('document.toc.update', 'Update Table of Contents'), commandId: 'doc.operation.update-table-of-contents', params: { tocId }, disabled, order: -1 },
    { id: 'doc-context-delete-table-of-contents', type: 'button', title: label('document.toc.delete', 'Remove Table of Contents'), commandId: 'doc.command.delete-table-of-contents', params: { tocId }, disabled, order: -0.5 },
  ]
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== null && !Array.isArray(value) }

/** Validate toolbar parameters before they enter the same command path as Canvas input. */
export function documentCommandFromRibbon(id: string, params: unknown): DocumentCommand | undefined {
  if (id === 'doc.command.insert-page-break' || id === 'doc.command.insert-column-break' || id === 'doc.command.insert-divider') return { id }
  if (id === 'doc.command.insert-table' && record(params)) {
    const { rows, cols } = params
    return Number.isInteger(rows) && Number.isInteger(cols) && Number(rows) >= 1 && Number(rows) <= 20 && Number(cols) >= 1 && Number(cols) <= 20
      ? { id, rows: Number(rows), cols: Number(cols) }
      : undefined
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
  if (!record(params.patch) || !Object.keys(params.patch).length) return undefined
  const patch = params.patch
  if (id === 'doc.command.format-text') {
    if (!Object.entries(patch).every(([key, value]) => ['bold', 'italic', 'underline', 'strike'].includes(key) ? typeof value === 'boolean' : key === 'verticalAlign' ? value === 'baseline' || value === 'subscript' || value === 'superscript' : key === 'fontSize' ? typeof value === 'number' && Number.isFinite(value) && value > 0 && value <= 400 : ['fontFamily', 'color', 'highlight', 'href'].includes(key) && typeof value === 'string')) return undefined
    return { id, patch }
  }
  if (id === 'doc.command.format-paragraph') {
    if (!Object.entries(patch).every(([key, value]) => key === 'align' ? ['left', 'center', 'right', 'justify'].includes(String(value)) : ['indent', 'lineHeight', 'spaceBefore', 'spaceAfter'].includes(key) && typeof value === 'number' && Number.isFinite(value) && value >= 0)) return undefined
    return { id, patch }
  }
  return undefined
}
