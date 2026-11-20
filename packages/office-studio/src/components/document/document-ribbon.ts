import type { Component } from 'vue'
import { AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, Clipboard, FileText, Highlighter, ImagePlus, IndentDecrease, IndentIncrease, Italic, Link2, List, ListOrdered, ListTree, Maximize, Minus, Plus, Printer, Redo2, Search, Strikethrough, Table2, Underline, Undo2 } from 'lucide-vue-next'
import type { DocumentCommand } from '../../features/document/commands/document-commands'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import type { DocumentToolbarState } from './document-toolbar-state'

interface DocumentRibbonContext {
  state: DocumentToolbarState
  canUndo: boolean
  canRedo: boolean
  outline: boolean
  focus: boolean
  paper: string
  orientation: string
  zoom: number
  contentControlsAvailable: boolean
}

const action = (id: string, title: string, icon: Component, commandId = id, params?: unknown, more: Partial<RibbonItem> = {}): RibbonItem => ({ id, type: 'button', title, icon, commandId, params, ...more })
const colors = [{ label: '黑色', value: '#1f2937' }, { label: '灰色', value: '#6b7280' }, { label: '红色', value: '#dc2626' }, { label: '橙色', value: '#ea580c' }, { label: '黄色', value: '#facc15' }, { label: '绿色', value: '#16a34a' }, { label: '蓝色', value: '#2563eb' }, { label: '紫色', value: '#9333ea' }]

/** Menu ordering follows docs-ui/menu/schema at the fixed Univer baseline. */
export function createDocumentRibbon(context: DocumentRibbonContext): RibbonTab[] {
  const { state } = context
  const format = (id: string, title: string, icon: Component, key: 'bold' | 'italic' | 'underline' | 'strike'): RibbonItem => action(id, title, icon, 'doc.command.format-text', { patch: { [key]: state.marks[key] !== true } }, { activated: state.marks[key] === true, tooltip: state.mixedMarks.includes(key) ? `${title}（混合格式）` : title })
  const alignments = [{ value: 'left', title: '左对齐', icon: AlignLeft }, { value: 'center', title: '居中', icon: AlignCenter }, { value: 'right', title: '右对齐', icon: AlignRight }, { value: 'justify', title: '两端对齐', icon: AlignJustify }]
  const palette = (id: string, title: string, key: 'color' | 'highlight', icon: Component): RibbonItem => ({
    id, type: 'menu', title, icon,
    children: [...colors.map((color) => action(`${id}-${color.value}`, color.label, icon, 'doc.command.format-text', { patch: { [key]: color.value } }, { activated: state.marks[key] === color.value })), action(`${id}-custom`, '自定义颜色', icon, 'doc.operation.custom-color', key)],
  })
  return [
    { id: 'home', title: '开始', order: 0, groups: [
      { id: 'doc-history', title: '历史与剪贴板', order: 0, items: [
        action('doc-undo', '撤销', Undo2, 'core.command.undo', undefined, { disabled: !context.canUndo, disabledReason: '没有可撤销的文档更改', shortcut: 'Ctrl+Z' }),
        action('doc-redo', '重做', Redo2, 'core.command.redo', undefined, { disabled: !context.canRedo, disabledReason: '没有可重做的文档更改', shortcut: 'Ctrl+Y' }),
        action('doc-copy', '复制', Clipboard, 'doc.command.copy', undefined, { disabled: !state.hasSelection, disabledReason: '请先选择要复制的文字' }),
      ] },
      { id: 'doc-format', title: '字体与样式', order: 1, items: [
        { id: 'doc-block-style', type: 'select', title: '段落样式', width: 96, value: state.blockStyle ?? null, placeholder: '混合样式', commandId: 'doc.command.set-block-type', options: [{ label: '正文', value: 'P' }, ...[1, 2, 3, 4, 5, 6].map((level) => ({ label: `标题 ${level}`, value: `H${level}` }))], valueParams: (value) => ({ type: value === 'P' ? 'paragraph' : 'heading', level: value === 'P' ? 0 : Number(String(value).slice(1)) }) },
        { id: 'doc-font-family', type: 'select', title: '字体', width: 112, value: state.marks.fontFamily ?? null, placeholder: '混合字体', commandId: 'doc.command.format-text', options: ['Arial', 'Aptos', 'Calibri', 'Georgia', 'Times New Roman', 'Microsoft YaHei', 'SimSun'].map((value) => ({ label: value === 'Microsoft YaHei' ? '微软雅黑' : value === 'SimSun' ? '宋体' : value, value })), valueParams: (value) => ({ patch: { fontFamily: value } }) },
        { id: 'doc-font-size', type: 'select', title: '字号', width: 60, value: state.marks.fontSize ?? null, placeholder: '混合', commandId: 'doc.command.format-text', options: [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72].map((value) => ({ label: String(value), value })), valueParams: (value) => ({ patch: { fontSize: Number(value) } }) },
        format('doc-bold', '加粗', Bold, 'bold'), format('doc-italic', '斜体', Italic, 'italic'), format('doc-underline', '下划线', Underline, 'underline'), format('doc-strike', '删除线', Strikethrough, 'strike'),
        palette('doc-text-color', '字体颜色', 'color', FileText), palette('doc-highlight', '文本高亮', 'highlight', Highlighter),
      ] },
      { id: 'doc-paragraph', title: '段落', order: 2, items: [
        { id: 'doc-align', type: 'menu', title: '段落对齐', icon: alignments.find((item) => item.value === state.paragraph.align)?.icon ?? AlignLeft, children: alignments.map((item) => action(`doc-align-${item.value}`, item.title, item.icon, 'doc.command.format-paragraph', { patch: { align: item.value } }, { activated: state.paragraph.align === item.value })) },
        action('doc-bullet', '无序列表', List, 'doc.command.set-block-type', { type: 'bullet' }, { activated: state.blockStyle === 'bullet' }),
        action('doc-numbered', '有序列表', ListOrdered, 'doc.command.set-block-type', { type: 'numbered' }, { activated: state.blockStyle === 'numbered' }),
        action('doc-indent-less', '减少缩进', IndentDecrease, 'doc.command.format-paragraph', { patch: { indent: Math.max(0, (state.paragraph.indent ?? 0) - 1) } }, { disabled: state.paragraph.indent === 0, disabledReason: '当前段落已经位于最左侧' }),
        action('doc-indent-more', '增加缩进', IndentIncrease, 'doc.command.format-paragraph', { patch: { indent: Math.min(20, (state.paragraph.indent ?? 0) + 1) } }),
        { id: 'doc-line-height', type: 'select', title: '行距', width: 76, value: state.paragraph.lineHeight ?? null, placeholder: '混合', commandId: 'doc.command.format-paragraph', options: [1, 1.15, 1.5, 2, 2.5, 3].map((value) => ({ label: `${value} 倍`, value })), valueParams: (value) => ({ patch: { lineHeight: Number(value) } }) },
      ] },
      { id: 'doc-editing', title: '编辑', order: 3, items: [action('doc-find', '查找', Search, 'doc.operation.toggle-find')] },
    ] },
    { id: 'insert', title: '插入', order: 1, groups: [{ id: 'doc-insert', title: '插入内容', items: [
      action('doc-table', '表格', Table2, 'doc.operation.open-table', undefined, { showLabel: true }),
      action('doc-image', '图片', ImagePlus, 'doc.operation.open-image', undefined, { showLabel: true }),
      action('doc-link', '链接', Link2, 'doc.operation.open-link', undefined, { showLabel: true }),
      action('doc-page-break', '分页符', Plus, 'doc.command.insert-page-break', undefined, { showLabel: true }),
      action('doc-divider', '分隔线', Minus, 'doc.command.insert-divider', undefined, { showLabel: true }),
      action('doc-controls', '内容控件', FileText, 'doc.operation.open-content-controls', undefined, { showLabel: true, disabled: !context.contentControlsAvailable, disabledReason: '内容控件需要连接文档命令服务' }),
    ] }] },
    { id: 'layout', title: '布局', order: 2, groups: [{ id: 'doc-paper', title: '页面设置', items: [
      { id: 'doc-paper-size', type: 'select', title: '纸张', width: 90, value: context.paper, commandId: 'doc.command.set-paper-size', options: ['A4', 'Letter'].map((value) => ({ label: value, value })) },
      { id: 'doc-orientation', type: 'select', title: '方向', width: 90, value: context.orientation, commandId: 'doc.command.set-orientation', options: [{ label: '纵向', value: 'portrait' }, { label: '横向', value: 'landscape' }] },
      action('doc-insert-page', '插入分页', Plus, 'doc.command.insert-page-break', undefined, { showLabel: true }),
    ] }] },
    { id: 'view', title: '视图', order: 3, groups: [{ id: 'doc-display', title: '显示', items: [
      action('doc-outline', '文档大纲', ListTree, 'doc.operation.toggle-outline', undefined, { activated: context.outline, showLabel: true }),
      action('doc-focus', '专注模式', Maximize, 'doc.operation.toggle-focus', undefined, { activated: context.focus, showLabel: true }),
      { id: 'doc-zoom', type: 'select', title: '缩放', width: 90, value: context.zoom, commandId: 'doc.operation.set-zoom', options: [25, 50, 75, 100, 125, 150, 200, 300, 400].map((value) => ({ label: `${value}%`, value })) },
      action('doc-print', '打印文档', Printer, 'doc.command.print', undefined, { showLabel: true }),
    ] }] },
  ]
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== null && !Array.isArray(value) }

/** Validate toolbar parameters before they enter the same command path as Canvas input. */
export function documentCommandFromRibbon(id: string, params: unknown): DocumentCommand | undefined {
  if (id === 'doc.command.insert-page-break' || id === 'doc.command.insert-divider') return { id }
  if (!record(params)) return undefined
  if (id === 'doc.command.set-block-type') {
    const type = params.type
    if (type !== 'paragraph' && type !== 'heading' && type !== 'bullet' && type !== 'numbered' && type !== 'quote') return undefined
    const level = params.level
    if (level !== undefined && (typeof level !== 'number' || !Number.isInteger(level) || level < 0 || level > 6)) return undefined
    return { id, type, level }
  }
  if (!record(params.patch) || !Object.keys(params.patch).length) return undefined
  const patch = params.patch
  if (id === 'doc.command.format-text') {
    if (!Object.entries(patch).every(([key, value]) => ['bold', 'italic', 'underline', 'strike'].includes(key) ? typeof value === 'boolean' : key === 'fontSize' ? typeof value === 'number' && Number.isFinite(value) && value > 0 && value <= 400 : ['fontFamily', 'color', 'highlight', 'href'].includes(key) && typeof value === 'string')) return undefined
    return { id, patch }
  }
  if (id === 'doc.command.format-paragraph') {
    if (!Object.entries(patch).every(([key, value]) => key === 'align' ? ['left', 'center', 'right', 'justify'].includes(String(value)) : ['indent', 'lineHeight', 'spaceBefore', 'spaceAfter'].includes(key) && typeof value === 'number' && Number.isFinite(value) && value >= 0)) return undefined
    return { id, patch }
  }
  return undefined
}
