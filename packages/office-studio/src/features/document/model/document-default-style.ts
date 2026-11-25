import { createDocumentDefaultStyle, type DocumentDefaultStyle, type DocumentModel } from '../../../domain/model'
import { normalizeDocumentHighlight } from './document-highlight'

/** New and older snapshots keep the workbench defaults; imported defaults are resolved in Rust. */
export function documentDefaultStyle(model?: Pick<DocumentModel, 'defaultStyle'>): DocumentDefaultStyle {
  const source = model?.defaultStyle
  if (source === undefined) return createDocumentDefaultStyle()
  if (!source || typeof source !== 'object' || Array.isArray(source) || source.version !== 1 || !source.text || !source.paragraph) throw new RangeError('文档默认格式版本无效。')
  const text = source.text, paragraph = source.paragraph
  if (typeof text.fontFamily !== 'string' || !text.fontFamily.trim()
    || typeof text.color !== 'string' || text.color.trim() !== text.color || !/^(?:auto|#?[\da-f]{6})$/i.test(text.color)
    || !Number.isFinite(text.fontSize) || text.fontSize <= 0
    || ['bold', 'italic', 'underline', 'strike'].some(key => typeof Reflect.get(text, key) !== 'boolean')
    || text.eastAsiaFontFamily !== undefined && (typeof text.eastAsiaFontFamily !== 'string' || !text.eastAsiaFontFamily.trim())
    || text.highlight !== undefined && (typeof text.highlight !== 'string'
      || text.highlight !== '' && text.highlight.toLowerCase() !== 'none' && !normalizeDocumentHighlight(text.highlight))
    || text.verticalAlign !== undefined && !['baseline', 'subscript', 'superscript'].includes(text.verticalAlign)) {
    throw new RangeError('文档默认字符格式无效。')
  }
  if (!['left', 'center', 'right', 'justify'].includes(paragraph.align)
    || !Number.isFinite(paragraph.indent) || !Number.isFinite(paragraph.lineHeight) || paragraph.lineHeight <= 0
    || !Number.isFinite(paragraph.spaceBefore) || paragraph.spaceBefore < 0
    || !Number.isFinite(paragraph.spaceAfter) || paragraph.spaceAfter < 0
    || ['indentStart', 'indentEnd', 'indentFirstLine', 'hanging'].some(key => {
      const value = Reflect.get(paragraph, key)
      return value !== undefined && (typeof value !== 'number' || !Number.isFinite(value))
    })
    || paragraph.spacingRule !== undefined && !['auto', 'atLeast', 'exact'].includes(paragraph.spacingRule)) {
    throw new RangeError('文档默认段落格式无效。')
  }
  if (['wordWrap', 'snapToGrid', 'keepLines', 'keepNext', 'widowControl', 'pageBreakBefore'].some(key => {
    const value = Reflect.get(paragraph, key)
    return value !== undefined && typeof value !== 'boolean'
  })) throw new RangeError('文档默认段落格式无效。')
  const resolved = {
    ...text, eastAsiaFontFamily: text.eastAsiaFontFamily ?? text.fontFamily,
    color: text.color.toLowerCase() === 'auto' ? 'auto' : text.color.startsWith('#') ? text.color : `#${text.color.toUpperCase()}`,
  }
  const highlight = normalizeDocumentHighlight(text.highlight)
  if (highlight) resolved.highlight = highlight
  else delete resolved.highlight
  return { version: 1, text: resolved, paragraph: { ...paragraph } }
}
