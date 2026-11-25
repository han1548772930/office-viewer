import type { DocumentParagraphStyle } from '../../domain/model'
import { DOCUMENT_PARAGRAPH_STYLE_FIELDS } from '../../features/document/services/document-settings'

export const DOCUMENT_PARAGRAPH_SETTING_FIELDS = DOCUMENT_PARAGRAPH_STYLE_FIELDS

export type DocumentParagraphSettingField = typeof DOCUMENT_PARAGRAPH_SETTING_FIELDS[number]
export type DocumentParagraphNumericField = 'indentStart' | 'indentEnd' | 'indentFirstLine' | 'hanging' | 'spaceBefore' | 'spaceAfter' | 'lineHeight'
export type DocumentParagraphSpacingRule = NonNullable<DocumentParagraphStyle['spacingRule']>
export interface DocumentParagraphLineMetrics { glyphLineHeight: number; renderedLineHeight: number }

export function documentParagraphSettingsStyle(style: Partial<DocumentParagraphStyle>): Required<DocumentParagraphStyle> {
  return {
    align: style.align ?? 'left', indent: style.indent ?? 0,
    lineHeight: style.lineHeight ?? 1.5, spaceBefore: style.spaceBefore ?? 0, spaceAfter: style.spaceAfter ?? 8,
    indentStart: style.indentStart ?? (style.indent ?? 0) * 24, indentEnd: style.indentEnd ?? 0,
    indentFirstLine: style.indentFirstLine ?? 0, hanging: style.hanging ?? 0,
    spacingRule: style.spacingRule ?? 'auto', wordWrap: style.wordWrap ?? false, snapToGrid: style.snapToGrid ?? false,
    keepLines: style.keepLines ?? false, keepNext: style.keepNext ?? false,
    widowControl: style.widowControl ?? false, pageBreakBefore: style.pageBreakBefore ?? false,
  }
}

export function isDocumentParagraphSpacingRule(value: unknown): value is DocumentParagraphSpacingRule {
  return value === 'auto' || value === 'atLeast' || value === 'exact'
}

/** Convert a displayed line's height once when its unit changes between px and a multiplier. */
export function changeDocumentParagraphSpacingRule(
  value: number, currentRule: DocumentParagraphSpacingRule, nextRule: DocumentParagraphSpacingRule,
  metrics?: DocumentParagraphLineMetrics, fontSize = 11,
): number {
  const currentValue = Number.isFinite(value) && value > 0 ? value : currentRule === 'auto' ? 1.5 : 1
  if (currentRule === nextRule) return currentValue
  const fallbackFontSize = Number.isFinite(fontSize) && fontSize > 0 ? fontSize : 11
  const glyphLineHeight = metrics && Number.isFinite(metrics.glyphLineHeight) && metrics.glyphLineHeight > 0
    ? metrics.glyphLineHeight : fallbackFontSize * 4 / 3
  const renderedLineHeight = metrics && Number.isFinite(metrics.renderedLineHeight) && metrics.renderedLineHeight > 0
    ? metrics.renderedLineHeight : currentRule === 'auto' ? glyphLineHeight * currentValue
      : currentRule === 'atLeast' ? Math.max(glyphLineHeight, currentValue) : currentValue
  if (nextRule === 'auto') return Math.round(Math.max(1, Math.min(5, renderedLineHeight / glyphLineHeight)) * 100) / 100
  return Math.round(Math.max(1, Math.min(100, renderedLineHeight)) * 100) / 100
}

export function documentParagraphLineSpacingInput(rule: DocumentParagraphSpacingRule) {
  return rule === 'auto' ? { min: 1, max: 5, step: 0.1, precision: 2 }
    : { min: 1, max: 100, step: 0.1, precision: 2 }
}
