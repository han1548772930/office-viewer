import type { CellScalar } from '../../domain/model.ts'
import { formatSheetValue } from './sheet-number-format.ts'

export interface CustomNumberFormatResult {
  pattern: string
  preview: string | undefined
}

export function validateCustomNumberFormat(input: string): CustomNumberFormatResult {
  const pattern = input.trim()
  if (!pattern) throw new Error('格式代码不能为空。')
  if (pattern.length > 255) throw new Error('格式代码不能超过 255 个字符。')

  let quoted = false
  let escaped = false
  let bracketDepth = 0
  let sectionCount = 1
  for (const character of pattern) {
    if (escaped) { escaped = false; continue }
    if (character === '\\') { escaped = true; continue }
    if (character === '"' && bracketDepth === 0) { quoted = !quoted; continue }
    if (quoted) continue
    if (character === '[') {
      bracketDepth += 1
      if (bracketDepth > 1) throw new Error('格式代码中的方括号不能嵌套。')
    } else if (character === ']') {
      bracketDepth -= 1
      if (bracketDepth < 0) throw new Error('格式代码中的方括号不匹配。')
    } else if (character === ';' && bracketDepth === 0) sectionCount += 1
    else if (character.charCodeAt(0) < 0x20) throw new Error('格式代码不能包含控制字符。')
  }
  if (quoted) throw new Error('格式代码中的引号不匹配。')
  if (escaped) throw new Error('格式代码不能以转义符结尾。')
  if (bracketDepth !== 0) throw new Error('格式代码中的方括号不匹配。')
  if (sectionCount > 4) throw new Error('格式代码最多支持四个分段。')
  const tokens = pattern.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '')
  if (!/[0#?@]/.test(tokens) && !/(?:y{2,4}|m{1,4}|d{1,4}|h{1,2}|s{1,2})/i.test(tokens)) {
    throw new Error('格式代码没有可识别的数字、日期或文本占位符。')
  }

  return { pattern, preview: formatSheetValue(1234.567, pattern) }
}

export function previewCustomNumberFormat(pattern: string, value: CellScalar, locale = 'zh-CN', date1904 = false): string | undefined {
  return formatSheetValue(value, pattern, date1904, locale)
}
