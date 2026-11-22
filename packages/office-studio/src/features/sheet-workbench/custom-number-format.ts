import type { CellScalar } from '../../domain/model.ts'
import { formatSheetValue } from './sheet-number-format.ts'
import { isValidSheetNumberFormat } from './number-format-render.ts'

export interface CustomNumberFormatResult {
  pattern: string
  preview: string | undefined
}

export function validateCustomNumberFormat(input: string): CustomNumberFormatResult {
  if (/[\u0000-\u001f\u007f-\u009f]/u.test(input)) throw new Error('格式代码不能包含控制字符。')
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
  }
  if (quoted) throw new Error('格式代码中的引号不匹配。')
  if (escaped) throw new Error('格式代码不能以转义符结尾。')
  if (bracketDepth !== 0) throw new Error('格式代码中的方括号不匹配。')
  if (sectionCount > 4) throw new Error('格式代码最多支持四个分段。')
  if (!isValidSheetNumberFormat(pattern)) throw new Error('格式代码无效。')

  return { pattern, preview: formatSheetValue(1234.567, pattern) }
}

export function previewCustomNumberFormat(pattern: string, value: CellScalar, locale = 'zh-CN', date1904 = false): string | undefined {
  return formatSheetValue(value, pattern, date1904, locale)
}
