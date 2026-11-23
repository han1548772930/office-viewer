export interface SheetClipboardContentAliases {
  text?: unknown
  textContent?: unknown
  html?: unknown
  htmlContent?: unknown
}

/** Undefined means omitted; an explicitly supplied empty string must never trigger a system read. */
export function normalizeSheetClipboardContentAliases(input: SheetClipboardContentAliases): { text?: string; html?: string } {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('剪贴板参数无效。')
  const result: { text?: string; html?: string } = {}
  for (const [key, alias] of [['text', 'textContent'], ['html', 'htmlContent']] as const) {
    const value = input[key], alternate = input[alias]
    if (value !== undefined && typeof value !== 'string' || alternate !== undefined && typeof alternate !== 'string') {
      throw new TypeError('剪贴板内容必须是文本。')
    }
    if (value !== undefined && alternate !== undefined && value !== alternate) throw new TypeError('剪贴板内容别名参数不一致。')
    const content = value ?? alternate
    if (content !== undefined) result[key] = content
  }
  return result
}
