/** Fixed docs-code-ui v1.6.0 language menu; grammar support is tracked separately. */
export const DOCUMENT_STRUCTURED_CODE_LANGUAGES = [
  { value: 'plaintext', label: '纯文本' }, { value: 'javascript', label: 'JavaScript' },
  { value: 'typescript', label: 'TypeScript' }, { value: 'json', label: 'JSON' },
  { value: 'markup', label: 'HTML / XML' }, { value: 'css', label: 'CSS' },
  { value: 'markdown', label: 'Markdown' }, { value: 'yaml', label: 'YAML' },
  { value: 'sql', label: 'SQL' }, { value: 'bash', label: 'Bash' },
  { value: 'python', label: 'Python' }, { value: 'java', label: 'Java' },
  { value: 'c', label: 'C' }, { value: 'cpp', label: 'C++' }, { value: 'csharp', label: 'C#' },
  { value: 'go', label: 'Go' }, { value: 'rust', label: 'Rust' }, { value: 'ruby', label: 'Ruby' },
  { value: 'php', label: 'PHP' }, { value: 'swift', label: 'Swift' }, { value: 'kotlin', label: 'Kotlin' },
] as const

export function documentStructuredCodeLanguageOptions(language: string): Array<{ value: string; label: string }> {
  const options: Array<{ value: string; label: string }> = DOCUMENT_STRUCTURED_CODE_LANGUAGES.map(item => ({ ...item }))
  if (!options.some(option => option.value === language)) options.push({ value: language, label: language === 'plain' ? '纯文本（plain）' : language })
  return options
}
