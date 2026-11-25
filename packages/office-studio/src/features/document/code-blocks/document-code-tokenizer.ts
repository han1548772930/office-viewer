import { DOCUMENT_CODE_PRISM_CONFIG } from './document-code-prism-config'
import Prism from 'prismjs'
import 'prismjs/components/prism-typescript'
import 'prismjs/components/prism-json'
import 'prismjs/components/prism-yaml'
import 'prismjs/components/prism-markdown'
import 'prismjs/components/prism-sql'
import 'prismjs/components/prism-bash'
import 'prismjs/components/prism-python'
import 'prismjs/components/prism-java'
import 'prismjs/components/prism-c'
import 'prismjs/components/prism-cpp'
import 'prismjs/components/prism-csharp'
import 'prismjs/components/prism-go'
import 'prismjs/components/prism-rust'
import 'prismjs/components/prism-ruby'
import 'prismjs/components/prism-markup-templating'
import 'prismjs/components/prism-php'
import 'prismjs/components/prism-swift'
import 'prismjs/components/prism-kotlin'
import { DOCUMENT_STRUCTURED_CODE_LANGUAGES, type DocumentStructuredCodeLanguage } from './document-code-languages'
import { normalizeDocumentCodeSource } from '../model/document-code-source'

// Use only the lexical API. Canvas owns rendering and the existing Worker protocol owns messages.
Prism.manual = DOCUMENT_CODE_PRISM_CONFIG.manual
Prism.disableWorkerMessageHandler = DOCUMENT_CODE_PRISM_CONFIG.disableWorkerMessageHandler

export type DocumentCodeLanguage = DocumentStructuredCodeLanguage | 'plain' | 'html'
export type CodeTokenKind = 'keyword' | 'string' | 'number' | 'comment' | 'operator' | 'punctuation' | 'plain'
export interface CodeToken { kind: CodeTokenKind; text: string; start: number; end: number }

const languages = new Set<string>(DOCUMENT_STRUCTURED_CODE_LANGUAGES.map(language => language.value))
const aliases: Record<string, DocumentStructuredCodeLanguage> = {
  plain: 'plaintext', text: 'plaintext', none: 'plaintext', html: 'markup', xml: 'markup', js: 'javascript', ts: 'typescript',
  md: 'markdown', yml: 'yaml', sh: 'bash', shell: 'bash', py: 'python', 'c++': 'cpp', 'c#': 'csharp', cs: 'csharp',
  dotnet: 'csharp', rb: 'ruby', kt: 'kotlin', kts: 'kotlin',
}

/** Resolution is view-only: unknown and author-supplied language values remain in the model. */
export function documentCodeGrammarLanguage(value: string): DocumentStructuredCodeLanguage | undefined {
  const normalized = value.trim().toLowerCase(), alias = Object.hasOwn(aliases, normalized) ? aliases[normalized] : undefined
  return alias ?? (languages.has(normalized) ? normalized as DocumentStructuredCodeLanguage : undefined)
}

export function documentCodeHasGrammar(value: string): boolean {
  const language = documentCodeGrammarLanguage(value)
  return language === 'plaintext' || !!(language && Prism.languages[language])
}

function tokenKind(type: string, alias: string | string[] | undefined, inherited: CodeTokenKind): CodeTokenKind {
  const names = [type, ...(Array.isArray(alias) ? alias : alias ? [alias] : [])]
  if (names.some(name => ['comment', 'prolog', 'doctype', 'cdata'].includes(name))) return 'comment'
  if (names.some(name => ['string', 'char', 'attr-value', 'url', 'regex', 'template-string'].includes(name))) return 'string'
  if (names.some(name => name === 'number')) return 'number'
  if (names.some(name => ['keyword', 'boolean', 'builtin', 'class-name', 'function', 'property', 'tag', 'attr-name', 'selector', 'atrule',
    'important', 'title', 'bold', 'italic', 'list', 'blockquote', 'inserted', 'deleted'].includes(name))) return 'keyword'
  if (names.some(name => name === 'operator')) return 'operator'
  if (names.some(name => name === 'punctuation')) return 'punctuation'
  if (names.some(name => name === 'interpolation' || name === 'embedded' || name.startsWith('language-'))) return 'plain'
  return inherited
}

/** Nested lexical tokens become a complete, non-overlapping UTF-16 partition of the authoritative source. */
export function tokenizeCode(source: string, languageValue: string): CodeToken[] {
  const text = normalizeDocumentCodeSource(source), language = documentCodeGrammarLanguage(languageValue)
  const grammar = language && language !== 'plaintext' ? Prism.languages[language] : undefined
  if (!grammar) return text ? [{ kind: 'plain', text, start: 0, end: text.length }] : []
  const output: CodeToken[] = []
  let cursor = 0
  const append = (value: string, kind: CodeTokenKind) => {
    if (!value) return
    const start = cursor; cursor += value.length
    const previous = output.at(-1)
    if (previous?.kind === kind && previous.end === start) { previous.text += value; previous.end = cursor }
    else output.push({ kind, text: value, start, end: cursor })
  }
  const visit = (value: string | Prism.Token | Array<string | Prism.Token>, inherited: CodeTokenKind) => {
    if (typeof value === 'string') append(value, inherited)
    else if (Array.isArray(value)) for (const child of value) visit(child, inherited)
    else visit(value.content, tokenKind(value.type, value.alias, inherited))
  }
  visit(Prism.tokenize(text, grammar), 'plain')
  return output
}
