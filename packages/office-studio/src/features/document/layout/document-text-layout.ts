import { createDocumentDefaultStyle, type DocumentBlock } from '../../../domain/model'
import { hasEastAsianLetters } from '../model/document-font'
import type { DocumentRun, TextTarget } from '../model/document-model'
import { getLatexFormula, LatexFormulaError, layoutLatexFormula, type LatexLayout } from '../latex/latex-formula'
import { documentImageAtom } from '../images/document-image-atom'
import { documentImageBox } from './document-image-layout'
import type { LayoutGlyph, LayoutLine, LayoutMeasurer, LayoutNoteGlyphDisplay, LayoutRightTabStop, LayoutTabGlyph, LayoutTabLeader } from './document-layout-types'

export function paragraphLineHeight(fontHeight: number, style: DocumentBlock['style']): number {
  if (style.spacingRule === 'exact') return Math.max(1, style.lineHeight)
  if (style.spacingRule === 'atLeast') return Math.max(fontHeight, style.lineHeight)
  return Math.max(1, fontHeight * Math.max(0.01, style.lineHeight))
}

export function runFont(run: DocumentRun, scale = 1, family = run.fontFamily): string {
  const verticalScale = run.verticalAlign === 'subscript' || run.verticalAlign === 'superscript' ? 0.65 : 1
  return `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${Math.max(1, run.fontSize) * 4 / 3 * verticalScale * scale}px ${family || 'Aptos'}, "Microsoft YaHei", sans-serif`
}

export function createTextMeasurer(hostDocument: Document = document): LayoutMeasurer {
  const canvas = hostDocument.createElement('canvas')
  const context = canvas.getContext('2d')
  const cache = new Map<string, number>()
  return { measure(text, font) {
    const key = `${font}\u0000${text}`
    const cached = cache.get(key)
    if (cached != null) return cached
    if (context) context.font = font
    const width = context?.measureText(text).width ?? text.length * 8
    if (cache.size > 40000) cache.clear()
    cache.set(key, width)
    return width
  } }
}

export type RawGlyph = Omit<LayoutGlyph, 'x' | 'y' | 'height'>
export interface RawLine { glyphs: RawGlyph[]; start: number; end: number; width: number; fontHeight: number; ascent?: number }
export interface TextWrappingOptions { softWrap?: boolean; tabSize?: number; maximumImageHeight?: number }

/** Leader text is measured decoration; the original tab remains one UTF-16 glyph. */
function tabLeaderGlyph(width: number, leader: LayoutTabLeader, font: string, measurer: LayoutMeasurer): LayoutTabGlyph {
  const character = leader === 'dots' ? '.' : leader === 'dashes' ? '-' : leader === 'underline' ? '_' : leader === 'middleDots' ? '·' : ''
  const advance = character ? measurer.measure(character, font) : 0
  let count = advance > 0 && Number.isFinite(advance) ? Math.floor(width / advance) : 0
  if (!Number.isSafeInteger(count) || count < 1) return { leader, leaderText: '', leaderWidth: 0 }
  let text = character.repeat(count), measured = measurer.measure(text, font)
  if (!Number.isFinite(measured) || measured > width) {
    let low = 0, high = count
    while (low < high) {
      const middle = Math.ceil((low + high) / 2), candidate = measurer.measure(character.repeat(middle), font)
      if (Number.isFinite(candidate) && candidate <= width) low = middle
      else high = middle - 1
    }
    count = low; text = character.repeat(count); measured = count ? measurer.measure(text, font) : 0
  }
  return { leader, leaderText: text, leaderWidth: Math.max(0, measured) }
}

export function wrapRuns(runs: readonly DocumentRun[], width: number | ((line: number) => number), measurer: LayoutMeasurer, wordWrap = false, minimumFontHeight = 14.67,
  rightTabs?: readonly LayoutRightTabStop[] | ((line: number) => readonly LayoutRightTabStop[]),
  display?: { references?: ReadonlyMap<string, LayoutNoteGlyphDisplay>; leading?: LayoutNoteGlyphDisplay },
  options: TextWrappingOptions = {}): RawLine[] {
  const tokens: RawGlyph[] = []
  const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' })
  let offset = 0
  if (display?.leading) {
    const label = display.leading, source = runs[0], defaults = createDocumentDefaultStyle().text
    // A display label has a stable layout identity and only text marks. It must
    // neither allocate content IDs nor inherit reference/field/control metadata.
    const run: DocumentRun = { ...defaults, id: `layout:${label.noteKind}:${label.noteId}:label`, text: '',
      fontFamily: source?.fontFamily ?? defaults.fontFamily, eastAsiaFontFamily: source?.eastAsiaFontFamily ?? defaults.eastAsiaFontFamily,
      fontSize: source?.fontSize ?? 10, color: source?.color ?? defaults.color,
      bold: source?.bold ?? false, italic: source?.italic ?? false, underline: source?.underline ?? false, strike: source?.strike ?? false,
      verticalAlign: 'superscript', ...label.textStyle }
    const font = runFont(run), gap = label.label ? measurer.measure(' ', font) : 0
    tokens.push({ text: label.label, offset: 0, end: 0, width: measurer.measure(label.label, font) + gap, run, font,
      note: { noteKind: label.noteKind, noteId: label.noteId, role: 'bodyLabel' } })
  }
  for (const run of runs) {
    const atom = documentImageAtom(run)
    if (atom) {
      const box = documentImageBox(atom.image, typeof width === 'number' ? width : width(0), options.maximumImageHeight ?? Number.MAX_SAFE_INTEGER)
      tokens.push({ text: run.text, offset, end: offset + run.text.length, width: (atom.image.wrappingStyle ?? 'inline') === 'inline' ? box.width : 0,
        run, font: runFont(run), imageAtom: { imageId: atom.imageId, image: atom.image, box } })
      offset += run.text.length
      continue
    }
    const reference = display?.references?.get(run.id)
    if (reference) {
      const visualRun = { ...run, ...reference.textStyle }, font = runFont(visualRun)
      tokens.push({ text: reference.label, offset, end: offset + run.text.length, width: measurer.measure(reference.label, font), run: visualRun, font,
        note: { noteKind: reference.noteKind, noteId: reference.noteId, role: 'reference' } })
      offset += run.text.length
      continue
    }
    const latinFont = runFont(run)
    const eastAsianFont = run.eastAsiaFontFamily && runFont(run, 1, run.eastAsiaFontFamily)
    for (const part of segmenter.segment(run.text)) {
      const text = part.segment
      const font = eastAsianFont && hasEastAsianLetters(text) ? eastAsianFont : latinFont
      let latex: LatexLayout | undefined
      let latexError: LayoutGlyph['latexError']
      const formula = text === '\uFFFC' ? getLatexFormula(run) : undefined
      if (formula) {
        try { latex = layoutLatexFormula(formula.latex, { fontSize: run.fontSize * 4 / 3, measure: measurer.measure }) }
        catch (cause) {
          if (!(cause instanceof LatexFormulaError)) throw cause
          latexError = { message: cause.message, offset: cause.offset }
          const preview = Array.from(formula.latex.replace(/\s+/g, ' ')).slice(0, 48).join('')
          const label = `⚠ ${preview}${Array.from(formula.latex).length > 48 ? '…' : ''} (${cause.offset + 1})`
          const size = Math.max(8, Math.min(96, run.fontSize * 4 / 3)), errorFont = `${size}px "Cambria Math", serif`
          latex = { width: measurer.measure(label, errorFont), height: size * 1.2, baseline: size * 0.9,
            primitives: [{ kind: 'text', x: 0, y: 0, text: label, font: errorFont }] }
        }
      }
      tokens.push({ text, offset: offset + part.index, end: offset + part.index + text.length, width: text === '\n' || text === '\u2028' ? 0 : latex?.width ?? measurer.measure(text === '\t' ? '    ' : text, font), run, font, latex, latexError })
    }
    offset += run.text.length
  }
  const lines: RawLine[] = []
  let glyphs: RawGlyph[] = []
  let lineStart = 0
  let lineWidth = 0
  const commit = (end: number) => {
    const textHeight = Math.max(minimumFontHeight, ...(glyphs.length ? glyphs.map(glyph => glyph.run.fontSize * 4 / 3)
      : [runs[0]?.fontSize ? runs[0].fontSize * 4 / 3 : minimumFontHeight]))
    let ascent = textHeight * 0.82, descent = textHeight * 0.18
    const hasObject = glyphs.some(glyph => glyph.latex || glyph.imageAtom && (glyph.imageAtom.image.wrappingStyle ?? 'inline') === 'inline')
    if (hasObject) for (const glyph of glyphs) {
      const offset = glyph.run.verticalAlign === 'superscript' ? -glyph.run.fontSize * 0.4 : glyph.run.verticalAlign === 'subscript' ? glyph.run.fontSize * 0.2 : 0
      const inlineImage = glyph.imageAtom && (glyph.imageAtom.image.wrappingStyle ?? 'inline') === 'inline' ? glyph.imageAtom : undefined
      ascent = Math.max(ascent, inlineImage ? inlineImage.box.height : (glyph.latex?.baseline ?? glyph.run.fontSize * 4 / 3 * 0.82) - offset)
      descent = Math.max(descent, inlineImage ? 0 : (glyph.latex ? glyph.latex.height - glyph.latex.baseline : glyph.run.fontSize * 4 / 3 * 0.18) + offset)
    }
    lines.push({ glyphs, start: lineStart, end, width: lineWidth, fontHeight: hasObject ? ascent + descent : textHeight,
      ...(hasObject ? { ascent } : {}) })
    glyphs = []; lineWidth = 0; lineStart = end
  }
  const availableWidth = () => Math.max(1, typeof width === 'number' ? width : width(lines.length))
  const currentRightTab = (offset: number): LayoutRightTabStop | undefined => {
    const stops = (typeof rightTabs === 'function' ? rightTabs(lines.length) : rightTabs ?? [])
      .filter(stop => stop.offset === offset && (stop.right === undefined || Number.isFinite(stop.right) && stop.right >= 0))
      .sort((a, b) => (a.right ?? availableWidth()) - (b.right ?? availableWidth()))
    return stops.find(stop => Math.min(stop.right ?? availableWidth(), availableWidth()) > lineWidth) ?? stops.at(-1)
  }
  const joinsWord = (token: RawGlyph) => !token.note && /^[\p{L}\p{M}\p{N}_]+$/u.test(token.text) && !/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}]/u.test(token.text)
  for (let index = 0; index < tokens.length;) {
    const token = tokens[index]!
    if (token.text === '\n' || token.text === '\u2028') { commit(token.offset); lineStart = token.end; index++; continue }
    if (token.text === '\t' && options.tabSize !== undefined) {
      const step = Math.max(1, measurer.measure(' '.repeat(options.tabSize), token.font))
      let advance = step - lineWidth % step
      if (options.softWrap !== false && lineWidth > 0 && lineWidth + advance > availableWidth()) { commit(token.offset); advance = step }
      glyphs.push({ ...token, width: advance, tab: { leader: 'none', leaderText: '', leaderWidth: 0 } })
      lineWidth += advance; index++; continue
    }
    let tab = token.text === '\t' ? currentRightTab(token.offset) : undefined
    if (tab) {
      let followingWidth = 0
      for (let following = index + 1; following < tokens.length; following++) {
        const glyph = tokens[following]!
        if (glyph.text === '\t' || glyph.text === '\n' || glyph.text === '\u2028') break
        followingWidth += glyph.width
      }
      let target = Math.min(tab.right ?? availableWidth(), availableWidth())
      if (lineWidth > 0 && target - lineWidth - followingWidth < -1e-6) {
        commit(token.offset)
        tab = currentRightTab(token.offset) ?? tab
        target = Math.min(tab.right ?? availableWidth(), availableWidth())
      }
      const gap = Math.max(0, target - lineWidth - followingWidth)
      glyphs.push({ ...token, width: gap, tab: tabLeaderGlyph(gap, tab.leader, token.font, measurer) })
      lineWidth += gap; index++
      continue
    }
    let end = index + 1
    if (!wordWrap && joinsWord(token)) while (end < tokens.length && joinsWord(tokens[end]!)) end++
    const word = tokens.slice(index, end)
    const wordWidth = word.reduce((sum, glyph) => sum + glyph.width, 0)
    if (options.softWrap !== false && lineWidth > 0 && lineWidth + wordWidth > availableWidth()) commit(token.offset)
    for (const glyph of word) {
      if (options.softWrap !== false && lineWidth > 0 && lineWidth + glyph.width > availableWidth()) commit(glyph.offset)
      glyphs.push(glyph); lineWidth += glyph.width
    }
    index = end
  }
  commit(offset)
  return lines
}

export function positionLine(raw: RawLine, target: TextTarget, x: number, y: number, width: number, lineHeight: number, page: number, align: DocumentBlock['style']['align'], justify: boolean, marker?: string): LayoutLine {
  const tabbed = raw.glyphs.some(glyph => glyph.tab !== undefined)
  const spare = Math.max(0, width - raw.width)
  const shift = !tabbed && align === 'center' ? spare / 2 : !tabbed && align === 'right' ? spare : 0
  const spaces = raw.glyphs.filter(glyph => glyph.text === ' ').length
  const gap = !tabbed && align === 'justify' && justify && spaces ? spare / spaces : 0
  const baseline = y + (lineHeight - raw.fontHeight) / 2 + (raw.ascent ?? raw.fontHeight * 0.82)
  let cursor = x + shift
  const glyphs = raw.glyphs.map(glyph => {
    const value: LayoutGlyph = { ...glyph, x: cursor, y, height: lineHeight, width: glyph.width + (glyph.text === ' ' ? gap : 0) }
    if (glyph.latex) {
      const offset = glyph.run.verticalAlign === 'superscript' ? -glyph.run.fontSize * 0.4 : glyph.run.verticalAlign === 'subscript' ? glyph.run.fontSize * 0.2 : 0
      value.y = baseline + offset - glyph.latex.baseline; value.height = glyph.latex.height
    }
    if (glyph.imageAtom && (glyph.imageAtom.image.wrappingStyle ?? 'inline') === 'inline') {
      value.y = baseline - glyph.imageAtom.box.height; value.height = glyph.imageAtom.box.height
    }
    cursor += value.width
    return value
  })
  return { x: x + shift, y, width: Math.max(2, cursor - x - shift), height: lineHeight, glyphLineHeight: raw.fontHeight, target, start: raw.start, end: raw.end, baseline, glyphs, page, marker }
}
