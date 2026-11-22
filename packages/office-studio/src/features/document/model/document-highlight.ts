const officeHighlights: Readonly<Record<string, string>> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00',
  magenta: '#FF00FF', red: '#FF0000', yellow: '#FFFF00', white: '#FFFFFF',
  darkblue: '#000080', darkcyan: '#008080', darkgreen: '#008000', darkmagenta: '#800080',
  darkred: '#800000', darkyellow: '#808000', darkgray: '#808080', lightgray: '#C0C0C0',
}

/** Office highlight names, domain defaults and run shading share the Canvas/writer RGB contract. */
export function normalizeDocumentHighlight(highlight: unknown, background?: unknown): string | undefined {
  const direct = typeof highlight === 'string' ? highlight.trim() : ''
  // "none" turns off the Office palette highlight; independent run shading
  // can still provide a background. The Rust adapter restores explicit nil.
  const source = direct && direct.toLowerCase() !== 'none' ? direct : typeof background === 'string' ? background.trim() : ''
  if (!source || source.toLowerCase() === 'none' || source.toLowerCase() === 'auto') return undefined
  const key = source.toLowerCase()
  const standard = Object.hasOwn(officeHighlights, key) ? officeHighlights[key] : undefined
  if (standard) return standard
  return /^#?[\da-f]{6}$/i.test(source) ? `#${source.replace(/^#/, '').toUpperCase()}` : undefined
}
