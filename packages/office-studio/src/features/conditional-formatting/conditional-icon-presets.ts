export interface ConditionalIconPreset { symbol: string; color: string; shape: 'arrow' | 'circle' | 'triangle' | 'diamond' | 'check' | 'warning' | 'cross' | 'flag' | 'rating' | 'progress' | 'face' | 'box' | 'star' }

const green = '#22a06b', gold = '#d9a300', red = '#d94841', gray = '#777777'
function presets(shape: ConditionalIconPreset['shape'], symbols: string[], colors: string[]): ConditionalIconPreset[] {
  return symbols.map((symbol, index) => ({ symbol, color: colors[index] ?? gray, shape }))
}

// Canvas renders these icon presets with text glyphs.
export const conditionalIconPresets: Record<string, ConditionalIconPreset[]> = {
  '3Arrows': presets('arrow', ['↑', '→', '↓'], [green, gold, red]),
  '3ArrowsGray': presets('arrow', ['↑', '→', '↓'], [gray, gray, gray]),
  '4Arrows': presets('arrow', ['↑', '↗', '↘', '↓'], [green, gold, gold, red]),
  '4ArrowsGray': presets('arrow', ['↑', '↗', '↘', '↓'], [gray, gray, gray, gray]),
  '5Arrows': presets('arrow', ['↑', '↗', '→', '↘', '↓'], [green, gold, gold, gold, red]),
  '5ArrowsGray': presets('arrow', ['↑', '↗', '→', '↘', '↓'], [gray, gray, gray, gray, gray]),
  '3Triangles': [{ symbol: '▲', color: green, shape: 'triangle' }, { symbol: '−', color: gold, shape: 'cross' }, { symbol: '▼', color: red, shape: 'triangle' }],
  '3TrafficLights1': presets('circle', ['●', '●', '●'], [green, gold, red]),
  '3Signs': [{ symbol: '●', color: green, shape: 'circle' }, { symbol: '▲', color: gold, shape: 'triangle' }, { symbol: '◆', color: red, shape: 'diamond' }],
  '3TrafficLights2': presets('circle', ['◉', '◉', '◉'], [green, gold, red]),
  '4RedToBlack': presets('circle', ['●', '●', '●', '●'], [red, '#ee9090', gray, '#222222']),
  '4TrafficLights': presets('circle', ['●', '●', '●', '●'], [green, gold, red, '#222222']),
  '3Symbols': [{ symbol: '✓', color: green, shape: 'check' }, { symbol: '!', color: gold, shape: 'warning' }, { symbol: '✕', color: red, shape: 'cross' }],
  '3Symbols2': [{ symbol: '✓', color: green, shape: 'check' }, { symbol: '!', color: gold, shape: 'warning' }, { symbol: '✕', color: red, shape: 'cross' }],
  '3Flags': presets('flag', ['⚑', '⚑', '⚑'], [green, gold, red]),
  '4Rating': presets('rating', ['▂', '▃', '▅', '▇'], [gray, gray, gray, gray]),
  '5Rating': presets('rating', ['·', '▂', '▃', '▅', '▇'], [gray, gray, gray, gray, gray]),
  '5Quarters': presets('progress', ['●', '◕', '◐', '◔', '○'], [gray, gray, gray, gray, gray]),
  '_5Felling': presets('face', ['☺', '☺', '⊖', '☹', '☹'], [green, green, gold, red, red]),
  '5Boxes': presets('box', ['■', '▣', '▤', '▥', '□'], [gray, gray, gray, gray, gray]),
  '3Stars': presets('star', ['★', '◐', '☆'], [gold, gold, gold]),
}

export function conditionalIconPreset(iconType: string, iconId: string): ConditionalIconPreset | undefined {
  return conditionalIconPresets[iconType]?.[Number(iconId)]
}
