import type { InjectionKey, Ref } from 'vue'
import type { RibbonItem, RibbonOption } from './ribbon-schema'

const defaultFamilies = [
  '宋体', 'Arial', 'Baskerville', 'Comic Sans MS', 'Courier New', 'Helvetica',
  'Times New Roman', 'Verdana', 'Apple LiGothic', 'Apple LiSung', 'Hei', 'Kai',
  'LiHei Pro', 'LiSong Pro', 'STFangsong', 'STHeiti', 'STKaiti', 'STSong',
  '报隶-繁', '报隶-简', '冬青黑体简体中文', '黑体-繁', '黑体-简', '华文仿宋',
  '华文楷体', '华文宋体', '楷体-繁', '楷体-简', '兰亭黑-繁', '兰亭黑-简',
  '隶变-繁', '隶变-简', '凌慧体-繁', '凌慧体-简', '翩翩体-繁', '翩翩体-简',
  '苹方-繁', '苹方-港', '苹方-简', '手札体-繁', '手札体-简', '宋体-繁',
  '宋体-简', '娃娃体-繁', '娃娃体-简', '魏碑-繁', '魏碑-简', '行楷-繁',
  '行楷-简', '雅痞-繁', '雅痞-简', '圆体-繁', '圆体-简', 'Arial Black',
  'Big Caslon', 'Courier', 'Didot', 'Futura', 'Georgia', 'Helvetica Neue',
  'Optima', 'Times', 'Tahoma', 'BIZ UDGothic', 'BIZ UDMincho', 'Hiragino Maru Gothic Pro',
  'Hiragino Maru Gothic ProN', 'Hiragino Mincho Pro', 'Hiragino Mincho ProN', 'Hiragino Sans', 'Klee', 'Osaka',
  'Toppan Bunkyu Gothic', 'Toppan Bunkyu Midashi Gothic', 'Toppan Bunkyu Midashi Mincho', 'Toppan Bunkyu Mincho', 'Tsukushi A Round Gothic', 'Tsukushi B Round Gothic',
  'YuKyokasho', 'YuKyokasho Yoko', 'YuMincho +36p Kana', 'American Typewriter', 'Andale Mono', 'Apple Chancery',
  'Apple SD Gothic Neo', 'Apple Symbols', 'AppleGothic', 'AppleMyungjo', 'Arial Unicode MS', 'Athelas',
  'Avenir', 'Avenir Next', 'Avenir Next Condensed', 'Ayuthaya', 'Bangla MN', 'Bangla Sangam MN',
  'Bauhaus 93', 'Bodoni 72', 'Bodoni 72 Oldstyle', 'Bodoni 72 Smallcaps', 'Bodoni Ornaments', 'Bookshelf Symbol 7',
  'Bradley Hand', 'Chalkboard', 'Chalkboard SE', 'Chalkduster', 'Charter', 'Cochin',
  'Copperplate', 'Devanagari Sangam MN', 'Din Alternate', 'Euphemia UCAS', 'Galvji', 'Geneva',
  'Gill Sans', 'Gujarati Sangam MN', 'GungSeo', 'Gurmukhi MN', 'Gurmukhi Sangam MN', 'HeadLineA',
  'Herculanum', 'Hiragino Kaku Gothic Pro', 'Hiragino Kaku Gothic ProN', 'Hiragino Kaku Gothic Std', 'Hiragino Kaku Gothic StdN', 'Hoefler Text',
  'Impact', 'InaiMathi', 'Iowan Old Style', 'Kannada MN', 'Kannada Sangam MN', 'Kefa',
  'Khmer MN', 'Khmer Sangam MN', 'Kohinoor Bangla', 'Kohinoor Devanagari', 'Kohinoor Telugu', 'Krungthep',
  'Lao MN', 'Lao Sangam MN', 'Lucida Grande', 'Luminari', 'Malayalam MN', 'Malayalam Sangam MN',
  'Marion', 'Marker Felt', 'Menlo', 'Microsoft Sans Serif', 'Modern No. 20', 'Monaco',
  'Mukta Mahee', 'Nanum Brush Script', 'Nanum Gothic', 'Nanum Myeongjo', 'Nanum Pen Script', 'Noteworthy',
  'Oriya MN', 'Oriya Sangam MN', 'Palatino', 'Papyrus', 'PCMyungjo', 'Phosphate',
  'PilGi', 'Plantagenet Cherokee', 'PT Mono', 'PT Sans', 'PT Sans Caption', 'PT Serif',
  'PT Serif Caption', 'Rockwell', 'Sathu', 'Savoye LET', 'Seravek', 'Shree Devanagari 714',
  'SignPainter', 'Silom', 'Sinhala MN', 'Sinhala Sangam MN', 'Skia', 'Snell Roundhand',
  'STIXGeneral', 'Sukhumvit Set', 'Superclarendon', 'Tamil MN', 'Tamil Sangam MN', 'Telugu MN',
  'Telugu Sangam MN', 'Trattatello', 'Trebuchet MS', 'Webdings', 'Wingdings', 'Wingdings 2',
  'Wingdings 3', 'YuGothic', 'YuGothic Light', 'YuGothic Medium', 'Zapfino',
]

export interface OfficeFontList { families: readonly { value: string; label?: string }[]; override?: boolean }
export const OfficeFontListKey: InjectionKey<Readonly<Ref<OfficeFontList>>> = Symbol('office-font-list')

const compatibilityFamilies = ['Aptos', 'Calibri', 'Microsoft YaHei', 'SimSun', '等线', '微软雅黑']
export function ribbonFontOptions(current?: string | null, config?: OfficeFontList): RibbonOption[] {
  const options = new Map<string, RibbonOption>()
  const add = (value: string, label = value) => { options.set(value, { label, value, preview: { fontFamily: value } }) }
  if (!config?.override) for (const value of [...defaultFamilies, ...compatibilityFamilies]) add(value)
  for (const font of config?.families ?? []) if (font.value.trim()) add(font.value, font.label)
  if (current && !options.has(current)) add(current)
  return [...options.values()]
}

export const fontSizeOptions = [9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map(value => ({ label: String(value), value }))

export function confirmRibbonDraft(item: RibbonItem, draft: string): string | number | undefined {
  const text = draft.trim()
  if (!text || item.disabled || item.hidden) return
  if (item.edit?.kind === 'font') return item.options?.find(option => !option.disabled && option.label.toLowerCase().includes(text.toLowerCase()))?.value
  if (item.edit?.kind !== 'number' || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)) return
  const number = Number(text)
  return Number.isFinite(number) ? Math.max(item.edit.min, Math.min(item.edit.max, number)) : undefined
}
