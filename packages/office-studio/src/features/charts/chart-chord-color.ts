/** Theme color values below are adapted from Univer v1.6.0
 * packages/themes/src/default.ts and dark-blue.ts (DreamNum Co., Ltd.,
 * Apache-2.0: https://www.apache.org/licenses/LICENSE-2.0).
 * Only color facts are retained; token lookup and host theme adaptation are local. */
import { chartColorChannels, chartSolidColorOpacity, validateChartColor } from './chart-color'

const LEVELS = [50, 100, 200, 300, 400, 500, 600, 700, 800, 900] as const
const LIGHT_ROWS: Readonly<Record<string, readonly string[]>> = {
  primary: ['#F3F5FF', '#E9EDFF', '#D2DAFA', '#BAC6F8', '#6280F9', '#466AF7', '#2C53F1', '#143EE3', '#083AD1', '#1033BF'],
  blue: ['#EBF5FF', '#E1EFFE', '#C3DDFD', '#A4CAFE', '#76A9FA', '#3F83F8', '#1C64F2', '#1A56DB', '#1E429F', '#233876'],
  red: ['#FDF2F2', '#FDE8E8', '#FBD5D5', '#F8B4B4', '#F98080', '#F05252', '#E02424', '#C81E1E', '#9B1C1C', '#771D1D'],
  orange: ['#FFF8F1', '#FEECDC', '#FCD9BD', '#FDBA8C', '#FF8A4C', '#FF5A1F', '#D03801', '#B43403', '#8A2C0D', '#771D1D'],
  yellow: ['#FDFCEA', '#FFF4B9', '#FCDC6A', '#FAC815', '#F1B312', '#D49D0F', '#AB7F0E', '#9A6D15', '#725213', '#634312'],
  green: ['#F3FAF7', '#DEF7EC', '#BCF0DA', '#84E1BC', '#31C48D', '#0DA471', '#057A55', '#046C4E', '#03543F', '#014737'],
  jiqing: ['#EDFAFA', '#D5F5F6', '#AFECEF', '#7EDCE2', '#16BDCA', '#0694A2', '#047481', '#036672', '#05505C', '#014451'],
  indigo: ['#F3F5FF', '#E9EDFF', '#D2DAFA', '#BAC6F8', '#6280F9', '#466AF7', '#2C53F1', '#143EE3', '#083AD1', '#1033BF'],
  purple: ['#F6F5FF', '#EDEBFE', '#DCD7FE', '#CABFFD', '#AC94FA', '#9061F9', '#7E3AF2', '#6C2BD9', '#5521B5', '#4A1D96'],
  pink: ['#FDF2F8', '#FCE8F3', '#FAD1E8', '#F8B4D9', '#F17EB8', '#E74694', '#D61F69', '#BF125D', '#99154B', '#751A3D'],
}
const DARK_ROWS: Readonly<Record<string, readonly string[]>> = {
  primary: ['#1E3A8A', '#1E40AF', '#1D4ED8', '#2563EB', '#3B82F6', '#60A5FA', '#93C5FD', '#BFDBFE', '#DBEAFE', '#EFF6FF'],
  blue: ['#0C4A6E', '#075985', '#0369A1', '#0284C7', '#0EA5E9', '#38BDF8', '#7DD3FC', '#BAE6FD', '#E0F2FE', '#F0F9FF'],
  red: ['#7F1D1D', '#991B1B', '#B91C1C', '#DC2626', '#EF4444', '#F87171', '#FCA5A5', '#FECACA', '#FEE2E2', '#FEF2F2'],
  orange: ['#7C2D12', '#9A3412', '#C2410C', '#EA580C', '#F97316', '#FB923C', '#FDBA74', '#FED7AA', '#FFEDD5', '#FFF7ED'],
  yellow: ['#78350F', '#92400E', '#B45309', '#D97706', '#F59E0B', '#FBBF24', '#FCD34D', '#FDE68A', '#FEF3C7', '#FFFBEB'],
  green: ['#064E3B', '#065F46', '#047857', '#059669', '#10B981', '#34D399', '#6EE7B7', '#A7F3D0', '#D1FAE5', '#ECFDF5'],
  jiqing: ['#134E4A', '#115E59', '#0F766E', '#0D9488', '#14B8A6', '#2DD4BF', '#5EEAD4', '#99F6E4', '#CCFBF1', '#F0FDFA'],
  indigo: ['#312E81', '#3730A3', '#4338CA', '#4F46E5', '#6366F1', '#818CF8', '#A5B4FC', '#C7D2FE', '#E0E7FF', '#EEF2FF'],
  purple: ['#4C1D95', '#5B21B6', '#6D28D9', '#7C3AED', '#8B5CF6', '#A78BFA', '#C4B5FD', '#DDD6FE', '#EDE9FE', '#F5F3FF'],
  pink: ['#831843', '#9D174D', '#BE185D', '#DB2777', '#EC4899', '#F472B6', '#F9A8D4', '#FBCFE8', '#FCE7F3', '#FDF2F8'],
}
const GRAY_LEVELS = [0, ...LEVELS, 1000]
const LIGHT_GRAY = ['#FFFFFF', '#F9FAFB', '#EEEFF1', '#E3E5EA', '#A8B0BD', '#7D8698', '#4E5565', '#31363F', '#272A2F', '#1F2124', '#1B1C1F', '#000000']
const DARK_GRAY = ['#07111F', '#111318', '#202329', '#343840', '#4A4F59', '#626873', '#838A95', '#ABB1BA', '#CFD3D9', '#E4E7EB', '#F5F6F7', '#F8FAFC']
const LIGHT_LOOP = ['purple.400', 'green.500', 'blue.500', 'yellow.400', 'pink.300', 'jiqing.600', 'orange.400', 'gray.800', 'indigo.500', 'red.300', 'green.600', 'yellow.700']
const DARK_LOOP = ['primary.600', 'green.600', 'purple.600', 'orange.600', 'pink.600', 'jiqing.600', 'yellow.600', 'red.600', 'indigo.600', 'blue.600', 'green.700', 'purple.700']
const LIGHT_HIGHLIGHT = ['purple.500', 'red.500', 'orange.400', 'green.400', 'blue.500', 'primary.500', 'pink.500', 'gray.400']
const DARK_HIGHLIGHT = ['primary.600', 'green.600', 'purple.600', 'orange.600', 'pink.600', 'jiqing.600', 'yellow.600', 'red.600']

function colors(rows: Readonly<Record<string, readonly string[]>>, gray: readonly string[], loop: readonly string[], highlights: readonly string[], dark: boolean): Readonly<Record<string, string>> {
  const result: Record<string, string> = {}
  for (const [group, row] of Object.entries(rows)) for (const [index, color] of row.entries()) result[`${group}.${LEVELS[index]}`] = color
  for (const [index, color] of gray.entries()) result[`gray.${GRAY_LEVELS[index]}`] = color
  for (const [index, token] of loop.entries()) result[`loop-color.${index + 1}`] = result[token]!
  for (let index = 0; index < 16; index++) result[`highlight.background.${index + 1}`] = chartSolidColorOpacity(result[highlights[index % 8]!]!, index < 8 ? dark ? 0.28 : 0.3 : dark ? 0.14 : 0.15)
  return Object.freeze(result)
}
export const CHART_CHORD_LIGHT_COLORS = colors(LIGHT_ROWS, LIGHT_GRAY, LIGHT_LOOP, LIGHT_HIGHLIGHT, false)
export const CHART_CHORD_DARK_COLORS = colors(DARK_ROWS, DARK_GRAY, DARK_LOOP, DARK_HIGHLIGHT, true)
export const CHART_CHORD_COLOR_TOKENS = Object.freeze(Object.keys(CHART_CHORD_LIGHT_COLORS))
export interface ChartChordColorTheme {
  backgroundColor?: string
  textColor?: string
  /** Hosts with a custom theme may override any fixed token using CSS colors. */
  tokenColors?: Readonly<Record<string, string>>
}

export function validateChartChordColor(value: unknown, name = '和弦颜色'): void {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${name}须为 CSS 颜色或固定主题颜色 token。`)
  if (Object.hasOwn(CHART_CHORD_LIGHT_COLORS, value)) return
  validateChartColor(value, name)
}

/** Authored tokens remain unchanged in saved options. Theme selection and CSS
 * conversion occur only at the paint boundary, including opaque borders. */
export function resolveChartChordColor(value: string | undefined, theme: ChartChordColorTheme = {}, fallback = 'gray.700'): string {
  const color = value ?? fallback
  validateChartChordColor(color)
  if (!Object.hasOwn(CHART_CHORD_LIGHT_COLORS, color)) return color
  const overridden = theme.tokenColors?.[color]
  if (overridden !== undefined) { validateChartColor(overridden); return overridden }
  if (color === 'gray.700' && theme.textColor) { validateChartColor(theme.textColor); return theme.textColor }
  if (color === 'gray.0' && theme.backgroundColor) { validateChartColor(theme.backgroundColor); return theme.backgroundColor }
  const background = chartColorChannels(theme.backgroundColor ?? '#FFFFFF')
  const dark = background[0] * 0.2126 + background[1] * 0.7152 + background[2] * 0.0722 < 0.5
  return (dark ? CHART_CHORD_DARK_COLORS : CHART_CHORD_LIGHT_COLORS)[color]!
}
