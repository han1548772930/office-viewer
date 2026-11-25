import type { CellScalar } from '../../domain/model.ts'

export interface SheetFillSeriesSeed { value: CellScalar; formula?: string | null; date?: boolean }
export interface SheetFillSeriesGroup {
  start: number
  length: number
  kind: 'copy' | 'number' | 'date' | 'extend' | 'chinese' | 'cycle'
  intercept: number
  step: number
  ratio?: number
  cycle?: readonly string[]
  prefix?: string
  suffix?: string
  padding?: readonly number[]
}
export interface SheetFillSeriesPattern {
  length: number
  groups: readonly SheetFillSeriesGroup[]
  supportsSeries: boolean
  preferredApplyType: 'COPY' | 'SERIES'
}

const CHINESE_DIGITS = '零一二三四五六七八九'
const CHINESE_SMALL_UNITS: Readonly<Record<string, number>> = { 十: 10, 百: 100, 千: 1000 }
const CYCLES: readonly (readonly string[])[] = [
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'],
  ['正月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '腊月'],
  ['子', '丑', '寅', '卯', '辰', '巳', '午', '未', '申', '酉', '戌', '亥'],
  ['子时', '丑时', '寅时', '卯时', '辰时', '巳时', '午时', '未时', '申时', '酉时', '戌时', '亥时'],
  ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛', '壬', '癸'],
  ['春', '夏', '秋', '冬'], ['春季', '夏季', '秋季', '冬季'],
  ['一', '二', '三', '四', '五', '六', '日'],
  ['周一', '周二', '周三', '周四', '周五', '周六', '周日'],
  ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'],
]
const modulo = (value: number, length: number): number => (value % length + length) % length
const close = (first: number, second: number): boolean => Math.abs(first - second) <= 1e-10 * Math.max(1, Math.abs(first), Math.abs(second))

function chineseNumber(text: string): number | undefined {
  if (!/^[零一二三四五六七八九十百千万亿]+$/u.test(text)) return undefined
  const sections = (value: string): number => {
    let result = 0, digit = 0
    for (const char of value) {
      const index = CHINESE_DIGITS.indexOf(char)
      if (index >= 0) digit = index
      else { result += (digit || 1) * (CHINESE_SMALL_UNITS[char] ?? 0); digit = 0 }
    }
    return result + digit
  }
  const tenThousand = (value: string): number => {
    const parts = value.split('万')
    return parts.length === 1 ? sections(value) : sections(parts[0]!) * 10000 + sections(parts[1] ?? '')
  }
  const parts = text.split('亿')
  if (parts.length > 2 || parts.some(part => part.split('万').length > 2)) return undefined
  const value = parts.length === 1 ? tenThousand(text) : tenThousand(parts[0]!) * 100000000 + tenThousand(parts[1] ?? '')
  return Number.isSafeInteger(value) ? value : undefined
}

function formatChineseNumber(value: number): string | undefined {
  if (!Number.isSafeInteger(value) || Math.abs(value) >= 1e16) return undefined
  if (value === 0) return '零'
  if (value < 0) { const positive = formatChineseNumber(-value); return positive === undefined ? undefined : '负' + positive }
  const section = (number: number): string => {
    let result = '', pendingZero = false
    for (let position = 3; position >= 0; position--) {
      const power = 10 ** position, digit = Math.floor(number / power) % 10
      if (!digit) { if (result) pendingZero = true; continue }
      if (pendingZero) result += '零'
      result += CHINESE_DIGITS[digit] + ['', '十', '百', '千'][position]!
      pendingZero = false
    }
    return result
  }
  let result = '', remaining = value
  for (const [power, unit] of [[1e12, '万亿'], [1e8, '亿'], [1e4, '万'], [1, '']] as const) {
    const part = Math.floor(remaining / power)
    if (!part) continue
    if (result && part < 1000) result += '零'
    result += section(part) + unit; remaining %= power
  }
  return result.replace(/^一十/u, '十')
}

interface ClassifiedSeed { kind: SheetFillSeriesGroup['kind']; key: string; value?: number; cycle?: readonly string[]; prefix?: string; suffix?: string; padding?: number }
function classify(seed: SheetFillSeriesSeed, preferredCycle?: readonly string[]): ClassifiedSeed {
  if (seed.formula) return { kind: 'copy', key: 'formula' }
  if (typeof seed.value === 'number' && Number.isFinite(seed.value)) return { kind: seed.date ? 'date' : 'number', key: seed.date ? 'date' : 'number', value: seed.value }
  if (typeof seed.value !== 'string') return { kind: 'copy', key: 'other' }
  const matches = [...seed.value.matchAll(/\d+/g)], match = matches.at(-1)
  if (match) {
    const value = Number(match[0]), prefix = seed.value.slice(0, match.index), suffix = seed.value.slice(match.index! + match[0].length)
    if (Number.isSafeInteger(value)) return { kind: 'extend', key: JSON.stringify(['extend', prefix, suffix]), value, prefix, suffix,
      padding: match[0].length }
  }
  const value = chineseNumber(seed.value)
  if (value !== undefined || seed.value === '日') return { kind: 'chinese', key: 'chinese', value: value ?? 7 }
  const cycle = preferredCycle?.includes(seed.value) ? preferredCycle : CYCLES.find(items => items.includes(seed.value as string))
  return cycle ? { kind: 'cycle', key: 'cycle:' + CYCLES.indexOf(cycle), value: cycle.indexOf(seed.value), cycle }
    : { kind: 'copy', key: 'other' }
}

/** Serializable lane patterns are shared by model fills and the legacy view adapter. */
export function createSheetFillSeriesPattern(seeds: readonly SheetFillSeriesSeed[]): SheetFillSeriesPattern {
  if (!seeds.length) throw new RangeError('填充源不能为空。')
  const texts = seeds.map(seed => seed.formula ? undefined : seed.value).filter((value): value is string => typeof value === 'string')
  const preferredCycle = texts.some(text => text === '正月' || text === '腊月') ? CYCLES[5] : undefined
  const classified = seeds.map(seed => classify(seed, preferredCycle)), groups: SheetFillSeriesGroup[] = []
  for (let start = 0; start < seeds.length;) {
    const first = classified[start]!, values: number[] = [], padding: number[] = []
    let end = start + 1
    while (end < seeds.length && classified[end]!.key === first.key) end++
    for (let index = start; index < end; index++) {
      let value = classified[index]!.value ?? 0
      if (first.cycle && values.length) {
        const previous = values.at(-1)!, length = first.cycle.length
        let difference = value - modulo(previous, length)
        if (difference < 0) difference += length
        value = previous + difference
      }
      values.push(value); padding.push(classified[index]!.padding ?? 0)
    }
    const length = end - start, group: SheetFillSeriesGroup = { start, length, kind: first.kind, intercept: values[0]!, step: 1,
      cycle: first.cycle, prefix: first.prefix, suffix: first.suffix, padding }
    // Fixed rules.ts also treats multi-seed 一/二 and 五/六 as weekday lanes.
    const hasWeek = seeds.slice(start, end).some(seed => seed.value === '日')
    const chineseWeek = length === 1 ? values[0]! > 0 && values[0]! < 7 || hasWeek
      : hasWeek || values.at(-1)! < 6 && values[0]! > 0 || values[0]! < 6 && values.at(-1)! > 0
    if (first.kind === 'chinese' && chineseWeek) {
      group.kind = 'cycle'; group.cycle = CYCLES[11]; group.intercept = values[0]! - 1
      for (let index = 1; index < values.length; index++) {
        if (values[index]! < modulo(values[index - 1]! - 1, 7) + 1) values[index] = values[index]! + Math.floor((values[index - 1]! - 1) / 7 + 1) * 7
      }
    }
    if (length > 1) {
      if (first.kind === 'number' || first.kind === 'date') {
        const ratio = values[0] === 0 ? NaN : values[1]! / values[0]!
        if (length > 2 && Number.isFinite(ratio) && ratio !== 0 && values.slice(1).every((value, index) => close(value, values[index]! * ratio))) group.ratio = ratio
        else {
          const meanX = (length - 1) / 2, meanY = values.reduce((sum, value) => sum + value, 0) / length
          let covariance = 0, variance = 0
          for (let index = 0; index < length; index++) { covariance += (index - meanX) * (values[index]! - meanY); variance += (index - meanX) ** 2 }
          group.step = covariance / variance; group.intercept = meanY - group.step * meanX
        }
      } else if (values.slice(1).every((value, index) => close(value - values[index]!, values[1]! - values[0]!))) group.step = values[1]! - values[0]!
      else group.kind = 'copy'
    }
    groups.push(group); start = end
  }
  const supportsSeries = groups.some(group => group.kind !== 'copy')
  const preferCopy = !supportsSeries || groups.every(group => group.kind === 'number' && group.length === 1)
  return { length: seeds.length, groups, supportsSeries, preferredApplyType: preferCopy ? 'COPY' : 'SERIES' }
}

export function sheetFillSeriesValueAt(pattern: SheetFillSeriesPattern, index: number): CellScalar | undefined {
  if (!Number.isSafeInteger(index)) throw new RangeError('填充序列位置无效。')
  const position = modulo(index, pattern.length), group = pattern.groups.find(item => position >= item.start && position < item.start + item.length)!
  if (group.kind === 'copy') return undefined
  const ordinal = Math.floor(index / pattern.length) * group.length + position - group.start
  const value = group.ratio === undefined ? group.intercept + group.step * ordinal : group.intercept * group.ratio ** ordinal
  if (!Number.isFinite(value)) return undefined
  if (group.kind === 'number' || group.kind === 'date') return group.ratio === undefined ? Math.round(value * 100000) / 100000 : value
  if (group.kind === 'cycle') return group.cycle![modulo(Math.round(value), group.cycle!.length)]
  if (group.kind === 'chinese') return formatChineseNumber(value)
  if (!Number.isSafeInteger(value)) return undefined
  const digits = String(Math.abs(value)).padStart(group.padding?.[position - group.start] ?? 0, '0')
  return `${group.prefix ?? ''}${value < 0 ? '-' : ''}${digits}${group.suffix ?? ''}`
}
