import { bodyRecord, type UniverBodyRecord } from './document-body-style'

/** Public numbering values are typed. Only an omitted value receives a default. */
export function documentBodyListInteger(value: unknown, fallback: number, minimum: number, maximum: number, label: string): number {
  if (value === undefined) return fallback
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < minimum || value > maximum) {
    throw new RangeError(`${label} 必须为 ${minimum} 到 ${maximum} 的安全整数。`)
  }
  return value
}

export function documentBodyListRecord(value: unknown, label: string): UniverBodyRecord {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype) {
    throw new TypeError(`${label} 必须为普通对象。`)
  }
  return value as UniverBodyRecord
}

export function documentBodyListLevels(value: unknown): UniverBodyRecord[] {
  if (value === undefined) return []
  if (!Array.isArray(value) || value.length > 9) throw new RangeError('列表层级必须为至多九项的对象数组。')
  const levels: UniverBodyRecord[] = []
  for (let ilvl = 0; ilvl < value.length; ilvl++) {
    if (!Object.hasOwn(value, ilvl) || value[ilvl] === undefined) throw new TypeError('公开列表层级不能包含数组空洞。')
    const level = documentBodyListRecord(value[ilvl], '列表层级')
    if (level.glyphType !== undefined && (typeof level.glyphType !== 'string' || !level.glyphType.trim())) {
      documentBodyListInteger(level.glyphType, 2, 0, 60, '列表编号格式')
    }
    for (const key of ['glyphSymbol', 'glyphFormat'] as const) {
      if (level[key] !== undefined && typeof level[key] !== 'string') throw new TypeError(`${key} 必须为文字。`)
    }
    documentBodyListInteger(level.startNumber, 0, -1, Number.MAX_SAFE_INTEGER - 1, '列表起始编号')
    documentBodyListInteger(level.followNumberWith, 0, 0, 2, '编号后的分隔符')
    documentBodyListInteger(level.bulletAlignment, 0, 0, 2, '编号对齐方式')
    if (level.isLegal !== undefined && typeof level.isLegal !== 'boolean' && level.isLegal !== 0 && level.isLegal !== 1) {
      throw new TypeError('列表十进制占位符标记必须为布尔值。')
    }
    documentBodyListRecord(level.textStyle, '编号文字样式')
    documentBodyListRecord(level.paragraphProperties, '编号段落属性')
    levels.push(level)
  }
  return levels
}

/** A native sparse ilvl array needs dense, usable public resources. The marked
 * default preserves absence on reopen until the caller edits or uses that level. */
export function documentBodyMissingListLevel(ilvl: number): UniverBodyRecord {
  return { glyphType: 2, glyphFormat: `%${ilvl + 1}.`, startNumber: 0, followNumberWith: 0, bulletAlignment: 0,
    officeStudio: { nativeMissingLevel: { version: 1, ilvl } } }
}

export function isDocumentBodyMissingListLevel(level: UniverBodyRecord, ilvl: number): boolean {
  const marker = bodyRecord(bodyRecord(level.officeStudio).nativeMissingLevel)
  return marker.version === 1 && marker.ilvl === ilvl && sameValue(level, documentBodyMissingListLevel(ilvl))
}

function sameValue(left: unknown, right: unknown): boolean {
  if (left === right) return true
  if (Array.isArray(left) || Array.isArray(right)) return Array.isArray(left) && Array.isArray(right)
    && left.length === right.length && left.every((value, index) => sameValue(value, right[index]))
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false
  const a = bodyRecord(left), b = bodyRecord(right)
  const keys = Object.keys(a).filter(key => a[key] !== undefined), other = Object.keys(b).filter(key => b[key] !== undefined)
  return keys.length === other.length && keys.every(key => Object.hasOwn(b, key) && sameValue(a[key], b[key]))
}
