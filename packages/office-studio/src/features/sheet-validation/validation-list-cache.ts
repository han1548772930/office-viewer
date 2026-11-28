import type { SheetDataValidationRule, WorkbookModel } from '../../domain/model.ts'
import { sheetValidationListOptions, type SheetValidationListOption } from './validation-list-input.ts'

interface CacheEntry { options: readonly SheetValidationListOption[]; bytes: number }
export interface SheetValidationListCacheLimits { maxRules?: number; maxBytes?: number; maxOptionsPerRule?: number }
export type SheetValidationListOptionsRead = (rule: SheetDataValidationRule) => readonly SheetValidationListOption[]

function immutableOptions(values: SheetValidationListOption[]): readonly SheetValidationListOption[] {
  for (const value of values) Object.freeze(value)
  return Object.freeze(values)
}

/** Inline lists have no formula dependencies and can survive unrelated content
 * revisions. Referenced/dynamic lists still use the sole formula reader on each
 * operation; caching them requires its future dependency/volatile epoch.
 * No workbook, rule, cell, permission object or previous snapshot is retained.
 */
export class SheetValidationListReadCache {
  private readonly maxRules: number
  private readonly maxBytes: number
  private readonly maxOptions: number
  private readonly entries = new Map<string, CacheEntry>()
  private bytes = 0
  private scope = ''
  private generation = 0

  constructor(limits: SheetValidationListCacheLimits = {}) {
    this.maxRules = limits.maxRules ?? 64
    this.maxBytes = limits.maxBytes ?? 8 * 1024 * 1024
    this.maxOptions = limits.maxOptionsPerRule ?? 65_536
    if (![this.maxRules, this.maxBytes, this.maxOptions].every(value => Number.isSafeInteger(value) && value > 0)) {
      throw new RangeError('列表候选缓存预算必须是正安全整数。')
    }
  }

  clear(): void { this.entries.clear(); this.bytes = 0; this.scope = ''; this.generation++ }
  get retained(): { rules: number; estimatedBytes: number } { return { rules: this.entries.size, estimatedBytes: this.bytes } }

  read(rule: SheetDataValidationRule, workbook: WorkbookModel, sheetId: string, userId?: string,
    region = 'zh-CN'): readonly SheetValidationListOption[] {
    const scope = JSON.stringify([workbook.id, sheetId, userId, region])
    if (scope !== this.scope) { this.clear(); this.scope = scope }
    const generation = this.generation
    const formula = rule.formula1?.trim()
    if ((rule.type !== 'list' && rule.type !== 'listMultiple') || !formula || formula.startsWith('=')) {
      return immutableOptions(sheetValidationListOptions(rule, workbook, sheetId, userId, region))
    }
    const signature = () => JSON.stringify([rule.id, rule.type, rule.formula1, rule.formula2])
    const key = signature()
    if (generation !== this.generation || scope !== this.scope) return immutableOptions(sheetValidationListOptions(rule, workbook, sheetId, userId, region))
    const cached = this.entries.get(key)
    if (cached) { this.entries.delete(key); this.entries.set(key, cached); return cached.options }
    const values = immutableOptions(sheetValidationListOptions(rule, workbook, sheetId, userId, region))
    // Conservative admission accounting for strings, objects and Map entries.
    // This bounds retention, not browser heap peaks or formula/source allocation.
    const bytes = key.length * 2 + values.reduce((sum, value) => sum + value.label.length * 2 + (value.color?.length ?? 0) * 2 + 96, 128)
    const latestKey = signature()
    if (generation !== this.generation || scope !== this.scope || key !== latestKey
      || values.length > this.maxOptions || bytes > this.maxBytes) return values
    while (this.entries.size >= this.maxRules || this.bytes + bytes > this.maxBytes) {
      const oldest = this.entries.entries().next().value
      if (!oldest) break
      this.entries.delete(oldest[0]); this.bytes -= oldest[1].bytes
    }
    this.entries.set(key, { options: values, bytes }); this.bytes += bytes
    return values
  }
}
