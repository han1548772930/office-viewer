import type { WorkbookModel } from '../../domain/model'

export type FormulaDefinedNameKey = 'definedNames' | 'localDefinedNames'
export interface FormulaDefinedName {
  name: string
  owner?: string
  defaultOwner: string
  key: FormulaDefinedNameKey
  record: Record<string, unknown>
  identity: string
  expression?: string
}
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)
export const formulaDefinedNameIdentity = (owner: string | undefined, name: string): string => `${owner ?? '*'}:${name.toLowerCase()}`

/** Native OOXML uses a zero-based localSheetId; Univer's public service uses
 * a sheet ID string and formulaOrRefString. Both resolve through sheet IDs. */
export function readFormulaDefinedNames(workbook: WorkbookModel, maximum = 8192): FormulaDefinedName[] {
  const result: FormulaDefinedName[] = []
  const sheetByName = new Map(workbook.sheets.map(sheet => [sheet.name.toLowerCase(), sheet.id]))
  const sheetIds = new Set(workbook.sheets.map(sheet => sheet.id))
  const read = (metadata: WorkbookModel['metadata'], defaultOwner: string, local: boolean): void => {
    for (const key of ['definedNames', 'localDefinedNames'] as const) {
      const values = metadata?.[key]
      if (!Array.isArray(values)) continue
      for (const value of values) {
        if (!object(value)) continue
        const name = typeof value.name === 'string' ? value.name : typeof value.id === 'string' ? value.id : undefined
        if (!name) continue
        if (result.length >= maximum) throw new RangeError('定义名称数量超出计算预算。')
        const declared = typeof value.sheetId === 'string' ? value.sheetId : typeof value.subUnitId === 'string' ? value.subUnitId
          : typeof value.sheet === 'string' ? sheetIds.has(value.sheet) ? value.sheet : sheetByName.get(value.sheet.toLowerCase())
            : typeof value.localSheetId === 'string' && value.localSheetId !== 'AllDefaultWorkbook' ? value.localSheetId
              : Number.isSafeInteger(value.localSheetId) ? workbook.sheets[Number(value.localSheetId)]?.id : undefined
        const hasScope = value.sheetId !== undefined || value.subUnitId !== undefined || value.sheet !== undefined
          || value.localSheetId !== undefined && value.localSheetId !== 'AllDefaultWorkbook'
        // Missing explicit scope cannot silently rebind to the active sheet.
        const explicitGlobal = value.localSheetId === 'AllDefaultWorkbook' && !hasScope
        const owner = declared ?? (hasScope ? `missing-scope:${String(value.localSheetId ?? value.sheet ?? value.sheetId ?? value.subUnitId)}`
          : !explicitGlobal && (local || key === 'localDefinedNames') ? defaultOwner : undefined)
        const expression = typeof value.refersTo === 'string' ? value.refersTo : typeof value.formulaOrRefString === 'string' ? value.formulaOrRefString : undefined
        result.push({ name, owner, defaultOwner: owner ?? defaultOwner, key, record: value, identity: formulaDefinedNameIdentity(owner, name), expression })
      }
    }
  }
  read(workbook.metadata, workbook.activeSheetId, false)
  for (const sheet of workbook.sheets) read(sheet.metadata, sheet.id, true)
  return result
}
export function indexFormulaDefinedNames(workbook: WorkbookModel): Map<string, FormulaDefinedName> {
  const result = new Map<string, FormulaDefinedName>()
  for (const name of readFormulaDefinedNames(workbook)) {
    const previous = result.get(name.identity)
    if (previous && previous.expression !== name.expression) throw new TypeError('同一作用域的定义名称存在不同表达式。')
    if (!previous) result.set(name.identity, name)
  }
  return result
}
export function resolveFormulaDefinedName(names: ReadonlyMap<string, FormulaDefinedName>, name: string, sheetId: string): FormulaDefinedName | undefined {
  return names.get(formulaDefinedNameIdentity(sheetId, name)) ?? names.get(formulaDefinedNameIdentity(undefined, name))
}

/** Update every formula-bearing name in its actual scope. The caller owns a
 * detached candidate; public and native expression fields stay synchronized. */
export function rewriteWorkbookFormulaDefinedNames(workbook: WorkbookModel,
  rewrite: (expression: string, ownerSheetName: string | undefined, name: FormulaDefinedName) => string): void {
  const sheets = new Map(workbook.sheets.map(sheet => [sheet.id, sheet.name]))
  for (const name of readFormulaDefinedNames(workbook)) {
    if (name.expression === undefined) continue
    const owner = sheets.get(name.defaultOwner)
    // An explicitly missing scope must not be treated as an unqualified local
    // reference to whichever sheet happens to be edited.
    if (name.owner && !owner) continue
    const expression = rewrite(name.expression, owner, name)
    if (name.record.refersTo !== undefined) name.record.refersTo = expression
    if (name.record.formulaOrRefString !== undefined) name.record.formulaOrRefString = expression
  }
}
