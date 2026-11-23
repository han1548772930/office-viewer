import type { HeaderFooterAction, HeaderFooterRequestScope, HeaderFooterRibbonKind } from './DocumentHeaderFooterRibbonControl.vue'

export interface HeaderFooterEditCapabilities {
  disabled: boolean
  contentDisabled?: boolean
  sectionDisabled?: boolean
}

export function headerFooterControlDisabled(kind: HeaderFooterRibbonKind, state: HeaderFooterEditCapabilities): boolean {
  if (kind === 'navigation') return false
  if (kind === 'position') return state.sectionDisabled ?? state.disabled
  return state.disabled || kind === 'insertion' && state.contentDisabled === true
}

/** Navigation reads a story; positions edit section geometry independently of story content. */
export function headerFooterActionDisabled(state: HeaderFooterEditCapabilities, action: HeaderFooterAction): boolean {
  if (action.action === 'navigate') return false
  if (action.action !== 'set-options') return state.disabled || state.contentDisabled === true
  const patch = action.patch
  const position = patch.marginHeader !== undefined || patch.marginFooter !== undefined
  const options = patch.linkToPrevious !== undefined || patch.differentFirstPage !== undefined || patch.differentOddEven !== undefined
  return position && (state.sectionDisabled ?? state.disabled) || options && state.disabled
}

/** A closed or invalidated popup cannot silently bind itself to a newer target. */
export function headerFooterRequestScopeMatches(state: HeaderFooterRequestScope, request: HeaderFooterRequestScope | undefined): boolean {
  return request !== undefined && state.requestScope === request.requestScope
}

export function parseHeaderFooterAction(raw: unknown): HeaderFooterAction | undefined {
  if (typeof raw !== 'string') return undefined
  let parsed: unknown
  try { parsed = JSON.parse(raw) } catch { return undefined }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return undefined
  const value = parsed as Record<string, unknown>
  if (value.requestScope !== undefined && (typeof value.requestScope !== 'number' || !Number.isSafeInteger(value.requestScope) || value.requestScope < 0)) return undefined
  const scope = value.requestScope === undefined ? {} : { requestScope: value.requestScope as number }
  if (value.action === 'navigate' && (value.area === 'header' || value.area === 'footer')) return { action: value.action, area: value.area, ...scope }
  if (value.action === 'insert-field' && (value.field === 'PAGE' || value.field === 'NUMPAGES')) return { action: value.action, field: value.field, ...scope }
  if (value.action === 'insert-date-time' && (value.format === 'short' || value.format === 'long' || value.format === 'time') && typeof value.automatic === 'boolean') {
    return { action: value.action, format: value.format, automatic: value.automatic, ...scope }
  }
  if (value.action === 'refresh-fields') return { action: value.action, ...scope }
  if (value.action !== 'set-options' || !value.patch || typeof value.patch !== 'object' || Array.isArray(value.patch)) return undefined
  const patch = value.patch as Record<string, unknown>
  for (const [key, setting] of Object.entries(patch)) {
    if (['linkToPrevious', 'differentFirstPage', 'differentOddEven'].includes(key)) {
      if (typeof setting !== 'boolean') return undefined
    } else if (key === 'marginHeader' || key === 'marginFooter') {
      if (typeof setting !== 'number' || !Number.isFinite(setting)) return undefined
    } else return undefined
  }
  return { action: value.action, patch: { ...patch }, ...scope } as HeaderFooterAction
}
