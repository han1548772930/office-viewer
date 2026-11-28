/** Matches v1.6.0 QuickSumShortcut; callers also require an inactive cell editor and a focused sheet Canvas. */
export function isSheetQuickSumShortcut(event: Pick<KeyboardEvent, 'key' | 'code' | 'altKey' | 'ctrlKey' | 'metaKey' | 'shiftKey' | 'isComposing' | 'keyCode' | 'defaultPrevented'>, mac = false): boolean {
  return !event.defaultPrevented && !event.isComposing && event.keyCode !== 229 && !event.shiftKey && event.altKey
    && (event.key === '=' || event.code === 'Equal')
    && (mac ? event.metaKey && !event.ctrlKey : !event.ctrlKey && !event.metaKey)
}
