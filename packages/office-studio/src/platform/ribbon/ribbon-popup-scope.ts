import { provide, ref, type InjectionKey, type Ref } from 'vue'

interface RibbonPopupScope {
  activeId: Ref<string | undefined>
  direction: Readonly<Ref<'ltr' | 'rtl'>>
  close: () => void
}

export const RibbonPopupScopeKey: InjectionKey<RibbonPopupScope> = Symbol('RibbonPopupScope')

/** Matches Univer ToolbarDropdownProvider while allowing nested menus their own scope. */
export function provideRibbonPopupScope(direction: Readonly<Ref<'ltr' | 'rtl'>> = ref('ltr')): RibbonPopupScope {
  const activeId = ref<string>()
  const scope = { activeId, direction, close: () => { activeId.value = undefined } }
  provide(RibbonPopupScopeKey, scope)
  return scope
}
