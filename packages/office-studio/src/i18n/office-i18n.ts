import type { InjectionKey, Ref } from 'vue'
import { createI18n } from 'vue-i18n'
import type { RibbonTab } from '../platform/ribbon/ribbon-schema.ts'
import { coreEnUS } from './core/locales/en-US.ts'
import { coreZhCN } from './core/locales/zh-CN.ts'
import { docsEnUS } from './docs/locales/en-US.ts'
import { docsZhCN } from './docs/locales/zh-CN.ts'
import { sheetsEnUS } from './sheets/locales/en-US.ts'
import { sheetsZhCN } from './sheets/locales/zh-CN.ts'

export type OfficeLocale = 'zh-CN' | 'en-US'
export type OfficeTranslator = (key: string, fallback: string) => string

export const OfficeLocaleKey: InjectionKey<Readonly<Ref<OfficeLocale>>> = Symbol('OfficeLocale')
export const OfficeTranslatorKey: InjectionKey<OfficeTranslator> = Symbol('OfficeTranslator')

const appLabels: Record<OfficeLocale, Record<string, string>> = {
  'zh-CN': { 'app.language': '界面语言', 'app.chinese': '简体中文', 'app.english': 'English' },
  'en-US': { 'app.language': 'Language', 'app.chinese': '简体中文', 'app.english': 'English' },
}

export function createOfficeMessages(): Record<OfficeLocale, Record<string, string>> {
  return {
    'zh-CN': { ...appLabels['zh-CN'], ...coreZhCN, ...sheetsZhCN, ...docsZhCN },
    'en-US': { ...appLabels['en-US'], ...coreEnUS, ...sheetsEnUS, ...docsEnUS },
  }
}

export function createOfficeI18n() {
  const stored = typeof localStorage === 'undefined' ? null : localStorage.getItem('office-studio.locale')
  const locale: OfficeLocale = stored === 'en-US' || stored === 'zh-CN' ? stored : 'zh-CN'
  return createI18n({
    legacy: false,
    locale,
    fallbackLocale: 'zh-CN',
    flatJson: true,
    messages: createOfficeMessages(),
  })
}

export function createOfficeTranslator(locale: OfficeLocale): OfficeTranslator {
  const messages = createOfficeMessages()[locale]
  return (key, fallback) => {
    const normalized = key.replace(/^(sheet|doc|core)\./, '')
    return messages[normalized] ?? fallback
  }
}

export function createRibbonTranslator(locale: OfficeLocale, namespace: 'sheet' | 'doc'): OfficeTranslator {
  const translate = createOfficeTranslator(locale)
  return (key, fallback) => translate(`${namespace}.${key}`, fallback)
}

export function localizeRibbon(tabs: RibbonTab[], namespace: 'sheet' | 'doc', translate: OfficeTranslator): RibbonTab[] {
  const localized = (key: string, fallback: string) => translate(`${namespace}.${key}`, fallback)
  return tabs.map(tab => ({
    ...tab,
    title: localized(`tabs.${tab.id}`, tab.title),
    groups: tab.groups.map(group => ({
      ...group,
      title: localized(`groups.${group.id}`, group.title),
      items: group.items.map(localizeItem),
    })),
  }))

  function localizeItem(item: RibbonTab['groups'][number]['items'][number]): RibbonTab['groups'][number]['items'][number] {
    const children = item.children?.map(localizeItem)
    const options = item.options?.map(option => ({
      ...option,
      label: localized(`options.${namespace}.${item.id}.${String(option.value)}`, option.label),
    }))
    return {
      ...item,
      title: localized(`items.${item.id}`, item.title),
      ...(children ? { children } : {}),
      ...(options ? { options } : {}),
    }
  }
}
