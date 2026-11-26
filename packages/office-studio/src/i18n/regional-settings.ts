import { computed, ref } from 'vue'

export const supportedOfficeLocales = [
  'ar-SA', 'ca-ES', 'de-DE', 'en-US', 'es-ES', 'fa-IR', 'fr-FR', 'id-ID', 'it-IT', 'ja-JP',
  'ko-KR', 'pl-PL', 'pt-BR', 'ru-RU', 'sk-SK', 'vi-VN', 'zh-CN', 'zh-HK', 'zh-TW',
] as const
export type OfficeLocale = (typeof supportedOfficeLocales)[number]
export type OfficeDirection = 'ltr' | 'rtl'

export function officeLocaleDirection(locale: OfficeLocale): OfficeDirection {
  return locale === 'ar-SA' || locale === 'fa-IR' ? 'rtl' : 'ltr'
}

function localeTag(value: string): OfficeLocale {
  const tag = supportedOfficeLocales.find(tag => tag === value || tag.replace('-', '') === value)
  if (!tag) throw new RangeError(`Unsupported office locale: ${value}`)
  return tag
}

/** Instance-scoped, independent of document snapshots and browser storage. */
export class OfficeRegionalSettings {
  private readonly language = ref<OfficeLocale>('zh-CN')
  private readonly explicitRegion = ref<OfficeLocale>()
  private readonly explicitDirection = ref<OfficeDirection>()
  readonly locale = computed(() => this.language.value)
  readonly region = computed(() => this.explicitRegion.value ?? this.language.value)
  readonly regionPreference = computed(() => this.explicitRegion.value ?? 'auto')
  readonly directionPreference = computed(() => this.explicitDirection.value)
  readonly direction = computed(() => this.explicitDirection.value ?? officeLocaleDirection(this.language.value))

  constructor(locale?: string, region?: string) {
    if (locale !== undefined) this.setLocale(locale)
    if (region !== undefined) this.setRegion(region)
  }

  setLocale(value: string) { this.language.value = localeTag(value) }
  setRegion(value: string) { this.explicitRegion.value = localeTag(value) }
  followLanguage() { this.explicitRegion.value = undefined }
  setDirection(value: OfficeDirection | undefined) { this.explicitDirection.value = value }
}
