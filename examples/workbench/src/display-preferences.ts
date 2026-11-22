import { computed, ref, watch } from 'vue'
import { isRibbonVariant, supportedOfficeLocales, WORKBENCH_ACCENTS, WORKBENCH_CHROME, WORKBENCH_ZOOMS, type WorkbenchAccent } from '@office-studio/office-studio'

const prefix = 'office-studio.'
export function readDisplayPreferences(storage: Pick<Storage, 'getItem'>) {
  const get = (key: string) => storage.getItem(prefix + key)
  const variant = get('ribbon-variant'), theme = get('theme')
  return {
    ribbon: isRibbonVariant(variant) ? variant : 'grid' as const,
    direction: get('direction') === 'rtl' ? 'rtl' as const : 'ltr' as const,
    dark: get('appearance') === 'dark',
    accent: theme && Object.hasOwn(WORKBENCH_ACCENTS, theme) ? theme as WorkbenchAccent : 'blue' as const,
    region: supportedOfficeLocales.find(value => value === get('region')) ?? 'auto' as const,
    chrome: WORKBENCH_CHROME.find(value => value === get('chrome')) ?? 'full',
    zoom: WORKBENCH_ZOOMS.find(value => value === Number(get('zoom'))) ?? 100,
  }
}

export function useDisplayPreferences(storage: Storage) {
  const p = ref(readDisplayPreferences(storage))
  watch(p, value => {
    for (const [key, setting] of Object.entries({ 'ribbon-variant': value.ribbon, direction: value.direction, appearance: value.dark ? 'dark' : 'light', theme: value.accent, region: value.region, chrome: value.chrome, zoom: value.zoom })) storage.setItem(prefix + key, String(setting))
  }, { deep: true, flush: 'sync' })
  function reset() {
    p.value = { ribbon: 'grid', direction: 'ltr', dark: false, accent: 'blue', region: 'auto', chrome: 'full', zoom: 100 }
    for (const key of ['ribbon-variant', 'direction', 'appearance', 'theme', 'region', 'chrome', 'zoom', 'locale']) storage.removeItem(prefix + key)
  }
  return { p, reset, display: computed(() => ({ chrome: p.value.chrome, zoom: p.value.zoom })) }
}
