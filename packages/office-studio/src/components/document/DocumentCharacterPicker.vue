<script setup lang="ts">
import { Alert, Button, ConfigProvider, Form, FormItem, Input, Listy, Popover, RadioButton, RadioGroup, Segmented, Select, SpaceCompact, Spin, Tooltip, theme } from 'antdv-next'
import type { ListyProps, ListyRef, SegmentedOptions } from 'antdv-next'
import { Activity, Apple, ChevronDown, Clock, Hash, Leaf, Lightbulb, MapPin, Search, Shuffle, Smile } from 'lucide-vue-next'
import { createPinia, getActivePinia } from 'pinia'
import { computed, h, inject, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue'
import { OfficeDirectionKey, OfficeLocaleKey, OfficeTranslatorKey } from '../../i18n/office-i18n'
import { DOCUMENT_EMOJI_CATEGORIES, DOCUMENT_EMOJI_SKIN_TONES, type DocumentCharacterItem, type DocumentCharacterLocale, type DocumentCharacterMode,
  type DocumentCharacterRow, type DocumentEmojiLocaleData, type DocumentEmojiSkinTone, type DocumentEmojiVariant } from '../../features/document/model/document-character-catalog'
import { activeDocumentCharacterSection, createDocumentCharacterRows, documentCharacterNeighbor, documentEmojiFamilyKey, documentEmojiSections,
  documentEmojiVariant, documentLocalizedEmojiTitle, documentSymbolSections, loadDocumentEmojiLocale, randomDocumentEmoji } from '../../features/document/services/document-special-characters'
import { useDocumentCharacterPreferencesStore, type DocumentCharacterPreferencesStore } from '../../features/document/stores/document-character-preferences-store'
import { documentCharacterUiLabels } from './document-character-picker-labels'

export interface DocumentCharacterInsertIntent {
  value: string
  keepOpen: boolean
  /** UI acknowledgment only; the workbench forwards just value to the public content command. */
  onResult?: (accepted: boolean) => void
}
const props = defineProps<{ mode: DocumentCharacterMode; disabled?: boolean; open?: boolean; preferences?: DocumentCharacterPreferencesStore }>()
const emit = defineEmits<{ insert: [params: DocumentCharacterInsertIntent, options?: { keepOpen?: boolean }] }>()
const inheritedPinia = getActivePinia()
const localPreferences = props.preferences ? undefined : useDocumentCharacterPreferencesStore(inheritedPinia ?? createPinia())
const ownsPreferences = !props.preferences && !inheritedPinia
const preferences = computed(() => props.preferences ?? localPreferences!)
const controlsDisabled = computed(() => props.disabled === true || props.open === false)
watch(preferences, store => store.ensureLoaded(), { immediate: true })
const locale = inject(OfficeLocaleKey, ref<DocumentCharacterLocale>('zh-CN'))
const direction = inject(OfficeDirectionKey, ref<'ltr' | 'rtl'>('ltr'))
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const labels = computed(() => documentCharacterUiLabels(locale.value))
const { token } = theme.useToken()
const pickerStyle = computed(() => ({ '--character-text': token.value.colorText, '--character-secondary-text': token.value.colorTextSecondary,
  '--character-surface': token.value.colorBgElevated, '--character-border': token.value.colorBorderSecondary,
  '--character-hover': token.value.controlItemBgHover, '--character-active': token.value.controlItemBgActive, '--character-primary': token.value.colorPrimary }))
const controlTheme = computed(() => ({ components: { Listy: { itemPaddingBlock: 0, itemPaddingInline: 0 }, Segmented: { itemSelectedBg: token.value.controlItemBgActive } } }))
const listStyles: ListyProps['styles'] = { item: { padding: 0, borderBottom: 'none' } }
const root = ref<HTMLElement>(), listRef = shallowRef<ListyRef | null>(null)
const query = ref(''), committedQuery = ref(''), searchComposing = ref(false), custom = ref(''), customComposing = ref(false)
const activeSection = ref(props.mode === 'emoji' ? 'recent' : 'mathematics'), focusKey = ref('')
const toneMenuOpen = ref(false), variantMenuKey = ref<string>(), localeData = shallowRef<DocumentEmojiLocaleData>()
const localeLoading = ref(false), localeError = ref(false)
let localeGeneration = 0, navigationGeneration = 0, disposed = false
let focusFrame: number | undefined, focusWindow: Window | undefined

function invalidateNavigation(): number {
  if (focusFrame !== undefined) focusWindow?.cancelAnimationFrame(focusFrame)
  focusFrame = undefined; focusWindow = undefined
  return ++navigationGeneration
}
function focusStillCurrent(owner: HTMLElement, generation: number, origin: Element | null): boolean {
  const active = owner.ownerDocument.activeElement
  return !disposed && root.value === owner && owner.isConnected && generation === navigationGeneration && !controlsDisabled.value
    && (active === origin || active === owner.ownerDocument.body)
}
async function restorePickerTrigger(attribute: 'toneTrigger' | 'variantsTrigger', key: string): Promise<void> {
  const owner = root.value
  if (!owner) return
  const origin = owner.ownerDocument.activeElement, generation = invalidateNavigation()
  await nextTick()
  if (!focusStillCurrent(owner, generation, origin) || toneMenuOpen.value || variantMenuKey.value) return
  const trigger = [...owner.querySelectorAll<HTMLButtonElement>('button')].find(button => button.dataset[attribute] === key)
  trigger?.focus({ preventScroll: true })
}
function closeToneMenu(restoreFocus = false): void {
  toneMenuOpen.value = false
  if (restoreFocus) void restorePickerTrigger('toneTrigger', 'tone')
}
function closeVariantMenu(key: string, restoreFocus = false): void {
  if (variantMenuKey.value !== key) return
  variantMenuKey.value = undefined
  if (restoreFocus) void restorePickerTrigger('variantsTrigger', key)
}

function text(key: string, fallback: string): string { return translate(`doc.character-picker.${key}`, fallback) }
async function loadLocale(): Promise<void> {
  const generation = ++localeGeneration, nextLocale = locale.value
  localeData.value = undefined; localeError.value = false
  if (props.mode !== 'emoji') { localeLoading.value = false; return }
  localeLoading.value = true
  try {
    const next = await loadDocumentEmojiLocale(nextLocale)
    if (!disposed && generation === localeGeneration && nextLocale === locale.value && props.mode === 'emoji') localeData.value = next
  } catch {
    if (!disposed && generation === localeGeneration) localeError.value = true
  } finally { if (!disposed && generation === localeGeneration) localeLoading.value = false }
}
watch([locale, () => props.mode], () => { invalidateNavigation(); toneMenuOpen.value = false; variantMenuKey.value = undefined; void loadLocale() }, { immediate: true })
const sections = computed(() => props.mode === 'emoji'
  ? documentEmojiSections(locale.value, preferences.value.recentEmojis, preferences.value.skinTone, committedQuery.value, localeData.value)
  : documentSymbolSections(labels.value.symbolGroups, labels.value.recents, preferences.value.recentSymbols, committedQuery.value))
const rowModel = computed(() => createDocumentCharacterRows(sections.value, `${props.mode}-${locale.value}`))
const characterRows = computed(() => rowModel.value.rows.filter((row): row is Extract<DocumentCharacterRow, { type: 'characters' }> => row.type === 'characters'))
const focusedKey = computed(() => characterRows.value.some(row => row.items.some(item => item.key === focusKey.value)) ? focusKey.value : characterRows.value[0]?.items[0]?.key)
// Options carry VNodes and must update with locale; snapshots contain only the user's values.
const categoryOptions = computed(() => {
  const icons = { recent: Clock, people: Smile, nature: Leaf, foods: Apple, activity: Activity, places: MapPin, objects: Lightbulb, symbols: Hash }
  const options: SegmentedOptions<string> = ['recent', ...DOCUMENT_EMOJI_CATEGORIES].map(value => {
    const title = value === 'recent' ? labels.value.recents : labels.value[value as (typeof DOCUMENT_EMOJI_CATEGORIES)[number]]
    return { value, label: h('span', { class: 'document-character-picker__sr-only' }, title), icon: h(icons[value as keyof typeof icons], { size: 16, 'aria-hidden': true }),
      tooltip: { title, trigger: ['hover', 'focus'] } }
  })
  return options
})
const symbolCategoryOptions = computed(() => sections.value.filter(section => section.items.length).map(section => ({ value: section.key, label: section.title })))
const skinToneOptions = computed(() => DOCUMENT_EMOJI_SKIN_TONES.map(value => {
  const item = documentEmojiVariant(`✋${value}`)
  return { ...item, value, title: documentLocalizedEmojiTitle(item, localeData.value) }
}))
const currentSkinTone = computed(() => skinToneOptions.value.find(item => item.value === preferences.value.skinTone) ?? skinToneOptions.value[0]!)

function commitSearch(): void { if (!searchComposing.value) committedQuery.value = query.value }
watch(query, commitSearch, { flush: 'sync' })
watch(committedQuery, () => {
  const generation = invalidateNavigation()
  focusKey.value = ''; variantMenuKey.value = undefined
  void nextTick(() => { if (!disposed && generation === navigationGeneration) listRef.value?.scrollTo(0) })
}, { flush: 'sync' })
watch(() => props.mode, () => {
  invalidateNavigation(); query.value = ''; committedQuery.value = ''; custom.value = ''; focusKey.value = ''
  searchComposing.value = false; customComposing.value = false; activeSection.value = props.mode === 'emoji' ? 'recent' : 'mathematics'
  toneMenuOpen.value = false; variantMenuKey.value = undefined
})
watch(controlsDisabled, disabled => { if (disabled) { toneMenuOpen.value = false; variantMenuKey.value = undefined; invalidateNavigation() } })
watch(() => rowModel.value.positions, positions => {
  if (!positions.some(position => position.key === activeSection.value) && !committedQuery.value.trim()) activeSection.value = positions[0]?.key ?? ''
})
function isActive(item: DocumentCharacterItem): boolean {
  return props.mode === 'emoji' ? !!preferences.value.activeEmoji && documentEmojiFamilyKey(preferences.value.activeEmoji) === item.familyKey
    : preferences.value.activeSymbol === item.value
}
function insert(value: string, keepOpen = true): void {
  if (controlsDisabled.value || !value.length || customComposing.value || searchComposing.value) return
  const store = preferences.value, mode = props.mode
  emit('insert', { value, keepOpen, onResult: accepted => { if (accepted) store.select(mode, value) } }, { keepOpen })
}
function insertCustom(event?: KeyboardEvent): void {
  if (event?.isComposing || event?.keyCode === 229 || customComposing.value) return
  if (event) { event.preventDefault(); event.stopPropagation() }
  insert(custom.value, false)
}
function insertVariant(value: DocumentEmojiVariant, key: string): void { if (!controlsDisabled.value) { closeVariantMenu(key, true); insert(value.emoji) } }
function insertRandom(): void { if (!controlsDisabled.value) insert(randomDocumentEmoji(Math.random, preferences.value.skinTone).emoji) }
function setSkinTone(value: DocumentEmojiSkinTone): void { if (!controlsDisabled.value) { preferences.value.setSkinTone(value); closeToneMenu(true); variantMenuKey.value = undefined } }
function openVariants(key: string, value: boolean): void { if (!value) closeVariantMenu(key); else if (!controlsDisabled.value) variantMenuKey.value = key }
function variantTitle(item: DocumentEmojiVariant): string { return documentLocalizedEmojiTitle(item, localeData.value) }
function variantActive(value: string): boolean { return preferences.value.activeEmoji === value }
function onScroll(event: Event): void {
  variantMenuKey.value = undefined
  if (committedQuery.value.trim()) return
  const target = event.currentTarget
  if (target && typeof target === 'object' && 'scrollTop' in target && typeof target.scrollTop === 'number') {
    activeSection.value = activeDocumentCharacterSection(rowModel.value.positions, target.scrollTop) ?? activeSection.value
  }
}
async function navigateSection(value: unknown): Promise<void> {
  if (controlsDisabled.value || typeof value !== 'string') return
  query.value = ''; committedQuery.value = ''; searchComposing.value = false; activeSection.value = value; variantMenuKey.value = undefined
  const generation = invalidateNavigation()
  await nextTick()
  if (disposed || generation !== navigationGeneration || controlsDisabled.value) return
  const position = rowModel.value.positions.find(item => item.key === value)
  if (position) listRef.value?.scrollTo({ key: position.rowKey, align: 'top' })
}
async function onCharacterKeydown(event: KeyboardEvent, item: DocumentCharacterItem): Promise<void> {
  if (controlsDisabled.value || event.isComposing || event.keyCode === 229) return
  const keyDirections = { ArrowLeft: 'left', ArrowRight: 'right', ArrowUp: 'up', ArrowDown: 'down', Home: 'first', End: 'last' } as const
  const movement = keyDirections[event.key as keyof typeof keyDirections]
  if (!movement) return
  const target = documentCharacterNeighbor(rowModel.value.rows, item.key, movement, direction.value === 'rtl')
  if (!target) return
  event.preventDefault(); event.stopPropagation(); focusKey.value = target.item.key
  const owner = root.value
  if (!owner) return
  const origin = owner.ownerDocument.activeElement, generation = invalidateNavigation()
  listRef.value?.scrollTo({ key: target.rowKey, align: 'auto' })
  await focusCharacter(owner, target.item.key, generation, origin)
}
async function focusCharacter(owner: HTMLElement, key: string, generation: number, origin: Element | null, attempts = 2): Promise<void> {
  await nextTick()
  if (!focusStillCurrent(owner, generation, origin) || focusedKey.value !== key) return
  const button = [...owner.querySelectorAll<HTMLButtonElement>('button[data-character-key]')].find(element => element.dataset.characterKey === key)
  if (button) { button.focus({ preventScroll: true }); return }
  // Listy measures and mounts virtual rows after scrolling. Retry for at most
  // two frames; a new query, navigation, locale, focus or owner cancels it.
  const view = owner.ownerDocument.defaultView
  if (!attempts || !view) return
  focusWindow = view
  focusFrame = view.requestAnimationFrame(() => {
    focusFrame = undefined; focusWindow = undefined
    void focusCharacter(owner, key, generation, origin, attempts - 1)
  })
}
onBeforeUnmount(() => { disposed = true; localeGeneration++; invalidateNavigation(); if (ownsPreferences) { localPreferences?.releaseStorage(); localPreferences?.$dispose() } })
</script>

<template>
  <ConfigProvider :theme="controlTheme">
    <section ref="root" class="document-character-picker" :style="pickerStyle" :dir="direction" :data-character-mode="mode" :aria-label="labels.list[mode]">
      <Form class="document-character-picker__search" :disabled="controlsDisabled" @submit.prevent>
        <FormItem no-style>
          <Input v-model:value="query" :aria-label="labels.search" :placeholder="labels.search" allow-clear :disabled="controlsDisabled"
            @compositionstart="searchComposing = true" @compositionend="searchComposing = false; commitSearch()">
            <template #prefix><Search :size="14" aria-hidden="true" /></template>
            <template #suffix><Spin v-if="localeLoading" size="small" /></template>
          </Input>
        </FormItem>
        <template v-if="mode === 'emoji'">
          <Tooltip :title="labels.random" :trigger="['hover', 'focus']"><Button type="text" class="document-character-picker__utility" :disabled="controlsDisabled" :aria-label="labels.random" @click="insertRandom"><Shuffle :size="16" /></Button></Tooltip>
          <Popover :open="toneMenuOpen && !controlsDisabled" trigger="click" placement="bottomRight" @open-change="value => { if (!controlsDisabled || !value) toneMenuOpen = value }">
            <Tooltip :title="text('skin-tone', labels.skinTone)" :trigger="['hover', 'focus']"><Button type="text" class="document-character-picker__utility" :disabled="controlsDisabled" :aria-label="text('skin-tone', labels.skinTone)" data-tone-trigger="tone">{{ currentSkinTone.emoji }}<ChevronDown :size="10" /></Button></Tooltip>
            <template #content>
              <Form class="document-character-picker__tone-menu" :disabled="controlsDisabled" :style="pickerStyle" @submit.prevent @keydown.esc.stop.prevent="closeToneMenu(true)">
                <FormItem :label="text('skin-tone', labels.skinTone)">
                  <RadioGroup :value="preferences.skinTone" :disabled="controlsDisabled" size="small" @update:value="setSkinTone">
                    <Tooltip v-for="tone in skinToneOptions" :key="tone.value" :title="tone.title" :trigger="['hover', 'focus']"><RadioButton :value="tone.value" :disabled="controlsDisabled"><span>{{ tone.emoji }}</span><span class="document-character-picker__sr-only">{{ tone.title }}</span></RadioButton></Tooltip>
                  </RadioGroup>
                </FormItem>
              </Form>
            </template>
          </Popover>
        </template>
      </Form>
      <Alert v-if="localeError" type="error" show-icon :title="text('load-error', labels.loadError)"><template #action><Button size="small" :disabled="controlsDisabled" @click="loadLocale">{{ text('retry', labels.retry) }}</Button></template></Alert>
      <Listy v-if="rowModel.rows.length" ref="listRef" class="document-character-picker__list" :items="rowModel.rows" row-key="key" :height="200" virtual :row-hoverable="false" :styles="listStyles" :aria-label="labels.list[mode]" @scroll="onScroll">
        <template #itemRender="row: DocumentCharacterRow">
          <div v-if="row.type === 'header'" class="document-character-picker__heading" :data-character-section="row.sectionKey">{{ row.title }}</div>
          <div v-else class="document-character-picker__row">
            <div v-for="item in row.items" :key="item.key" class="document-character-picker__cell">
              <Tooltip :title="item.title" :trigger="['hover', 'focus']">
                <Button type="text" :disabled="controlsDisabled" :aria-label="item.title" :aria-pressed="isActive(item)" :tabindex="focusedKey === item.key ? 0 : -1"
                  :data-character-key="item.key" :data-character-value="item.value" :data-emoji-family="item.familyKey" :class="{ 'is-active': isActive(item) }"
                  @focus="focusKey = item.key" @click="insert(item.value)" @keydown="onCharacterKeydown($event, item)">{{ item.value }}</Button>
              </Tooltip>
              <Popover v-if="item.variants?.length" :open="variantMenuKey === item.key && !controlsDisabled" trigger="click" placement="bottomRight" @open-change="value => openVariants(item.key, value)">
                <Tooltip :title="`${item.title}, ${labels.more}`" :trigger="['hover', 'focus']"><Button type="text" class="document-character-picker__variants-trigger" :disabled="controlsDisabled" :aria-label="`${item.title}, ${labels.more}`" :data-variants-trigger="item.key"><ChevronDown :size="10" /></Button></Tooltip>
                <template #content>
                  <div class="document-character-picker__variants" :style="pickerStyle" :aria-label="item.title" @keydown.esc.stop.prevent="closeVariantMenu(item.key, true)">
                    <Tooltip v-for="variant in item.variants" :key="variant.emoji" :title="variantTitle(variant)" :trigger="['hover', 'focus']">
                      <Button type="text" :disabled="controlsDisabled" :aria-label="variantTitle(variant)" :aria-pressed="variantActive(variant.emoji)" :data-character-variant="variant.emoji" :class="{ 'is-active': variantActive(variant.emoji) }" @click="insertVariant(variant, item.key)">{{ variant.emoji }}</Button>
                    </Tooltip>
                  </div>
                </template>
              </Popover>
            </div>
          </div>
        </template>
      </Listy>
      <div v-else class="document-character-picker__empty" role="status">{{ mode === 'emoji' ? labels.noResults : text('no-symbols', labels.noSymbols) }}</div>
      <Form class="document-character-picker__categories" :disabled="controlsDisabled" @submit.prevent>
        <FormItem no-style>
          <Segmented v-if="mode === 'emoji'" :value="committedQuery.trim() ? undefined : activeSection" :options="categoryOptions" block :disabled="controlsDisabled" :aria-label="text('category', labels.category)" @change="navigateSection" />
          <Select v-else :value="activeSection" :options="symbolCategoryOptions" :disabled="controlsDisabled" :aria-label="text('category', labels.category)" @change="navigateSection" />
        </FormItem>
      </Form>
      <Form class="document-character-picker__custom" layout="vertical" :disabled="controlsDisabled" @submit.prevent>
        <FormItem :label="text('custom-character', labels.custom)">
          <SpaceCompact block>
            <Input v-model:value="custom" :disabled="controlsDisabled" :aria-label="text('custom-character', labels.custom)" :placeholder="text('custom-placeholder', labels.customPlaceholder)"
              @compositionstart="customComposing = true" @compositionend="customComposing = false" @press-enter="insertCustom" />
            <Button type="primary" :disabled="controlsDisabled || !custom.length || customComposing || searchComposing" @click="insertCustom()">{{ text('insert', labels.insert) }}</Button>
          </SpaceCompact>
        </FormItem>
      </Form>
    </section>
  </ConfigProvider>
</template>

<style scoped>
.document-character-picker { display: flex; flex-direction: column; gap: 6px; width: 420px; max-width: calc(100vw - 40px); color: var(--character-text); }
.document-character-picker__search { display: flex; align-items: center; gap: 4px; margin: 0; min-width: 0; }
.document-character-picker__search > :deep(.ant-input-affix-wrapper) { min-width: 0; flex: 1; }
.document-character-picker__utility.ant-btn { display: inline-flex; align-items: center; justify-content: center; flex: none; height: 32px; min-width: 32px; padding-inline: 4px; }
.document-character-picker__list { min-width: 0; height: 200px; overflow: hidden; }
.document-character-picker__heading { display: flex; align-items: center; height: 24px; padding-inline: 4px; color: var(--character-secondary-text); font-size: 12px; line-height: 24px; }
.document-character-picker__row { display: grid; grid-template-columns: repeat(10, minmax(0, 1fr)); align-items: center; gap: 4px; height: 32px; }
.document-character-picker__cell { position: relative; display: flex; align-items: center; justify-content: center; min-width: 0; height: 32px; }
.document-character-picker__cell :deep(.ant-btn), .document-character-picker__variants :deep(.ant-btn) { display: inline-flex; align-items: center; justify-content: center; width: 32px; max-width: 100%; height: 32px; padding: 0; border: 0; border-radius: 6px; font: 18px/1 'Segoe UI Emoji', 'Segoe UI Symbol', 'Microsoft YaHei', sans-serif; color: var(--character-text); }
.document-character-picker__cell :deep(.ant-btn:not(:disabled):hover), .document-character-picker__variants :deep(.ant-btn:not(:disabled):hover) { background: var(--character-hover); }
.document-character-picker__cell :deep(.ant-btn.is-active), .document-character-picker__variants :deep(.ant-btn.is-active) { background: var(--character-active); }
.document-character-picker__cell :deep(.ant-btn:not(:disabled):focus-visible), .document-character-picker__variants :deep(.ant-btn:not(:disabled):focus-visible) { outline: 1px solid var(--character-primary); outline-offset: -1px; }
.document-character-picker__cell :deep(.document-character-picker__variants-trigger.ant-btn) { position: absolute; inset-inline-end: 0; bottom: 0; width: 12px; min-width: 12px; height: 12px; padding: 0; border-radius: 3px; color: var(--character-secondary-text); background: var(--character-surface); }
.document-character-picker__variants { display: grid; grid-template-columns: repeat(6, 32px); gap: 4px; max-height: 272px; overflow-y: auto; }
.document-character-picker__tone-menu { margin: 0; }
.document-character-picker__tone-menu :deep(.ant-form-item) { margin-bottom: 0; }
.document-character-picker__tone-menu :deep(.ant-radio-button-wrapper) { padding-inline: 8px; font-family: 'Segoe UI Emoji', sans-serif; }
.document-character-picker__categories { padding-top: 4px; border-top: 1px solid var(--character-border); margin: 0; }
.document-character-picker__categories :deep(.ant-select) { width: 100%; }
.document-character-picker__categories :deep(.ant-segmented-item-label) { display: flex; align-items: center; justify-content: center; padding-inline: 6px; min-width: 28px; }
.document-character-picker__categories :deep(.ant-segmented-item-icon) { display: flex; align-items: center; margin-inline-end: 0; }
.document-character-picker__custom { margin: 0; }
.document-character-picker__custom :deep(.ant-form-item) { margin-bottom: 0; }
.document-character-picker__custom :deep(.ant-form-item-label) { padding-bottom: 3px; }
.document-character-picker__custom :deep(.ant-form-item-label > label) { font-size: 12px; }
.document-character-picker__custom :deep(.ant-input) { min-width: 0; }
.document-character-picker__empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--character-secondary-text); font-size: 14px; }
.document-character-picker__sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; border: 0; }
</style>
