<script setup lang="ts">
import { Button, Checkbox, ColorPicker, Form, FormItem, Input, Modal, Select } from 'antdv-next'
import { computed, inject, reactive, ref, watch } from 'vue'
import type { DocumentContentControl, DocumentContentControlType } from '../../domain/model'
import type { FDocument } from '../../api/facade'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

type ContentControlProperties = Partial<DocumentContentControl> & {
  date?: { format?: string; locale?: string; calendar?: string; storage?: string }
}

const props = defineProps<{
  open: boolean
  selectedControlId?: string
  documentApi?: FDocument
}>()
const emit = defineEmits<{
  'update:open': [value: boolean]
  error: [message: string]
}>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)

const properties = reactive<ContentControlProperties>({})
const listText = ref('')
const selected = computed(() => {
  if (!props.open || !props.selectedControlId || !props.documentApi) return undefined
  return props.documentApi.getContentControl(props.selectedControlId)?.getInfo()
})
const contentControlTypes = computed<Array<{ label: string; value: DocumentContentControlType }>>(() => [
  { label: t('doc.content-control.type.richText', '富文本'), value: 'richText' },
  { label: t('doc.content-control.type.text', '纯文本'), value: 'text' },
  { label: t('doc.content-control.type.checkbox', '复选框'), value: 'checkbox' },
  { label: t('doc.content-control.type.date', '日期'), value: 'date' },
  { label: t('doc.content-control.type.dropDownList', '下拉列表'), value: 'dropDownList' },
  { label: t('doc.content-control.type.comboBox', '组合框'), value: 'comboBox' },
  { label: t('doc.content-control.type.picture', '图片'), value: 'picture' },
  { label: t('doc.content-control.type.group', '分组'), value: 'group' },
  { label: t('doc.content-control.type.repeatingSection', '重复节'), value: 'repeatingSection' },
  { label: t('doc.content-control.type.repeatingSectionItem', '重复节项'), value: 'repeatingSectionItem' },
  { label: t('doc.content-control.type.docPartObj', '文档部件'), value: 'docPartObj' },
])
const appearances = computed(() => [
  { label: t('doc.content-control.appearance.boundingBox', '边框'), value: 'boundingBox' },
  { label: t('doc.content-control.appearance.tags', '标签'), value: 'tags' },
  { label: t('doc.content-control.appearance.hidden', '隐藏'), value: 'hidden' },
])
const locks = computed(() => [
  { label: t('doc.content-control.lock.unlocked', '不锁定'), value: 'unlocked' },
  { label: t('doc.content-control.lock.sdtLocked', '锁定控件'), value: 'sdtLocked' },
  { label: t('doc.content-control.lock.contentLocked', '锁定内容'), value: 'contentLocked' },
  { label: t('doc.content-control.lock.sdtContentLocked', '内容和控件都锁定'), value: 'sdtContentLocked' },
])
const dateStorageOptions = computed(() => [
  { label: t('doc.content-control.date-storage.text', '文本'), value: 'text' },
  { label: t('doc.content-control.date-storage.date', '日期'), value: 'date' },
  { label: t('doc.content-control.date-storage.dateTime', '日期和时间'), value: 'dateTime' },
])
const hasPlaceholder = computed(() => selected.value != null && ![
  'checkbox', 'repeatingSection', 'repeatingSectionItem',
].includes(selected.value.type))
const hasDataBinding = computed(() => selected.value?.dataBinding != null)
const isDate = computed(() => selected.value?.type === 'date')
const isList = computed(() => selected.value?.type === 'dropDownList' || selected.value?.type === 'comboBox')
const isRepeatingSection = computed(() => selected.value?.type === 'repeatingSection')
const isDocPart = computed(() => selected.value?.type === 'docPartObj')
const isText = computed(() => selected.value?.type === 'text' || selected.value?.type === 'richText')
const isPlainText = computed(() => selected.value?.type === 'text')

watch(selected, (control) => {
  for (const key of Object.keys(properties) as Array<keyof ContentControlProperties>) delete properties[key]
  if (control) {
    Object.assign(properties, control)
    properties.date = {
      format: control.dateFormat,
      locale: control.dateLocale,
      calendar: control.dateCalendar,
      storage: control.dateStorage ?? 'text',
    }
    listText.value = (control.items ?? [])
      .map((item) => `${item.displayText ?? item.value}|${item.value}`)
      .join('\n')
  } else {
    listText.value = ''
  }
}, { immediate: true })

function patch(next: Partial<DocumentContentControl>): void {
  Object.assign(properties, next)
}

function patchDate(patch: NonNullable<ContentControlProperties['date']>): void {
  properties.date = { ...properties.date, ...patch }
}

function checkboxValue(value: boolean | { target?: { checked?: boolean } }): boolean {
  return typeof value === 'boolean' ? value : Boolean(value.target?.checked)
}

function updateListItems(value: string): void {
  listText.value = value
  patch({ items: value.split(/\r?\n/).filter(Boolean).map((line) => {
    const separator = line.indexOf('|')
    return separator < 0
      ? { displayText: line, value: line }
      : { displayText: line.slice(0, separator), value: line.slice(separator + 1) }
  }) })
}

function close(): void {
  emit('update:open', false)
}

function save(): void {
  const control = selected.value
  if (!control || !props.documentApi) return
  try {
    const date = properties.date
    props.documentApi.getContentControl(control.id)?.update({
      alias: properties.alias,
      tag: properties.tag,
      appearance: properties.appearance,
      lock: properties.lock,
      color: properties.color,
      placeholder: properties.placeholder,
      temporary: properties.temporary,
      multiLine: properties.multiLine,
      dateFormat: date?.format,
      dateLocale: date?.locale,
      dateCalendar: date?.calendar,
      dateStorage: date?.storage,
      items: properties.items,
      repeatingSection: properties.repeatingSection,
      docPart: properties.docPart,
      dataBinding: properties.dataBinding,
    })
    close()
  } catch (error) {
    emit('error', error instanceof Error ? error.message : String(error))
  }
}
</script>

<template>
  <Modal
    :open="open && selected != null"
    :title="t('doc.content-control.properties', '内容控件属性')"
    :width="640"
    :footer="null"
    destroy-on-hidden
    @cancel="close"
  >
    <Form v-if="selected" layout="vertical" class="content-control-properties" @submit.prevent="save">
      <section class="properties-grid">
        <FormItem :label="t('doc.content-control.type', '类型')">
          <Select :value="selected.type" :options="contentControlTypes" disabled />
        </FormItem>
        <FormItem :label="t('doc.content-control.ooxml-id', 'OOXML 标识')">
          <Input :value="String(selected.metadata?.ooxmlId ?? '')" disabled />
        </FormItem>
        <FormItem :label="t('doc.content-control.label-title', '标题')">
          <Input :value="properties.alias ?? selected.title ?? ''" @update:value="value => patch({ alias: value })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.tag', '标记')">
          <Input :value="properties.tag ?? ''" @update:value="value => patch({ tag: value })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.appearance', '外观')">
          <Select :value="properties.appearance ?? 'boundingBox'" :options="appearances" @change="value => patch({ appearance: String(value) })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.lock', '锁定')">
          <Select :value="properties.lock ?? 'unlocked'" :options="locks" @change="value => patch({ lock: String(value) as DocumentContentControl['lock'] })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.color', '颜色')">
          <div class="color-input">
            <ColorPicker
              :value="`#${(properties.color ?? '4472C4').replace(/^#/, '')}`"
              value-format="hex"
              :aria-label="t('doc.content-control.color', '颜色')"
              @update:value="value => patch({ color: value.replace(/^#/, '').toUpperCase() })"
            />
            <Input :value="properties.color ?? '4472C4'" @update:value="value => patch({ color: value.replace(/^#/, '').toUpperCase() })" />
          </div>
        </FormItem>
        <FormItem v-if="hasPlaceholder" :label="t('doc.content-control.placeholder', '占位文本')">
          <Input :value="properties.placeholder ?? ''" :disabled="selected.lock === 'contentLocked' || selected.lock === 'sdtContentLocked'" @update:value="value => patch({ placeholder: value })" />
        </FormItem>
      </section>

      <section v-if="isText" class="properties-section">
        <Checkbox v-if="isPlainText" :checked="properties.multiLine ?? false" @change="value => patch({ multiLine: checkboxValue(value) })">
          {{ t('doc.content-control.multiline', '允许多行') }}
        </Checkbox>
        <Checkbox :checked="properties.temporary ?? false" @change="value => patch({ temporary: checkboxValue(value) })">
          {{ t('doc.content-control.temporary', '临时控件') }}
        </Checkbox>
      </section>

      <section v-if="isDate" class="properties-grid properties-section">
        <FormItem :label="t('doc.content-control.date-format', '日期格式')">
          <Input :value="properties.date?.format ?? ''" @update:value="value => patchDate({ format: value })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.date-locale', '区域设置')">
          <Input :value="properties.date?.locale ?? ''" @update:value="value => patchDate({ locale: value })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.date-calendar', '日历')">
          <Input :value="properties.date?.calendar ?? ''" @update:value="value => patchDate({ calendar: value })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.date-storage', '存储格式')">
          <Select :value="properties.date?.storage ?? 'text'" :options="dateStorageOptions" @change="value => patchDate({ storage: String(value) })" />
        </FormItem>
      </section>

      <FormItem v-if="isList" class="list-items" :label="t('doc.content-control.list-items', '列表项')">
        <Input.TextArea :value="listText" :auto-size="{ minRows: 4, maxRows: 8 }" @update:value="updateListItems" />
      </FormItem>

      <section v-if="isRepeatingSection" class="properties-grid properties-section">
        <FormItem :label="t('doc.content-control.section-title', '重复节标题')">
          <Input :value="properties.repeatingSection?.sectionTitle ?? ''" @update:value="value => patch({ repeatingSection: { ...properties.repeatingSection, sectionTitle: value } })" />
        </FormItem>
        <Checkbox
          :checked="properties.repeatingSection?.doNotAllowInsertDelete ?? false"
          @change="value => patch({ repeatingSection: { ...properties.repeatingSection, doNotAllowInsertDelete: checkboxValue(value) } })"
        >
          {{ t('doc.content-control.prevent-repeat-changes', '禁止插入或删除重复项') }}
        </Checkbox>
      </section>

      <section v-if="isDocPart" class="properties-grid properties-section">
        <FormItem :label="t('doc.content-control.gallery', '文档部件库')">
          <Input :value="properties.docPart?.gallery ?? ''" @update:value="value => patch({ docPart: { ...properties.docPart, gallery: value } })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.category', '类别')">
          <Input :value="properties.docPart?.category ?? ''" @update:value="value => patch({ docPart: { ...properties.docPart, category: value } })" />
        </FormItem>
        <Checkbox
          :checked="properties.docPart?.unique ?? false"
          @change="value => patch({ docPart: { ...properties.docPart, unique: checkboxValue(value) } })"
        >
          {{ t('doc.content-control.unique', '唯一部件') }}
        </Checkbox>
      </section>

      <section v-if="hasDataBinding" class="properties-grid properties-section">
        <FormItem :label="t('doc.content-control.binding-xpath', '绑定 XPath')">
          <Input :value="properties.dataBinding?.xpath ?? ''" @update:value="value => patch({ dataBinding: { ...properties.dataBinding, xpath: value, dirty: true } })" />
        </FormItem>
        <FormItem :label="t('doc.content-control.binding-store', '数据存储 ID')">
          <Input :value="properties.dataBinding?.storeItemID ?? ''" @update:value="value => patch({ dataBinding: { ...properties.dataBinding, storeItemID: value, dirty: true } })" />
        </FormItem>
      </section>

      <footer class="properties-footer">
        <Button html-type="button" @click="close">{{ t('common.cancel', '取消') }}</Button>
        <Button type="primary" html-type="submit">{{ t('common.save', '保存') }}</Button>
      </footer>
    </Form>
  </Modal>
</template>

<style scoped>
.content-control-properties { display: flex; max-height: 72vh; flex-direction: column; gap: 16px; overflow: auto; padding: 4px; }
.properties-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.properties-grid :deep(.ant-form-item), .list-items { min-width: 0; margin: 0; }
.properties-grid :deep(.ant-form-item-label), .list-items :deep(.ant-form-item-label) { padding: 0 0 6px; }
.color-input { display: flex; min-width: 0; align-items: center; gap: 8px; }
.properties-section { border-top: 1px solid #e5e8ed; padding-top: 14px; }
.properties-footer { display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid #e5e8ed; padding-top: 14px; }
@media (max-width: 560px) { .properties-grid { grid-template-columns: 1fr; } }
</style>
