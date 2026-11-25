import { defineStore } from 'pinia'
import { ref, shallowRef } from 'vue'
import type { Pinia } from 'pinia'
import { cloneModel } from '../../domain/serialization'
import type { ChartObjectClipboard } from './chart-object-clipboard'

export const useChartObjectClipboardStore = defineStore('office-chart-object-clipboard', () => {
  const clipboard = shallowRef<ChartObjectClipboard>()
  const systemWritten = ref(false), request = ref(0)
  const pendingUnit = ref<string>()
  function beginCopy(unitId: string): number { pendingUnit.value = unitId; return ++request.value }
  function publish(ticket: number, value: ChartObjectClipboard, written: boolean): boolean {
    if (ticket !== request.value) return false
    clipboard.value = cloneModel(value); systemWritten.value = written; pendingUnit.value = undefined
    return true
  }
  function clear(): void { ++request.value; clipboard.value = undefined; systemWritten.value = false; pendingUnit.value = undefined }
  function clearUnit(unitId: string): void { if (clipboard.value?.source.unitId === unitId || pendingUnit.value === unitId) clear() }
  return { clipboard, systemWritten, request, beginCopy, publish, clear, clearUnit }
})

export function chartObjectClipboardStore(pinia: Pinia): ReturnType<typeof useChartObjectClipboardStore> {
  return useChartObjectClipboardStore(pinia)
}
