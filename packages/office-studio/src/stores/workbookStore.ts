import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import { createBlankWorkbook, type WorkbookModel } from '../domain/model'

export const useWorkbookStore = defineStore('workbook-model', () => {
  const model = shallowRef<WorkbookModel>(createBlankWorkbook())
  function replace(value: WorkbookModel): void { model.value = value }
  return { model, replace }
})
