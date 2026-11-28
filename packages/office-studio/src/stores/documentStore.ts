import { defineStore } from 'pinia'
import { shallowRef, triggerRef } from 'vue'
import { createBlankDocument, type DocumentModel } from '../domain/model'

export const useDocumentStore = defineStore('document-model', () => {
  const model = shallowRef<DocumentModel>(createBlankDocument())
  function replace(value: DocumentModel): void { model.value = value }
  function notifyViewChange(): void { triggerRef(model) }
  return { model, replace, notifyViewChange }
})
