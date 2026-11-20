import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import { createBlankDocument, type DocumentModel } from '../domain/model'

export const useDocumentStore = defineStore('document-model', () => {
  const model = shallowRef<DocumentModel>(createBlankDocument())
  function replace(value: DocumentModel): void { model.value = value }
  return { model, replace }
})
