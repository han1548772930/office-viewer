import { defineStore } from 'pinia'
import { ref } from 'vue'
import type { ImportedOfficeModel } from '../parser/contracts'
import type { ParserDiagnostic } from '../domain/model'
import { ParserError, parseOfficeFile } from '../services/wasmAdapter'

export const useParserStore = defineStore('parse-task', () => {
  const status = ref<'idle' | 'loading' | 'ready' | 'error'>('idle')
  const message = ref('')
  const diagnostics = ref<ParserDiagnostic[]>([])
  let controller: AbortController | undefined
  let generation = 0
  async function parse(file: File): Promise<ImportedOfficeModel | undefined> {
    controller?.abort()
    const token = ++generation
    const task = new AbortController()
    controller = task
    status.value = 'loading'
    message.value = '正在读取 '+file.name
    diagnostics.value = []
    try {
      const result = await parseOfficeFile(file, { signal: task.signal, onProgress: (progress) => {
        if (token === generation) message.value = progress.message
      } })
      if (token !== generation) return
      status.value = 'ready'
      message.value = '解析完成'
      diagnostics.value = result.diagnostics
      return result
    } catch (error: unknown) {
      if (token !== generation) return
      if (error instanceof ParserError && (error.code === 'CANCELLED' || error.code === 'SUPERSEDED')) {
        status.value = 'idle'
        message.value = '已取消'
        return
      }
      status.value = 'error'
      message.value = error instanceof Error ? error.message : '文件解析失败'
      diagnostics.value = error instanceof ParserError ? error.diagnostics : []
      throw error
    } finally { if (token === generation) controller = undefined }
  }
  function cancel(): void {
    generation++
    controller?.abort()
    controller = undefined
    status.value = 'idle'
    message.value = '已取消'
  }
  return { status, message, diagnostics, parse, cancel }
})
