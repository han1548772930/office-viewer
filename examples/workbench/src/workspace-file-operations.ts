export type WorkspaceKind = 'workbook' | 'document'
export type FileAction = 'save' | 'export'
export type NoticeType = 'success' | 'error' | 'info'

export function workspaceFileNameError(value: string): string | undefined {
  const name = value.trim()
  if (!name) return '文件名不能为空'
  if (/^\.+$/u.test(name)) return '文件名不能只包含句点'
  if (name.length > 255 || /[\u0000-\u001f<>:"/\\|?*]/u.test(name)) return '文件名不能超过255个字符，且不能包含 <>:"/\\|?* 或控制字符'
  return undefined
}

export interface FileSavepoint {
  kind: WorkspaceKind
  unitId: string
  revision: number
  generation: number
}

export interface FileSnapshot<Model> {
  kind: WorkspaceKind
  model: Model
  fileName: string
  savepoint: FileSavepoint
}

export interface ExportTask {
  action: FileAction
  fileName: string
  message: string
}

interface ExportResult {
  bytes: Uint8Array
  warnings: ReadonlyArray<{ code: string; message: string }>
}

export interface WorkspaceFileServices<Model, FileInput extends { name: string }> {
  importFile: (file: FileInput) => Promise<boolean>
  cancelImport: () => void
  capture: (kind?: WorkspaceKind) => FileSnapshot<Model>
  write: (snapshot: FileSnapshot<Model>, options: {
    signal: AbortSignal
    onProgress: (progress: { message: string }) => void
  }) => Promise<ExportResult>
  download: (bytes: Uint8Array, snapshot: FileSnapshot<Model>) => void
  markSaved: (savepoint: FileSavepoint) => boolean
  onExportTask: (task: ExportTask | null) => void
  notify: (type: NoticeType, message: string) => void
}

/** Coordinate file operations without owning a second document or history model. */
export function createWorkspaceFileOperations<Model, FileInput extends { name: string }>(services: WorkspaceFileServices<Model, FileInput>) {
  let importGeneration = 0
  let exportGeneration = 0
  let exportController: AbortController | undefined
  let disposed = false

  async function importFile(file: FileInput): Promise<boolean> {
    if (disposed) return false
    const generation = ++importGeneration
    try {
      const applied = await services.importFile(file)
      if (disposed || generation !== importGeneration || !applied) return false
      services.notify('success', `${file.name} 已解析并载入`)
      return true
    } catch (error: unknown) {
      if (!disposed && generation === importGeneration) services.notify('error', error instanceof Error ? error.message : '文件解析失败')
      return false
    }
  }

  function cancelImport(): void {
    importGeneration++
    services.cancelImport()
  }

  async function exportFile(action: FileAction = 'save', kind?: WorkspaceKind): Promise<boolean> {
    if (disposed) return false
    const generation = ++exportGeneration
    exportController?.abort()
    const controller = new AbortController()
    exportController = controller
    try {
      const snapshot = services.capture(kind)
      services.onExportTask({ action, fileName: snapshot.fileName, message: '正在准备文件' })
      const result = await services.write(snapshot, {
        signal: controller.signal,
        onProgress: ({ message }) => {
          if (!disposed && generation === exportGeneration) services.onExportTask({ action, fileName: snapshot.fileName, message })
        },
      })
      if (disposed || controller.signal.aborted || generation !== exportGeneration) return false
      services.download(result.bytes, snapshot)
      // A lossy Office export must not clear the original document's dirty state.
      if (action === 'save' && result.warnings.length === 0) services.markSaved(snapshot.savepoint)
      const format = snapshot.kind === 'workbook' ? 'XLSX' : 'DOCX'
      if (result.warnings.length) {
        services.notify('info', `${format} 已下载，有 ${result.warnings.length} 项内容未完整保留：${result.warnings[0]!.message}。请保留原文件。`)
      } else {
        services.notify('success', `${snapshot.fileName} ${action === 'save' ? '保存副本' : '导出文件'}已下载`)
      }
      return true
    } catch (error: unknown) {
      if (!disposed && generation === exportGeneration && !controller.signal.aborted) services.notify('error', error instanceof Error ? error.message : 'Office 文件导出失败')
      return false
    } finally {
      if (!disposed && generation === exportGeneration) {
        exportController = undefined
        services.onExportTask(null)
      }
    }
  }

  function cancelExport(): boolean {
    if (!exportController) return false
    exportGeneration++
    exportController.abort()
    exportController = undefined
    services.onExportTask(null)
    if (!disposed) services.notify('info', '导出任务已取消')
    return true
  }

  function dispose(): void {
    if (disposed) return
    disposed = true
    importGeneration++
    exportGeneration++
    exportController?.abort()
    exportController = undefined
    services.cancelImport()
  }

  return { importFile, cancelImport, exportFile, cancelExport, dispose }
}
