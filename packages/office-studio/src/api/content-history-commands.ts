import type { OfficeKind } from '../domain/model'

export const CONTENT_HISTORY_COMMANDS = {
  undo: 'core.command.undo',
  redo: 'core.command.redo',
} as const

export type ContentHistoryDirection = 'undo' | 'redo'

export class ContentHistoryCanceledError extends Error {
  readonly code = 'CONTENT_HISTORY_CANCELED'
  constructor(readonly direction: ContentHistoryDirection) {
    super(`${direction === 'undo' ? 'Undo' : 'Redo'} was canceled`)
    this.name = 'ContentHistoryCanceledError'
  }
}

const HISTORY_ACTIONS: Readonly<Record<string, { direction: ContentHistoryDirection; kind?: OfficeKind }>> = {
  [CONTENT_HISTORY_COMMANDS.undo]: { direction: 'undo' },
  [CONTENT_HISTORY_COMMANDS.redo]: { direction: 'redo' },
  'univer.command.undo': { direction: 'undo' },
  'univer.command.redo': { direction: 'redo' },
  'sheet.command.undo': { direction: 'undo', kind: 'workbook' },
  'sheet.command.redo': { direction: 'redo', kind: 'workbook' },
  'doc.command.undo': { direction: 'undo', kind: 'document' },
  'doc.command.redo': { direction: 'redo', kind: 'document' },
}

export const CONTENT_HISTORY_COMMAND_IDS = Object.freeze(Object.keys(HISTORY_ACTIONS))

export interface ContentHistoryCommandInput {
  unitId?: string
  kind?: OfficeKind
}

export interface ContentHistoryTarget {
  unitId: string
  kind: OfficeKind
}

export function contentHistoryCommandAction(id: string): { direction: ContentHistoryDirection; kind?: OfficeKind } | undefined {
  const action = Object.hasOwn(HISTORY_ACTIONS, id) ? HISTORY_ACTIONS[id] : undefined
  return action ? { ...action } : undefined
}

/** Resolve the requested unit before replay emits any cancellable history events. */
export function resolveContentHistoryTarget(
  commandId: string,
  input: unknown,
  active: { kind: OfficeKind; workbookId: string; documentId: string },
  hasUnit: (kind: OfficeKind, unitId: string) => boolean,
): ContentHistoryTarget {
  const action = contentHistoryCommandAction(commandId)
  if (!action) throw new Error(`Unknown history command: ${commandId}`)
  if (input !== undefined && (input === null || typeof input !== 'object' || Array.isArray(input))) {
    throw new TypeError('历史命令参数必须是对象。')
  }
  const source = input === undefined ? {} : input as Record<string, unknown>
  const requestedKind = source.kind, requestedId = source.unitId
  if (requestedKind !== undefined && requestedKind !== 'workbook' && requestedKind !== 'document') {
    throw new TypeError('历史命令的文档类型无效。')
  }
  if (requestedId !== undefined && (typeof requestedId !== 'string' || !requestedId.trim())) {
    throw new TypeError('历史命令的文档标识无效。')
  }
  if (action.kind && requestedKind !== undefined && requestedKind !== action.kind) {
    throw new TypeError('历史命令与目标文档类型不一致。')
  }
  let kind = requestedKind ?? action.kind
  if (!kind && requestedId !== undefined) {
    const workbook = hasUnit('workbook', requestedId), document = hasUnit('document', requestedId)
    if (workbook && document) throw new TypeError('文档标识对应多个类型，不能共用内容历史。')
    kind = workbook ? 'workbook' : document ? 'document' : undefined
    if (!kind) throw new Error(`History unit not found: ${requestedId}`)
  }
  kind ??= active.kind
  const unitId = requestedId ?? (kind === 'workbook' ? active.workbookId : active.documentId)
  if (hasUnit('workbook', unitId) && hasUnit('document', unitId)) {
    throw new TypeError('文档标识对应多个类型，不能共用内容历史。')
  }
  if (!hasUnit(kind, unitId)) throw new Error(`History ${kind} not found: ${unitId}`)
  return { kind, unitId }
}
