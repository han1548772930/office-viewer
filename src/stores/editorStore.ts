import { computed, reactive } from 'vue'
import { defineStore } from 'pinia'
import {
  createBlankDocument,
  createBlankWorkbook,
  type CellRange,
  type DocumentModel,
  type EditorDocumentState,
  type OfficeKind,
  type ParseEnvelope,
  type SpreadsheetCell,
  type SpreadsheetSheet,
  type WorkbookModel,
} from '../domain/model'
import { CommandBus, type EditorCommand } from '../services/commandBus'
import { parseOfficeFile } from '../services/wasmAdapter'
import type { ExcelWorkbookView } from '../services/excelViewAdapter'
import { applyExcelWorkbookView } from '../services/excelViewAdapter'

export const useEditorStore = defineStore('editor', () => {
const state = reactive<EditorDocumentState>({
  activeKind: 'workbook',
  workbook: createBlankWorkbook(),
  document: createBlankDocument(),
  fileName: '未命名工作簿',
  parserStatus: 'idle',
  parserMessage: '',
  diagnostics: [],
  selectedRange: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
})

const commandBus = new CommandBus<EditorDocumentState>()
const activeSheet = computed(() => state.workbook.sheets.find((sheet) => sheet.id === state.workbook.activeSheetId) ?? state.workbook.sheets[0])

function setActiveKind(kind: OfficeKind): void {
  state.activeKind = kind
}

function execute(command: EditorCommand<EditorDocumentState>): void {
  commandBus.execute(command, state)
  if (state.activeKind === 'workbook') state.workbook.modified = true
  else state.document.modified = true
}

function undo(): string | undefined {
  return commandBus.undo(state)?.label
}

function redo(): string | undefined {
  return commandBus.redo(state)?.label
}

function setCell(sheet: SpreadsheetSheet, row: number, col: number, value: string): void {
  const id = `${row}:${col}`
  const oldCell = sheet.cells[id]
  const formula = value.startsWith('=') ? value : null
  const nextCell: SpreadsheetCell = oldCell
    ? { ...oldCell, value, formula, display: value }
    : {
        id,
        row,
        col,
        value,
        formula,
        display: value,
        style: {
          font: { family: 'Aptos', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#1f2937' },
          fill: '#ffffff',
          horizontal: 'left',
          vertical: 'middle',
          wrap: false,
          numberFormat: 'General',
        },
      }
  sheet.cells[id] = nextCell
}

function editCell(row: number, col: number, value: string): void {
  const sheet = activeSheet.value
  if (!sheet) return
  const id = `${row}:${col}`
  const previous = sheet.cells[id] ? { ...sheet.cells[id] } : undefined
  execute({
    id: 'cell.edit',
    label: '编辑单元格',
    execute: () => setCell(sheet, row, col, value),
    undo: () => {
      if (previous) sheet.cells[id] = previous
      else delete sheet.cells[id]
    },
  })
}

function selectSheet(sheetId: string): void {
  if (state.workbook.sheets.some((sheet) => sheet.id === sheetId)) state.workbook.activeSheetId = sheetId
}

function addSheet(): SpreadsheetSheet {
  const sheet: SpreadsheetSheet = {
    id: `sheet-${Date.now()}-${state.workbook.sheets.length}`,
    name: `Sheet${state.workbook.sheets.length + 1}`,
    visible: true,
    cells: {},
    mergedRanges: [],
    columnWidths: {},
    rowHeights: {},
    frozenRows: 0,
    frozenCols: 0,
    showGridlines: true,
  }
  state.workbook.sheets.push(sheet)
  state.workbook.activeSheetId = sheet.id
  state.workbook.modified = true
  return sheet
}

function removeSheet(sheetId: string): void {
  if (state.workbook.sheets.length <= 1) return
  const index = state.workbook.sheets.findIndex((sheet) => sheet.id === sheetId)
  if (index < 0) return
  state.workbook.sheets.splice(index, 1)
  if (state.workbook.activeSheetId === sheetId) state.workbook.activeSheetId = state.workbook.sheets[Math.max(0, index - 1)].id
  state.workbook.modified = true
}

function setSelection(range: CellRange): void {
  state.selectedRange = { ...range }
}

function replaceWorkbookFromView(view: ExcelWorkbookView): void {
  const before = structuredClone(state.workbook)
  const after = applyExcelWorkbookView(state.workbook, view)
  after.modified = true
  commandBus.execute({
    id: 'workbook.update',
    label: '编辑工作簿',
    execute: (context) => { context.workbook = structuredClone(after) },
    undo: (context) => { context.workbook = structuredClone(before) },
  }, state)
}

function replaceDocumentFromView(document: DocumentModel): void {
  const before = structuredClone(state.document)
  const after = structuredClone(document)
  after.modified = true
  commandBus.execute({
    id: 'document.update',
    label: '编辑文档',
    execute: (context) => { context.document = structuredClone(after) },
    undo: (context) => { context.document = structuredClone(before) },
  }, state)
}

function newWorkbook(): void {
  state.workbook = createBlankWorkbook()
  state.workbook.name = '未命名工作簿'
  state.fileName = state.workbook.name
  state.activeKind = 'workbook'
  state.parserStatus = 'idle'
  state.diagnostics = []
  state.selectedRange = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }
  commandBus.clear()
}

function newDocument(): void {
  state.document = createBlankDocument()
  state.activeKind = 'document'
  state.fileName = state.document.name
  state.parserStatus = 'idle'
  state.diagnostics = []
  commandBus.clear()
}

function applyEnvelope(envelope: ParseEnvelope): void {
  state.fileName = envelope.name
  state.diagnostics = envelope.diagnostics
  state.parserStatus = 'ready'
  state.parserMessage = '解析完成'
  if (envelope.kind === 'xlsx' || envelope.kind === 'xls' || envelope.kind === 'xlsb' || envelope.kind === 'ods') {
    state.activeKind = 'workbook'
    state.workbook = normalizeWorkbook(envelope)
  } else {
    state.activeKind = 'document'
    state.document = normalizeDocument(envelope)
  }
}

async function importFile(file: File): Promise<void> {
  state.parserStatus = 'loading'
  state.parserMessage = `正在解析 ${file.name}`
  state.diagnostics = []
  try {
    applyEnvelope(await parseOfficeFile(file))
  } catch (error: unknown) {
    state.parserStatus = 'error'
    state.parserMessage = error instanceof Error ? error.message : '解析失败'
    throw error
  }
}

function normalizeWorkbook(envelope: ParseEnvelope): WorkbookModel {
  const payload = envelope.payload as { sheets?: Array<Record<string, unknown>>; date1904?: boolean; activeSheet?: number }
  const sheets = (payload.sheets ?? []).map((raw, index) => normalizeSheet(raw, index))
  const normalizedSheets = sheets.length > 0 ? sheets : createBlankWorkbook().sheets
  return {
    kind: 'workbook',
    name: envelope.name,
    sheets: normalizedSheets,
    activeSheetId: normalizedSheets[payload.activeSheet ?? 0]?.id ?? normalizedSheets[0].id,
    date1904: payload.date1904 ?? false,
    modified: false,
  }
}

function normalizeSheet(raw: Record<string, unknown>, index: number): SpreadsheetSheet {
  const sheet: SpreadsheetSheet = {
    id: `sheet-imported-${index}-${Date.now()}`,
    name: typeof raw.name === 'string' ? raw.name : `Sheet${index + 1}`,
    visible: raw.visible !== 'hidden' && raw.visible !== false,
    cells: {},
    mergedRanges: Array.isArray(raw.mergedRanges) ? raw.mergedRanges.map((range) => ({
      startRow: Number((range as Record<string, unknown>).firstRow ?? 0),
      startCol: Number((range as Record<string, unknown>).firstCol ?? 0),
      endRow: Number((range as Record<string, unknown>).lastRow ?? 0),
      endCol: Number((range as Record<string, unknown>).lastCol ?? 0),
    })) : [],
    columnWidths: {},
    rowHeights: {},
    frozenRows: 0,
    frozenCols: 0,
    showGridlines: raw.gridlines !== false,
  }
  if (Array.isArray(raw.cells)) {
    for (const source of raw.cells) {
      const cell = source as Record<string, unknown>
      const row = Number(cell.row ?? 0)
      const col = Number(cell.col ?? 0)
      const id = `${row}:${col}`
      sheet.cells[id] = {
        id,
        row,
        col,
        value: (cell.value as SpreadsheetCell['value']) ?? '',
        formula: typeof cell.formula === 'string' ? cell.formula : null,
        display: typeof cell.display === 'string' ? cell.display : String(cell.value ?? ''),
        style: {
          font: { family: 'Aptos', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#1f2937' },
          fill: '#ffffff', horizontal: 'left', vertical: 'middle', wrap: false, numberFormat: 'General',
        },
      }
    }
  }
  return sheet
}

function normalizeDocument(envelope: ParseEnvelope): DocumentModel {
  const document = createBlankDocument()
  document.name = envelope.name
  document.modified = false
  document.blocks[0].runs[0].text = typeof envelope.payload === 'string' ? envelope.payload : '已导入文档内容，请从解析结果继续编辑。'
  document.wordCount = document.blocks.reduce((total, block) => total + block.runs.reduce((count, run) => count + run.text.length, 0), 0)
  return document
}

  return {
    state,
    activeSheet,
    canUndo: computed(() => commandBus.canUndo()),
    canRedo: computed(() => commandBus.canRedo()),
    setActiveKind,
    execute,
    undo,
    redo,
    editCell,
    selectSheet,
    addSheet,
    removeSheet,
    setSelection,
    replaceWorkbookFromView,
    replaceDocumentFromView,
    newWorkbook,
    newDocument,
    importFile,
  }
})
