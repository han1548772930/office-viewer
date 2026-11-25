import type { SheetDrawingLeafRef } from '../../domain/sheet-drawings.ts'
import type { DrawCommand } from '../contracts.ts'

/** Stable within each owner, including its pane fragments. Selection controls stay above content. */
export function orderSheetDrawingCommands(commands: readonly DrawCommand[], order?: readonly SheetDrawingLeafRef[]): DrawCommand[] {
  if (!order) return [...commands]
  const ranks = new Map(order.map((ref, index) => [`${ref.kind}\0${ref.id}`, index]))
  const content: Array<{ command: DrawCommand; index: number; rank: number }> = [], overlays: DrawCommand[] = []
  commands.forEach((command, index) => {
    if (!command.drawingRef || command.drawingOverlay) { overlays.push(command); return }
    content.push({ command, index, rank: ranks.get(`${command.drawingRef.kind}\0${command.drawingRef.id}`) ?? order.length })
  })
  content.sort((left, right) => left.rank - right.rank || left.index - right.index)
  return [...content.map(entry => entry.command), ...overlays]
}
