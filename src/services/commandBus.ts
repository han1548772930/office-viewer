export interface EditorCommand<TContext> {
  id: string
  label: string
  execute(context: TContext): void
  undo(context: TContext): void
}

export class CommandBus<TContext> {
  private readonly undoStack: EditorCommand<TContext>[] = []
  private readonly redoStack: EditorCommand<TContext>[] = []

  execute(command: EditorCommand<TContext>, context: TContext): void {
    command.execute(context)
    this.undoStack.push(command)
    this.redoStack.length = 0
  }

  undo(context: TContext): EditorCommand<TContext> | undefined {
    const command = this.undoStack.pop()
    if (!command) return undefined
    command.undo(context)
    this.redoStack.push(command)
    return command
  }

  redo(context: TContext): EditorCommand<TContext> | undefined {
    const command = this.redoStack.pop()
    if (!command) return undefined
    command.execute(context)
    this.undoStack.push(command)
    return command
  }

  canUndo(): boolean {
    return this.undoStack.length > 0
  }

  canRedo(): boolean {
    return this.redoStack.length > 0
  }

  clear(): void {
    this.undoStack.length = 0
    this.redoStack.length = 0
  }
}
