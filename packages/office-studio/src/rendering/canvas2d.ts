import type { RenderDriver, RenderScene } from './contracts'

export function createCanvas2DDriver(canvas: HTMLCanvasElement): RenderDriver {
  const context = canvas.getContext('2d', { alpha: false })
  if (!context) throw new Error('无法创建 Canvas2D 绘图上下文')
  let scaleX = 1, scaleY = 1
  return {
    backend: 'canvas2d',
    resize(width, height, pixelRatio) {
      const safePixelRatio = Number.isFinite(pixelRatio) && pixelRatio > 0 ? pixelRatio : 1
      const cssWidth = Math.max(0, width), cssHeight = Math.max(0, height)
      canvas.width = Math.max(1, Math.round(cssWidth * safePixelRatio))
      canvas.height = Math.max(1, Math.round(cssHeight * safePixelRatio))
      // Rounding the backing store makes the effective ratio differ slightly from
      // the requested DPR. Keep drawing coordinates aligned with its actual size.
      scaleX = cssWidth > 0 ? canvas.width / cssWidth : 1
      scaleY = cssHeight > 0 ? canvas.height / cssHeight : 1
    },
    render(scene: RenderScene) {
      context.setTransform(scaleX, 0, 0, scaleY, 0, 0)
      context.fillStyle = scene.background ?? '#ffffff'
      context.fillRect(0, 0, scene.width, scene.height)
      for (const command of scene.commands) {
        context.save()
        if (command.clip) { context.beginPath(); context.rect(command.clip.x, command.clip.y, command.clip.width, command.clip.height); context.clip() }
        switch (command.kind) {
          case 'rect':
            context.fillStyle = command.fill; context.fillRect(command.x, command.y, command.width, command.height)
            if (command.stroke) { context.strokeStyle = command.stroke; context.lineWidth = command.strokeWidth ?? 1; context.strokeRect(command.x, command.y, command.width, command.height) }
            break
          case 'line':
            context.strokeStyle = command.color; context.lineWidth = command.width ?? 1; context.beginPath(); context.moveTo(command.x1, command.y1); context.lineTo(command.x2, command.y2); context.stroke(); break
          case 'text': {
            context.font = command.font; context.fillStyle = command.color; context.textAlign = command.align ?? 'left'; context.textBaseline = command.baseline ?? 'alphabetic'
            if (command.maxWidth !== undefined) {
              const left = command.x - (command.align === 'center' ? command.maxWidth / 2 : command.align === 'right' ? command.maxWidth : 0)
              context.beginPath(); context.rect(left, -1e7, command.maxWidth, 2e7); context.clip()
            }
            context.fillText(command.text, command.x, command.y)
            break
          }
          case 'image':
            context.globalAlpha = command.opacity ?? 1
            context.drawImage(command.source, command.x, command.y, command.width, command.height)
            break
        }
        context.restore()
      }
    },
    dispose() { canvas.width = 1; canvas.height = 1 },
  }
}
