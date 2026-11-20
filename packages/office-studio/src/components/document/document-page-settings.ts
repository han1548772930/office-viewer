import type { DocumentPageSetup } from '../../domain/model'

export const DOCUMENT_PAPER_SIZES = { A4: { width: 794, height: 1123 }, Letter: { width: 816, height: 1056 } } as const
export type DocumentPaperSize = keyof typeof DOCUMENT_PAPER_SIZES | 'Custom'

export function documentPaperSize(page: DocumentPageSetup): DocumentPaperSize {
  const width = Math.min(page.width, page.height)
  const height = Math.max(page.width, page.height)
  for (const size of ['A4', 'Letter'] as const) {
    const expected = DOCUMENT_PAPER_SIZES[size]
    if (Math.abs(width - expected.width) < 1 && Math.abs(height - expected.height) < 1) return size
  }
  return 'Custom'
}

export function documentPageSettingsError(page: DocumentPageSetup): string | undefined {
  if (!Object.values(page).every((value) => typeof value === 'number' && Number.isFinite(value))) return '纸张尺寸和页边距必须是有效数字'
  if (page.width < 200 || page.height < 240) return '纸张宽度至少为 200 像素，高度至少为 240 像素'
  if ([page.marginTop, page.marginRight, page.marginBottom, page.marginLeft].some((value) => value < 0)) return '页边距不能为负数'
  if (page.marginLeft + page.marginRight >= page.width || page.marginTop + page.marginBottom >= page.height) return '页边距之和必须小于对应纸张尺寸'
  // The current shared layout kernel clamps individual margins at 42%.
  // Reject this case explicitly instead of silently accepting a different rendering.
  if (page.marginLeft > page.width * 0.42 || page.marginRight > page.width * 0.42 || page.marginTop > page.height * 0.42 || page.marginBottom > page.height * 0.42) return '当前排版要求单侧页边距不超过纸张对应尺寸的 42%'
  return undefined
}

/** Changing orientation preserves imported custom geometry and all page margins. */
export function documentOrientedPage(page: DocumentPageSetup, orientation: 'portrait' | 'landscape', paper: DocumentPaperSize): DocumentPageSetup {
  const size = paper === 'Custom' ? { width: Math.min(page.width, page.height), height: Math.max(page.width, page.height) } : DOCUMENT_PAPER_SIZES[paper]
  return { ...page, width: orientation === 'portrait' ? size.width : size.height, height: orientation === 'portrait' ? size.height : size.width }
}
