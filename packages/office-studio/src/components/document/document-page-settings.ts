import type { DocumentPageSetup } from '../../domain/model'

export const DOCUMENT_PAPER_SIZES = {
  A3: { width: 1123, height: 1587 },
  A4: { width: 794, height: 1124 },
  A5: { width: 559, height: 794 },
  B4: { width: 944, height: 1344 },
  B5: { width: 665, height: 944 },
  Executive: { width: 696, height: 1008 },
  Folio: { width: 816, height: 1248 },
  Legal: { width: 816, height: 1344 },
  Letter: { width: 816, height: 1056 },
  Statement: { width: 528, height: 816 },
  Tabloid: { width: 1056, height: 1632 },
} as const
export type DocumentPaperSize = keyof typeof DOCUMENT_PAPER_SIZES | 'Custom'
export const DOCUMENT_MODERN_WIDTHS = { narrow: 794, medium: 960, wide: 1123 } as const
export type DocumentPageMode = 'traditional' | 'modern'

export function documentPaperSize(page: DocumentPageSetup): DocumentPaperSize {
  const width = Math.min(page.width, page.height)
  const height = Math.max(page.width, page.height)
  for (const [size, expected] of Object.entries(DOCUMENT_PAPER_SIZES)) {
    if (Math.abs(width - expected.width) < 1 && Math.abs(height - expected.height) < 1) return size as keyof typeof DOCUMENT_PAPER_SIZES
  }
  return 'Custom'
}

export function documentPageSettingsError(page: DocumentPageSetup): string | undefined {
  if (!Object.values(page).every((value) => typeof value === 'number' && Number.isFinite(value))) return '纸张尺寸和页边距必须是有效数字'
  if (page.width < 200 || page.height < 240) return '纸张宽度至少为 200 像素，高度至少为 240 像素'
  if ([page.marginTop, page.marginRight, page.marginBottom, page.marginLeft].some((value) => value < 0)) return '页边距不能为负数'
  if (page.marginLeft + page.marginRight >= page.width || page.marginTop + page.marginBottom >= page.height) return '页边距之和必须小于对应纸张尺寸'
  if (page.marginLeft > page.width / 2 || page.marginRight > page.width / 2 || page.marginTop > page.height / 2 || page.marginBottom > page.height / 2) return '页边距不能超过对应纸张尺寸的一半'
  return undefined
}

export function documentModernWidth(page: DocumentPageSetup): keyof typeof DOCUMENT_MODERN_WIDTHS {
  let nearest: keyof typeof DOCUMENT_MODERN_WIDTHS = 'medium'
  for (const mode of ['narrow', 'medium', 'wide'] as const) {
    if (Math.abs(DOCUMENT_MODERN_WIDTHS[mode] - page.width) < Math.abs(DOCUMENT_MODERN_WIDTHS[nearest] - page.width)) nearest = mode
  }
  return nearest
}

/** Changing orientation preserves imported custom geometry and all page margins. */
export function documentOrientedPage(page: DocumentPageSetup, orientation: 'portrait' | 'landscape', paper: DocumentPaperSize): DocumentPageSetup {
  const size = paper === 'Custom' ? { width: Math.min(page.width, page.height), height: Math.max(page.width, page.height) } : DOCUMENT_PAPER_SIZES[paper]
  return { ...page, width: orientation === 'portrait' ? size.width : size.height, height: orientation === 'portrait' ? size.height : size.width }
}
