import type { OfficeLocale } from '../../office-i18n'
import { DOCUMENT_LATEX_CATEGORY_LABELS, DOCUMENT_LATEX_TEMPLATE_CATEGORIES, DOCUMENT_LATEX_TEMPLATES, type DocumentLatexTemplateCategory } from '../../../features/document/latex/latex-templates'

const categoryNamesZhTW: Record<DocumentLatexTemplateCategory, string> = {
  fractions: '分式', scripts: '上下標', radicals: '根式', integrals: '積分', 'large-operators': '大型運算子',
  brackets: '括號', functions: '函數', accents: '重音', 'limits-logarithms': '極限與對數', operators: '運算子', matrices: '矩陣',
}
const templateNamesZhTW: Record<string, string> = {
  'fraction-basic': '基本分式', 'fraction-expression': '運算式分式', 'fraction-derivative': '導數分式', 'fraction-partial': '偏導數分式',
  'script-superscript': '上標', 'script-subscript': '下標', 'script-combined': '上下標', 'script-left-right': '左右上下標',
  'radical-square': '平方根', 'radical-nth': 'n 次方根', 'radical-nested': '巢狀根式',
  'integral-basic': '不定積分', 'integral-bounded': '定積分', 'integral-double': '二重積分', 'integral-triple': '三重積分',
  'integral-contour': '環路積分', 'integral-contour-bounded': '帶路徑積分',
  'large-sum': '求和', 'large-product': '乘積', 'large-coproduct': '餘積', 'large-union': '大型聯集', 'large-intersection': '大型交集', 'large-logical-product': '大型邏輯且',
  'bracket-parentheses': '圓括號', 'bracket-square': '方括號', 'bracket-braces': '大括號', 'bracket-angle': '角括號',
  'bracket-absolute': '絕對值', 'bracket-norm': '範數', 'bracket-floor': '向下取整', 'bracket-ceiling': '向上取整',
  'function-sine': '正弦', 'function-cosine': '餘弦', 'function-tangent': '正切', 'function-logarithm': '對數',
  'function-natural-logarithm': '自然對數', 'function-exponential': '指數函數', 'function-minimum': '最小值', 'function-maximum': '最大值',
  'accent-hat': '帽重音', 'accent-bar': '橫線重音', 'accent-vector': '向量箭頭', 'accent-dot': '點重音',
  'accent-double-dot': '雙點重音', 'accent-overline': '上劃線', 'accent-underline': '下劃線', 'accent-overbrace': '上大括號',
  'limit-basic': '基本極限', 'limit-one-sided': '單側極限', 'limit-infinity': '無窮極限', 'logarithm-base': '指定底數對數',
  'minimum-conditioned': '條件最小值', 'maximum-conditioned': '條件最大值',
  'operator-plus-minus': '正負號', 'operator-multiply': '乘號', 'operator-divide': '除號', 'operator-not-equal': '不等於',
  'operator-approximately': '約等於', 'operator-less-equal': '小於等於', 'operator-greater-equal': '大於等於', 'operator-element': '屬於',
  'operator-not-element': '不屬於', 'operator-subset': '子集', 'operator-union': '聯集', 'operator-intersection': '交集',
  'matrix-2x2': '2×2 矩陣', 'matrix-3x3': '3×3 矩陣', 'matrix-parenthesized': '圓括號矩陣',
  'matrix-bracketed': '方括號矩陣', 'matrix-braced': '大括號矩陣', 'matrix-determinant': '行列式',
}
function templateMessages(locale: 'en-US' | 'zh-CN' | 'zh-TW'): Record<string, string> {
  return Object.fromEntries([
    ...DOCUMENT_LATEX_TEMPLATE_CATEGORIES.map(category => ['document.latex.category.' + category,
      locale === 'zh-TW' ? categoryNamesZhTW[category] : DOCUMENT_LATEX_CATEGORY_LABELS[category][locale === 'en-US' ? 1 : 0]]),
    ...DOCUMENT_LATEX_TEMPLATES.map(entry => ['document.latex.template.' + entry.id,
      locale === 'en-US' ? entry.englishLabel : locale === 'zh-TW' ? templateNamesZhTW[entry.id] ?? entry.label : entry.label]),
  ])
}

export const documentLatexEnUS = {
  ...templateMessages('en-US'),
  'document.latex.menu-formula': 'LaTeX formula', 'document.latex.floating-tools': 'Formula selection tools', 'document.latex.quick-formula': 'Insert formula',
  'document.latex.templates': 'Formula templates', 'document.latex.template-category': 'Template category',
  'document.latex.formula-tab': 'Formula', 'document.latex.insert-template': 'Insert {name}',
  'document.latex.edit': 'Edit formula', 'document.latex.remove': 'Delete formula', 'document.latex.replace-text': 'Convert formula to source text',
  'document.latex.stale': 'The document or formula selection has changed. Reopen the formula editor to continue.',
  'document.latex.failed': 'The formula could not be applied. Your source text has been kept.',
  'document.latex.read-only': 'This formula selection is read-only.',
  'document.latex.invalid-offset': '{message} (source position {offset})',
}
export const documentLatexLocaleMessages: Partial<Record<OfficeLocale, Record<string, string>>> = {
  'zh-CN': {
    ...templateMessages('zh-CN'),
    'document.latex.menu-formula': 'LaTeX 公式', 'document.latex.floating-tools': '公式选区操作', 'document.latex.quick-formula': '插入公式',
    'document.latex.templates': '公式模板', 'document.latex.template-category': '模板类别',
    'document.latex.formula-tab': '公式', 'document.latex.insert-template': '插入 {name}',
    'document.latex.edit': '编辑公式', 'document.latex.remove': '删除公式', 'document.latex.replace-text': '将公式转为源文本',
    'document.latex.stale': '文档或公式选区已变化，请重新打开公式编辑器后继续。',
    'document.latex.failed': '公式未能应用，已保留当前源文本。', 'document.latex.read-only': '当前公式选区只读。',
    'document.latex.invalid-offset': '{message}（源文本第 {offset} 个位置）',
  },
  'zh-TW': {
    ...templateMessages('zh-TW'),
    'document.latex.menu-formula': 'LaTeX 公式', 'document.latex.floating-tools': '公式選取操作', 'document.latex.quick-formula': '插入公式',
    'document.latex.templates': '公式範本', 'document.latex.template-category': '範本類別',
    'document.latex.formula-tab': '公式', 'document.latex.insert-template': '插入 {name}',
    'document.latex.edit': '編輯公式', 'document.latex.remove': '刪除公式', 'document.latex.replace-text': '將公式轉為原始文字',
    'document.latex.stale': '文件或公式選取範圍已變更，請重新開啟公式編輯器後繼續。',
    'document.latex.failed': '無法套用公式，已保留目前原始文字。', 'document.latex.read-only': '目前公式選取範圍唯讀。',
    'document.latex.invalid-offset': '{message}（原始文字第 {offset} 個位置）',
  },
}
documentLatexLocaleMessages['zh-HK'] = documentLatexLocaleMessages['zh-TW']
