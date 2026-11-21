import { parseFormula } from '../formula/formula-parser.ts'

export interface SheetFunctionEntry { name: string; label: string; category: string; example: string }

/** This UI catalog lists only functions currently implemented by the authoritative engine. */
export const SHEET_FUNCTION_CATALOG: readonly SheetFunctionEntry[] = [
  { name: 'SUM', label: '求和', category: '常用与统计', example: 'A1:A10' },
  { name: 'AVERAGE', label: '平均值', category: '常用与统计', example: 'A1:A10' },
  { name: 'COUNT', label: '数值计数', category: '常用与统计', example: 'A1:A10' },
  { name: 'COUNTA', label: '非空计数', category: '常用与统计', example: 'A1:A10' },
  { name: 'MIN', label: '最小值', category: '常用与统计', example: 'A1:A10' },
  { name: 'MAX', label: '最大值', category: '常用与统计', example: 'A1:A10' },
  { name: 'IF', label: '条件判断', category: '逻辑', example: 'A1>0,"是","否"' },
  { name: 'IFERROR', label: '错误时返回', category: '逻辑', example: 'A1/B1,0' },
  { name: 'AND', label: '全部为真', category: '逻辑', example: 'A1>0,B1>0' },
  { name: 'OR', label: '任一为真', category: '逻辑', example: 'A1>0,B1>0' },
  { name: 'NOT', label: '逻辑取反', category: '逻辑', example: 'A1>0' },
  { name: 'ABS', label: '绝对值', category: '数学', example: '-5' },
  { name: 'SQRT', label: '平方根', category: '数学', example: '16' },
  { name: 'INT', label: '向下取整', category: '数学', example: '3.9' },
  { name: 'ROUND', label: '四舍五入', category: '数学', example: 'A1,2' },
  { name: 'MOD', label: '余数', category: '数学', example: '10,3' },
  { name: 'POWER', label: '乘方', category: '数学', example: '2,3' },
  { name: 'MEDIAN', label: '中位数', category: '统计', example: 'A1:A10' },
  { name: 'LARGE', label: '第 k 大值', category: '统计', example: 'A1:A10,2' },
  { name: 'SMALL', label: '第 k 小值', category: '统计', example: 'A1:A10,2' },
  { name: 'VAR.S', label: '样本方差', category: '统计', example: 'A1:A10' },
  { name: 'STDEV.S', label: '样本标准差', category: '统计', example: 'A1:A10' },
  { name: 'MATCH', label: '查找位置', category: '查找与引用', example: '"East",A1:A10,0' },
  { name: 'INDEX', label: '按位置取值', category: '查找与引用', example: 'A1:C10,2,3' },
  { name: 'VLOOKUP', label: '纵向查找', category: '查找与引用', example: '"East",A1:C10,3,FALSE' },
  { name: 'PMT', label: '贷款定期付款', category: '财务', example: '5%/12,36,10000' },
  { name: 'PV', label: '现值', category: '财务', example: '5%/12,36,-300' },
  { name: 'FV', label: '未来值', category: '财务', example: '5%/12,36,-300' },
  { name: 'NPV', label: '净现值', category: '财务', example: '10%,B1:B5' },
  { name: 'LEN', label: '文本长度', category: '文本', example: '"文本"' },
  { name: 'UPPER', label: '转大写', category: '文本', example: '"hello"' },
  { name: 'LOWER', label: '转小写', category: '文本', example: '"HELLO"' },
  { name: 'TRIM', label: '整理空格', category: '文本', example: '" 文本 "' },
  { name: 'CONCAT', label: '合并文本', category: '文本', example: '"甲","乙"' },
  { name: 'DATE', label: '创建日期', category: '日期与时间', example: '2026,11,20' },
  { name: 'YEAR', label: '年份', category: '日期与时间', example: 'DATE(2026,11,20)' },
  { name: 'MONTH', label: '月份', category: '日期与时间', example: 'DATE(2026,11,20)' },
  { name: 'DAY', label: '日期中的日', category: '日期与时间', example: 'DATE(2026,11,20)' },
  { name: 'TIME', label: '创建时间', category: '日期与时间', example: '12,30,0' },
  { name: 'HOUR', label: '小时', category: '日期与时间', example: 'TIME(12,30,0)' },
  { name: 'MINUTE', label: '分钟', category: '日期与时间', example: 'TIME(12,30,0)' },
  { name: 'SECOND', label: '秒', category: '日期与时间', example: 'TIME(12,30,15)' },
]

export function buildSheetFunctionFormula(name: string, argumentsText: string): string {
  if (!SHEET_FUNCTION_CATALOG.some(entry => entry.name === name)) throw new Error('请选择当前函数目录中的函数。')
  if (!argumentsText.trim()) throw new Error('请输入函数参数或单元格区域。')
  const formula = `=${name}(${argumentsText.trim()})`
  try { parseFormula(formula) } catch { throw new Error('参数语法无效，请检查括号、引号和单元格引用。') }
  return formula
}
