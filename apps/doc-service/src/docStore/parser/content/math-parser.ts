import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface MathElement {
  type: 'oMath' | 'oMathPara'
  content: string
  structure?: any
}

export class MathParser {
  /**
   * 解析文档中的所有数学公式
   */
  parse(xml: string): MathElement[] {
    const mathElements: MathElement[] = []
    const doc = xmlParser.parse(xml)
    
    this.extractMath(doc, mathElements)
    return mathElements
  }

  /**
   * 解析单个数学公式
   */
  parseMath(mathNode: XmlNode): MathElement | null {
    if (!mathNode) return null

    const type = mathNode['m:oMath'] ? 'oMath' : 'oMathPara'
    const content = this.extractMathText(mathNode)
    const structure = this.parseMathStructure(mathNode)

    return {
      type,
      content,
      structure
    }
  }

  /**
   * 解析数学公式结构
   */
  private parseMathStructure(node: XmlNode): any {
    const structure: any = {}

    // 分数 (m:f)
    if (node['m:f']) {
      structure.type = 'fraction'
      structure.numerator = this.extractMathText(node['m:f']['m:num'])
      structure.denominator = this.extractMathText(node['m:f']['m:den'])
    }

    // 上标 (m:sSup)
    if (node['m:sSup']) {
      structure.type = 'superscript'
      structure.base = this.extractMathText(node['m:sSup']['m:e'])
      structure.superscript = this.extractMathText(node['m:sSup']['m:sup'])
    }

    // 下标 (m:sSub)
    if (node['m:sSub']) {
      structure.type = 'subscript'
      structure.base = this.extractMathText(node['m:sSub']['m:e'])
      structure.subscript = this.extractMathText(node['m:sSub']['m:sub'])
    }

    // 上下标 (m:sSubSup)
    if (node['m:sSubSup']) {
      structure.type = 'subsuperscript'
      structure.base = this.extractMathText(node['m:sSubSup']['m:e'])
      structure.subscript = this.extractMathText(node['m:sSubSup']['m:sub'])
      structure.superscript = this.extractMathText(node['m:sSubSup']['m:sup'])
    }

    // 根号 (m:rad)
    if (node['m:rad']) {
      structure.type = 'radical'
      structure.degree = this.extractMathText(node['m:rad']['m:deg'])
      structure.base = this.extractMathText(node['m:rad']['m:e'])
    }

    // 矩阵 (m:m)
    if (node['m:m']) {
      structure.type = 'matrix'
      structure.rows = this.parseMatrixRows(node['m:m'])
    }

    // 括号 (m:d)
    if (node['m:d']) {
      structure.type = 'delimiters'
      structure.begin = node['m:d']['m:dPr']?.['m:begChr']?.['@_m:val'] || '('
      structure.end = node['m:d']['m:dPr']?.['m:endChr']?.['@_m:val'] || ')'
      structure.content = this.extractMathText(node['m:d']['m:e'])
    }

    // 函数 (m:func)
    if (node['m:func']) {
      structure.type = 'function'
      structure.name = this.extractMathText(node['m:func']['m:fName'])
      structure.argument = this.extractMathText(node['m:func']['m:e'])
    }

    // 积分 (m:nary)
    if (node['m:nary']) {
      structure.type = 'nary'
      structure.operator = node['m:nary']['m:naryPr']?.['m:chr']?.['@_m:val'] || '∫'
      structure.lowerLimit = this.extractMathText(node['m:nary']['m:sub'])
      structure.upperLimit = this.extractMathText(node['m:nary']['m:sup'])
      structure.base = this.extractMathText(node['m:nary']['m:e'])
    }

    return structure
  }

  /**
   * 解析矩阵行
   */
  private parseMatrixRows(matrixNode: XmlNode): any[] {
    const rows: any[] = []
    
    if (matrixNode['m:mr']) {
      const rowArray = Array.isArray(matrixNode['m:mr'])
        ? matrixNode['m:mr']
        : [matrixNode['m:mr']]
      
      for (const row of rowArray) {
        const cells: string[] = []
        if (row['m:e']) {
          const cellArray = Array.isArray(row['m:e']) ? row['m:e'] : [row['m:e']]
          for (const cell of cellArray) {
            cells.push(this.extractMathText(cell))
          }
        }
        rows.push(cells)
      }
    }
    
    return rows
  }

  /**
   * 提取数学公式文本
   */
  private extractMathText(node: any): string {
    if (!node) return ''
    
    const texts: string[] = []

    // 处理文本运行 (m:r)
    if (node['m:r']) {
      const runArray = Array.isArray(node['m:r']) ? node['m:r'] : [node['m:r']]
      for (const run of runArray) {
        if (run['m:t']) {
          const textNodes = Array.isArray(run['m:t']) ? run['m:t'] : [run['m:t']]
          for (const t of textNodes) {
            texts.push(this.getTextContent(t))
          }
        }
      }
    }

    // 递归处理子元素
    if (typeof node === 'object') {
      for (const key in node) {
        if (key.startsWith('m:') && key !== 'm:r' && typeof node[key] === 'object') {
          texts.push(this.extractMathText(node[key]))
        }
      }
    }

    return texts.join('')
  }

  /**
   * 获取文本节点内容
   */
  private getTextContent(node: any): string {
    if (typeof node === 'string') return node
    if (node['#text']) return node['#text']
    return ''
  }

  /**
   * 递归提取所有数学公式
   */
  private extractMath(node: any, mathElements: MathElement[]): void {
    if (!node || typeof node !== 'object') return

    // 处理 m:oMath
    if (node['m:oMath']) {
      const mathArray = Array.isArray(node['m:oMath'])
        ? node['m:oMath']
        : [node['m:oMath']]
      
      for (const math of mathArray) {
        const element = this.parseMath({ 'm:oMath': math })
        if (element) mathElements.push(element)
      }
    }

    // 处理 m:oMathPara
    if (node['m:oMathPara']) {
      const mathParaArray = Array.isArray(node['m:oMathPara'])
        ? node['m:oMathPara']
        : [node['m:oMathPara']]
      
      for (const mathPara of mathParaArray) {
        const element = this.parseMath({ 'm:oMathPara': mathPara })
        if (element) mathElements.push(element)
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractMath(node[key], mathElements)
      }
    }
  }
}

export const mathParser = new MathParser()
