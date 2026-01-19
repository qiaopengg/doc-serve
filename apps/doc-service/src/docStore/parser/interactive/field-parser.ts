import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { FieldRun } from '../types.js'

export interface Field {
  id: string
  code: string
  result?: string
  dirty?: boolean
  locked?: boolean
  type?: string // 域类型：PAGE, DATE, TOC, REF 等
  arguments?: string[]
}

export class FieldParser {
  /**
   * 解析文档中的所有域代码
   */
  parse(xml: string): Field[] {
    const fields: Field[] = []
    const doc = xmlParser.parse(xml)
    
    this.extractFields(doc, fields)
    return fields
  }

  /**
   * 解析简单域
   */
  parseSimpleField(fldSimpleNode: XmlNode): Field | null {
    if (!fldSimpleNode) return null

    const instr = fldSimpleNode['@_w:instr']
    if (!instr) return null

    const field: Field = {
      id: this.generateFieldId(),
      code: instr,
      dirty: fldSimpleNode['@_w:dirty'] === '1',
      locked: fldSimpleNode['@_w:lock'] === '1'
    }

    // 解析域代码
    this.parseFieldCode(field)

    // 提取域结果
    const textNodes = this.extractText(fldSimpleNode)
    if (textNodes.length > 0) {
      field.result = textNodes.join('')
    }

    return field
  }

  /**
   * 解析复杂域（begin/separate/end）
   */
  parseComplexField(runs: any[]): Field | null {
    let fieldCode = ''
    let fieldResult = ''
    let inFieldCode = false
    let inFieldResult = false
    let dirty = false
    let locked = false

    for (const run of runs) {
      const fldChar = run['w:fldChar']
      
      if (fldChar) {
        const fldCharType = fldChar['@_w:fldCharType']
        
        if (fldCharType === 'begin') {
          inFieldCode = true
          dirty = fldChar['@_w:dirty'] === '1'
          locked = fldChar['@_w:lock'] === '1'
        } else if (fldCharType === 'separate') {
          inFieldCode = false
          inFieldResult = true
        } else if (fldCharType === 'end') {
          inFieldResult = false
          break
        }
      }

      const instrText = run['w:instrText']
      if (instrText && inFieldCode) {
        fieldCode += this.getTextContent(instrText)
      }

      const text = run['w:t']
      if (text && inFieldResult) {
        fieldResult += this.getTextContent(text)
      }
    }

    if (!fieldCode) return null

    const field: Field = {
      id: this.generateFieldId(),
      code: fieldCode.trim(),
      result: fieldResult || undefined,
      dirty,
      locked
    }

    this.parseFieldCode(field)
    return field
  }

  /**
   * 解析域代码，提取类型和参数
   */
  private parseFieldCode(field: Field): void {
    const code = field.code.trim()
    const parts = code.split(/\s+/)
    
    if (parts.length > 0) {
      field.type = parts[0].toUpperCase()
      field.arguments = parts.slice(1)
    }
  }

  /**
   * 递归提取所有域
   */
  private extractFields(node: any, fields: Field[]): void {
    if (!node || typeof node !== 'object') return

    // 处理简单域
    if (node['w:fldSimple']) {
      const fldSimpleArray = Array.isArray(node['w:fldSimple'])
        ? node['w:fldSimple']
        : [node['w:fldSimple']]

      for (const fldSimple of fldSimpleArray) {
        const field = this.parseSimpleField(fldSimple)
        if (field) fields.push(field)
      }
    }

    // 处理复杂域（需要在段落级别处理）
    if (node['w:p']) {
      const paragraphs = Array.isArray(node['w:p']) ? node['w:p'] : [node['w:p']]
      
      for (const para of paragraphs) {
        const runs = this.extractRuns(para)
        const field = this.parseComplexField(runs)
        if (field) fields.push(field)
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractFields(node[key], fields)
      }
    }
  }

  /**
   * 提取段落中的所有运行
   */
  private extractRuns(paragraph: XmlNode): any[] {
    const runs: any[] = []
    
    if (paragraph['w:r']) {
      const runArray = Array.isArray(paragraph['w:r'])
        ? paragraph['w:r']
        : [paragraph['w:r']]
      runs.push(...runArray)
    }

    return runs
  }

  /**
   * 提取文本内容
   */
  private extractText(node: any): string[] {
    const texts: string[] = []

    if (node['w:t']) {
      const textArray = Array.isArray(node['w:t']) ? node['w:t'] : [node['w:t']]
      for (const t of textArray) {
        texts.push(this.getTextContent(t))
      }
    }

    if (node['w:r']) {
      const runs = Array.isArray(node['w:r']) ? node['w:r'] : [node['w:r']]
      for (const run of runs) {
        texts.push(...this.extractText(run))
      }
    }

    return texts
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
   * 生成唯一域ID
   */
  private fieldIdCounter = 0
  private generateFieldId(): string {
    return `field_${++this.fieldIdCounter}`
  }
}

export const fieldParser = new FieldParser()
