import { XMLParser } from 'fast-xml-parser'
import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import { asArray } from '../core/utils.js'
import type { NumberingDefinition, NumberingLevel, NumberingFormat } from '../types.js'

export class NumberingParser {
  /**
   * 解析编号定义（完整版本）
   */
  parse(xml: string): Map<number, NumberingDefinition> {
    const doc = xmlParser.parse(xml)
    const numberingRoot = doc['w:numbering']
    if (!numberingRoot) return new Map()
    
    // 解析抽象编号定义
    const abstractNums = this.parseAbstractNums(numberingRoot)
    
    // 解析编号实例
    const numInstances = this.parseNumInstances(numberingRoot)
    
    // 合并抽象定义和实例
    const result = new Map<number, NumberingDefinition>()
    for (const [numId, abstractNumId] of numInstances.entries()) {
      const abstractNum = abstractNums.get(abstractNumId)
      if (abstractNum) {
        result.set(numId, abstractNum)
      }
    }
    
    return result
  }
  
  /**
   * 解析编号定义（简化版本，兼容旧接口）
   * 返回 Map<string, Map<number, { format?: string; text?: string }>>
   */
  parseSimple(xml: string): Map<string, Map<number, { format?: string; text?: string }>> {
    const parser = new XMLParser({ ignoreAttributes: false })
    const obj: any = parser.parse(xml)
    const numberingRoot = obj?.["w:numbering"]
    
    const abstractNums = new Map<string, Map<number, { format?: string; text?: string }>>()
    const numMap = new Map<string, string>()
    
    // 解析抽象编号定义
    for (const abstractNum of asArray<any>(numberingRoot?.["w:abstractNum"])) {
      const abstractNumId = String(abstractNum?.["@_w:abstractNumId"] ?? "")
      if (!abstractNumId) continue
      
      const levels = new Map<number, { format?: string; text?: string }>()
      
      for (const lvl of asArray<any>(abstractNum?.["w:lvl"])) {
        const ilvl = Number.parseInt(String(lvl?.["@_w:ilvl"] ?? ""), 10)
        if (!Number.isFinite(ilvl)) continue
        
        const numFmt = lvl?.["w:numFmt"]?.["@_w:val"]
        const lvlText = lvl?.["w:lvlText"]?.["@_w:val"]
        
        levels.set(ilvl, {
          format: typeof numFmt === "string" ? numFmt : undefined,
          text: typeof lvlText === "string" ? lvlText : undefined
        })
      }
      
      abstractNums.set(abstractNumId, levels)
    }
    
    // 解析编号实例
    for (const num of asArray<any>(numberingRoot?.["w:num"])) {
      const numId = String(num?.["@_w:numId"] ?? "")
      const abstractNumId = String(num?.["w:abstractNumId"]?.["@_w:val"] ?? "")
      if (numId && abstractNumId) {
        numMap.set(numId, abstractNumId)
      }
    }
    
    // 构建最终映射：numId -> levels
    const result = new Map<string, Map<number, { format?: string; text?: string }>>()
    for (const [numId, abstractNumId] of numMap.entries()) {
      const levels = abstractNums.get(abstractNumId)
      if (levels) result.set(numId, levels)
    }
    
    return result
  }
  
  /**
   * 解析抽象编号定义
   */
  private parseAbstractNums(numberingRoot: XmlNode): Map<number, NumberingDefinition> {
    const abstractNums = new Map<number, NumberingDefinition>()
    
    const abstractNumNodes = xmlParser.getChildren(numberingRoot, 'w:abstractNum')
    for (const abstractNum of abstractNumNodes) {
      const abstractNumId = xmlParser.parseInt(xmlParser.getAttr(abstractNum, 'w:abstractNumId'))
      if (abstractNumId === undefined) continue
      
      const levels = this.parseLevels(abstractNum)
      
      abstractNums.set(abstractNumId, {
        abstractNumId,
        levels
      })
    }
    
    return abstractNums
  }
  
  /**
   * 解析编号级别
   */
  private parseLevels(abstractNum: XmlNode): NumberingLevel[] {
    const levels: NumberingLevel[] = []
    
    const lvlNodes = xmlParser.getChildren(abstractNum, 'w:lvl')
    for (const lvl of lvlNodes) {
      const ilvl = xmlParser.parseInt(xmlParser.getAttr(lvl, 'w:ilvl'))
      if (ilvl === undefined) continue
      
      const level: NumberingLevel = {
        level: ilvl
      }
      
      // 起始编号
      const start = xmlParser.getChild(lvl, 'w:start')
      if (start) {
        level.start = xmlParser.parseInt(xmlParser.getAttr(start, 'w:val'))
      }
      
      // 编号格式
      const numFmt = xmlParser.getChild(lvl, 'w:numFmt')
      if (numFmt) {
        level.format = xmlParser.getAttr(numFmt, 'w:val') as NumberingFormat
      }
      
      // 编号文本
      const lvlText = xmlParser.getChild(lvl, 'w:lvlText')
      if (lvlText) {
        level.text = xmlParser.getAttr(lvlText, 'w:val')
      }
      
      // 对齐
      const lvlJc = xmlParser.getChild(lvl, 'w:lvlJc')
      if (lvlJc) {
        level.alignment = xmlParser.getAttr(lvlJc, 'w:val') as any
      }
      
      // 段落属性
      const pPr = xmlParser.getChild(lvl, 'w:pPr')
      if (pPr) {
        level.paragraphProperties = this.parseParagraphProperties(pPr)
      }
      
      // 运行属性
      const rPr = xmlParser.getChild(lvl, 'w:rPr')
      if (rPr) {
        level.runProperties = this.parseRunProperties(rPr)
      }
      
      // 重启编号
      const lvlRestart = xmlParser.getChild(lvl, 'w:lvlRestart')
      if (lvlRestart) {
        level.restart = xmlParser.parseInt(xmlParser.getAttr(lvlRestart, 'w:val'))
      }
      
      // 后缀
      const suff = xmlParser.getChild(lvl, 'w:suff')
      if (suff) {
        level.suffix = xmlParser.getAttr(suff, 'w:val') as any
      }
      
      // 法律编号
      const isLgl = xmlParser.getChild(lvl, 'w:isLgl')
      if (isLgl !== undefined) {
        const val = xmlParser.getAttr(isLgl, 'w:val')
        level.isLegal = xmlParser.parseBool(val, true)
      }
      
      levels.push(level)
    }
    
    return levels
  }
  
  /**
   * 解析编号实例
   */
  private parseNumInstances(numberingRoot: XmlNode): Map<number, number> {
    const instances = new Map<number, number>()
    
    const numNodes = xmlParser.getChildren(numberingRoot, 'w:num')
    for (const num of numNodes) {
      const numId = xmlParser.parseInt(xmlParser.getAttr(num, 'w:numId'))
      if (numId === undefined) continue
      
      const abstractNumId = xmlParser.getChild(num, 'w:abstractNumId')
      if (abstractNumId) {
        const abstractId = xmlParser.parseInt(xmlParser.getAttr(abstractNumId, 'w:val'))
        if (abstractId !== undefined) {
          instances.set(numId, abstractId)
        }
      }
    }
    
    return instances
  }
  
  private parseParagraphProperties(pPr: XmlNode): any {
    // 简化实现，只提取关键属性
    const props: any = {}
    
    const ind = xmlParser.getChild(pPr, 'w:ind')
    if (ind) {
      props.indentation = {
        left: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:left')),
        hanging: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:hanging')),
        firstLine: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:firstLine'))
      }
    }
    
    return props
  }
  
  private parseRunProperties(rPr: XmlNode): any {
    // 简化实现，只提取关键属性
    const props: any = {}
    
    const rFonts = xmlParser.getChild(rPr, 'w:rFonts')
    if (rFonts) {
      props.fonts = {
        ascii: xmlParser.getAttr(rFonts, 'w:ascii'),
        hAnsi: xmlParser.getAttr(rFonts, 'w:hAnsi')
      }
    }
    
    const sz = xmlParser.getChild(rPr, 'w:sz')
    if (sz) {
      props.fontSize = xmlParser.parseHalfPoint(xmlParser.getAttr(sz, 'w:val'))
    }
    
    const b = xmlParser.getChild(rPr, 'w:b')
    if (b !== undefined) {
      const val = xmlParser.getAttr(b, 'w:val')
      props.bold = xmlParser.parseBool(val, true)
    }
    
    return props
  }
}

/**
 * 解析 numbering.xml 获取编号格式定义（兼容旧接口）
 * 返回 Map<string, Map<number, { format?: string; text?: string }>>
 */
export function parseNumbering(xml: string): Map<string, Map<number, { format?: string; text?: string }>> {
  const parser = new NumberingParser()
  return parser.parseSimple(xml)
}
