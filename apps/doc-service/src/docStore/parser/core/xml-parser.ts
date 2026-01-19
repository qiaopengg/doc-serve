/**
 * XML 解析基础设施
 * 提供统一的 XML 解析接口和工具函数
 */

import { XMLParser, XMLBuilder } from 'fast-xml-parser'

export interface XmlNode {
  [key: string]: any
}

export interface XmlAttributes {
  [key: string]: string | number | boolean
}

export class XmlParserCore {
  private parser: XMLParser
  private builder: XMLBuilder
  
  // 安全配置常量
  private readonly MAX_XML_SIZE = 50 * 1024 * 1024 // 50MB
  private readonly MAX_NESTING_DEPTH = 100
  private readonly MAX_ENTITY_EXPANSION = 1000

  constructor() {
    this.parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      allowBooleanAttributes: true,
      parseAttributeValue: false,
      trimValues: true,
      // 安全配置
      processEntities: false,              // 禁用实体处理，防止 XXE 攻击
      stopNodes: ['*.script', '*.style'],  // 阻止脚本节点
    })

    this.builder = new XMLBuilder({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      format: false
    })
  }

  parse(xml: string): XmlNode {
    // 安全检查 1: 文件大小
    if (xml.length > this.MAX_XML_SIZE) {
      throw new Error(`XML file too large: ${xml.length} bytes (max: ${this.MAX_XML_SIZE})`)
    }
    
    // 安全检查 2: 嵌套深度
    const depth = this.calculateNestingDepth(xml)
    if (depth > this.MAX_NESTING_DEPTH) {
      throw new Error(`XML nesting too deep: ${depth} levels (max: ${this.MAX_NESTING_DEPTH})`)
    }
    
    // 安全检查 3: 实体扩展
    if (this.hasEntityExpansion(xml)) {
      throw new Error('XML entity expansion detected - possible XML bomb attack')
    }
    
    return this.parser.parse(xml)
  }

  build(obj: XmlNode): string {
    return this.builder.build(obj)
  }

  /**
   * 获取节点的属性值
   */
  getAttr(node: XmlNode, attrName: string): string | undefined {
    const key = `@_${attrName}`
    return node?.[key] as string | undefined
  }

  /**
   * 获取子节点（单个）
   */
  getChild(node: XmlNode, childName: string): XmlNode | undefined {
    return node?.[childName]
  }

  /**
   * 获取子节点（数组）
   */
  getChildren(node: XmlNode, childName: string): XmlNode[] {
    const child = node?.[childName]
    if (!child) return []
    return Array.isArray(child) ? child : [child]
  }

  /**
   * 获取所有子节点
   */
  getAllChildren(node: XmlNode): XmlNode[] {
    const children: XmlNode[] = []
    for (const key in node) {
      if (key.startsWith('@_')) continue
      const value = node[key]
      if (Array.isArray(value)) {
        children.push(...value)
      } else if (typeof value === 'object') {
        children.push(value)
      }
    }
    return children
  }

  /**
   * 解析布尔值（on/off, true/false, 1/0）
   */
  parseBool(value: any, defaultValue?: boolean): boolean | undefined {
    if (value === undefined || value === null) {
      return defaultValue
    }
    const str = String(value).toLowerCase().trim()
    if (str === '0' || str === 'false' || str === 'off' || str === 'none') {
      return false
    }
    if (str === '1' || str === 'true' || str === 'on') {
      return true
    }
    return defaultValue !== undefined ? defaultValue : true
  }

  /**
   * 解析整数
   */
  parseInt(value: any): number | undefined {
    const num = Number.parseInt(String(value ?? ''), 10)
    return Number.isFinite(num) ? num : undefined
  }

  /**
   * 解析十六进制颜色
   */
  parseColor(value: any): string | undefined {
    if (!value) return undefined
    const str = String(value).trim().replace(/^#/, '').toUpperCase()
    if (str === 'AUTO' || str === 'NONE') return undefined
    if (/^[0-9A-F]{6}$/.test(str)) return str
    return undefined
  }

  /**
   * 解析 Twip 值（1/20 point）
   */
  parseTwip(value: any): number | undefined {
    return this.parseInt(value)
  }

  /**
   * 解析半点值（字号）
   */
  parseHalfPoint(value: any): number | undefined {
    const raw = this.parseInt(value)
    return raw !== undefined ? raw / 2 : undefined
  }
  
  /**
   * 计算 XML 嵌套深度（防止 XML 炸弹）
   */
  private calculateNestingDepth(xml: string): number {
    let maxDepth = 0
    let currentDepth = 0
    let inTag = false
    let isClosingTag = false
    
    for (let i = 0; i < xml.length; i++) {
      const char = xml[i]
      
      if (char === '<') {
        inTag = true
        isClosingTag = xml[i + 1] === '/'
      } else if (char === '>' && inTag) {
        if (!isClosingTag && xml[i - 1] !== '/') {
          currentDepth++
          maxDepth = Math.max(maxDepth, currentDepth)
        } else if (isClosingTag) {
          currentDepth--
        }
        inTag = false
        isClosingTag = false
      }
    }
    
    return maxDepth
  }
  
  /**
   * 检测 XML 实体扩展（防止 Billion Laughs 攻击）
   */
  private hasEntityExpansion(xml: string): boolean {
    // 检查 DOCTYPE 声明
    if (xml.includes('<!DOCTYPE') || xml.includes('<!ENTITY')) {
      return true
    }
    
    // 检查实体引用
    const entityPattern = /&[a-zA-Z0-9]+;/g
    const matches = xml.match(entityPattern)
    if (matches && matches.length > this.MAX_ENTITY_EXPANSION) {
      return true
    }
    
    return false
  }
}

// 单例导出
export const xmlParser = new XmlParserCore()
