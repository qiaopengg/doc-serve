/**
 * 关系解析器
 * 解析 .rels 文件，建立文件间的关系映射
 */

import { xmlParser, type XmlNode } from './xml-parser.js'

export interface Relationship {
  id: string
  type: string
  target: string
  targetMode?: 'Internal' | 'External'
}

export class RelationshipParser {
  private readonly ALLOWED_PROTOCOLS = ['http:', 'https:', 'mailto:', 'ftp:']
  
  /**
   * 验证外部 URL 安全性
   */
  private isValidExternalUrl(url: string): boolean {
    try {
      const parsed = new URL(url)
      
      // 只允许安全协议
      if (!this.ALLOWED_PROTOCOLS.includes(parsed.protocol)) {
        return false
      }
      
      // 阻止 javascript: 和 data: 协议
      if (parsed.protocol === 'javascript:' || parsed.protocol === 'data:') {
        return false
      }
      
      return true
    } catch {
      return false
    }
  }
  
  /**
   * 解析关系文件
   */
  parse(xml: string): Map<string, Relationship> {
    const doc = xmlParser.parse(xml)
    const relationships = new Map<string, Relationship>()
    
    const rels = doc['Relationships']?.['Relationship']
    if (!rels) return relationships
    
    const relArray = Array.isArray(rels) ? rels : [rels]
    
    for (const rel of relArray) {
      const id = xmlParser.getAttr(rel, 'Id')
      const type = xmlParser.getAttr(rel, 'Type')
      const target = xmlParser.getAttr(rel, 'Target')
      const targetMode = xmlParser.getAttr(rel, 'TargetMode') as 'Internal' | 'External' | undefined
      
      if (id && type && target) {
        // 验证外部链接
        if (targetMode === 'External' && !this.isValidExternalUrl(target)) {
          console.warn(`Blocked unsafe external URL in relationship ${id}: ${target}`)
          continue
        }
        
        relationships.set(id, { id, type, target, targetMode })
      }
    }
    
    return relationships
  }
  
  /**
   * 根据类型筛选关系
   */
  filterByType(relationships: Map<string, Relationship>, typePattern: string): Relationship[] {
    const result: Relationship[] = []
    for (const rel of relationships.values()) {
      if (rel.type.includes(typePattern)) {
        result.push(rel)
      }
    }
    return result
  }
  
  /**
   * 获取超链接关系
   */
  getHyperlinks(relationships: Map<string, Relationship>): Map<string, string> {
    const hyperlinks = new Map<string, string>()
    for (const [id, rel] of relationships.entries()) {
      if (rel.type.includes('/hyperlink')) {
        hyperlinks.set(id, rel.target)
      }
    }
    return hyperlinks
  }
  
  /**
   * 获取图片关系
   */
  getImages(relationships: Map<string, Relationship>): Map<string, string> {
    const images = new Map<string, string>()
    for (const [id, rel] of relationships.entries()) {
      if (rel.type.includes('/image')) {
        images.set(id, rel.target)
      }
    }
    return images
  }
}
