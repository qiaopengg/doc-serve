/**
 * SmartArt 解析器
 * 解析 SmartArt 图形
 */

import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface SmartArt {
  id: string
  type: string
  layout: string
  colorScheme?: string
  style?: string
  nodes: SmartArtNode[]
  connections?: SmartArtConnection[]
}

export interface SmartArtNode {
  id: string
  type: 'node' | 'assistant' | 'shape'
  text: string
  level: number
  children?: SmartArtNode[]
  style?: SmartArtNodeStyle
}

export interface SmartArtNodeStyle {
  fill?: string
  line?: string
  font?: string
  fontSize?: number
}

export interface SmartArtConnection {
  from: string
  to: string
  type?: 'line' | 'arrow' | 'curve'
}

export class SmartArtParser {
  private idCounter = 0
  
  /**
   * 解析文档中的所有 SmartArt
   */
  parse(xml: string): Map<string, SmartArt> {
    const smartArts = new Map<string, SmartArt>()
    const doc = xmlParser.parse(xml)
    
    this.extractSmartArts(doc, smartArts)
    return smartArts
  }
  
  /**
   * 解析单个 SmartArt
   */
  parseSmartArt(smartArtNode: XmlNode, dataXml?: string, layoutXml?: string): SmartArt | null {
    if (!smartArtNode) return null
    
    const smartArt: SmartArt = {
      id: this.generateId(),
      type: 'unknown',
      layout: 'unknown',
      nodes: []
    }
    
    // 如果有独立的数据和布局 XML，解析它们
    if (dataXml) {
      const dataModel = this.parseDataModel(dataXml)
      if (dataModel) {
        smartArt.nodes = dataModel.nodes
      }
    }
    
    if (layoutXml) {
      const layoutInfo = this.parseLayout(layoutXml)
      if (layoutInfo) {
        smartArt.layout = layoutInfo.type
      }
    }
    
    // 从节点解析基本信息
    smartArt.type = this.detectSmartArtType(smartArtNode)
    
    return smartArt
  }
  
  /**
   * 解析 SmartArt 数据模型
   */
  private parseDataModel(xml: string): { nodes: SmartArtNode[] } | null {
    try {
      const doc = xmlParser.parse(xml)
      const dataModel = doc['dgm:dataModel']
      if (!dataModel) return null
      
      const ptLst = dataModel['dgm:ptLst']
      if (!ptLst) return null
      
      const nodes: SmartArtNode[] = []
      const pts = xmlParser.getChildren(ptLst, 'dgm:pt')
      
      for (const pt of pts) {
        const node = this.parseDataPoint(pt)
        if (node) {
          nodes.push(node)
        }
      }
      
      // 构建层次结构
      this.buildHierarchy(nodes, ptLst)
      
      return { nodes }
    } catch (error) {
      console.error('Failed to parse SmartArt data model:', error)
      return null
    }
  }
  
  /**
   * 解析数据点
   */
  private parseDataPoint(ptNode: XmlNode): SmartArtNode | null {
    const modelId = xmlParser.getAttr(ptNode, 'modelId')
    if (!modelId) return null
    
    const type = xmlParser.getAttr(ptNode, 'type') || 'node'
    
    const node: SmartArtNode = {
      id: modelId,
      type: type as any,
      text: '',
      level: 0
    }
    
    // 解析文本
    const prSet = ptNode['dgm:prSet']
    if (prSet) {
      const t = prSet['dgm:t']
      if (t) {
        const val = t['@_val']
        if (val) {
          node.text = String(val)
        }
      }
    }
    
    return node
  }
  
  /**
   * 构建层次结构
   */
  private buildHierarchy(nodes: SmartArtNode[], ptLst: XmlNode): void {
    const cxnLst = ptLst['dgm:cxnLst']
    if (!cxnLst) return
    
    const cxns = xmlParser.getChildren(cxnLst, 'dgm:cxn')
    const nodeMap = new Map<string, SmartArtNode>()
    
    // 创建节点映射
    for (const node of nodes) {
      nodeMap.set(node.id, node)
    }
    
    // 建立父子关系
    for (const cxn of cxns) {
      const type = xmlParser.getAttr(cxn, 'type')
      if (type !== 'parOf') continue
      
      const srcId = xmlParser.getAttr(cxn, 'srcId')
      const destId = xmlParser.getAttr(cxn, 'destId')
      
      if (srcId && destId) {
        const child = nodeMap.get(srcId)
        const parent = nodeMap.get(destId)
        
        if (child && parent) {
          if (!parent.children) {
            parent.children = []
          }
          parent.children.push(child)
          child.level = parent.level + 1
        }
      }
    }
  }
  
  /**
   * 解析布局
   */
  private parseLayout(xml: string): { type: string } | null {
    try {
      const doc = xmlParser.parse(xml)
      const layoutDef = doc['dgm:layoutDef']
      if (!layoutDef) return null
      
      const uniqueId = xmlParser.getAttr(layoutDef, 'uniqueId')
      
      return {
        type: uniqueId || 'unknown'
      }
    } catch (error) {
      console.error('Failed to parse SmartArt layout:', error)
      return null
    }
  }
  
  /**
   * 检测 SmartArt 类型
   */
  private detectSmartArtType(smartArtNode: XmlNode): string {
    // 从节点属性推断类型
    // 这是简化实现
    return 'diagram'
  }
  
  /**
   * 递归提取所有 SmartArt
   */
  private extractSmartArts(node: any, smartArts: Map<string, SmartArt>): void {
    if (!node || typeof node !== 'object') return
    
    // 处理 SmartArt 节点
    // SmartArt 通常在 drawing 中，作为 diagram 类型
    if (node['dgm:relIds']) {
      const smartArt = this.parseSmartArt(node)
      if (smartArt) {
        smartArts.set(smartArt.id, smartArt)
      }
    }
    
    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractSmartArts(node[key], smartArts)
      }
    }
  }
  
  /**
   * 生成唯一 ID
   */
  private generateId(): string {
    return `smartart_${++this.idCounter}`
  }
}

export const smartArtParser = new SmartArtParser()
