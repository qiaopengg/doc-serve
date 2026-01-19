import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface EmbeddedObject {
  id: string
  type: 'oleObject' | 'package'
  progId?: string
  shapeId?: string
  relationshipId?: string
  drawAspect?: 'content' | 'icon'
  updateMode?: 'always' | 'onCall'
}

export class ObjectParser {
  /**
   * 解析文档中的所有嵌入对象
   */
  parse(xml: string): Map<string, EmbeddedObject> {
    const objects = new Map<string, EmbeddedObject>()
    const doc = xmlParser.parse(xml)
    
    this.extractObjects(doc, objects)
    return objects
  }

  /**
   * 解析单个嵌入对象
   */
  parseObject(objectNode: XmlNode): EmbeddedObject | null {
    if (!objectNode) return null

    const oleObject = objectNode['o:OLEObject']
    if (!oleObject) return null

    const type = oleObject['@_Type'] || 'oleObject'
    const progId = oleObject['@_ProgID']
    const shapeId = oleObject['@_ShapeID']
    const relationshipId = oleObject['@_r:id']
    const drawAspect = oleObject['@_DrawAspect']
    const updateMode = oleObject['@_UpdateMode']

    const obj: EmbeddedObject = {
      id: shapeId || this.generateId(),
      type: type === 'Embed' ? 'oleObject' : 'package',
      progId,
      shapeId,
      relationshipId,
      drawAspect: drawAspect as any,
      updateMode: updateMode as any
    }

    return obj
  }

  /**
   * 递归提取所有嵌入对象
   */
  private extractObjects(node: any, objects: Map<string, EmbeddedObject>): void {
    if (!node || typeof node !== 'object') return

    // 处理 w:object
    if (node['w:object']) {
      const objectArray = Array.isArray(node['w:object'])
        ? node['w:object']
        : [node['w:object']]
      
      for (const obj of objectArray) {
        const parsed = this.parseObject(obj)
        if (parsed) {
          objects.set(parsed.id, parsed)
        }
      }
    }

    // 处理 o:OLEObject（直接在文档中）
    if (node['o:OLEObject']) {
      const oleArray = Array.isArray(node['o:OLEObject'])
        ? node['o:OLEObject']
        : [node['o:OLEObject']]
      
      for (const ole of oleArray) {
        const parsed = this.parseObject({ 'o:OLEObject': ole })
        if (parsed) {
          objects.set(parsed.id, parsed)
        }
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractObjects(node[key], objects)
      }
    }
  }

  /**
   * 生成唯一 ID
   */
  private idCounter = 0
  private generateId(): string {
    return `object_${++this.idCounter}`
  }
}

export const objectParser = new ObjectParser()
