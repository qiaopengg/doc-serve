import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface Revision {
  id: string
  type: 'insert' | 'delete' | 'moveFrom' | 'moveTo' | 'propertyChange'
  author: string
  date?: Date
  content?: string
  properties?: any
}

export class RevisionParser {
  /**
   * 解析文档中的所有修订追踪
   */
  parse(xml: string): Revision[] {
    const revisions: Revision[] = []
    const doc = xmlParser.parse(xml)
    
    this.extractRevisions(doc, revisions)
    return revisions
  }

  /**
   * 解析插入修订
   */
  parseInsert(insNode: XmlNode): Revision | null {
    if (!insNode) return null

    const id = insNode['@_w:id']
    const author = insNode['@_w:author']
    const dateStr = insNode['@_w:date']

    if (!id || !author) return null

    const revision: Revision = {
      id,
      type: 'insert',
      author,
      content: this.extractText(insNode)
    }

    if (dateStr) {
      revision.date = new Date(dateStr)
    }

    return revision
  }

  /**
   * 解析删除修订
   */
  parseDelete(delNode: XmlNode): Revision | null {
    if (!delNode) return null

    const id = delNode['@_w:id']
    const author = delNode['@_w:author']
    const dateStr = delNode['@_w:date']

    if (!id || !author) return null

    const revision: Revision = {
      id,
      type: 'delete',
      author,
      content: this.extractText(delNode)
    }

    if (dateStr) {
      revision.date = new Date(dateStr)
    }

    return revision
  }

  /**
   * 解析移动修订
   */
  parseMoveFrom(moveFromNode: XmlNode): Revision | null {
    if (!moveFromNode) return null

    const id = moveFromNode['@_w:id']
    const author = moveFromNode['@_w:author']
    const dateStr = moveFromNode['@_w:date']

    if (!id || !author) return null

    const revision: Revision = {
      id,
      type: 'moveFrom',
      author,
      content: this.extractText(moveFromNode)
    }

    if (dateStr) {
      revision.date = new Date(dateStr)
    }

    return revision
  }

  parseMoveTo(moveToNode: XmlNode): Revision | null {
    if (!moveToNode) return null

    const id = moveToNode['@_w:id']
    const author = moveToNode['@_w:author']
    const dateStr = moveToNode['@_w:date']

    if (!id || !author) return null

    const revision: Revision = {
      id,
      type: 'moveTo',
      author,
      content: this.extractText(moveToNode)
    }

    if (dateStr) {
      revision.date = new Date(dateStr)
    }

    return revision
  }

  /**
   * 解析属性更改修订
   */
  parsePropertyChange(changeNode: XmlNode): Revision | null {
    if (!changeNode) return null

    const id = changeNode['@_w:id']
    const author = changeNode['@_w:author']
    const dateStr = changeNode['@_w:date']

    if (!id || !author) return null

    const revision: Revision = {
      id,
      type: 'propertyChange',
      author,
      properties: changeNode
    }

    if (dateStr) {
      revision.date = new Date(dateStr)
    }

    return revision
  }

  /**
   * 递归提取所有修订
   */
  private extractRevisions(node: any, revisions: Revision[]): void {
    if (!node || typeof node !== 'object') return

    // 插入修订
    if (node['w:ins']) {
      const insArray = Array.isArray(node['w:ins']) ? node['w:ins'] : [node['w:ins']]
      for (const ins of insArray) {
        const revision = this.parseInsert(ins)
        if (revision) revisions.push(revision)
      }
    }

    // 删除修订
    if (node['w:del']) {
      const delArray = Array.isArray(node['w:del']) ? node['w:del'] : [node['w:del']]
      for (const del of delArray) {
        const revision = this.parseDelete(del)
        if (revision) revisions.push(revision)
      }
    }

    // 移动修订
    if (node['w:moveFrom']) {
      const moveFromArray = Array.isArray(node['w:moveFrom']) 
        ? node['w:moveFrom'] 
        : [node['w:moveFrom']]
      for (const moveFrom of moveFromArray) {
        const revision = this.parseMoveFrom(moveFrom)
        if (revision) revisions.push(revision)
      }
    }

    if (node['w:moveTo']) {
      const moveToArray = Array.isArray(node['w:moveTo']) 
        ? node['w:moveTo'] 
        : [node['w:moveTo']]
      for (const moveTo of moveToArray) {
        const revision = this.parseMoveTo(moveTo)
        if (revision) revisions.push(revision)
      }
    }

    // 属性更改
    if (node['w:rPrChange']) {
      const changeArray = Array.isArray(node['w:rPrChange']) 
        ? node['w:rPrChange'] 
        : [node['w:rPrChange']]
      for (const change of changeArray) {
        const revision = this.parsePropertyChange(change)
        if (revision) revisions.push(revision)
      }
    }

    if (node['w:pPrChange']) {
      const changeArray = Array.isArray(node['w:pPrChange']) 
        ? node['w:pPrChange'] 
        : [node['w:pPrChange']]
      for (const change of changeArray) {
        const revision = this.parsePropertyChange(change)
        if (revision) revisions.push(revision)
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractRevisions(node[key], revisions)
      }
    }
  }

  /**
   * 提取文本内容
   */
  private extractText(node: any): string {
    const texts: string[] = []

    if (node['w:r']) {
      const runs = Array.isArray(node['w:r']) ? node['w:r'] : [node['w:r']]
      for (const run of runs) {
        if (run['w:t']) {
          const textNodes = Array.isArray(run['w:t']) ? run['w:t'] : [run['w:t']]
          for (const t of textNodes) {
            texts.push(this.getTextContent(t))
          }
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
}

export const revisionParser = new RevisionParser()
