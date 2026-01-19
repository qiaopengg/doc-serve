import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface Bookmark {
  id: string
  name: string
  startPosition?: number
  endPosition?: number
  colFirst?: number
  colLast?: number
}

export class BookmarkParser {
  /**
   * 解析文档中的所有书签
   */
  parse(xml: string): Map<string, Bookmark> {
    const bookmarks = new Map<string, Bookmark>()
    const doc = xmlParser.parse(xml)
    
    this.extractBookmarks(doc, bookmarks)
    return bookmarks
  }

  /**
   * 解析书签开始标记
   */
  parseBookmarkStart(bookmarkStartNode: XmlNode): Partial<Bookmark> | null {
    if (!bookmarkStartNode) return null

    const id = bookmarkStartNode['@_w:id']
    const name = bookmarkStartNode['@_w:name']
    const colFirst = bookmarkStartNode['@_w:colFirst']
    const colLast = bookmarkStartNode['@_w:colLast']

    if (!id || !name) return null

    return {
      id,
      name,
      colFirst: colFirst ? parseInt(colFirst) : undefined,
      colLast: colLast ? parseInt(colLast) : undefined
    }
  }

  /**
   * 解析书签结束标记
   */
  parseBookmarkEnd(bookmarkEndNode: XmlNode): string | null {
    if (!bookmarkEndNode) return null
    return bookmarkEndNode['@_w:id']
  }

  /**
   * 递归提取所有书签
   */
  private extractBookmarks(node: any, bookmarks: Map<string, Bookmark>): void {
    if (!node || typeof node !== 'object') return

    // 处理书签开始
    if (node['w:bookmarkStart']) {
      const startArray = Array.isArray(node['w:bookmarkStart'])
        ? node['w:bookmarkStart']
        : [node['w:bookmarkStart']]

      for (const start of startArray) {
        const bookmark = this.parseBookmarkStart(start)
        if (bookmark && bookmark.id) {
          bookmarks.set(bookmark.id, bookmark as Bookmark)
        }
      }
    }

    // 处理书签结束
    if (node['w:bookmarkEnd']) {
      const endArray = Array.isArray(node['w:bookmarkEnd'])
        ? node['w:bookmarkEnd']
        : [node['w:bookmarkEnd']]

      for (const end of endArray) {
        const id = this.parseBookmarkEnd(end)
        if (id && bookmarks.has(id)) {
          // 书签结束标记已找到，可以记录位置信息
          const bookmark = bookmarks.get(id)!
          bookmark.endPosition = this.getCurrentPosition()
        }
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractBookmarks(node[key], bookmarks)
      }
    }
  }

  /**
   * 获取当前位置（简化实现）
   */
  private positionCounter = 0
  private getCurrentPosition(): number {
    return ++this.positionCounter
  }
}

export const bookmarkParser = new BookmarkParser()
