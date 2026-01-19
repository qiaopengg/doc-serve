import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface Comment {
  id: string
  author: string
  date?: Date
  initials?: string
  content: string
  paragraphs?: any[]
}

export class CommentParser {
  /**
   * 解析注释文件（comments.xml）
   */
  parse(xml: string): Map<string, Comment> {
    const comments = new Map<string, Comment>()
    const doc = xmlParser.parse(xml)
    
    const commentsRoot = doc['w:comments']
    if (!commentsRoot) return comments

    const commentArray = commentsRoot['w:comment']
    if (!commentArray) return comments

    const commentNodes = Array.isArray(commentArray) ? commentArray : [commentArray]

    for (const commentNode of commentNodes) {
      const comment = this.parseComment(commentNode)
      if (comment) {
        comments.set(comment.id, comment)
      }
    }

    return comments
  }

  /**
   * 解析单个注释
   */
  parseComment(commentNode: XmlNode): Comment | null {
    if (!commentNode) return null

    const id = commentNode['@_w:id']
    const author = commentNode['@_w:author']
    const dateStr = commentNode['@_w:date']
    const initials = commentNode['@_w:initials']

    if (!id || !author) return null

    const comment: Comment = {
      id,
      author,
      initials,
      content: ''
    }

    // 解析日期
    if (dateStr) {
      comment.date = new Date(dateStr)
    }

    // 提取注释内容
    comment.content = this.extractCommentText(commentNode)

    // 提取段落结构（如果需要）
    if (commentNode['w:p']) {
      const paragraphs = Array.isArray(commentNode['w:p'])
        ? commentNode['w:p']
        : [commentNode['w:p']]
      comment.paragraphs = paragraphs
    }

    return comment
  }

  /**
   * 提取注释文本内容
   */
  private extractCommentText(node: any): string {
    const texts: string[] = []

    if (node['w:p']) {
      const paragraphs = Array.isArray(node['w:p']) ? node['w:p'] : [node['w:p']]
      
      for (const para of paragraphs) {
        if (para['w:r']) {
          const runs = Array.isArray(para['w:r']) ? para['w:r'] : [para['w:r']]
          
          for (const run of runs) {
            if (run['w:t']) {
              const textNodes = Array.isArray(run['w:t']) ? run['w:t'] : [run['w:t']]
              for (const t of textNodes) {
                texts.push(this.getTextContent(t))
              }
            }
          }
        }
      }
    }

    return texts.join(' ')
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

export const commentParser = new CommentParser()
