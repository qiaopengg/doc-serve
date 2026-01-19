import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface Hyperlink {
  id: string
  relationshipId?: string
  anchor?: string
  tooltip?: string
  text: string
  history?: boolean
}

export class HyperlinkParser {
  private readonly ALLOWED_PROTOCOLS = ['http:', 'https:', 'mailto:', 'ftp:']
  private readonly BLOCKED_HOSTS = ['localhost', '127.0.0.1', '0.0.0.0', '::1']
  
  /**
   * 验证 URL 安全性
   */
  private isValidUrl(url: string): boolean {
    try {
      const parsed = new URL(url)
      
      // 只允许安全协议
      if (!this.ALLOWED_PROTOCOLS.includes(parsed.protocol)) {
        console.warn(`Blocked unsafe protocol: ${parsed.protocol}`)
        return false
      }
      
      // 阻止本地文件访问
      if (this.BLOCKED_HOSTS.includes(parsed.hostname.toLowerCase())) {
        console.warn(`Blocked local host access: ${parsed.hostname}`)
        return false
      }
      
      // 阻止内网 IP
      if (this.isPrivateIP(parsed.hostname)) {
        console.warn(`Blocked private IP: ${parsed.hostname}`)
        return false
      }
      
      return true
    } catch {
      // 可能是相对 URL 或锚点，允许通过
      return true
    }
  }
  
  /**
   * 检查是否为内网 IP
   */
  private isPrivateIP(hostname: string): boolean {
    // 检查 IPv4 内网地址
    const ipv4Pattern = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/
    const match = hostname.match(ipv4Pattern)
    
    if (match) {
      const [, a, b, c, d] = match.map(Number)
      
      // 10.0.0.0/8
      if (a === 10) return true
      
      // 172.16.0.0/12
      if (a === 172 && b >= 16 && b <= 31) return true
      
      // 192.168.0.0/16
      if (a === 192 && b === 168) return true
      
      // 169.254.0.0/16 (链路本地)
      if (a === 169 && b === 254) return true
    }
    
    return false
  }
  
  /**
   * 解析文档中的所有超链接
   */
  parse(xml: string): Map<string, Hyperlink> {
    const hyperlinks = new Map<string, Hyperlink>()
    const doc = xmlParser.parse(xml)
    
    this.extractHyperlinks(doc, hyperlinks)
    return hyperlinks
  }

  /**
   * 解析单个超链接节点
   */
  parseHyperlink(hyperlinkNode: XmlNode): Hyperlink | null {
    if (!hyperlinkNode) return null

    const id = hyperlinkNode['@_w:id'] || this.generateHyperlinkId()
    const relationshipId = hyperlinkNode['@_r:id']
    const anchor = hyperlinkNode['@_w:anchor']
    const tooltip = hyperlinkNode['@_w:tooltip']
    const history = hyperlinkNode['@_w:history'] === '1'

    // 提取超链接文本
    const text = this.extractText(hyperlinkNode)

    const hyperlink: Hyperlink = {
      id,
      relationshipId,
      anchor,
      tooltip,
      text,
      history
    }
    
    // 验证 anchor（如果是完整 URL）
    if (anchor && anchor.includes('://')) {
      if (!this.isValidUrl(anchor)) {
        console.warn(`Blocked unsafe anchor URL: ${anchor}`)
        return null
      }
    }

    return hyperlink
  }

  /**
   * 递归提取所有超链接
   */
  private extractHyperlinks(node: any, hyperlinks: Map<string, Hyperlink>): void {
    if (!node || typeof node !== 'object') return

    if (node['w:hyperlink']) {
      const hyperlinkArray = Array.isArray(node['w:hyperlink'])
        ? node['w:hyperlink']
        : [node['w:hyperlink']]

      for (const hyperlink of hyperlinkArray) {
        const parsed = this.parseHyperlink(hyperlink)
        if (parsed) {
          hyperlinks.set(parsed.id, parsed)
        }
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractHyperlinks(node[key], hyperlinks)
      }
    }
  }

  /**
   * 提取超链接文本
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

  /**
   * 生成唯一超链接ID
   */
  private hyperlinkIdCounter = 0
  private generateHyperlinkId(): string {
    return `hyperlink_${++this.hyperlinkIdCounter}`
  }
}

export const hyperlinkParser = new HyperlinkParser()
