import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface DocumentProtection {
  edit?: 'none' | 'readOnly' | 'comments' | 'trackedChanges' | 'forms'
  enforcement?: boolean
  formatting?: boolean
  cryptProviderType?: string
  cryptAlgorithmClass?: string
  cryptAlgorithmType?: string
  cryptAlgorithmSid?: number
  cryptSpinCount?: number
  hash?: string
  salt?: string
}

export interface WriteProtection {
  recommended?: boolean
  cryptProviderType?: string
  cryptAlgorithmClass?: string
  cryptAlgorithmType?: string
  cryptAlgorithmSid?: number
  cryptSpinCount?: number
  hash?: string
  salt?: string
}

export class ProtectionParser {
  /**
   * 解析文档保护设置
   */
  parse(xml: string): { documentProtection?: DocumentProtection; writeProtection?: WriteProtection } | null {
    const doc = xmlParser.parse(xml)
    
    const settings = doc['w:settings']
    if (!settings) return null

    const result: any = {}

    // 文档保护
    const docProtection = settings['w:documentProtection']
    if (docProtection) {
      result.documentProtection = this.parseDocumentProtection(docProtection)
    }

    // 写保护
    const writeProtection = settings['w:writeProtection']
    if (writeProtection) {
      result.writeProtection = this.parseWriteProtection(writeProtection)
    }

    return Object.keys(result).length > 0 ? result : null
  }

  /**
   * 解析文档保护
   */
  private parseDocumentProtection(node: XmlNode): DocumentProtection {
    const protection: DocumentProtection = {}

    const edit = node['@_w:edit']
    if (edit) {
      protection.edit = edit as any
    }

    const enforcement = node['@_w:enforcement']
    if (enforcement) {
      protection.enforcement = enforcement === '1' || enforcement === 'true'
    }

    const formatting = node['@_w:formatting']
    if (formatting) {
      protection.formatting = formatting === '1' || formatting === 'true'
    }

    // 加密相关
    protection.cryptProviderType = node['@_w:cryptProviderType']
    protection.cryptAlgorithmClass = node['@_w:cryptAlgorithmClass']
    protection.cryptAlgorithmType = node['@_w:cryptAlgorithmType']
    
    const cryptAlgorithmSid = node['@_w:cryptAlgorithmSid']
    if (cryptAlgorithmSid) {
      protection.cryptAlgorithmSid = parseInt(cryptAlgorithmSid)
    }

    const cryptSpinCount = node['@_w:cryptSpinCount']
    if (cryptSpinCount) {
      protection.cryptSpinCount = parseInt(cryptSpinCount)
    }

    protection.hash = node['@_w:hash']
    protection.salt = node['@_w:salt']

    return protection
  }

  /**
   * 解析写保护
   */
  private parseWriteProtection(node: XmlNode): WriteProtection {
    const protection: WriteProtection = {}

    const recommended = node['@_w:recommended']
    if (recommended) {
      protection.recommended = recommended === '1' || recommended === 'true'
    }

    // 加密相关
    protection.cryptProviderType = node['@_w:cryptProviderType']
    protection.cryptAlgorithmClass = node['@_w:cryptAlgorithmClass']
    protection.cryptAlgorithmType = node['@_w:cryptAlgorithmType']
    
    const cryptAlgorithmSid = node['@_w:cryptAlgorithmSid']
    if (cryptAlgorithmSid) {
      protection.cryptAlgorithmSid = parseInt(cryptAlgorithmSid)
    }

    const cryptSpinCount = node['@_w:cryptSpinCount']
    if (cryptSpinCount) {
      protection.cryptSpinCount = parseInt(cryptSpinCount)
    }

    protection.hash = node['@_w:hash']
    protection.salt = node['@_w:salt']

    return protection
  }
}

export const protectionParser = new ProtectionParser()
