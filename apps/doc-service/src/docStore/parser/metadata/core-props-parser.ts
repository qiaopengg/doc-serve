import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface CoreProperties {
  title?: string
  subject?: string
  creator?: string
  keywords?: string
  description?: string
  lastModifiedBy?: string
  revision?: string
  created?: Date
  modified?: Date
  category?: string
  contentStatus?: string
}

export class CorePropsParser {
  parse(xml: string): CoreProperties {
    const doc = xmlParser.parse(xml)
    const coreProps = doc['cp:coreProperties'] || doc['coreProperties']
    if (!coreProps) return {}
    
    const props: CoreProperties = {}
    
    // 标题
    const title = coreProps['dc:title']
    if (title) {
      props.title = String(title)
    }
    
    // 主题
    const subject = coreProps['dc:subject']
    if (subject) {
      props.subject = String(subject)
    }
    
    // 作者
    const creator = coreProps['dc:creator']
    if (creator) {
      props.creator = String(creator)
    }
    
    // 关键词
    const keywords = coreProps['cp:keywords']
    if (keywords) {
      props.keywords = String(keywords)
    }
    
    // 描述
    const description = coreProps['dc:description']
    if (description) {
      props.description = String(description)
    }
    
    // 最后修改者
    const lastModifiedBy = coreProps['cp:lastModifiedBy']
    if (lastModifiedBy) {
      props.lastModifiedBy = String(lastModifiedBy)
    }
    
    // 修订号
    const revision = coreProps['cp:revision']
    if (revision) {
      props.revision = String(revision)
    }
    
    // 创建时间
    const created = coreProps['dcterms:created']
    if (created) {
      props.created = this.parseDate(created)
    }
    
    // 修改时间
    const modified = coreProps['dcterms:modified']
    if (modified) {
      props.modified = this.parseDate(modified)
    }
    
    // 类别
    const category = coreProps['cp:category']
    if (category) {
      props.category = String(category)
    }
    
    // 内容状态
    const contentStatus = coreProps['cp:contentStatus']
    if (contentStatus) {
      props.contentStatus = String(contentStatus)
    }
    
    return props
  }
  
  private parseDate(value: any): Date | undefined {
    if (!value) return undefined
    const d = new Date(String(value))
    return isNaN(d.getTime()) ? undefined : d
  }
}
