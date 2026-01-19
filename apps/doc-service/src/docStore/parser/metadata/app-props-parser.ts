import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface AppProperties {
  application?: string
  appVersion?: string
  totalTime?: number
  pages?: number
  words?: number
  characters?: number
  paragraphs?: number
  lines?: number
  company?: string
  manager?: string
}

export class AppPropsParser {
  parse(xml: string): AppProperties {
    const doc = xmlParser.parse(xml)
    const props = doc['Properties'] || doc['ap:Properties']
    if (!props) return {}
    
    const appProps: AppProperties = {}
    
    // 应用程序
    if (props['Application']) {
      appProps.application = String(props['Application'])
    }
    
    // 应用版本
    if (props['AppVersion']) {
      appProps.appVersion = String(props['AppVersion'])
    }
    
    // 总编辑时间
    if (props['TotalTime']) {
      appProps.totalTime = this.parseNumber(props['TotalTime'])
    }
    
    // 页数
    if (props['Pages']) {
      appProps.pages = this.parseNumber(props['Pages'])
    }
    
    // 字数
    if (props['Words']) {
      appProps.words = this.parseNumber(props['Words'])
    }
    
    // 字符数
    if (props['Characters']) {
      appProps.characters = this.parseNumber(props['Characters'])
    }
    
    // 段落数
    if (props['Paragraphs']) {
      appProps.paragraphs = this.parseNumber(props['Paragraphs'])
    }
    
    // 行数
    if (props['Lines']) {
      appProps.lines = this.parseNumber(props['Lines'])
    }
    
    // 公司
    if (props['Company']) {
      appProps.company = String(props['Company'])
    }
    
    // 经理
    if (props['Manager']) {
      appProps.manager = String(props['Manager'])
    }
    
    return appProps
  }
  
  private parseNumber(value: any): number | undefined {
    const num = Number(value)
    return isNaN(num) ? undefined : num
  }
}
