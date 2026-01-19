import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface ContentControl {
  id: string
  tag?: string
  alias?: string
  lock?: 'unlocked' | 'sdtLocked' | 'contentLocked' | 'sdtContentLocked'
  type?: 'richText' | 'plainText' | 'picture' | 'comboBox' | 'dropDownList' | 'date' | 'checkbox'
  content?: any[]
  properties?: any
}

export class SdtParser {
  /**
   * 解析文档中的所有内容控件（Structured Document Tags）
   */
  parse(xml: string): ContentControl[] {
    const contentControls: ContentControl[] = []
    const doc = xmlParser.parse(xml)
    
    this.extractContentControls(doc, contentControls)
    return contentControls
  }

  /**
   * 解析单个内容控件
   */
  parseContentControl(sdtNode: XmlNode): ContentControl | null {
    if (!sdtNode) return null

    const sdtPr = sdtNode['w:sdtPr']
    const sdtContent = sdtNode['w:sdtContent']

    if (!sdtPr) return null

    const id = sdtPr['w:id']?.['@_w:val'] || this.generateId()
    const tag = sdtPr['w:tag']?.['@_w:val']
    const alias = sdtPr['w:alias']?.['@_w:val']
    const lock = sdtPr['w:lock']?.['@_w:val']

    // 确定内容控件类型
    let type: ContentControl['type'] = 'richText'
    if (sdtPr['w:text']) {
      type = 'plainText'
    } else if (sdtPr['w:picture']) {
      type = 'picture'
    } else if (sdtPr['w:comboBox']) {
      type = 'comboBox'
    } else if (sdtPr['w:dropDownList']) {
      type = 'dropDownList'
    } else if (sdtPr['w:date']) {
      type = 'date'
    } else if (sdtPr['w:checkbox']) {
      type = 'checkbox'
    }

    const contentControl: ContentControl = {
      id,
      tag,
      alias,
      lock: lock as any,
      type,
      properties: this.parseProperties(sdtPr),
      content: []
    }

    // 提取内容
    if (sdtContent) {
      contentControl.content = this.extractContent(sdtContent)
    }

    return contentControl
  }

  /**
   * 解析内容控件属性
   */
  private parseProperties(sdtPr: XmlNode): any {
    const props: any = {}

    // 下拉列表项
    if (sdtPr['w:dropDownList']) {
      const listItems: any[] = []
      const listItem = sdtPr['w:dropDownList']['w:listItem']
      if (listItem) {
        const items = Array.isArray(listItem) ? listItem : [listItem]
        for (const item of items) {
          listItems.push({
            displayText: item['@_w:displayText'],
            value: item['@_w:value']
          })
        }
      }
      props.listItems = listItems
    }

    // 日期格式
    if (sdtPr['w:date']) {
      props.dateFormat = sdtPr['w:date']['w:dateFormat']?.['@_w:val']
      props.calendar = sdtPr['w:date']['w:calendar']?.['@_w:val']
    }

    // 复选框
    if (sdtPr['w:checkbox']) {
      props.checked = sdtPr['w:checkbox']['w:checked']?.['@_w:val'] === '1'
      props.checkedState = sdtPr['w:checkbox']['w:checkedState']?.['@_w:val']
      props.uncheckedState = sdtPr['w:checkbox']['w:uncheckedState']?.['@_w:val']
    }

    // 占位符
    if (sdtPr['w:placeholder']) {
      props.placeholder = sdtPr['w:placeholder']['w:docPart']?.['@_w:val']
    }

    // 数据绑定
    if (sdtPr['w:dataBinding']) {
      props.dataBinding = {
        xpath: sdtPr['w:dataBinding']['@_w:xpath'],
        storeItemID: sdtPr['w:dataBinding']['@_w:storeItemID']
      }
    }

    return props
  }

  /**
   * 提取内容控件的内容
   */
  private extractContent(sdtContent: XmlNode): any[] {
    const content: any[] = []

    // 段落
    if (sdtContent['w:p']) {
      const pArray = Array.isArray(sdtContent['w:p']) 
        ? sdtContent['w:p'] 
        : [sdtContent['w:p']]
      content.push(...pArray)
    }

    // 表格
    if (sdtContent['w:tbl']) {
      const tblArray = Array.isArray(sdtContent['w:tbl'])
        ? sdtContent['w:tbl']
        : [sdtContent['w:tbl']]
      content.push(...tblArray)
    }

    // 运行
    if (sdtContent['w:r']) {
      const rArray = Array.isArray(sdtContent['w:r'])
        ? sdtContent['w:r']
        : [sdtContent['w:r']]
      content.push(...rArray)
    }

    return content
  }

  /**
   * 递归提取所有内容控件
   */
  private extractContentControls(node: any, contentControls: ContentControl[]): void {
    if (!node || typeof node !== 'object') return

    if (node['w:sdt']) {
      const sdtArray = Array.isArray(node['w:sdt']) ? node['w:sdt'] : [node['w:sdt']]
      for (const sdt of sdtArray) {
        const cc = this.parseContentControl(sdt)
        if (cc) contentControls.push(cc)
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractContentControls(node[key], contentControls)
      }
    }
  }

  /**
   * 生成唯一ID
   */
  private idCounter = 0
  private generateId(): string {
    return `sdt_${++this.idCounter}`
  }
}

export const sdtParser = new SdtParser()
