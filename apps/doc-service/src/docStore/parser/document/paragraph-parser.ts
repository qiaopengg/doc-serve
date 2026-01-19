import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { Paragraph, ParagraphProperties, Indentation, Spacing, ParagraphBorders, Tab, NumberingReference, FrameProperties } from '../types.js'

export class ParagraphParser {
  /**
   * 解析段落节点
   */
  parseParagraph(pNode: XmlNode): Paragraph {
    const pPr = xmlParser.getChild(pNode, 'w:pPr')
    const properties = pPr ? this.parseParagraphProperties(pPr) : undefined
    
    // 解析运行（runs）
    const runs: any[] = []
    const rNodes = xmlParser.getChildren(pNode, 'w:r')
    for (const rNode of rNodes) {
      runs.push({ text: this.extractText(rNode) })
    }
    
    // 解析超链接中的运行
    const hyperlinkNodes = xmlParser.getChildren(pNode, 'w:hyperlink')
    for (const hlNode of hyperlinkNodes) {
      const hlRuns = xmlParser.getChildren(hlNode, 'w:r')
      for (const rNode of hlRuns) {
        runs.push({ text: this.extractText(rNode) })
      }
    }
    
    const id = pPr ? xmlParser.getAttr(pPr, 'w14:paraId') : undefined
    
    return {
      properties,
      runs,
      id
    }
  }
  
  /**
   * 解析段落属性
   */
  private parseParagraphProperties(pPr: XmlNode): ParagraphProperties {
    const props: ParagraphProperties = {}
    
    // 样式ID
    const pStyle = xmlParser.getChild(pPr, 'w:pStyle')
    if (pStyle) {
      props.styleId = xmlParser.getAttr(pStyle, 'w:val')
    }
    
    // 对齐
    const jc = xmlParser.getChild(pPr, 'w:jc')
    if (jc) {
      const val = xmlParser.getAttr(jc, 'w:val')
      if (val) props.alignment = val as any
    }
    
    // 缩进
    const ind = xmlParser.getChild(pPr, 'w:ind')
    if (ind) {
      props.indentation = this.parseIndentation(ind)
    }
    
    // 间距
    const spacing = xmlParser.getChild(pPr, 'w:spacing')
    if (spacing) {
      props.spacing = this.parseSpacing(spacing)
    }
    
    // 边框
    const pBdr = xmlParser.getChild(pPr, 'w:pBdr')
    if (pBdr) {
      props.borders = this.parseParagraphBorders(pBdr)
    }
    
    // 底纹
    const shd = xmlParser.getChild(pPr, 'w:shd')
    if (shd) {
      props.shading = {
        fill: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:fill')),
        color: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:color')),
        pattern: xmlParser.getAttr(shd, 'w:val') as any
      }
    }
    
    // 制表位
    const tabs = xmlParser.getChild(pPr, 'w:tabs')
    if (tabs) {
      props.tabs = this.parseTabs(tabs)
    }
    
    // 分页控制
    props.keepNext = this.parseBool(xmlParser.getChild(pPr, 'w:keepNext'))
    props.keepLines = this.parseBool(xmlParser.getChild(pPr, 'w:keepLines'))
    props.pageBreakBefore = this.parseBool(xmlParser.getChild(pPr, 'w:pageBreakBefore'))
    props.widowControl = this.parseBool(xmlParser.getChild(pPr, 'w:widowControl'))
    
    // 编号
    const numPr = xmlParser.getChild(pPr, 'w:numPr')
    if (numPr) {
      props.numbering = this.parseNumberingReference(numPr)
    }
    
    // 框架属性
    const framePr = xmlParser.getChild(pPr, 'w:framePr')
    if (framePr) {
      props.framePr = this.parseFrameProperties(framePr)
    }
    
    // 文本方向
    const textDirection = xmlParser.getChild(pPr, 'w:textDirection')
    if (textDirection) {
      props.textDirection = xmlParser.getAttr(textDirection, 'w:val') as any
    }
    
    // 文本对齐
    const textAlignment = xmlParser.getChild(pPr, 'w:textAlignment')
    if (textAlignment) {
      props.textAlignment = xmlParser.getAttr(textAlignment, 'w:val') as any
    }
    
    // 双向文本
    props.bidi = this.parseBool(xmlParser.getChild(pPr, 'w:bidi'))
    
    // 网格和间距
    props.snapToGrid = this.parseBool(xmlParser.getChild(pPr, 'w:snapToGrid'))
    props.contextualSpacing = this.parseBool(xmlParser.getChild(pPr, 'w:contextualSpacing'))
    props.mirrorIndents = this.parseBool(xmlParser.getChild(pPr, 'w:mirrorIndents'))
    
    // 禁则处理
    props.suppressLineNumbers = this.parseBool(xmlParser.getChild(pPr, 'w:suppressLineNumbers'))
    props.suppressAutoHyphens = this.parseBool(xmlParser.getChild(pPr, 'w:suppressAutoHyphens'))
    props.kinsoku = this.parseBool(xmlParser.getChild(pPr, 'w:kinsoku'))
    props.wordWrap = this.parseBool(xmlParser.getChild(pPr, 'w:wordWrap'))
    props.overflowPunct = this.parseBool(xmlParser.getChild(pPr, 'w:overflowPunct'))
    props.topLinePunct = this.parseBool(xmlParser.getChild(pPr, 'w:topLinePunct'))
    props.autoSpaceDE = this.parseBool(xmlParser.getChild(pPr, 'w:autoSpaceDE'))
    props.autoSpaceDN = this.parseBool(xmlParser.getChild(pPr, 'w:autoSpaceDN'))
    
    // 大纲级别
    const outlineLvl = xmlParser.getChild(pPr, 'w:outlineLvl')
    if (outlineLvl) {
      props.outlineLevel = xmlParser.parseInt(xmlParser.getAttr(outlineLvl, 'w:val'))
    }
    
    return props
  }
  
  private parseIndentation(ind: XmlNode): Indentation {
    return {
      left: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:left')),
      right: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:right')),
      start: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:start')),
      end: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:end')),
      firstLine: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:firstLine')),
      hanging: xmlParser.parseTwip(xmlParser.getAttr(ind, 'w:hanging'))
    }
  }
  
  private parseSpacing(spacing: XmlNode): Spacing {
    return {
      before: xmlParser.parseTwip(xmlParser.getAttr(spacing, 'w:before')),
      after: xmlParser.parseTwip(xmlParser.getAttr(spacing, 'w:after')),
      line: xmlParser.parseTwip(xmlParser.getAttr(spacing, 'w:line')),
      lineRule: xmlParser.getAttr(spacing, 'w:lineRule') as any,
      beforeAutoSpacing: xmlParser.parseBool(xmlParser.getAttr(spacing, 'w:beforeAutospacing')),
      afterAutoSpacing: xmlParser.parseBool(xmlParser.getAttr(spacing, 'w:afterAutospacing'))
    }
  }
  
  private parseParagraphBorders(pBdr: XmlNode): ParagraphBorders {
    return {
      top: this.parseBorder(xmlParser.getChild(pBdr, 'w:top')),
      bottom: this.parseBorder(xmlParser.getChild(pBdr, 'w:bottom')),
      left: this.parseBorder(xmlParser.getChild(pBdr, 'w:left')),
      right: this.parseBorder(xmlParser.getChild(pBdr, 'w:right')),
      between: this.parseBorder(xmlParser.getChild(pBdr, 'w:between')),
      bar: this.parseBorder(xmlParser.getChild(pBdr, 'w:bar'))
    }
  }
  
  private parseBorder(bdr: XmlNode | undefined): any {
    if (!bdr) return undefined
    return {
      style: xmlParser.getAttr(bdr, 'w:val'),
      color: xmlParser.parseColor(xmlParser.getAttr(bdr, 'w:color')),
      size: xmlParser.parseInt(xmlParser.getAttr(bdr, 'w:sz')),
      space: xmlParser.parseInt(xmlParser.getAttr(bdr, 'w:space')),
      shadow: xmlParser.parseBool(xmlParser.getAttr(bdr, 'w:shadow')),
      frame: xmlParser.parseBool(xmlParser.getAttr(bdr, 'w:frame'))
    }
  }
  
  private parseTabs(tabs: XmlNode): Tab[] {
    const tabNodes = xmlParser.getChildren(tabs, 'w:tab')
    return tabNodes.map(tab => ({
      position: xmlParser.parseTwip(xmlParser.getAttr(tab, 'w:pos')) || 0,
      alignment: xmlParser.getAttr(tab, 'w:val') as any,
      leader: xmlParser.getAttr(tab, 'w:leader') as any
    }))
  }
  
  private parseNumberingReference(numPr: XmlNode): NumberingReference | undefined {
    const numId = xmlParser.getChild(numPr, 'w:numId')
    const ilvl = xmlParser.getChild(numPr, 'w:ilvl')
    
    const numIdVal = numId ? xmlParser.parseInt(xmlParser.getAttr(numId, 'w:val')) : undefined
    const ilvlVal = ilvl ? xmlParser.parseInt(xmlParser.getAttr(ilvl, 'w:val')) : undefined
    
    if (numIdVal !== undefined && ilvlVal !== undefined) {
      return { numId: numIdVal, ilvl: ilvlVal }
    }
    return undefined
  }
  
  private parseFrameProperties(framePr: XmlNode): FrameProperties {
    return {
      dropCap: xmlParser.getAttr(framePr, 'w:dropCap') as any,
      lines: xmlParser.parseInt(xmlParser.getAttr(framePr, 'w:lines')),
      width: xmlParser.parseTwip(xmlParser.getAttr(framePr, 'w:w')),
      height: xmlParser.parseTwip(xmlParser.getAttr(framePr, 'w:h')),
      vAnchor: xmlParser.getAttr(framePr, 'w:vAnchor') as any,
      hAnchor: xmlParser.getAttr(framePr, 'w:hAnchor') as any,
      x: xmlParser.parseTwip(xmlParser.getAttr(framePr, 'w:x')),
      xAlign: xmlParser.getAttr(framePr, 'w:xAlign') as any,
      y: xmlParser.parseTwip(xmlParser.getAttr(framePr, 'w:y')),
      yAlign: xmlParser.getAttr(framePr, 'w:yAlign') as any,
      hRule: xmlParser.getAttr(framePr, 'w:hRule') as any,
      wrap: xmlParser.parseBool(xmlParser.getAttr(framePr, 'w:wrap')),
      hSpace: xmlParser.parseTwip(xmlParser.getAttr(framePr, 'w:hSpace')),
      vSpace: xmlParser.parseTwip(xmlParser.getAttr(framePr, 'w:vSpace'))
    }
  }
  
  private extractText(rNode: XmlNode): string {
    const tNodes = xmlParser.getChildren(rNode, 'w:t')
    return tNodes.map(t => {
      // 文本可能直接在节点中，或在子节点中
      if (typeof t === 'string') return t
      if (t['#text']) return String(t['#text'])
      return ''
    }).join('')
  }
  
  private parseBool(node: XmlNode | undefined): boolean | undefined {
    if (!node) return undefined
    return xmlParser.parseBool(xmlParser.getAttr(node, 'w:val'), true)
  }
  
  parse(xml: string): Paragraph[] {
    const doc = xmlParser.parse(xml)
    const body = doc['w:document']?.['w:body']
    if (!body) return []
    
    const pNodes = xmlParser.getChildren(body, 'w:p')
    return pNodes.map(p => this.parseParagraph(p))
  }
}
