import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { TextRun, RunProperties, FontInfo, EastAsianLayout, LanguageInfo } from '../types.js'

export class RunParser {
  /**
   * 解析文本运行
   */
  parseRun(rNode: XmlNode): TextRun {
    const rPr = xmlParser.getChild(rNode, 'w:rPr')
    const properties = rPr ? this.parseRunProperties(rPr) : undefined
    
    const text = this.extractText(rNode)
    
    return {
      text,
      properties
    }
  }
  
  /**
   * 解析运行属性
   */
  private parseRunProperties(rPr: XmlNode): RunProperties {
    const props: RunProperties = {}
    
    // 样式ID
    const rStyle = xmlParser.getChild(rPr, 'w:rStyle')
    if (rStyle) {
      props.styleId = xmlParser.getAttr(rStyle, 'w:val')
    }
    
    // 基础样式
    props.bold = this.parseBool(xmlParser.getChild(rPr, 'w:b'))
    props.italic = this.parseBool(xmlParser.getChild(rPr, 'w:i'))
    
    // 下划线
    const u = xmlParser.getChild(rPr, 'w:u')
    if (u) {
      const val = xmlParser.getAttr(u, 'w:val')
      props.underline = val === 'none' ? 'none' : (val as any || 'single')
    }
    
    props.strike = this.parseBool(xmlParser.getChild(rPr, 'w:strike'))
    props.doubleStrike = this.parseBool(xmlParser.getChild(rPr, 'w:dstrike'))
    
    // 字体
    const rFonts = xmlParser.getChild(rPr, 'w:rFonts')
    if (rFonts) {
      props.fonts = this.parseFontInfo(rFonts)
    }
    
    // 字号
    const sz = xmlParser.getChild(rPr, 'w:sz')
    if (sz) {
      props.fontSize = xmlParser.parseHalfPoint(xmlParser.getAttr(sz, 'w:val'))
    }
    
    const szCs = xmlParser.getChild(rPr, 'w:szCs')
    if (szCs) {
      props.fontSizeCs = xmlParser.parseHalfPoint(xmlParser.getAttr(szCs, 'w:val'))
    }
    
    // 颜色
    const color = xmlParser.getChild(rPr, 'w:color')
    if (color) {
      props.color = xmlParser.parseColor(xmlParser.getAttr(color, 'w:val'))
    }
    
    // 高亮
    const highlight = xmlParser.getChild(rPr, 'w:highlight')
    if (highlight) {
      props.highlight = xmlParser.getAttr(highlight, 'w:val') as any
    }
    
    // 底纹
    const shd = xmlParser.getChild(rPr, 'w:shd')
    if (shd) {
      props.shading = {
        fill: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:fill')),
        color: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:color')),
        pattern: xmlParser.getAttr(shd, 'w:val') as any
      }
    }
    
    // 边框
    const bdr = xmlParser.getChild(rPr, 'w:bdr')
    if (bdr) {
      props.border = {
        style: xmlParser.getAttr(bdr, 'w:val') as any,
        color: xmlParser.parseColor(xmlParser.getAttr(bdr, 'w:color')),
        size: xmlParser.parseInt(xmlParser.getAttr(bdr, 'w:sz')),
        space: xmlParser.parseInt(xmlParser.getAttr(bdr, 'w:space')),
        shadow: xmlParser.parseBool(xmlParser.getAttr(bdr, 'w:shadow')),
        frame: xmlParser.parseBool(xmlParser.getAttr(bdr, 'w:frame'))
      }
    }
    
    // 位置和间距
    const position = xmlParser.getChild(rPr, 'w:position')
    if (position) {
      props.position = xmlParser.parseInt(xmlParser.getAttr(position, 'w:val'))
    }
    
    const spacing = xmlParser.getChild(rPr, 'w:spacing')
    if (spacing) {
      props.spacing = xmlParser.parseInt(xmlParser.getAttr(spacing, 'w:val'))
    }
    
    const w = xmlParser.getChild(rPr, 'w:w')
    if (w) {
      props.scale = xmlParser.parseInt(xmlParser.getAttr(w, 'w:val'))
    }
    
    const kern = xmlParser.getChild(rPr, 'w:kern')
    if (kern) {
      props.kern = xmlParser.parseInt(xmlParser.getAttr(kern, 'w:val'))
    }
    
    // 上下标
    const vertAlign = xmlParser.getChild(rPr, 'w:vertAlign')
    if (vertAlign) {
      props.verticalAlign = xmlParser.getAttr(vertAlign, 'w:val') as any
    }
    
    // 大小写
    props.smallCaps = this.parseBool(xmlParser.getChild(rPr, 'w:smallCaps'))
    props.allCaps = this.parseBool(xmlParser.getChild(rPr, 'w:caps'))
    
    // 隐藏
    props.hidden = this.parseBool(xmlParser.getChild(rPr, 'w:vanish'))
    props.webHidden = this.parseBool(xmlParser.getChild(rPr, 'w:webHidden'))
    props.specVanish = this.parseBool(xmlParser.getChild(rPr, 'w:specVanish'))
    
    // 效果
    props.emboss = this.parseBool(xmlParser.getChild(rPr, 'w:emboss'))
    props.imprint = this.parseBool(xmlParser.getChild(rPr, 'w:imprint'))
    props.outline = this.parseBool(xmlParser.getChild(rPr, 'w:outline'))
    props.shadow = this.parseBool(xmlParser.getChild(rPr, 'w:shadow'))
    
    // 文本效果
    const effect = xmlParser.getChild(rPr, 'w:effect')
    if (effect) {
      props.effect = xmlParser.getAttr(effect, 'w:val') as any
    }
    
    // 东亚排版
    const eastAsianLayout = xmlParser.getChild(rPr, 'w:eastAsianLayout')
    if (eastAsianLayout) {
      props.eastAsianLayout = this.parseEastAsianLayout(eastAsianLayout)
    }
    
    // 着重号
    const em = xmlParser.getChild(rPr, 'w:em')
    if (em) {
      props.emphasis = xmlParser.getAttr(em, 'w:val') as any
    }
    
    // 语言
    const lang = xmlParser.getChild(rPr, 'w:lang')
    if (lang) {
      props.lang = this.parseLanguageInfo(lang)
    }
    
    // 适应文本
    const fitText = xmlParser.getChild(rPr, 'w:fitText')
    if (fitText) {
      const width = xmlParser.parseTwip(xmlParser.getAttr(fitText, 'w:w'))
      const id = xmlParser.parseInt(xmlParser.getAttr(fitText, 'w:id'))
      if (width !== undefined && id !== undefined) {
        props.fitText = { width, id }
      }
    }
    
    // 右到左
    props.rtl = this.parseBool(xmlParser.getChild(rPr, 'w:rtl'))
    
    return props
  }
  
  private parseFontInfo(rFonts: XmlNode): FontInfo {
    return {
      ascii: xmlParser.getAttr(rFonts, 'w:ascii'),
      hAnsi: xmlParser.getAttr(rFonts, 'w:hAnsi'),
      eastAsia: xmlParser.getAttr(rFonts, 'w:eastAsia'),
      cs: xmlParser.getAttr(rFonts, 'w:cs'),
      hint: xmlParser.getAttr(rFonts, 'w:hint') as any
    }
  }
  
  private parseEastAsianLayout(eal: XmlNode): EastAsianLayout {
    return {
      id: xmlParser.parseInt(xmlParser.getAttr(eal, 'w:id')),
      combine: xmlParser.parseBool(xmlParser.getAttr(eal, 'w:combine')),
      combineBrackets: xmlParser.getAttr(eal, 'w:combineBrackets') as any,
      vert: xmlParser.parseBool(xmlParser.getAttr(eal, 'w:vert')),
      vertCompress: xmlParser.parseBool(xmlParser.getAttr(eal, 'w:vertCompress'))
    }
  }
  
  private parseLanguageInfo(lang: XmlNode): LanguageInfo {
    return {
      val: xmlParser.getAttr(lang, 'w:val'),
      eastAsia: xmlParser.getAttr(lang, 'w:eastAsia'),
      bidi: xmlParser.getAttr(lang, 'w:bidi')
    }
  }
  
  private extractText(rNode: XmlNode): string {
    let text = ''
    
    // 普通文本
    const tNodes = xmlParser.getChildren(rNode, 'w:t')
    for (const t of tNodes) {
      if (typeof t === 'string') {
        text += t
      } else if (t['#text']) {
        text += String(t['#text'])
      }
    }
    
    // 制表符
    const tabNodes = xmlParser.getChildren(rNode, 'w:tab')
    text += '\t'.repeat(tabNodes.length)
    
    // 换行符
    const brNodes = xmlParser.getChildren(rNode, 'w:br')
    text += '\n'.repeat(brNodes.length)
    
    // 回车符
    const crNodes = xmlParser.getChildren(rNode, 'w:cr')
    text += '\n'.repeat(crNodes.length)
    
    return text
  }
  
  private parseBool(node: XmlNode | undefined): boolean | undefined {
    if (!node) return undefined
    return xmlParser.parseBool(xmlParser.getAttr(node, 'w:val'), true)
  }
  
  parse(xml: string): TextRun[] {
    const doc = xmlParser.parse(xml)
    const body = doc['w:document']?.['w:body']
    if (!body) return []
    
    const runs: TextRun[] = []
    const pNodes = xmlParser.getChildren(body, 'w:p')
    
    for (const p of pNodes) {
      const rNodes = xmlParser.getChildren(p, 'w:r')
      for (const r of rNodes) {
        runs.push(this.parseRun(r))
      }
    }
    
    return runs
  }
}
