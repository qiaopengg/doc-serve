import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { Style, RunProperties, ParagraphProperties } from '../types.js'

export class StyleParser {
  /**
   * 解析样式定义
   */
  parse(xml: string): Style[] {
    const styles: Style[] = []
    const doc = xmlParser.parse(xml)
    
    const stylesRoot = doc['w:styles']
    if (!stylesRoot) return styles

    const styleArray = stylesRoot['w:style']
    if (!styleArray) return styles

    const styleNodes = Array.isArray(styleArray) ? styleArray : [styleArray]

    for (const styleNode of styleNodes) {
      const style = this.parseStyle(styleNode)
      if (style) {
        styles.push(style)
      }
    }

    return styles
  }

  /**
   * 解析默认样式
   */
  parseDefaults(xml: string): any {
    const doc = xmlParser.parse(xml)
    const stylesRoot = doc['w:styles']
    if (!stylesRoot) return {}

    const docDefaults = stylesRoot['w:docDefaults']
    if (!docDefaults) return {}

    const defaults: any = {}

    // 默认运行属性
    if (docDefaults['w:rPrDefault']) {
      const rPr = docDefaults['w:rPrDefault']['w:rPr']
      if (rPr) {
        defaults.runProperties = this.parseRunProperties(rPr)
      }
    }

    // 默认段落属性
    if (docDefaults['w:pPrDefault']) {
      const pPr = docDefaults['w:pPrDefault']['w:pPr']
      if (pPr) {
        defaults.paragraphProperties = this.parseParagraphProperties(pPr)
      }
    }

    return defaults
  }

  /**
   * 解析单个样式
   */
  private parseStyle(styleNode: XmlNode): Style | null {
    if (!styleNode) return null

    const styleId = styleNode['@_w:styleId']
    const type = styleNode['@_w:type']
    
    if (!styleId || !type) return null

    const style: Style = {
      styleId,
      type: type as any
    }

    // 样式名称
    const name = styleNode['w:name']
    if (name) {
      style.name = name['@_w:val']
    }

    // 基于样式
    const basedOn = styleNode['w:basedOn']
    if (basedOn) {
      style.basedOn = basedOn['@_w:val']
    }

    // 后续样式
    const next = styleNode['w:next']
    if (next) {
      style.next = next['@_w:val']
    }

    // 链接样式
    const link = styleNode['w:link']
    if (link) {
      style.link = link['@_w:val']
    }

    // 自动重定义
    if (styleNode['w:autoRedefine']) {
      style.autoRedefine = true
    }

    // 隐藏
    if (styleNode['w:hidden']) {
      style.hidden = true
    }

    // UI 优先级
    const uiPriority = styleNode['w:uiPriority']
    if (uiPriority) {
      style.uiPriority = parseInt(uiPriority['@_w:val'] || '0')
    }

    // 半隐藏
    if (styleNode['w:semiHidden']) {
      style.semiHidden = true
    }

    // 使用时显示
    if (styleNode['w:unhideWhenUsed']) {
      style.unhideWhenUsed = true
    }

    // 快速样式
    if (styleNode['w:qFormat']) {
      style.qFormat = true
    }

    // 锁定
    if (styleNode['w:locked']) {
      style.locked = true
    }

    // 个人样式
    if (styleNode['w:personal']) {
      style.personal = true
    }

    if (styleNode['w:personalCompose']) {
      style.personalCompose = true
    }

    if (styleNode['w:personalReply']) {
      style.personalReply = true
    }

    // 修订 ID
    const rsid = styleNode['w:rsid']
    if (rsid) {
      style.rsid = rsid['@_w:val']
    }

    // 段落属性
    const pPr = styleNode['w:pPr']
    if (pPr) {
      style.paragraphProperties = this.parseParagraphProperties(pPr)
    }

    // 运行属性
    const rPr = styleNode['w:rPr']
    if (rPr) {
      style.runProperties = this.parseRunProperties(rPr)
    }

    // 表格属性
    const tblPr = styleNode['w:tblPr']
    if (tblPr) {
      style.tableProperties = tblPr
    }

    const trPr = styleNode['w:trPr']
    if (trPr) {
      style.tableRowProperties = trPr
    }

    const tcPr = styleNode['w:tcPr']
    if (tcPr) {
      style.tableCellProperties = tcPr
    }

    return style
  }

  /**
   * 解析运行属性（简化版）
   */
  private parseRunProperties(rPr: XmlNode): Partial<RunProperties> {
    const props: Partial<RunProperties> = {}

    if (rPr['w:b']) {
      props.bold = rPr['w:b']['@_w:val'] !== '0'
    }

    if (rPr['w:i']) {
      props.italic = rPr['w:i']['@_w:val'] !== '0'
    }

    if (rPr['w:u']) {
      props.underline = rPr['w:u']['@_w:val'] as any
    }

    if (rPr['w:sz']) {
      props.fontSize = parseInt(rPr['w:sz']['@_w:val'] || '0')
    }

    if (rPr['w:color']) {
      props.color = rPr['w:color']['@_w:val']
    }

    if (rPr['w:rFonts']) {
      props.fonts = {
        ascii: rPr['w:rFonts']['@_w:ascii'],
        hAnsi: rPr['w:rFonts']['@_w:hAnsi'],
        eastAsia: rPr['w:rFonts']['@_w:eastAsia'],
        cs: rPr['w:rFonts']['@_w:cs']
      }
    }

    return props
  }

  /**
   * 解析段落属性（简化版）
   */
  private parseParagraphProperties(pPr: XmlNode): Partial<ParagraphProperties> {
    const props: Partial<ParagraphProperties> = {}

    if (pPr['w:jc']) {
      props.alignment = pPr['w:jc']['@_w:val'] as any
    }

    if (pPr['w:ind']) {
      props.indentation = {
        left: parseInt(pPr['w:ind']['@_w:left'] || '0'),
        right: parseInt(pPr['w:ind']['@_w:right'] || '0'),
        firstLine: parseInt(pPr['w:ind']['@_w:firstLine'] || '0'),
        hanging: parseInt(pPr['w:ind']['@_w:hanging'] || '0')
      }
    }

    if (pPr['w:spacing']) {
      props.spacing = {
        before: parseInt(pPr['w:spacing']['@_w:before'] || '0'),
        after: parseInt(pPr['w:spacing']['@_w:after'] || '0'),
        line: parseInt(pPr['w:spacing']['@_w:line'] || '0'),
        lineRule: pPr['w:spacing']['@_w:lineRule'] as any
      }
    }

    return props
  }
}

export const styleParser = new StyleParser()
