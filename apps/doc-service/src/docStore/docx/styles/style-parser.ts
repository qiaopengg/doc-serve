import { XMLParser } from 'fast-xml-parser'
import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { Style, RunProperties, ParagraphProperties, StyleMap, RunStyle, DocxParagraph } from '../types.js'
import { normalizeColor, normalizeAlignment, parseBooleanOnOff, detectHeadingLevel, mergeDefined, asArray } from '../core/utils.js'

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

  parseStyles(xml: string): { styles: StyleMap; docDefaultRun: Partial<RunStyle>; docDefaultPara: Pick<DocxParagraph, "alignment"> } {
    const parser = new XMLParser({ ignoreAttributes: false, allowBooleanAttributes: true })
    const obj: any = parser.parse(xml)
    const stylesRoot = obj?.["w:styles"]

    const styles: StyleMap = new Map()

    const docDefaults = stylesRoot?.["w:docDefaults"]
    const docDefaultRun = this.parseRunPropsFromObj(docDefaults?.["w:rPrDefault"]?.["w:rPr"])
    const docDefaultPara = { alignment: this.parseParaAlignmentFromObj(docDefaults?.["w:pPrDefault"]?.["w:pPr"]) }

    for (const st of asArray<any>(stylesRoot?.["w:style"])) {
      const styleId = st?.["@_w:styleId"]
      if (typeof styleId !== "string" || !styleId) continue

      const name = st?.["w:name"]?.["@_w:val"]
      const basedOn = st?.["w:basedOn"]?.["@_w:val"]
      const type = st?.["@_w:type"]

      const para = {
        alignment: this.parseParaAlignmentFromObj(st?.["w:pPr"]),
        headingLevel: detectHeadingLevel(styleId, name)
      }

      styles.set(styleId, {
        type: typeof type === "string" ? type : undefined,
        basedOn: typeof basedOn === "string" ? basedOn : undefined,
        name: typeof name === "string" ? name : undefined,
        run: this.parseRunPropsFromObj(st?.["w:rPr"]),
        para
      })
    }

    return { styles, docDefaultRun, docDefaultPara }
  }

  parseRunPropsFromObj(rPr: any): Partial<RunStyle> {
    if (!rPr || typeof rPr !== "object") return {}

    const props: Partial<RunStyle> = {}

    if (rPr["w:b"] != null) {
      const v = rPr["w:b"]?.["@_w:val"]
      const on = parseBooleanOnOff(v, true)
      if (on !== undefined) props.bold = on
    }
    if (rPr["w:i"] != null) {
      const v = rPr["w:i"]?.["@_w:val"]
      const on = parseBooleanOnOff(v, true)
      if (on !== undefined) props.italic = on
    }
    if (rPr["w:u"] != null) {
      const v = String(rPr["w:u"]?.["@_w:val"] ?? "").trim().toLowerCase()
      if (!v) {
        props.underline = true
      } else if (v === "none") {
        props.underline = false
      } else {
        props.underline = true
      }
    }
    const sz = rPr["w:sz"]?.["@_w:val"]
    const szCs = rPr["w:szCs"]?.["@_w:val"]
    const sizeRaw = sz != null ? sz : szCs
    if (sizeRaw != null) {
      const raw = Number.parseInt(String(sizeRaw), 10)
      if (Number.isFinite(raw)) props.fontSize = raw / 2
    }
    const col = normalizeColor(rPr["w:color"]?.["@_w:val"])
    if (col) props.color = col
    const fonts = rPr["w:rFonts"]
    if (fonts && typeof fonts === "object") {
      const font = fonts["@_w:eastAsia"] ?? fonts["@_w:ascii"] ?? fonts["@_w:hAnsi"]
      if (typeof font === "string" && font) props.font = font
    }

    return props
  }

  parseParaAlignmentFromObj(pPr: any): "left" | "center" | "right" | "justify" | undefined {
    const v = pPr?.["w:jc"]?.["@_w:val"]
    return normalizeAlignment(v)
  }

  resolveStyleChain(styleId: string | undefined, styles: StyleMap, kind: "run" | "para", visited = new Set<string>()): any {
    if (!styleId) return {}
    if (visited.has(styleId)) return {}
    visited.add(styleId)
    const s = styles.get(styleId)
    if (!s) return {}
    const base = this.resolveStyleChain(s.basedOn, styles, kind, visited)
    const own = kind === "run" ? (s.run ?? {}) : (s.para ?? {})
    return mergeDefined(base, own)
  }
}

export const styleParser = new StyleParser()

