import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { SectionProperties, PageSize, PageMargin, HeaderFooterReference, PageNumberType, ColumnProperties, LineNumberType, SectionBorders, SectionPropertiesSpec, OrderedXmlNode } from '../types.js'
import { childOf, attrsOf, attrOf, childrenNamed } from '../core/utils.js'

export class SectionParser {
  /**
   * 解析分节属性
   */
  parseSectionProperties(sectPr: XmlNode): SectionProperties {
    const props: SectionProperties = {}
    
    // 页面大小
    const pgSz = xmlParser.getChild(sectPr, 'w:pgSz')
    if (pgSz) {
      props.pageSize = this.parsePageSize(pgSz)
      const orient = xmlParser.getAttr(pgSz, 'w:orient')
      if (orient) {
        props.pageOrientation = orient as any
      }
    }
    
    // 页边距
    const pgMar = xmlParser.getChild(sectPr, 'w:pgMar')
    if (pgMar) {
      props.pageMargin = this.parsePageMargin(pgMar)
    }
    
    // 页眉引用
    const headerRefs = xmlParser.getChildren(sectPr, 'w:headerReference')
    if (headerRefs.length > 0) {
      props.headerReference = headerRefs.map(ref => this.parseHeaderFooterReference(ref))
    }
    
    // 页脚引用
    const footerRefs = xmlParser.getChildren(sectPr, 'w:footerReference')
    if (footerRefs.length > 0) {
      props.footerReference = footerRefs.map(ref => this.parseHeaderFooterReference(ref))
    }
    
    // 页码格式
    const pgNumType = xmlParser.getChild(sectPr, 'w:pgNumType')
    if (pgNumType) {
      props.pageNumberType = this.parsePageNumberType(pgNumType)
    }
    
    // 分栏
    const cols = xmlParser.getChild(sectPr, 'w:cols')
    if (cols) {
      props.columns = this.parseColumnProperties(cols)
    }
    
    // 行号
    const lnNumType = xmlParser.getChild(sectPr, 'w:lnNumType')
    if (lnNumType) {
      props.lineNumberType = this.parseLineNumberType(lnNumType)
    }
    
    // 分节类型
    const type = xmlParser.getChild(sectPr, 'w:type')
    if (type) {
      props.sectionType = xmlParser.getAttr(type, 'w:val') as any
    }
    
    // 文本方向
    const textDirection = xmlParser.getChild(sectPr, 'w:textDirection')
    if (textDirection) {
      props.textDirection = xmlParser.getAttr(textDirection, 'w:val') as any
    }
    
    // 垂直对齐
    const vAlign = xmlParser.getChild(sectPr, 'w:vAlign')
    if (vAlign) {
      props.verticalAlign = xmlParser.getAttr(vAlign, 'w:val') as any
    }
    
    // 边框
    const pgBorders = xmlParser.getChild(sectPr, 'w:pgBorders')
    if (pgBorders) {
      props.borders = this.parseSectionBorders(pgBorders)
    }
    
    // 首页不同
    const titlePg = xmlParser.getChild(sectPr, 'w:titlePg')
    if (titlePg) {
      props.titlePage = xmlParser.parseBool(xmlParser.getAttr(titlePg, 'w:val'), true)
    }
    
    // 右侧装订线
    const rtlGutter = xmlParser.getChild(sectPr, 'w:rtlGutter')
    if (rtlGutter) {
      props.rtlGutter = xmlParser.parseBool(xmlParser.getAttr(rtlGutter, 'w:val'), true)
    }
    
    // 表单保护
    const formProt = xmlParser.getChild(sectPr, 'w:formProt')
    if (formProt) {
      props.formProtection = xmlParser.parseBool(xmlParser.getAttr(formProt, 'w:val'), true)
    }
    
    // 双向文本
    const bidi = xmlParser.getChild(sectPr, 'w:bidi')
    if (bidi) {
      props.bidi = xmlParser.parseBool(xmlParser.getAttr(bidi, 'w:val'), true)
    }
    
    return props
  }
  
  private parsePageSize(pgSz: XmlNode): PageSize {
    return {
      width: xmlParser.parseTwip(xmlParser.getAttr(pgSz, 'w:w')) || 0,
      height: xmlParser.parseTwip(xmlParser.getAttr(pgSz, 'w:h')) || 0,
      code: xmlParser.parseInt(xmlParser.getAttr(pgSz, 'w:code'))
    }
  }
  
  private parsePageMargin(pgMar: XmlNode): PageMargin {
    return {
      top: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:top')) || 0,
      right: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:right')) || 0,
      bottom: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:bottom')) || 0,
      left: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:left')) || 0,
      header: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:header')) || 0,
      footer: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:footer')) || 0,
      gutter: xmlParser.parseTwip(xmlParser.getAttr(pgMar, 'w:gutter')) || 0
    }
  }
  
  private parseHeaderFooterReference(ref: XmlNode): HeaderFooterReference {
    return {
      type: xmlParser.getAttr(ref, 'w:type') as any || 'default',
      relationshipId: xmlParser.getAttr(ref, 'r:id') || ''
    }
  }
  
  private parsePageNumberType(pgNumType: XmlNode): PageNumberType {
    return {
      format: xmlParser.getAttr(pgNumType, 'w:fmt') as any,
      start: xmlParser.parseInt(xmlParser.getAttr(pgNumType, 'w:start')),
      chapStyle: xmlParser.parseInt(xmlParser.getAttr(pgNumType, 'w:chapStyle')),
      chapSep: xmlParser.getAttr(pgNumType, 'w:chapSep') as any
    }
  }
  
  private parseColumnProperties(cols: XmlNode): ColumnProperties {
    const props: ColumnProperties = {
      count: xmlParser.parseInt(xmlParser.getAttr(cols, 'w:num')),
      space: xmlParser.parseTwip(xmlParser.getAttr(cols, 'w:space')),
      equalWidth: xmlParser.parseBool(xmlParser.getAttr(cols, 'w:equalWidth'), true),
      separator: xmlParser.parseBool(xmlParser.getAttr(cols, 'w:sep'))
    }
    
    const colNodes = xmlParser.getChildren(cols, 'w:col')
    if (colNodes.length > 0) {
      props.columns = colNodes.map(col => ({
        width: xmlParser.parseTwip(xmlParser.getAttr(col, 'w:w')) || 0,
        space: xmlParser.parseTwip(xmlParser.getAttr(col, 'w:space')) || 0
      }))
    }
    
    return props
  }
  
  private parseLineNumberType(lnNumType: XmlNode): LineNumberType {
    return {
      countBy: xmlParser.parseInt(xmlParser.getAttr(lnNumType, 'w:countBy')),
      start: xmlParser.parseInt(xmlParser.getAttr(lnNumType, 'w:start')),
      restart: xmlParser.getAttr(lnNumType, 'w:restart') as any,
      distance: xmlParser.parseTwip(xmlParser.getAttr(lnNumType, 'w:distance'))
    }
  }
  
  private parseSectionBorders(pgBorders: XmlNode): SectionBorders {
    return {
      top: this.parseBorder(xmlParser.getChild(pgBorders, 'w:top')),
      bottom: this.parseBorder(xmlParser.getChild(pgBorders, 'w:bottom')),
      left: this.parseBorder(xmlParser.getChild(pgBorders, 'w:left')),
      right: this.parseBorder(xmlParser.getChild(pgBorders, 'w:right')),
      offsetFrom: xmlParser.getAttr(pgBorders, 'w:offsetFrom') as any,
      zOrder: xmlParser.getAttr(pgBorders, 'w:zOrder') as any,
      display: xmlParser.getAttr(pgBorders, 'w:display') as any
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
  
  parse(xml: string): SectionProperties[] {
    const doc = xmlParser.parse(xml)
    const body = doc['w:document']?.['w:body']
    if (!body) return []
    
    const sections: SectionProperties[] = []
    
    // 文档级分节属性
    const sectPr = xmlParser.getChild(body, 'w:sectPr')
    if (sectPr) {
      sections.push(this.parseSectionProperties(sectPr))
    }
    
    // 段落中的分节属性
    const pNodes = xmlParser.getChildren(body, 'w:p')
    for (const p of pNodes) {
      const pPr = xmlParser.getChild(p, 'w:pPr')
      if (pPr) {
        const pSectPr = xmlParser.getChild(pPr, 'w:sectPr')
        if (pSectPr) {
          sections.push(this.parseSectionProperties(pSectPr))
        }
      }
    }
    
    return sections
  }

  parseSectionPropertiesFromSectPr(sectPr: OrderedXmlNode | undefined): SectionPropertiesSpec | undefined {
    if (!sectPr) return undefined

    const pgSz = childOf(sectPr, "w:pgSz")
    const pgMar = childOf(sectPr, "w:pgMar")
    const cols = childOf(sectPr, "w:cols")

    const sizeAttrs = pgSz ? attrsOf(pgSz) : undefined
    const widthRaw = sizeAttrs ? Number.parseInt(String(attrOf(sizeAttrs, "w:w") ?? ""), 10) : NaN
    const heightRaw = sizeAttrs ? Number.parseInt(String(attrOf(sizeAttrs, "w:h") ?? ""), 10) : NaN
    const orientRaw = sizeAttrs ? String(attrOf(sizeAttrs, "w:orient") ?? "").trim().toLowerCase() : ""
    const orientation: "portrait" | "landscape" | undefined =
      orientRaw === "landscape" ? "landscape" : orientRaw === "portrait" ? "portrait" : undefined

    const marAttrs = pgMar ? attrsOf(pgMar) : undefined
    const topRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:top") ?? ""), 10) : NaN
    const rightRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:right") ?? ""), 10) : NaN
    const bottomRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:bottom") ?? ""), 10) : NaN
    const leftRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:left") ?? ""), 10) : NaN
    const headerRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:header") ?? ""), 10) : NaN
    const footerRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:footer") ?? ""), 10) : NaN
    const gutterRaw = marAttrs ? Number.parseInt(String(attrOf(marAttrs, "w:gutter") ?? ""), 10) : NaN

    const colsAttrs = cols ? attrsOf(cols) : undefined
    const colCountRaw = colsAttrs ? Number.parseInt(String(attrOf(colsAttrs, "w:num") ?? ""), 10) : NaN
    const colSpaceRaw = colsAttrs ? Number.parseInt(String(attrOf(colsAttrs, "w:space") ?? ""), 10) : NaN

    const page: SectionPropertiesSpec["page"] = {}
    const size: NonNullable<SectionPropertiesSpec["page"]>["size"] = {}
    if (Number.isFinite(widthRaw)) size.width = widthRaw
    if (Number.isFinite(heightRaw)) size.height = heightRaw
    if (orientation) size.orientation = orientation
    if (Object.keys(size).length) page.size = size

    const margin: NonNullable<SectionPropertiesSpec["page"]>["margin"] = {}
    if (Number.isFinite(topRaw)) margin.top = topRaw
    if (Number.isFinite(rightRaw)) margin.right = rightRaw
    if (Number.isFinite(bottomRaw)) margin.bottom = bottomRaw
    if (Number.isFinite(leftRaw)) margin.left = leftRaw
    if (Number.isFinite(headerRaw)) margin.header = headerRaw
    if (Number.isFinite(footerRaw)) margin.footer = footerRaw
    if (Number.isFinite(gutterRaw)) margin.gutter = gutterRaw
    if (Object.keys(margin).length) page.margin = margin

    const out: SectionPropertiesSpec = {}
    if (Object.keys(page).length) out.page = page
    if (Number.isFinite(colCountRaw) || Number.isFinite(colSpaceRaw)) {
      const column: SectionPropertiesSpec["column"] = {}
      if (Number.isFinite(colCountRaw)) column.count = colCountRaw
      if (Number.isFinite(colSpaceRaw)) column.space = colSpaceRaw
      out.column = column
    }

    return Object.keys(out).length ? out : undefined
  }
}
