import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { Table, TableProperties, TableRow, TableCell, TableCellProperties, TableRowProperties } from '../types.js'

export class TableParser {
  /**
   * 解析表格
   */
  parseTable(tblNode: XmlNode): Table {
    const tblPr = xmlParser.getChild(tblNode, 'w:tblPr')
    const properties = tblPr ? this.parseTableProperties(tblPr) : undefined
    
    const tblGrid = xmlParser.getChild(tblNode, 'w:tblGrid')
    const grid = tblGrid ? this.parseTableGrid(tblGrid) : undefined
    
    const trNodes = xmlParser.getChildren(tblNode, 'w:tr')
    const rows = trNodes.map(tr => this.parseTableRow(tr))
    
    return {
      properties,
      grid,
      rows
    }
  }
  
  /**
   * 解析表格属性
   */
  private parseTableProperties(tblPr: XmlNode): TableProperties {
    const props: TableProperties = {}
    
    // 样式ID
    const tblStyle = xmlParser.getChild(tblPr, 'w:tblStyle')
    if (tblStyle) {
      props.styleId = xmlParser.getAttr(tblStyle, 'w:val')
    }
    
    // 宽度
    const tblW = xmlParser.getChild(tblPr, 'w:tblW')
    if (tblW) {
      props.width = {
        type: xmlParser.getAttr(tblW, 'w:type') as any || 'auto',
        value: xmlParser.parseInt(xmlParser.getAttr(tblW, 'w:w')) || 0
      }
    }
    
    // 对齐
    const jc = xmlParser.getChild(tblPr, 'w:jc')
    if (jc) {
      props.alignment = xmlParser.getAttr(jc, 'w:val') as any
    }
    
    // 缩进
    const tblInd = xmlParser.getChild(tblPr, 'w:tblInd')
    if (tblInd) {
      props.indent = xmlParser.parseTwip(xmlParser.getAttr(tblInd, 'w:w'))
    }
    
    // 边框
    const tblBorders = xmlParser.getChild(tblPr, 'w:tblBorders')
    if (tblBorders) {
      props.borders = {
        top: this.parseBorder(xmlParser.getChild(tblBorders, 'w:top')),
        bottom: this.parseBorder(xmlParser.getChild(tblBorders, 'w:bottom')),
        left: this.parseBorder(xmlParser.getChild(tblBorders, 'w:left')),
        right: this.parseBorder(xmlParser.getChild(tblBorders, 'w:right')),
        insideH: this.parseBorder(xmlParser.getChild(tblBorders, 'w:insideH')),
        insideV: this.parseBorder(xmlParser.getChild(tblBorders, 'w:insideV'))
      }
    }
    
    // 底纹
    const shd = xmlParser.getChild(tblPr, 'w:shd')
    if (shd) {
      props.shading = {
        fill: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:fill')),
        color: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:color')),
        pattern: xmlParser.getAttr(shd, 'w:val') as any
      }
    }
    
    // 布局
    const tblLayout = xmlParser.getChild(tblPr, 'w:tblLayout')
    if (tblLayout) {
      props.layout = xmlParser.getAttr(tblLayout, 'w:type') as any
    }
    
    // 单元格间距
    const tblCellSpacing = xmlParser.getChild(tblPr, 'w:tblCellSpacing')
    if (tblCellSpacing) {
      props.cellSpacing = xmlParser.parseTwip(xmlParser.getAttr(tblCellSpacing, 'w:w'))
    }
    
    // 单元格边距
    const tblCellMar = xmlParser.getChild(tblPr, 'w:tblCellMar')
    if (tblCellMar) {
      const top = xmlParser.getChild(tblCellMar, 'w:top')
      const bottom = xmlParser.getChild(tblCellMar, 'w:bottom')
      const left = xmlParser.getChild(tblCellMar, 'w:left')
      const right = xmlParser.getChild(tblCellMar, 'w:right')
      const start = xmlParser.getChild(tblCellMar, 'w:start')
      const end = xmlParser.getChild(tblCellMar, 'w:end')
      
      props.cellMargin = {
        top: top ? xmlParser.parseTwip(xmlParser.getAttr(top, 'w:w')) : undefined,
        bottom: bottom ? xmlParser.parseTwip(xmlParser.getAttr(bottom, 'w:w')) : undefined,
        left: left ? xmlParser.parseTwip(xmlParser.getAttr(left, 'w:w')) : undefined,
        right: right ? xmlParser.parseTwip(xmlParser.getAttr(right, 'w:w')) : undefined,
        start: start ? xmlParser.parseTwip(xmlParser.getAttr(start, 'w:w')) : undefined,
        end: end ? xmlParser.parseTwip(xmlParser.getAttr(end, 'w:w')) : undefined
      }
    }
    
    // 表格样式覆盖
    const tblLook = xmlParser.getChild(tblPr, 'w:tblLook')
    if (tblLook) {
      props.look = {
        firstRow: xmlParser.parseBool(xmlParser.getAttr(tblLook, 'w:firstRow')),
        lastRow: xmlParser.parseBool(xmlParser.getAttr(tblLook, 'w:lastRow')),
        firstColumn: xmlParser.parseBool(xmlParser.getAttr(tblLook, 'w:firstColumn')),
        lastColumn: xmlParser.parseBool(xmlParser.getAttr(tblLook, 'w:lastColumn')),
        noHBand: xmlParser.parseBool(xmlParser.getAttr(tblLook, 'w:noHBand')),
        noVBand: xmlParser.parseBool(xmlParser.getAttr(tblLook, 'w:noVBand'))
      }
    }
    
    // 标题和描述
    const tblCaption = xmlParser.getChild(tblPr, 'w:tblCaption')
    if (tblCaption) {
      props.caption = xmlParser.getAttr(tblCaption, 'w:val')
    }
    
    const tblDescription = xmlParser.getChild(tblPr, 'w:tblDescription')
    if (tblDescription) {
      props.description = xmlParser.getAttr(tblDescription, 'w:val')
    }
    
    // 双向文本
    const bidiVisual = xmlParser.getChild(tblPr, 'w:bidiVisual')
    if (bidiVisual) {
      props.bidiVisual = xmlParser.parseBool(xmlParser.getAttr(bidiVisual, 'w:val'))
    }
    
    // 重叠
    const tblOverlap = xmlParser.getChild(tblPr, 'w:tblOverlap')
    if (tblOverlap) {
      props.overlap = xmlParser.getAttr(tblOverlap, 'w:val') as any
    }
    
    return props
  }
  
  /**
   * 解析表格网格
   */
  private parseTableGrid(tblGrid: XmlNode): any {
    const gridColNodes = xmlParser.getChildren(tblGrid, 'w:gridCol')
    const columns = gridColNodes.map(col => ({
      width: xmlParser.parseTwip(xmlParser.getAttr(col, 'w:w')) || 0
    }))
    
    return { columns }
  }
  
  /**
   * 解析表格行
   */
  private parseTableRow(trNode: XmlNode): TableRow {
    const trPr = xmlParser.getChild(trNode, 'w:trPr')
    const properties = trPr ? this.parseTableRowProperties(trPr) : undefined
    
    const tcNodes = xmlParser.getChildren(trNode, 'w:tc')
    const cells = tcNodes.map(tc => this.parseTableCell(tc))
    
    return {
      properties,
      cells
    }
  }
  
  /**
   * 解析表格行属性
   */
  private parseTableRowProperties(trPr: XmlNode): TableRowProperties {
    const props: TableRowProperties = {}
    
    // 行不可拆分
    const cantSplit = xmlParser.getChild(trPr, 'w:cantSplit')
    if (cantSplit) {
      props.cantSplit = xmlParser.parseBool(xmlParser.getAttr(cantSplit, 'w:val'))
    }
    
    // 行高
    const trHeight = xmlParser.getChild(trPr, 'w:trHeight')
    if (trHeight) {
      props.height = {
        value: xmlParser.parseTwip(xmlParser.getAttr(trHeight, 'w:val')) || 0,
        rule: xmlParser.getAttr(trHeight, 'w:hRule') as any
      }
    }
    
    // 标题行重复
    const tblHeader = xmlParser.getChild(trPr, 'w:tblHeader')
    if (tblHeader) {
      props.header = xmlParser.parseBool(xmlParser.getAttr(tblHeader, 'w:val'))
    }
    
    // 行前后网格
    const gridBefore = xmlParser.getChild(trPr, 'w:gridBefore')
    if (gridBefore) {
      props.gridBefore = xmlParser.parseInt(xmlParser.getAttr(gridBefore, 'w:val'))
    }
    
    const gridAfter = xmlParser.getChild(trPr, 'w:gridAfter')
    if (gridAfter) {
      props.gridAfter = xmlParser.parseInt(xmlParser.getAttr(gridAfter, 'w:val'))
    }
    
    // 隐藏
    const hidden = xmlParser.getChild(trPr, 'w:hidden')
    if (hidden) {
      props.hidden = xmlParser.parseBool(xmlParser.getAttr(hidden, 'w:val'))
    }
    
    return props
  }
  
  /**
   * 解析表格单元格
   */
  private parseTableCell(tcNode: XmlNode): TableCell {
    const tcPr = xmlParser.getChild(tcNode, 'w:tcPr')
    const properties = tcPr ? this.parseTableCellProperties(tcPr) : undefined
    
    // 解析单元格内容（段落）
    const pNodes = xmlParser.getChildren(tcNode, 'w:p')
    const content: any[] = pNodes.map(() => ({ runs: [] })) // 简化处理
    
    return {
      properties,
      content
    }
  }
  
  /**
   * 解析表格单元格属性
   */
  private parseTableCellProperties(tcPr: XmlNode): TableCellProperties {
    const props: TableCellProperties = {}
    
    // 宽度
    const tcW = xmlParser.getChild(tcPr, 'w:tcW')
    if (tcW) {
      props.width = {
        type: xmlParser.getAttr(tcW, 'w:type') as any || 'auto',
        value: xmlParser.parseInt(xmlParser.getAttr(tcW, 'w:w')) || 0
      }
    }
    
    // 水平合并
    const gridSpan = xmlParser.getChild(tcPr, 'w:gridSpan')
    if (gridSpan) {
      props.gridSpan = xmlParser.parseInt(xmlParser.getAttr(gridSpan, 'w:val'))
    }
    
    // 垂直合并
    const vMerge = xmlParser.getChild(tcPr, 'w:vMerge')
    if (vMerge) {
      const val = xmlParser.getAttr(vMerge, 'w:val')
      props.vMerge = val === 'restart' ? 'restart' : 'continue'
    }
    
    // 边框
    const tcBorders = xmlParser.getChild(tcPr, 'w:tcBorders')
    if (tcBorders) {
      props.borders = {
        top: this.parseBorder(xmlParser.getChild(tcBorders, 'w:top')),
        bottom: this.parseBorder(xmlParser.getChild(tcBorders, 'w:bottom')),
        left: this.parseBorder(xmlParser.getChild(tcBorders, 'w:left')),
        right: this.parseBorder(xmlParser.getChild(tcBorders, 'w:right')),
        insideH: this.parseBorder(xmlParser.getChild(tcBorders, 'w:insideH')),
        insideV: this.parseBorder(xmlParser.getChild(tcBorders, 'w:insideV')),
        tl2br: this.parseBorder(xmlParser.getChild(tcBorders, 'w:tl2br')),
        tr2bl: this.parseBorder(xmlParser.getChild(tcBorders, 'w:tr2bl'))
      }
    }
    
    // 底纹
    const shd = xmlParser.getChild(tcPr, 'w:shd')
    if (shd) {
      props.shading = {
        fill: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:fill')),
        color: xmlParser.parseColor(xmlParser.getAttr(shd, 'w:color')),
        pattern: xmlParser.getAttr(shd, 'w:val') as any
      }
    }
    
    // 单元格边距
    const tcMar = xmlParser.getChild(tcPr, 'w:tcMar')
    if (tcMar) {
      const top = xmlParser.getChild(tcMar, 'w:top')
      const bottom = xmlParser.getChild(tcMar, 'w:bottom')
      const left = xmlParser.getChild(tcMar, 'w:left')
      const right = xmlParser.getChild(tcMar, 'w:right')
      const start = xmlParser.getChild(tcMar, 'w:start')
      const end = xmlParser.getChild(tcMar, 'w:end')
      
      props.margins = {
        top: top ? xmlParser.parseTwip(xmlParser.getAttr(top, 'w:w')) : undefined,
        bottom: bottom ? xmlParser.parseTwip(xmlParser.getAttr(bottom, 'w:w')) : undefined,
        left: left ? xmlParser.parseTwip(xmlParser.getAttr(left, 'w:w')) : undefined,
        right: right ? xmlParser.parseTwip(xmlParser.getAttr(right, 'w:w')) : undefined,
        start: start ? xmlParser.parseTwip(xmlParser.getAttr(start, 'w:w')) : undefined,
        end: end ? xmlParser.parseTwip(xmlParser.getAttr(end, 'w:w')) : undefined
      }
    }
    
    // 垂直对齐
    const vAlign = xmlParser.getChild(tcPr, 'w:vAlign')
    if (vAlign) {
      props.verticalAlign = xmlParser.getAttr(vAlign, 'w:val') as any
    }
    
    // 文本方向
    const textDirection = xmlParser.getChild(tcPr, 'w:textDirection')
    if (textDirection) {
      props.textDirection = xmlParser.getAttr(textDirection, 'w:val') as any
    }
    
    // 适应文本
    const tcFitText = xmlParser.getChild(tcPr, 'w:tcFitText')
    if (tcFitText) {
      props.fitText = xmlParser.parseBool(xmlParser.getAttr(tcFitText, 'w:val'))
    }
    
    // 不换行
    const noWrap = xmlParser.getChild(tcPr, 'w:noWrap')
    if (noWrap) {
      props.noWrap = xmlParser.parseBool(xmlParser.getAttr(noWrap, 'w:val'))
    }
    
    // 隐藏
    const hideMark = xmlParser.getChild(tcPr, 'w:hideMark')
    if (hideMark) {
      props.hidden = xmlParser.parseBool(xmlParser.getAttr(hideMark, 'w:val'))
    }
    
    return props
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
  
  parse(xml: string): Table[] {
    const doc = xmlParser.parse(xml)
    const body = doc['w:document']?.['w:body']
    if (!body) return []
    
    const tblNodes = xmlParser.getChildren(body, 'w:tbl')
    return tblNodes.map(tbl => this.parseTable(tbl))
  }
}
