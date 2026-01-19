import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, UnderlineType, IRunOptions, PageOrientation } from "docx"
import { XMLParser } from "fast-xml-parser"
import { readZipEntry, replaceZipEntry } from "./core/zipReader.js"
import { parseXml, buildXml } from "./core/xmlParser.js"

export type BorderSpec = {
  style?: string
  size?: number
  color?: string
}

export type TableBordersSpec = {
  top?: BorderSpec
  bottom?: BorderSpec
  left?: BorderSpec
  right?: BorderSpec
  insideHorizontal?: BorderSpec
  insideVertical?: BorderSpec
}

export type CellBordersSpec = {
  top?: BorderSpec
  bottom?: BorderSpec
  left?: BorderSpec
  right?: BorderSpec
}

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  fontSize?: number
  color?: string
  fill?: string  // 背景色
  alignment?: "left" | "center" | "right" | "justify"
  gridSpan?: number  // 合并列数
  rowSpan?: number
  colIndex?: number
  skip?: boolean
  font?: string  // 字体名称
  verticalAlign?: "top" | "center" | "bottom"  // 垂直对齐
  borders?: CellBordersSpec
}

// 图片/图形
export interface ImageSpec {
  type: "image"
  relationshipId?: string
  imageData?: string  // base64 或 URL
  width?: number
  height?: number
  description?: string
  title?: string
  hyperlink?: string
}

// 列表/编号
export interface NumberingSpec {
  numId?: number
  level?: number
  format?: string  // bullet, decimal, lowerLetter, etc.
  text?: string  // 自定义编号文本
}

// 书签
export interface BookmarkSpec {
  id: string
  name: string
  type: "start" | "end"
}

// 域代码（字段）
export interface FieldSpec {
  type: "field"
  code?: string  // 域代码，如 "TOC \o \"1-3\""
  result?: string  // 域结果
  fieldType?: "toc" | "pageref" | "ref" | "hyperlink" | "date" | "time" | "formula" | "other"
}

// 脚注/尾注
export interface NoteSpec {
  type: "footnote" | "endnote"
  id: string
  content?: DocxParagraph[]
}

// 注释
export interface CommentSpec {
  id: string
  author?: string
  date?: string
  content?: string
  rangeType: "start" | "end"
}

// 文本框
export interface TextBoxSpec {
  type: "textbox"
  content?: DocxParagraph[]
  width?: number
  height?: number
  positioning?: {
    x?: number
    y?: number
    anchor?: "page" | "paragraph" | "margin"
  }
}

// 形状
export interface ShapeSpec {
  type: "shape"
  shapeType?: string  // rect, ellipse, line, etc.
  width?: number
  height?: number
  fill?: string
  stroke?: string
  strokeWidth?: number
  text?: string
}

// 修订标记
export interface RevisionSpec {
  type: "insert" | "delete" | "format"
  author?: string
  date?: string
  content?: string
}

// 数学公式 (OMML)
export interface MathSpec {
  type: "math"
  omml?: string  // Office Math Markup Language
  latex?: string  // 转换后的 LaTeX（可选）
}

// 嵌套表格
export interface NestedTableSpec {
  type: "nestedTable"
  table: DocxParagraph  // 指向表格段落
}

export interface RunStyle {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  color?: string
  font?: string
  highlight?: string  // 高亮颜色
  strikethrough?: boolean
  doubleStrikethrough?: boolean
  subscript?: boolean
  superscript?: boolean
  smallCaps?: boolean
  allCaps?: boolean
  emboss?: boolean
  imprint?: boolean
  shadow?: boolean
  outline?: boolean
}

export type ParagraphSpacing = {
  before?: number
  after?: number
  line?: number
  lineRule?: "auto" | "exact" | "atLeast"
}

export type SectionPropertiesSpec = {
  page?: {
    size?: {
      width?: number
      height?: number
      orientation?: "portrait" | "landscape"
    }
    margin?: {
      top?: number
      right?: number
      bottom?: number
      left?: number
      header?: number
      footer?: number
      gutter?: number
    }
  }
  column?: {
    count?: number
    space?: number
  }
}

export interface DocxParagraph {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  color?: string
  font?: string
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6
  alignment?: "left" | "center" | "right" | "justify"
  spacing?: ParagraphSpacing
  isTable?: boolean
  tableData?: string[][]
  tableCellStyles?: CellStyle[][]  // 单元格样式
  tableLayout?: "fixed" | "autofit"
  tableGridCols?: number[]
  tableBorders?: TableBordersSpec
  tableStyleId?: string  // 表格样式引用
  link?: string
  runs?: RunStyle[]  // 支持多个 run（用于处理混合样式段落）
  sectionProperties?: SectionPropertiesSpec
  
  // 新增：扩展元素
  images?: ImageSpec[]
  numbering?: NumberingSpec
  bookmarks?: BookmarkSpec[]
  fields?: FieldSpec[]
  notes?: NoteSpec[]
  comments?: CommentSpec[]
  textBoxes?: TextBoxSpec[]
  shapes?: ShapeSpec[]
  revisions?: RevisionSpec[]
  math?: MathSpec[]
  indent?: {
    left?: number
    right?: number
    firstLine?: number
    hanging?: number
  }
  keepNext?: boolean  // 与下一段保持在同一页
  keepLines?: boolean  // 段落内不分页
  pageBreakBefore?: boolean
  widowControl?: boolean
  outlineLevel?: number
  styleId?: string
  styleName?: string
}

type StyleMap = Map<string, { type?: string; basedOn?: string; name?: string; run?: Partial<RunStyle>; para?: Pick<DocxParagraph, "alignment" | "headingLevel"> }>

type OrderedXmlNode = Record<string, any>

function asArray<T>(v: T | T[] | undefined | null): T[] {
  if (!v) return []
  return Array.isArray(v) ? v : [v]
}

function parseBooleanOnOff(v: unknown, defaultWhenMissingVal?: true): boolean | undefined {
  if (v == null) return defaultWhenMissingVal ? true : undefined
  const s = String(v).trim().toLowerCase()
  if (s === "0" || s === "false" || s === "off" || s === "none") return false
  return true
}

function normalizeColor(v: unknown): string | undefined {
  if (!v) return undefined
  const s = String(v).trim()
  if (!s || s.toLowerCase() === "auto") return undefined
  const hex = s.replace(/^#/, "").toUpperCase()
  // 支持 6 位和 8 位十六进制颜色
  if (/^[0-9A-F]{6}$/.test(hex)) return hex
  if (/^[0-9A-F]{8}$/.test(hex)) return hex.substring(0, 6) // 去掉 alpha 通道
  return undefined
}

function normalizeBorderColor(v: unknown): string | undefined {
  if (!v) return undefined
  const s = String(v).trim()
  if (!s) return undefined
  if (s.toLowerCase() === "auto") return "auto"
  const hex = s.replace(/^#/, "").toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(hex)) return undefined
  return hex
}

function normalizeAlignment(v: unknown): "left" | "center" | "right" | "justify" | undefined {
  const s = String(v ?? "").trim().toLowerCase()
  if (!s) return undefined
  if (s === "center") return "center"
  if (s === "right" || s === "end") return "right"
  if (s === "left" || s === "start") return "left"
  if (s === "both" || s === "justify") return "justify"
  return undefined
}

function tagNameOf(node: OrderedXmlNode): string | undefined {
  const keys = Object.keys(node)
  for (const k of keys) {
    if (k === ":@" || k === "#text") continue
    return k
  }
  if (keys.includes("#text")) return "#text"
  return undefined
}

function attrsOf(node: OrderedXmlNode): Record<string, any> {
  const attrs = node[":@"]
  if (!attrs || typeof attrs !== "object") return {}
  return attrs
}

function attrOf(attrs: Record<string, any>, name: string): any {
  if (name in attrs) return attrs[name]
  const alt = `@_${name}`
  if (alt in attrs) return attrs[alt]
  return undefined
}

function childrenOf(node: OrderedXmlNode): OrderedXmlNode[] {
  const tn = tagNameOf(node)
  if (!tn || tn === "#text") return []
  const v = node[tn]
  return Array.isArray(v) ? v : []
}

function childOf(node: OrderedXmlNode, name: string): OrderedXmlNode | undefined {
  for (const c of childrenOf(node)) {
    if (tagNameOf(c) === name) return c
  }
  return undefined
}

function childrenNamed(node: OrderedXmlNode, name: string): OrderedXmlNode[] {
  return childrenOf(node).filter((c) => tagNameOf(c) === name)
}

function textFromOrdered(node: OrderedXmlNode): string {
  const tn = tagNameOf(node)
  if (tn === "#text") return String((node as any)["#text"] ?? "")

  if (tn === "w:t") {
    const children = childrenOf(node)
    if (children.length === 0) return ""
    return children.map(textFromOrdered).join("")
  }

  if (tn === "w:tab") return "\t"
  if (tn === "w:br" || tn === "w:cr") return "\n"

  return childrenOf(node).map(textFromOrdered).join("")
}

function parseRunPropsFromOrdered(rPr: OrderedXmlNode | undefined): { props: Partial<RunStyle>; charStyleId?: string } {
  if (!rPr) return { props: {} }

  const props: Partial<RunStyle> = {}
  let charStyleId: string | undefined

  for (const c of childrenOf(rPr)) {
    const tn = tagNameOf(c)
    const attrs = attrsOf(c)

    if (tn === "w:rStyle") {
      const v = attrOf(attrs, "w:val")
      if (typeof v === "string" && v) charStyleId = v
    } else if (tn === "w:b") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.bold = on
    } else if (tn === "w:i") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.italic = on
    } else if (tn === "w:u") {
      const v = String(attrOf(attrs, "w:val") ?? "").trim().toLowerCase()
      if (!v) {
        props.underline = true
      } else if (v === "none") {
        props.underline = false
      } else {
        props.underline = true
      }
    } else if (tn === "w:color") {
      const col = normalizeColor(attrOf(attrs, "w:val"))
      if (col) props.color = col
    } else if (tn === "w:sz") {
      const raw = Number.parseInt(String(attrOf(attrs, "w:val") ?? ""), 10)
      if (Number.isFinite(raw)) props.fontSize = raw / 2
    } else if (tn === "w:szCs") {
      if (props.fontSize === undefined) {
        const raw = Number.parseInt(String(attrOf(attrs, "w:val") ?? ""), 10)
        if (Number.isFinite(raw)) props.fontSize = raw / 2
      }
    } else if (tn === "w:rFonts") {
      const font = (attrOf(attrs, "w:eastAsia") ?? attrOf(attrs, "w:ascii") ?? attrOf(attrs, "w:hAnsi")) as unknown
      if (typeof font === "string" && font) props.font = font
    } else if (tn === "w:highlight") {
      const col = normalizeColor(attrOf(attrs, "w:val"))
      if (col) props.highlight = col
    } else if (tn === "w:strike") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.strikethrough = on
    } else if (tn === "w:dstrike") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.doubleStrikethrough = on
    } else if (tn === "w:vertAlign") {
      const v = String(attrOf(attrs, "w:val") ?? "").trim().toLowerCase()
      if (v === "subscript") props.subscript = true
      if (v === "superscript") props.superscript = true
    } else if (tn === "w:smallCaps") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.smallCaps = on
    } else if (tn === "w:caps") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.allCaps = on
    } else if (tn === "w:emboss") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.emboss = on
    } else if (tn === "w:imprint") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.imprint = on
    } else if (tn === "w:shadow") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.shadow = on
    } else if (tn === "w:outline") {
      const on = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
      if (on !== undefined) props.outline = on
    }
  }

  return { props, charStyleId }
}

function parseParaPropsFromOrdered(pPr: OrderedXmlNode | undefined): {
  alignment?: "left" | "center" | "right" | "justify"
  paraStyleId?: string
  runDefaults?: Partial<RunStyle>
  spacing?: ParagraphSpacing
  numbering?: NumberingSpec
  indent?: { left?: number; right?: number; firstLine?: number; hanging?: number }
  keepNext?: boolean
  keepLines?: boolean
  pageBreakBefore?: boolean
  widowControl?: boolean
  outlineLevel?: number
} {
  if (!pPr) return {}

  let alignment: "left" | "center" | "right" | "justify" | undefined
  let paraStyleId: string | undefined
  let runDefaults: Partial<RunStyle> | undefined
  let spacing: ParagraphSpacing | undefined
  let numbering: NumberingSpec | undefined
  let indent: { left?: number; right?: number; firstLine?: number; hanging?: number } | undefined
  let keepNext: boolean | undefined
  let keepLines: boolean | undefined
  let pageBreakBefore: boolean | undefined
  let widowControl: boolean | undefined
  let outlineLevel: number | undefined

  for (const c of childrenOf(pPr)) {
    const tn = tagNameOf(c)
    const attrs = attrsOf(c)
    if (tn === "w:jc") {
      alignment = normalizeAlignment(attrOf(attrs, "w:val"))
    } else if (tn === "w:pStyle") {
      const v = attrOf(attrs, "w:val")
      if (typeof v === "string" && v) paraStyleId = v
    } else if (tn === "w:rPr") {
      runDefaults = parseRunPropsFromOrdered(c).props
    } else if (tn === "w:spacing") {
      const beforeRaw = Number.parseInt(String(attrOf(attrs, "w:before") ?? ""), 10)
      const afterRaw = Number.parseInt(String(attrOf(attrs, "w:after") ?? ""), 10)
      const lineRaw = Number.parseInt(String(attrOf(attrs, "w:line") ?? ""), 10)
      const lineRuleRaw = String(attrOf(attrs, "w:lineRule") ?? "").trim().toLowerCase()

      const s: ParagraphSpacing = {}
      if (Number.isFinite(beforeRaw)) s.before = beforeRaw
      if (Number.isFinite(afterRaw)) s.after = afterRaw
      if (Number.isFinite(lineRaw)) s.line = lineRaw
      if (lineRuleRaw === "auto" || lineRuleRaw === "exact" || lineRuleRaw === "atleast") {
        s.lineRule = lineRuleRaw === "atleast" ? "atLeast" : (lineRuleRaw as any)
      }
      if (Object.keys(s).length) spacing = s
    } else if (tn === "w:numPr") {
      const numIdNode = childOf(c, "w:numId")
      const ilvlNode = childOf(c, "w:ilvl")
      const numIdRaw = numIdNode ? Number.parseInt(String(attrOf(attrsOf(numIdNode), "w:val") ?? ""), 10) : NaN
      const ilvlRaw = ilvlNode ? Number.parseInt(String(attrOf(attrsOf(ilvlNode), "w:val") ?? ""), 10) : NaN
      numbering = {
        numId: Number.isFinite(numIdRaw) ? numIdRaw : undefined,
        level: Number.isFinite(ilvlRaw) ? ilvlRaw : undefined
      }
    } else if (tn === "w:ind") {
      const leftRaw = Number.parseInt(String(attrOf(attrs, "w:left") ?? attrOf(attrs, "w:start") ?? ""), 10)
      const rightRaw = Number.parseInt(String(attrOf(attrs, "w:right") ?? attrOf(attrs, "w:end") ?? ""), 10)
      const firstLineRaw = Number.parseInt(String(attrOf(attrs, "w:firstLine") ?? ""), 10)
      const hangingRaw = Number.parseInt(String(attrOf(attrs, "w:hanging") ?? ""), 10)
      indent = {}
      if (Number.isFinite(leftRaw)) indent.left = leftRaw
      if (Number.isFinite(rightRaw)) indent.right = rightRaw
      if (Number.isFinite(firstLineRaw)) indent.firstLine = firstLineRaw
      if (Number.isFinite(hangingRaw)) indent.hanging = hangingRaw
      if (!Object.keys(indent).length) indent = undefined
    } else if (tn === "w:keepNext") {
      keepNext = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
    } else if (tn === "w:keepLines") {
      keepLines = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
    } else if (tn === "w:pageBreakBefore") {
      pageBreakBefore = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
    } else if (tn === "w:widowControl") {
      widowControl = parseBooleanOnOff(attrOf(attrs, "w:val"), true)
    } else if (tn === "w:outlineLvl") {
      const lvl = Number.parseInt(String(attrOf(attrs, "w:val") ?? ""), 10)
      if (Number.isFinite(lvl)) outlineLevel = lvl
    }
  }

  return { alignment, paraStyleId, runDefaults, spacing, numbering, indent, keepNext, keepLines, pageBreakBefore, widowControl, outlineLevel }
}

function parseRunPropsFromObj(rPr: any): Partial<RunStyle> {
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

function parseParaAlignmentFromObj(pPr: any): "left" | "center" | "right" | "justify" | undefined {
  const v = pPr?.["w:jc"]?.["@_w:val"]
  return normalizeAlignment(v)
}

function detectHeadingLevel(styleId: string | undefined, styleName: string | undefined): 1 | 2 | 3 | 4 | 5 | 6 | undefined {
  const s = String(styleId ?? styleName ?? "").toLowerCase()
  if (!s) return undefined
  const m = s.match(/heading\s*([1-6])/i) || s.match(/heading_([1-6])/i) || s.match(/heading([1-6])/i)
  if (!m) return undefined
  const n = Number.parseInt(m[1]!, 10)
  if (n >= 1 && n <= 6) return n as any
  return undefined
}

function mergeDefined<T extends Record<string, any>>(...parts: Array<T | undefined>): T {
  const out: any = {}
  for (const p of parts) {
    if (!p) continue
    for (const [k, v] of Object.entries(p)) {
      if (v !== undefined) out[k] = v
    }
  }
  return out
}

// 解析图片/图形
function parseDrawingNode(drawingNode: OrderedXmlNode, hyperlinkRels: Map<string, string>): ImageSpec[] {
  const images: ImageSpec[] = []
  
  // 处理 w:drawing (现代图片格式)
  const inlineNodes = childrenNamed(drawingNode, "wp:inline")
  const anchorNodes = childrenNamed(drawingNode, "wp:anchor")
  
  for (const node of [...inlineNodes, ...anchorNodes]) {
    const graphicNode = childOf(node, "a:graphic")
    if (!graphicNode) continue
    
    const graphicDataNode = childOf(graphicNode, "a:graphicData")
    if (!graphicDataNode) continue
    
    const picNode = childOf(graphicDataNode, "pic:pic")
    if (!picNode) continue
    
    const blipFillNode = childOf(picNode, "pic:blipFill")
    const blipNode = blipFillNode ? childOf(blipFillNode, "a:blip") : undefined
    const embedId = blipNode ? attrOf(attrsOf(blipNode), "r:embed") : undefined
    
    const spPrNode = childOf(picNode, "pic:spPr")
    const xfrmNode = spPrNode ? childOf(spPrNode, "a:xfrm") : undefined
    const extNode = xfrmNode ? childOf(xfrmNode, "a:ext") : undefined
    
    const cxRaw = extNode ? Number.parseInt(String(attrOf(attrsOf(extNode), "cx") ?? ""), 10) : NaN
    const cyRaw = extNode ? Number.parseInt(String(attrOf(attrsOf(extNode), "cy") ?? ""), 10) : NaN
    
    const nvPicPrNode = childOf(picNode, "pic:nvPicPr")
    const cNvPrNode = nvPicPrNode ? childOf(nvPicPrNode, "pic:cNvPr") : undefined
    const descr = cNvPrNode ? attrOf(attrsOf(cNvPrNode), "descr") : undefined
    const title = cNvPrNode ? attrOf(attrsOf(cNvPrNode), "title") : undefined
    
    images.push({
      type: "image",
      relationshipId: typeof embedId === "string" ? embedId : undefined,
      width: Number.isFinite(cxRaw) ? cxRaw / 914400 * 72 : undefined, // EMU to points
      height: Number.isFinite(cyRaw) ? cyRaw / 914400 * 72 : undefined,
      description: typeof descr === "string" ? descr : undefined,
      title: typeof title === "string" ? title : undefined
    })
  }
  
  return images
}

// 解析旧版图片 (v:shape, w:pict)
function parsePictNode(pictNode: OrderedXmlNode): ImageSpec[] {
  const images: ImageSpec[] = []
  
  // 简化处理：提取基本信息
  const shapeNodes = childrenNamed(pictNode, "v:shape")
  for (const shapeNode of shapeNodes) {
    const imageDataNodes = childrenNamed(shapeNode, "v:imagedata")
    for (const imgNode of imageDataNodes) {
      const attrs = attrsOf(imgNode)
      const rid = attrOf(attrs, "r:id")
      const title = attrOf(attrs, "o:title")
      
      images.push({
        type: "image",
        relationshipId: typeof rid === "string" ? rid : undefined,
        title: typeof title === "string" ? title : undefined
      })
    }
  }
  
  return images
}

// 解析书签
function parseBookmarkNodes(pNode: OrderedXmlNode): BookmarkSpec[] {
  const bookmarks: BookmarkSpec[] = []
  
  for (const c of childrenOf(pNode)) {
    const tn = tagNameOf(c)
    const attrs = attrsOf(c)
    
    if (tn === "w:bookmarkStart") {
      const id = String(attrOf(attrs, "w:id") ?? "")
      const name = String(attrOf(attrs, "w:name") ?? "")
      if (id) bookmarks.push({ id, name, type: "start" })
    } else if (tn === "w:bookmarkEnd") {
      const id = String(attrOf(attrs, "w:id") ?? "")
      if (id) bookmarks.push({ id, name: "", type: "end" })
    }
  }
  
  return bookmarks
}

// 解析域代码（字段）
function parseFieldNodes(pNode: OrderedXmlNode): FieldSpec[] {
  const fields: FieldSpec[] = []
  
  // 简单域 (w:fldSimple)
  for (const c of childrenOf(pNode)) {
    if (tagNameOf(c) === "w:fldSimple") {
      const attrs = attrsOf(c)
      const instr = String(attrOf(attrs, "w:instr") ?? "")
      const text = textFromOrdered(c)
      
      fields.push({
        type: "field",
        code: instr,
        result: text,
        fieldType: detectFieldType(instr)
      })
    }
  }
  
  // 复杂域 (w:fldChar) - 需要状态机解析
  let inField = false
  let fieldCode = ""
  let fieldResult = ""
  
  for (const rNode of childrenNamed(pNode, "w:r")) {
    for (const c of childrenOf(rNode)) {
      const tn = tagNameOf(c)
      
      if (tn === "w:fldChar") {
        const fldCharType = String(attrOf(attrsOf(c), "w:fldCharType") ?? "")
        if (fldCharType === "begin") {
          inField = true
          fieldCode = ""
          fieldResult = ""
        } else if (fldCharType === "separate") {
          // 开始收集结果
        } else if (fldCharType === "end") {
          if (inField) {
            fields.push({
              type: "field",
              code: fieldCode.trim(),
              result: fieldResult.trim(),
              fieldType: detectFieldType(fieldCode)
            })
          }
          inField = false
        }
      } else if (tn === "w:instrText") {
        if (inField) fieldCode += textFromOrdered(c)
      } else if (tn === "w:t") {
        if (inField && fieldCode) fieldResult += textFromOrdered(c)
      }
    }
  }
  
  return fields
}

function detectFieldType(code: string): FieldSpec["fieldType"] {
  const c = code.trim().toUpperCase()
  if (c.startsWith("TOC")) return "toc"
  if (c.startsWith("PAGEREF")) return "pageref"
  if (c.startsWith("REF")) return "ref"
  if (c.startsWith("HYPERLINK")) return "hyperlink"
  if (c.startsWith("DATE")) return "date"
  if (c.startsWith("TIME")) return "time"
  if (c.startsWith("=")) return "formula"
  return "other"
}

// 解析注释标记
function parseCommentNodes(pNode: OrderedXmlNode): CommentSpec[] {
  const comments: CommentSpec[] = []
  
  for (const c of childrenOf(pNode)) {
    const tn = tagNameOf(c)
    const attrs = attrsOf(c)
    
    if (tn === "w:commentRangeStart") {
      const id = String(attrOf(attrs, "w:id") ?? "")
      if (id) comments.push({ id, rangeType: "start" })
    } else if (tn === "w:commentRangeEnd") {
      const id = String(attrOf(attrs, "w:id") ?? "")
      if (id) comments.push({ id, rangeType: "end" })
    }
  }
  
  return comments
}

// 解析脚注/尾注引用
function parseNoteReferences(pNode: OrderedXmlNode): NoteSpec[] {
  const notes: NoteSpec[] = []
  
  for (const rNode of childrenNamed(pNode, "w:r")) {
    for (const c of childrenOf(rNode)) {
      const tn = tagNameOf(c)
      const attrs = attrsOf(c)
      
      if (tn === "w:footnoteReference") {
        const id = String(attrOf(attrs, "w:id") ?? "")
        if (id) notes.push({ type: "footnote", id })
      } else if (tn === "w:endnoteReference") {
        const id = String(attrOf(attrs, "w:id") ?? "")
        if (id) notes.push({ type: "endnote", id })
      }
    }
  }
  
  return notes
}

// 解析数学公式 (OMML)
function parseMathNodes(pNode: OrderedXmlNode): MathSpec[] {
  const maths: MathSpec[] = []
  
  for (const c of childrenOf(pNode)) {
    if (tagNameOf(c) === "m:oMath" || tagNameOf(c) === "m:oMathPara") {
      // 简化：保存原始 XML
      const omml = buildXml([c], { ignoreAttributes: false, preserveOrder: true, format: false })
      maths.push({ type: "math", omml })
    }
  }
  
  return maths
}

function resolveStyleChain(styleId: string | undefined, styles: StyleMap, kind: "run" | "para", visited = new Set<string>()): any {
  if (!styleId) return {}
  if (visited.has(styleId)) return {}
  visited.add(styleId)
  const s = styles.get(styleId)
  if (!s) return {}
  const base = resolveStyleChain(s.basedOn, styles, kind, visited)
  const own = kind === "run" ? (s.run ?? {}) : (s.para ?? {})
  return mergeDefined(base, own)
}

// 解析 numbering.xml 获取编号格式定义
function parseNumbering(xml: string): Map<string, Map<number, { format?: string; text?: string }>> {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  const numberingRoot = obj?.["w:numbering"]
  
  const abstractNums = new Map<string, Map<number, { format?: string; text?: string }>>()
  const numMap = new Map<string, string>() // numId -> abstractNumId
  
  // 解析抽象编号定义
  for (const abstractNum of asArray<any>(numberingRoot?.["w:abstractNum"])) {
    const abstractNumId = String(abstractNum?.["@_w:abstractNumId"] ?? "")
    if (!abstractNumId) continue
    
    const levels = new Map<number, { format?: string; text?: string }>()
    
    for (const lvl of asArray<any>(abstractNum?.["w:lvl"])) {
      const ilvl = Number.parseInt(String(lvl?.["@_w:ilvl"] ?? ""), 10)
      if (!Number.isFinite(ilvl)) continue
      
      const numFmt = lvl?.["w:numFmt"]?.["@_w:val"]
      const lvlText = lvl?.["w:lvlText"]?.["@_w:val"]
      
      levels.set(ilvl, {
        format: typeof numFmt === "string" ? numFmt : undefined,
        text: typeof lvlText === "string" ? lvlText : undefined
      })
    }
    
    abstractNums.set(abstractNumId, levels)
  }
  
  // 解析编号实例
  for (const num of asArray<any>(numberingRoot?.["w:num"])) {
    const numId = String(num?.["@_w:numId"] ?? "")
    const abstractNumId = String(num?.["w:abstractNumId"]?.["@_w:val"] ?? "")
    if (numId && abstractNumId) {
      numMap.set(numId, abstractNumId)
    }
  }
  
  // 构建最终映射：numId -> levels
  const result = new Map<string, Map<number, { format?: string; text?: string }>>()
  for (const [numId, abstractNumId] of numMap.entries()) {
    const levels = abstractNums.get(abstractNumId)
    if (levels) result.set(numId, levels)
  }
  
  return result
}

function parseRelationships(xml: string): Map<string, string> {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  const rels = obj?.Relationships?.Relationship
  const out = new Map<string, string>()
  for (const r of asArray<any>(rels)) {
    const id = r?.["@_Id"]
    const type = r?.["@_Type"]
    const target = r?.["@_Target"]
    if (typeof id === "string" && typeof target === "string" && typeof type === "string") {
      if (type.includes("/hyperlink")) out.set(id, target)
      // 扩展：支持图片、页眉页脚等关系
      if (type.includes("/image")) out.set(id, target)
      if (type.includes("/header")) out.set(id, target)
      if (type.includes("/footer")) out.set(id, target)
      if (type.includes("/footnotes")) out.set(id, target)
      if (type.includes("/endnotes")) out.set(id, target)
      if (type.includes("/comments")) out.set(id, target)
    }
  }
  return out
}

function parseStyles(xml: string): { styles: StyleMap; docDefaultRun: Partial<RunStyle>; docDefaultPara: Pick<DocxParagraph, "alignment"> } {
  const parser = new XMLParser({ ignoreAttributes: false, allowBooleanAttributes: true })
  const obj: any = parser.parse(xml)
  const stylesRoot = obj?.["w:styles"]

  const styles: StyleMap = new Map()

  const docDefaults = stylesRoot?.["w:docDefaults"]
  const docDefaultRun = parseRunPropsFromObj(docDefaults?.["w:rPrDefault"]?.["w:rPr"])
  const docDefaultPara = { alignment: parseParaAlignmentFromObj(docDefaults?.["w:pPrDefault"]?.["w:pPr"]) }

  for (const st of asArray<any>(stylesRoot?.["w:style"])) {
    const styleId = st?.["@_w:styleId"]
    if (typeof styleId !== "string" || !styleId) continue

    const name = st?.["w:name"]?.["@_w:val"]
    const basedOn = st?.["w:basedOn"]?.["@_w:val"]
    const type = st?.["@_w:type"]

    const para = {
      alignment: parseParaAlignmentFromObj(st?.["w:pPr"]),
      headingLevel: detectHeadingLevel(styleId, name)
    }

    styles.set(styleId, {
      type: typeof type === "string" ? type : undefined,
      basedOn: typeof basedOn === "string" ? basedOn : undefined,
      name: typeof name === "string" ? name : undefined,
      run: parseRunPropsFromObj(st?.["w:rPr"]),
      para
    })
  }

  return { styles, docDefaultRun, docDefaultPara }
}

function buildParagraphFromRuns(
  runs: RunStyle[],
  paraInfo: { alignment?: "left" | "center" | "right" | "justify"; headingLevel?: 1 | 2 | 3 | 4 | 5 | 6; link?: string; spacing?: ParagraphSpacing }
): DocxParagraph {
  const text = runs.map((r) => r.text).join("")

  const base: DocxParagraph = {
    text,
    alignment: paraInfo.alignment,
    headingLevel: paraInfo.headingLevel,
    link: paraInfo.link,
    spacing: paraInfo.spacing,
    runs: runs.length ? runs : undefined
  }

  if (runs.length) {
    const keys: Array<keyof RunStyle> = ["bold", "italic", "underline", "fontSize", "color", "font"]
    const first = runs[0]!
    const uniform = keys.every((k) => runs.every((r) => r[k] === first[k]))
    if (uniform) {
      base.bold = first.bold
      base.italic = first.italic
      base.underline = first.underline
      base.fontSize = first.fontSize
      base.color = first.color
      base.font = first.font
    }
  }

  return base
}

function parseParagraphNode(
  pNode: OrderedXmlNode,
  styles: StyleMap,
  docDefaultRun: Partial<RunStyle>,
  docDefaultPara: Pick<DocxParagraph, "alignment">,
  hyperlinkRels: Map<string, string>
): DocxParagraph {
  const pPr = childOf(pNode, "w:pPr")
  const paraProps = parseParaPropsFromOrdered(pPr)

  const paraStyleId = paraProps.paraStyleId
  const paraStyleRun = resolveStyleChain(paraStyleId, styles, "run")
  const paraStylePara = resolveStyleChain(paraStyleId, styles, "para")

  const alignment = paraProps.alignment ?? paraStylePara.alignment ?? docDefaultPara.alignment
  const headingLevel = paraStylePara.headingLevel
  const spacing = paraProps.spacing
  const numbering = paraProps.numbering
  const indent = paraProps.indent

  const runs: RunStyle[] = []
  const hyperlinkTargetsInPara = new Set<string>()
  const images: ImageSpec[] = []
  const bookmarks: BookmarkSpec[] = []
  const fields: FieldSpec[] = []
  const notes: NoteSpec[] = []
  const comments: CommentSpec[] = []

  // 解析书签
  bookmarks.push(...parseBookmarkNodes(pNode))
  
  // 解析注释标记
  comments.push(...parseCommentNodes(pNode))
  
  // 解析脚注/尾注引用
  notes.push(...parseNoteReferences(pNode))
  
  // 解析域代码
  fields.push(...parseFieldNodes(pNode))

  for (const c of childrenOf(pNode)) {
    const tn = tagNameOf(c)
    if (tn === "w:pPr") continue

    // 解析超链接
    if (tn === "w:hyperlink") {
      const attrs = attrsOf(c)
      const rid = attrOf(attrs, "r:id")
      const href = typeof rid === "string" ? hyperlinkRels.get(rid) : undefined
      if (href) hyperlinkTargetsInPara.add(href)

      for (const hc of childrenOf(c)) {
        if (tagNameOf(hc) === "w:r") {
          const rr = parseRunNode(hc, styles, docDefaultRun, paraStyleRun, paraProps.runDefaults)
          if (rr.text) runs.push(rr)
        }
      }
      continue
    }

    // 解析普通 run
    if (tn === "w:r") {
      // 检查 run 中的图片
      for (const rc of childrenOf(c)) {
        const rctn = tagNameOf(rc)
        if (rctn === "w:drawing") {
          images.push(...parseDrawingNode(rc, hyperlinkRels))
        } else if (rctn === "w:pict") {
          images.push(...parsePictNode(rc))
        }
      }
      
      const rr = parseRunNode(c, styles, docDefaultRun, paraStyleRun, paraProps.runDefaults)
      if (rr.text) runs.push(rr)
      continue
    }
  }

  const link = hyperlinkTargetsInPara.size === 1 ? [...hyperlinkTargetsInPara][0] : undefined

  if (runs.length === 0) {
    const effective = mergeDefined(docDefaultRun, paraStyleRun, paraProps.runDefaults)
    const emptyRun: RunStyle = {
      text: "",
      bold: effective.bold,
      italic: effective.italic,
      underline: effective.underline,
      fontSize: effective.fontSize,
      color: effective.color,
      font: effective.font
    }
    const para = buildParagraphFromRuns([emptyRun], { alignment, headingLevel, link, spacing })
    
    // 添加扩展属性
    if (images.length) para.images = images
    if (numbering) para.numbering = numbering
    if (bookmarks.length) para.bookmarks = bookmarks
    if (fields.length) para.fields = fields
    if (notes.length) para.notes = notes
    if (comments.length) para.comments = comments
    if (indent) para.indent = indent
    if (paraProps.keepNext) para.keepNext = paraProps.keepNext
    if (paraProps.keepLines) para.keepLines = paraProps.keepLines
    if (paraProps.pageBreakBefore) para.pageBreakBefore = paraProps.pageBreakBefore
    if (paraProps.widowControl !== undefined) para.widowControl = paraProps.widowControl
    if (paraProps.outlineLevel !== undefined) para.outlineLevel = paraProps.outlineLevel
    if (paraStyleId) para.styleId = paraStyleId
    
    return para
  }

  const para = buildParagraphFromRuns(runs, { alignment, headingLevel, link, spacing })
  
  // 添加扩展属性
  if (images.length) para.images = images
  if (numbering) para.numbering = numbering
  if (bookmarks.length) para.bookmarks = bookmarks
  if (fields.length) para.fields = fields
  if (notes.length) para.notes = notes
  if (comments.length) para.comments = comments
  if (indent) para.indent = indent
  if (paraProps.keepNext) para.keepNext = paraProps.keepNext
  if (paraProps.keepLines) para.keepLines = paraProps.keepLines
  if (paraProps.pageBreakBefore) para.pageBreakBefore = paraProps.pageBreakBefore
  if (paraProps.widowControl !== undefined) para.widowControl = paraProps.widowControl
  if (paraProps.outlineLevel !== undefined) para.outlineLevel = paraProps.outlineLevel
  if (paraStyleId) para.styleId = paraStyleId
  
  return para
}

function parseRunNode(
  rNode: OrderedXmlNode,
  styles: StyleMap,
  docDefaultRun: Partial<RunStyle>,
  paraStyleRun: Partial<RunStyle>,
  paraRunDefaults: Partial<RunStyle> | undefined
): RunStyle {
  const rPr = childOf(rNode, "w:rPr")
  const { props: direct, charStyleId } = parseRunPropsFromOrdered(rPr)
  const charStyleRun = resolveStyleChain(charStyleId, styles, "run")

  const effective = mergeDefined(docDefaultRun, paraStyleRun, paraRunDefaults, charStyleRun, direct)

  const text = childrenOf(rNode)
    .filter((c) => tagNameOf(c) !== "w:rPr")
    .map(textFromOrdered)
    .join("")

  return {
    text,
    bold: effective.bold,
    italic: effective.italic,
    underline: effective.underline,
    fontSize: effective.fontSize,
    color: effective.color,
    font: effective.font
  }
}

function parseBorderSpecFromNode(n: OrderedXmlNode | undefined): BorderSpec | undefined {
  if (!n) return undefined
  const attrs = attrsOf(n)
  const styleRaw = String(attrOf(attrs, "w:val") ?? "").trim().toLowerCase()
  const sizeRaw = Number.parseInt(String(attrOf(attrs, "w:sz") ?? ""), 10)
  const colorRaw = normalizeBorderColor(attrOf(attrs, "w:color"))

  const out: BorderSpec = {}
  if (styleRaw) out.style = styleRaw
  if (Number.isFinite(sizeRaw)) out.size = sizeRaw
  if (colorRaw) out.color = colorRaw
  return Object.keys(out).length ? out : undefined
}

function parseTableBordersFromNode(bordersNode: OrderedXmlNode | undefined): TableBordersSpec | undefined {
  if (!bordersNode) return undefined
  const out: TableBordersSpec = {
    top: parseBorderSpecFromNode(childOf(bordersNode, "w:top")),
    bottom: parseBorderSpecFromNode(childOf(bordersNode, "w:bottom")),
    left: parseBorderSpecFromNode(childOf(bordersNode, "w:left")),
    right: parseBorderSpecFromNode(childOf(bordersNode, "w:right")),
    insideHorizontal: parseBorderSpecFromNode(childOf(bordersNode, "w:insideH")),
    insideVertical: parseBorderSpecFromNode(childOf(bordersNode, "w:insideV"))
  }
  return Object.values(out).some((v) => v != null) ? out : undefined
}

function parseCellBordersFromNode(bordersNode: OrderedXmlNode | undefined): CellBordersSpec | undefined {
  if (!bordersNode) return undefined
  const out: CellBordersSpec = {
    top: parseBorderSpecFromNode(childOf(bordersNode, "w:top")),
    bottom: parseBorderSpecFromNode(childOf(bordersNode, "w:bottom")),
    left: parseBorderSpecFromNode(childOf(bordersNode, "w:left")),
    right: parseBorderSpecFromNode(childOf(bordersNode, "w:right"))
  }
  return Object.values(out).some((v) => v != null) ? out : undefined
}

function parseSectionPropertiesFromSectPr(sectPr: OrderedXmlNode | undefined): SectionPropertiesSpec | undefined {
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

function parseTableNode(
  tblNode: OrderedXmlNode,
  styles: StyleMap,
  docDefaultRun: Partial<RunStyle>,
  docDefaultPara: Pick<DocxParagraph, "alignment">,
  hyperlinkRels: Map<string, string>
): DocxParagraph {
  const tblPr = childOf(tblNode, "w:tblPr")
  const tblLayoutNode = tblPr ? childOf(tblPr, "w:tblLayout") : undefined
  const layoutTypeRaw = tblLayoutNode ? String(attrOf(attrsOf(tblLayoutNode), "w:type") ?? "").trim().toLowerCase() : ""
  const tableLayout: "fixed" | "autofit" | undefined = layoutTypeRaw === "fixed" ? "fixed" : layoutTypeRaw === "autofit" ? "autofit" : undefined

  const tblBordersNode = tblPr ? childOf(tblPr, "w:tblBorders") : undefined
  const tableBorders = parseTableBordersFromNode(tblBordersNode)
  
  // 解析表格样式引用
  const tblStyleNode = tblPr ? childOf(tblPr, "w:tblStyle") : undefined
  const tableStyleId = tblStyleNode ? String(attrOf(attrsOf(tblStyleNode), "w:val") ?? "") : undefined

  const tblGrid = childOf(tblNode, "w:tblGrid")
  const tableGridColsRaw = tblGrid
    ? childrenNamed(tblGrid, "w:gridCol").map((col) => Number.parseInt(String(attrOf(attrsOf(col), "w:w") ?? ""), 10))
    : []
  const tableGridCols = tableGridColsRaw.length ? tableGridColsRaw.filter((n) => Number.isFinite(n) && n > 0) : undefined

  const trs = childrenNamed(tblNode, "w:tr")
  const colCountFromGrid = tableGridCols?.length

  let maxCols = typeof colCountFromGrid === "number" && colCountFromGrid > 0 ? colCountFromGrid : 0
  if (maxCols === 0) {
    for (const tr of trs) {
      const trPr = childOf(tr, "w:trPr")
      const gridBeforeNode = trPr ? childOf(trPr, "w:gridBefore") : undefined
      const gridAfterNode = trPr ? childOf(trPr, "w:gridAfter") : undefined
      const gridBefore = Number.parseInt(String(attrOf(attrsOf(gridBeforeNode ?? {}), "w:val") ?? ""), 10)
      const gridAfter = Number.parseInt(String(attrOf(attrsOf(gridAfterNode ?? {}), "w:val") ?? ""), 10)
      let cols = 0
      if (Number.isFinite(gridBefore)) cols += gridBefore
      for (const tc of childrenNamed(tr, "w:tc")) {
        const tcPr = childOf(tc, "w:tcPr")
        const gridSpanNode = tcPr ? childOf(tcPr, "w:gridSpan") : undefined
        const spanRaw = Number.parseInt(String(attrOf(attrsOf(gridSpanNode ?? {}), "w:val") ?? ""), 10)
        cols += Number.isFinite(spanRaw) && spanRaw > 0 ? spanRaw : 1
      }
      if (Number.isFinite(gridAfter)) cols += gridAfter
      if (cols > maxCols) maxCols = cols
    }
  }

  const rows: string[][] = []
  const cellStyles: CellStyle[][] = []

  const mergeTracker = new Map<string, CellStyle>()

  for (const tr of trs) {
    const rowTexts = new Array(maxCols).fill("")
    const rowStyles = new Array(maxCols).fill(undefined).map(() => ({} as CellStyle))

    const trPr = childOf(tr, "w:trPr")
    const gridBeforeNode = trPr ? childOf(trPr, "w:gridBefore") : undefined
    const gridAfterNode = trPr ? childOf(trPr, "w:gridAfter") : undefined
    const gridBefore = Number.parseInt(String(attrOf(attrsOf(gridBeforeNode ?? {}), "w:val") ?? ""), 10)
    const gridAfter = Number.parseInt(String(attrOf(attrsOf(gridAfterNode ?? {}), "w:val") ?? ""), 10)

    const occupied = new Array(maxCols).fill(false)
    for (const [key, top] of mergeTracker.entries()) {
      const [colStartStr, spanStr] = key.split(":")
      const colStart = Number.parseInt(colStartStr ?? "", 10)
      const span = Number.parseInt(spanStr ?? "", 10)
      if (!Number.isFinite(colStart) || !Number.isFinite(span)) continue
      for (let c = colStart; c < Math.min(maxCols, colStart + span); c += 1) occupied[c] = true
      if (rowStyles[colStart]) rowStyles[colStart].skip = true
      for (let c = colStart + 1; c < Math.min(maxCols, colStart + span); c += 1) {
        if (rowStyles[c]) rowStyles[c].skip = true
      }
    }

    const beforeCols = Number.isFinite(gridBefore) && gridBefore > 0 ? gridBefore : 0
    for (let c = 0; c < Math.min(maxCols, beforeCols); c += 1) {
      occupied[c] = true
      rowStyles[c].skip = true
    }

    let colCursor = beforeCols

    for (const tc of childrenNamed(tr, "w:tc")) {
      while (colCursor < maxCols && occupied[colCursor]) colCursor += 1
      if (colCursor >= maxCols) break

      const tcPr = childOf(tc, "w:tcPr")
      const shading = tcPr ? childOf(tcPr, "w:shd") : undefined
      const tcBordersNode = tcPr ? childOf(tcPr, "w:tcBorders") : undefined
      const gridSpanNode = tcPr ? childOf(tcPr, "w:gridSpan") : undefined
      const vAlign = tcPr ? childOf(tcPr, "w:vAlign") : undefined
      const vMerge = tcPr ? childOf(tcPr, "w:vMerge") : undefined

      const fill = shading ? normalizeColor(attrOf(attrsOf(shading), "w:fill")) : undefined
      const borders = parseCellBordersFromNode(tcBordersNode)
      const spanRaw = Number.parseInt(String(attrOf(attrsOf(gridSpanNode ?? {}), "w:val") ?? ""), 10)
      const span = Number.isFinite(spanRaw) && spanRaw > 0 ? spanRaw : 1
      const verticalVal = vAlign ? String(attrOf(attrsOf(vAlign), "w:val") ?? "").trim().toLowerCase() : ""
      const vMergeValRaw = vMerge ? attrOf(attrsOf(vMerge), "w:val") : undefined
      const vMergeVal = vMerge ? String(vMergeValRaw ?? "").trim().toLowerCase() : undefined

      const colStart = colCursor
      const colEnd = Math.min(maxCols, colStart + span)
      for (let c = colStart; c < colEnd; c += 1) occupied[c] = true

      const mergeKey = `${colStart}:${span}`

      if (vMerge != null && vMergeVal !== "restart") {
        const top = mergeTracker.get(mergeKey)
        if (top) {
          top.rowSpan = (top.rowSpan ?? 1) + 1
          // 继承顶部单元格的样式到合并的单元格
          rowStyles[colStart] = {
            skip: true,
            fill: top.fill,
            borders: top.borders,
            bold: top.bold,
            italic: top.italic,
            fontSize: top.fontSize,
            color: top.color,
            font: top.font,
            alignment: top.alignment,
            verticalAlign: top.verticalAlign
          }
        } else {
          rowStyles[colStart].skip = true
        }
        for (let c = colStart + 1; c < colEnd; c += 1) rowStyles[c].skip = true
        colCursor = colEnd
        continue
      }

      mergeTracker.delete(mergeKey)

      const paras = childrenNamed(tc, "w:p").map((p) =>
        parseParagraphNode(p, styles, docDefaultRun, docDefaultPara, hyperlinkRels)
      )

      const cellText = paras.map((p) => p.text).join("\n")
      rowTexts[colStart] = cellText

      const firstRun = paras.find((p) => p.runs?.length)?.runs?.[0]
      const style: CellStyle = {
        fill,
        gridSpan: span > 1 ? span : undefined,
        colIndex: colStart,
        verticalAlign:
          verticalVal === "center" ? "center" : verticalVal === "bottom" ? "bottom" : verticalVal ? "top" : undefined,
        alignment: paras.find((p) => p.alignment)?.alignment,
        borders
      }

      if (firstRun) {
        style.bold = firstRun.bold
        style.italic = firstRun.italic
        style.fontSize = firstRun.fontSize
        style.color = firstRun.color
        style.font = firstRun.font
      }

      if (vMergeVal === "restart") {
        style.rowSpan = 1
        mergeTracker.set(mergeKey, style)
      }

      rowStyles[colStart] = style
      for (let c = colStart + 1; c < colEnd; c += 1) rowStyles[c] = { skip: true }

      colCursor = colEnd
    }

    const afterCols = Number.isFinite(gridAfter) && gridAfter > 0 ? gridAfter : 0
    if (afterCols > 0) {
      const start = Math.max(0, maxCols - afterCols)
      for (let c = start; c < maxCols; c += 1) {
        rowStyles[c].skip = true
      }
    }

    rows.push(rowTexts)
    cellStyles.push(rowStyles)
  }

  return { text: "", isTable: true, tableData: rows, tableCellStyles: cellStyles, tableLayout, tableGridCols, tableBorders, tableStyleId }
}

export async function extractParagraphsFromDocx(docxBuffer: Buffer): Promise<DocxParagraph[]> {
  if (!docxBuffer || docxBuffer.length === 0) return []

  const [stylesXmlBuf, documentXmlBuf, relsBuf, numberingXmlBuf, commentsXmlBuf, footnotesXmlBuf, endnotesXmlBuf] = await Promise.all([
    readZipEntry(docxBuffer, "word/styles.xml"),
    readZipEntry(docxBuffer, "word/document.xml"),
    readZipEntry(docxBuffer, "word/_rels/document.xml.rels"),
    readZipEntry(docxBuffer, "word/numbering.xml"),
    readZipEntry(docxBuffer, "word/comments.xml"),
    readZipEntry(docxBuffer, "word/footnotes.xml"),
    readZipEntry(docxBuffer, "word/endnotes.xml")
  ])

  if (!documentXmlBuf) return []

  const { styles, docDefaultRun, docDefaultPara } = stylesXmlBuf
    ? parseStyles(stylesXmlBuf.toString("utf-8"))
    : { styles: new Map(), docDefaultRun: {}, docDefaultPara: { alignment: undefined } }

  const hyperlinkRels = relsBuf ? parseRelationships(relsBuf.toString("utf-8")) : new Map()
  
  const numberingDefs = numberingXmlBuf ? parseNumbering(numberingXmlBuf.toString("utf-8")) : new Map()

  const parser = new XMLParser({ ignoreAttributes: false, preserveOrder: true })
  const ordered: any[] = parser.parse(documentXmlBuf.toString("utf-8"))

  const docNode = ordered.find((n) => tagNameOf(n) === "w:document")
  if (!docNode) return []
  const bodyNode = childOf(docNode, "w:body")
  if (!bodyNode) return []

  const bodyChildren = childrenOf(bodyNode)

  const outReversed: DocxParagraph[] = []
  let currentSectPr: OrderedXmlNode | undefined

  for (let i = bodyChildren.length - 1; i >= 0; i -= 1) {
    const c = bodyChildren[i] as OrderedXmlNode
    const tn = tagNameOf(c)
    
    if (tn === "w:sectPr") {
      currentSectPr = c
      continue
    }
    
    if (tn === "w:p") {
      const pPr = childOf(c, "w:pPr")
      const s = pPr ? childOf(pPr, "w:sectPr") : undefined
      if (s) currentSectPr = s
      const para = parseParagraphNode(c, styles, docDefaultRun, docDefaultPara, hyperlinkRels)
      
      // 补充编号格式信息
      if (para.numbering?.numId !== undefined) {
        const numId = String(para.numbering.numId)
        const level = para.numbering.level ?? 0
        const levelDef = numberingDefs.get(numId)?.get(level)
        if (levelDef) {
          para.numbering.format = levelDef.format
          para.numbering.text = levelDef.text
        }
      }
      
      const sectionProperties = parseSectionPropertiesFromSectPr(currentSectPr)
      if (sectionProperties) para.sectionProperties = sectionProperties
      outReversed.push(para)
      continue
    }
    
    if (tn === "w:tbl") {
      const tbl = parseTableNode(c, styles, docDefaultRun, docDefaultPara, hyperlinkRels)
      const sectionProperties = parseSectionPropertiesFromSectPr(currentSectPr)
      if (sectionProperties) tbl.sectionProperties = sectionProperties
      outReversed.push(tbl)
      continue
    }
    
    // 处理其他未识别的元素（保留为占位符）
    // 这样可以保持文档结构的完整性
    if (tn && tn !== "w:bookmarkStart" && tn !== "w:bookmarkEnd") {
      // 创建一个占位段落来表示未解析的元素
      const placeholder: DocxParagraph = {
        text: `[未解析元素: ${tn}]`,
        color: "999999",
        italic: true,
        fontSize: 9
      }
      outReversed.push(placeholder)
    }
  }

  outReversed.reverse()
  return outReversed
}

/**
 * 扁平化段落列表，将表格行展开为独立的流式单元
 * 这样可以逐行流式输出表格内容
 */
export interface FlattenedElement {
  type: "paragraph" | "table-row"
  paragraph?: DocxParagraph
  tableContext?: {
    tableIndex: number
    rowIndex: number
    totalRows: number
    tableParagraph: DocxParagraph
  }
}

export function flattenParagraphsForStreaming(paragraphs: DocxParagraph[]): FlattenedElement[] {
  const flattened: FlattenedElement[] = []
  let tableIndex = 0
  
  for (const para of paragraphs) {
    if (para.isTable && para.tableData) {
      // 将表格的每一行作为独立的流式单元
      const totalRows = para.tableData.length
      for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
        flattened.push({
          type: "table-row",
          tableContext: {
            tableIndex,
            rowIndex,
            totalRows,
            tableParagraph: para
          }
        })
      }
      tableIndex++
    } else {
      // 普通段落直接添加
      flattened.push({
        type: "paragraph",
        paragraph: para
      })
    }
  }
  
  return flattened
}

/**
 * 根据扁平化的元素列表生成 docx 切片
 * 支持逐行流式输出表格
 */
export async function createDocxFromFlattenedSlice(
  sourceDocxBuffer: Buffer,
  flattenedElements: FlattenedElement[],
  endIndex: number
): Promise<Buffer> {
  if (!sourceDocxBuffer || sourceDocxBuffer.length === 0) return Buffer.alloc(0)
  if (endIndex <= 0) return Buffer.alloc(0)

  const documentXmlBuf = await readZipEntry(sourceDocxBuffer, "word/document.xml")
  if (!documentXmlBuf) return Buffer.alloc(0)

  const parser = new XMLParser({ ignoreAttributes: false, preserveOrder: true })
  const ordered: any[] = parser.parse(documentXmlBuf.toString("utf-8"))

  const docNode = ordered.find((n) => tagNameOf(n) === "w:document")
  if (!docNode) return Buffer.alloc(0)
  const bodyNode = childOf(docNode, "w:body")
  if (!bodyNode) return Buffer.alloc(0)

  const bodyChildren = childrenOf(bodyNode)
  const cloneNode = <T,>(v: T): T => JSON.parse(JSON.stringify(v)) as T

  const findSectPrFromIndex = (startIndex: number): OrderedXmlNode | undefined => {
    for (let i = Math.max(0, startIndex); i < bodyChildren.length; i += 1) {
      const n = bodyChildren[i] as OrderedXmlNode
      const tn = tagNameOf(n)
      if (tn === "w:sectPr") return n
      if (tn === "w:p") {
        const pPr = childOf(n, "w:pPr")
        const s = pPr ? childOf(pPr, "w:sectPr") : undefined
        if (s) return s
      }
    }
    return undefined
  }

  const newBodyChildren: OrderedXmlNode[] = []
  let flatIndex = 0
  let bodyIndex = 0
  let lastIncludedBodyIndex = -1
  const tableRowsToInclude = new Map<number, number>() // tableBodyIndex -> rowCount

  // 第一遍：确定每个表格需要包含多少行
  for (let i = 0; i < Math.min(endIndex, flattenedElements.length); i++) {
    const elem = flattenedElements[i]!
    if (elem.type === "table-row" && elem.tableContext) {
      const key = elem.tableContext.tableIndex
      const currentCount = tableRowsToInclude.get(key) || 0
      tableRowsToInclude.set(key, currentCount + 1)
    }
  }

  // 第二遍：构建新的 body
  let currentTableIndex = -1
  for (let i = 0; i < bodyChildren.length && flatIndex < endIndex; i++) {
    const n = bodyChildren[i] as OrderedXmlNode
    const tn = tagNameOf(n)
    if (tn === "w:sectPr") continue

    if (tn === "w:tbl") {
      currentTableIndex++
      const rowsToInclude = tableRowsToInclude.get(currentTableIndex)
      
      if (rowsToInclude && rowsToInclude > 0) {
        // 构建包含部分行的表格
        const tblPr = childOf(n, "w:tblPr")
        const tblGrid = childOf(n, "w:tblGrid")
        const allRows = childrenNamed(n, "w:tr")
        
        const partialTable = cloneNode(n)
        const tableKey = "w:tbl"
        const tableChildren: OrderedXmlNode[] = []
        
        if (tblPr) tableChildren.push(cloneNode(tblPr))
        if (tblGrid) tableChildren.push(cloneNode(tblGrid))
        
        // 只包含需要的行数
        for (let r = 0; r < Math.min(rowsToInclude, allRows.length); r++) {
          tableChildren.push(cloneNode(allRows[r]!))
          flatIndex++
        }
        
        partialTable[tableKey] = tableChildren
        newBodyChildren.push(partialTable)
        lastIncludedBodyIndex = i
      }
      
      bodyIndex++
      continue
    }

    if (tn === "w:p") {
      newBodyChildren.push(cloneNode(n))
      lastIncludedBodyIndex = i
      flatIndex++
      bodyIndex++
    }
  }

  const sectPrForChunk = findSectPrFromIndex(lastIncludedBodyIndex)
  if (sectPrForChunk) newBodyChildren.push(cloneNode(sectPrForChunk))

  const bodyKey = "w:body"
  bodyNode[bodyKey] = newBodyChildren

  const newDocumentXml = buildXml(ordered, { ignoreAttributes: false, preserveOrder: true, format: false })

  return await replaceZipEntry(sourceDocxBuffer, "word/document.xml", Buffer.from(newDocumentXml, "utf-8"))
}

/**
 * 从源 docx 创建切片，支持逐行流式输出表格
 * 
 * @param sourceDocxBuffer 源 docx 文件的 Buffer
 * @param bodyElementCount 要包含的 body 元素数量（段落按1计，表格的每一行也按1计）
 * @returns 包含指定数量元素的 docx Buffer
 */
export async function createDocxFromSourceDocxSlice(sourceDocxBuffer: Buffer, bodyElementCount: number): Promise<Buffer> {
  if (!sourceDocxBuffer || sourceDocxBuffer.length === 0) return Buffer.alloc(0)

  const documentXmlBuf = await readZipEntry(sourceDocxBuffer, "word/document.xml")
  if (!documentXmlBuf) return Buffer.alloc(0)

  const parser = new XMLParser({ ignoreAttributes: false, preserveOrder: true })
  const ordered: any[] = parser.parse(documentXmlBuf.toString("utf-8"))

  const docNode = ordered.find((n) => tagNameOf(n) === "w:document")
  if (!docNode) return Buffer.alloc(0)
  const bodyNode = childOf(docNode, "w:body")
  if (!bodyNode) return Buffer.alloc(0)

  const bodyChildren = childrenOf(bodyNode)
  const cloneNode = <T,>(v: T): T => JSON.parse(JSON.stringify(v)) as T

  const findSectPrFromIndex = (startIndex: number): OrderedXmlNode | undefined => {
    for (let i = Math.max(0, startIndex); i < bodyChildren.length; i += 1) {
      const n = bodyChildren[i] as OrderedXmlNode
      const tn = tagNameOf(n)
      if (tn === "w:sectPr") return n
      if (tn === "w:p") {
        const pPr = childOf(n, "w:pPr")
        const s = pPr ? childOf(pPr, "w:sectPr") : undefined
        if (s) return s
      }
    }
    return undefined
  }

  const newBodyChildren: OrderedXmlNode[] = []
  let included = 0
  let lastIncludedBodyIndex = -1
  
  for (let i = 0; i < bodyChildren.length && included < bodyElementCount; i += 1) {
    const n = bodyChildren[i] as OrderedXmlNode
    const tn = tagNameOf(n)
    if (tn === "w:sectPr") continue
    
    // 处理表格：逐行包含
    if (tn === "w:tbl") {
      const tblPr = childOf(n, "w:tblPr")
      const tblGrid = childOf(n, "w:tblGrid")
      const allRows = childrenNamed(n, "w:tr")
      
      // 计算还能包含多少行
      const remainingCount = bodyElementCount - included
      const rowsToInclude = Math.min(remainingCount, allRows.length)
      
      if (rowsToInclude > 0) {
        // 构建包含部分行的表格
        const partialTable = cloneNode(n)
        const tableKey = "w:tbl"
        const tableChildren: OrderedXmlNode[] = []
        
        // 添加表格属性和网格定义
        if (tblPr) tableChildren.push(cloneNode(tblPr))
        if (tblGrid) tableChildren.push(cloneNode(tblGrid))
        
        // 添加指定数量的行
        for (let r = 0; r < rowsToInclude; r++) {
          tableChildren.push(cloneNode(allRows[r]!))
        }
        
        partialTable[tableKey] = tableChildren
        newBodyChildren.push(partialTable)
        lastIncludedBodyIndex = i
        included += rowsToInclude
      }
      
      continue
    }
    
    // 处理普通段落
    if (tn === "w:p") {
      newBodyChildren.push(cloneNode(n))
      lastIncludedBodyIndex = i
      included += 1
    }
  }
  
  const sectPrForChunk = findSectPrFromIndex(lastIncludedBodyIndex)
  if (sectPrForChunk) newBodyChildren.push(cloneNode(sectPrForChunk))

  const bodyKey = "w:body"
  bodyNode[bodyKey] = newBodyChildren

  const newDocumentXml = buildXml(ordered, { ignoreAttributes: false, preserveOrder: true, format: false })

  return await replaceZipEntry(sourceDocxBuffer, "word/document.xml", Buffer.from(newDocumentXml, "utf-8"))
}

/**
 * 根据段落生成完整的 docx 文件
 * 
 * 改进点（参考 Python improved_replacer.py）：
 * 1. 支持多 run 段落（处理混合样式）
 * 2. 更精确的表格单元格样式处理
 * 3. 保留空段落的格式信息
 * 4. 更好的字体和颜色处理
 */
export async function createDocxFromParagraphs(paragraphs: DocxParagraph[]): Promise<Buffer> {
  const children: (Paragraph | Table)[] = []

  const toDocxBorderStyle = (style: string | undefined) => {
    const s = String(style ?? "").trim().toLowerCase()
    if (!s || s === "single") return BorderStyle.SINGLE
    if (s === "none" || s === "nil") return BorderStyle.NONE
    if (s === "dotted") return BorderStyle.DOTTED
    if (s === "dashed") return BorderStyle.DASHED
    if (s === "double") return BorderStyle.DOUBLE
    if (s === "thick") return BorderStyle.THICK
    return BorderStyle.SINGLE
  }

  const toDocxBorder = (spec: BorderSpec | undefined): any => {
    if (!spec) return undefined
    const style = toDocxBorderStyle(spec.style)
    const size = Number.isFinite(spec.size as number) ? (spec.size as number) : style === BorderStyle.NONE ? 0 : 4
    const color = spec.color ? spec.color : "auto"
    return { style, size, color }
  }

  for (const para of paragraphs) {
    // 处理表格
    if (para.isTable && para.tableData) {
      const widthTypeDxa = (WidthType as any).DXA ?? WidthType.PERCENTAGE
      const gridCols = (para.tableGridCols ?? []).filter((n) => Number.isFinite(n) && n > 0)
      const hasGrid = gridCols.length > 0
      const totalGridWidth = hasGrid ? gridCols.reduce((a, b) => a + b, 0) : 0

      const tableRows = para.tableData.map((rowData, rowIndex) => {
        const cells: TableCell[] = []
        const rowStyles = para.tableCellStyles?.[rowIndex] || []
        const rowLen = Array.isArray(rowData) ? rowData.length : 0

        for (let cellIndex = 0; cellIndex < rowLen; cellIndex += 1) {
          const cellText = rowData[cellIndex]
          const cellStyle = rowStyles[cellIndex] || {}
          if (cellStyle.skip) continue

          const paragraphAlignment =
            cellStyle.alignment === "center"
              ? AlignmentType.CENTER
              : cellStyle.alignment === "right"
              ? AlignmentType.RIGHT
              : cellStyle.alignment === "justify"
              ? ((AlignmentType as any).JUSTIFIED ?? (AlignmentType as any).BOTH ?? AlignmentType.LEFT)
              : AlignmentType.LEFT

          const lines = String(cellText ?? "").split("\n")
          const cellRuns = lines.map((line, lineIndex) => {
            const runOptions: any = {
              text: line,
              bold: cellStyle.bold ?? false,
              italics: cellStyle.italic ?? false,
              size: cellStyle.fontSize ? cellStyle.fontSize * 2 : 20,
              color: cellStyle.color,
              font: cellStyle.font ? { name: cellStyle.font } : undefined
            }
            if (lineIndex > 0) runOptions.break = 1
            return new TextRun(runOptions as IRunOptions)
          })

          const span = typeof cellStyle.gridSpan === "number" && cellStyle.gridSpan > 0 ? cellStyle.gridSpan : 1
          const colIndex = typeof cellStyle.colIndex === "number" && cellStyle.colIndex >= 0 ? cellStyle.colIndex : cellIndex

          let cellWidth: number | undefined
          if (hasGrid) {
            const seg = gridCols.slice(colIndex, colIndex + span)
            const w = seg.reduce((a, b) => a + b, 0)
            if (w > 0) cellWidth = w
          }

          const cellOptions: any = {
            children: [
              new Paragraph({
                children: cellRuns,
                alignment: paragraphAlignment
              })
            ],
            verticalAlign:
              cellStyle.verticalAlign === "center" ? "center" : cellStyle.verticalAlign === "bottom" ? "bottom" : "top"
          }

          if (cellWidth != null) {
            cellOptions.width = { size: cellWidth, type: widthTypeDxa }
          } else {
            const denom = rowLen || 1
            cellOptions.width = { size: Math.max(1, Math.floor(100 / denom)), type: WidthType.PERCENTAGE }
          }

          if (cellStyle.fill) {
            cellOptions.shading = { fill: cellStyle.fill }
          }

          if (cellStyle.borders) {
            const borders: any = {}
            const top = toDocxBorder(cellStyle.borders.top)
            const bottom = toDocxBorder(cellStyle.borders.bottom)
            const left = toDocxBorder(cellStyle.borders.left)
            const right = toDocxBorder(cellStyle.borders.right)
            if (top) borders.top = top
            if (bottom) borders.bottom = bottom
            if (left) borders.left = left
            if (right) borders.right = right
            if (Object.keys(borders).length) cellOptions.borders = borders
          }

          if (span > 1) cellOptions.columnSpan = span
          if (typeof cellStyle.rowSpan === "number" && cellStyle.rowSpan > 1) (cellOptions as any).rowSpan = cellStyle.rowSpan

          cells.push(new TableCell(cellOptions))
        }

        return new TableRow({ children: cells })
      })

      const defaultBorders: any = {
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" }
      }

      const tableBorders: any = para.tableBorders
        ? {
            top: toDocxBorder(para.tableBorders.top) ?? defaultBorders.top,
            bottom: toDocxBorder(para.tableBorders.bottom) ?? defaultBorders.bottom,
            left: toDocxBorder(para.tableBorders.left) ?? defaultBorders.left,
            right: toDocxBorder(para.tableBorders.right) ?? defaultBorders.right,
            insideHorizontal: toDocxBorder(para.tableBorders.insideHorizontal) ?? defaultBorders.insideHorizontal,
            insideVertical: toDocxBorder(para.tableBorders.insideVertical) ?? defaultBorders.insideVertical
          }
        : defaultBorders

      children.push(
        new Table({
          rows: tableRows,
          width: hasGrid ? { size: totalGridWidth, type: widthTypeDxa } : { size: 100, type: WidthType.PERCENTAGE },
          borders: tableBorders
        })
      )
      continue
    }

    // 处理普通段落 - 支持多 run（混合样式）
    let textRuns: TextRun[]
    
    if (para.runs && para.runs.length > 0) {
      // 使用多个 run（用于混合样式段落）
      textRuns = para.runs.map(run => {
        const runOptions: IRunOptions = {
          text: run.text,
          bold: run.bold ?? false,
          italics: run.italic ?? false,
          underline: run.underline ? { type: UnderlineType.SINGLE } : undefined,
          size: run.fontSize ? run.fontSize * 2 : 22,
          color: run.color,
          font: run.font ? { name: run.font } : undefined
        }
        return new TextRun(runOptions)
      })
    } else if (para.link) {
      // 超链接样式
      textRuns = [
        new TextRun({
          text: para.text,
          bold: para.bold ?? false,
          italics: para.italic ?? false,
          size: para.fontSize ? para.fontSize * 2 : 22,
          style: "Hyperlink",
          color: "0563C1",
          underline: { type: UnderlineType.SINGLE },
          font: para.font ? { name: para.font } : undefined
        })
      ]
    } else {
      // 单一样式段落
      const runOptions: IRunOptions = {
        text: para.text,
        bold: para.bold ?? false,
        italics: para.italic ?? false,
        underline: para.underline ? { type: UnderlineType.SINGLE } : undefined,
        size: para.fontSize ? para.fontSize * 2 : 22,
        color: para.color,
        font: para.font ? { name: para.font } : undefined
      }
      textRuns = [new TextRun(runOptions)]
    }

    const paragraphOptions: any = {
      children: textRuns
    }
    if (para.spacing) paragraphOptions.spacing = para.spacing

    if (para.headingLevel) {
      paragraphOptions.heading = HeadingLevel[`HEADING_${para.headingLevel}`]
    }

    if (para.alignment) {
      paragraphOptions.alignment =
        para.alignment === "center"
          ? AlignmentType.CENTER
          : para.alignment === "right"
          ? AlignmentType.RIGHT
          : para.alignment === "justify"
          ? ((AlignmentType as any).JUSTIFIED ?? (AlignmentType as any).BOTH ?? AlignmentType.LEFT)
          : AlignmentType.LEFT
    }

    children.push(new Paragraph(paragraphOptions))
  }

  let sectionSpec: SectionPropertiesSpec | undefined
  for (const p of paragraphs) {
    if (p.sectionProperties) sectionSpec = p.sectionProperties
  }
  const sectionProperties: any = {}
  if (sectionSpec?.page?.size || sectionSpec?.page?.margin) {
    const page: any = {}
    if (sectionSpec.page?.size) {
      const size: any = {}
      if (typeof sectionSpec.page.size.width === "number") size.width = sectionSpec.page.size.width
      if (typeof sectionSpec.page.size.height === "number") size.height = sectionSpec.page.size.height
      if (sectionSpec.page.size.orientation === "landscape") size.orientation = PageOrientation.LANDSCAPE
      if (sectionSpec.page.size.orientation === "portrait") size.orientation = PageOrientation.PORTRAIT
      if (Object.keys(size).length) page.size = size
    }
    if (sectionSpec.page?.margin) {
      const margin: any = {}
      const m = sectionSpec.page.margin
      if (typeof m.top === "number") margin.top = m.top
      if (typeof m.right === "number") margin.right = m.right
      if (typeof m.bottom === "number") margin.bottom = m.bottom
      if (typeof m.left === "number") margin.left = m.left
      if (typeof m.header === "number") margin.header = m.header
      if (typeof m.footer === "number") margin.footer = m.footer
      if (typeof m.gutter === "number") margin.gutter = m.gutter
      if (Object.keys(margin).length) page.margin = margin
    }
    if (Object.keys(page).length) sectionProperties.page = page
  }
  if (sectionSpec?.column) {
    const column: any = {}
    if (typeof sectionSpec.column.count === "number") column.count = sectionSpec.column.count
    if (typeof sectionSpec.column.space === "number") column.space = sectionSpec.column.space
    if (Object.keys(column).length) sectionProperties.column = column
  }

  const doc = new Document({
    sections: [
      {
        properties: sectionProperties,
        children
      }
    ]
  })

  return Buffer.from(await Packer.toBuffer(doc))
}
