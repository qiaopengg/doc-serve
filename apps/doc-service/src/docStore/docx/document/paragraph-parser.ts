import type { 
  DocxParagraph, 
  RunStyle, 
  StyleMap, 
  OrderedXmlNode,
  ImageSpec,
  BookmarkSpec,
  FieldSpec,
  NoteSpec,
  CommentSpec,
  MathSpec,
  ParagraphSpacing,
  NumberingSpec
} from "../types.js"
import { 
  tagNameOf, 
  attrsOf, 
  attrOf, 
  childrenOf, 
  childOf, 
  childrenNamed, 
  textFromOrdered,
  mergeDefined,
  normalizeColor,
  normalizeAlignment,
  parseBooleanOnOff
} from "../core/utils.js"

export function parseParagraphNode(
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

  bookmarks.push(...parseBookmarkNodes(pNode))
  comments.push(...parseCommentNodes(pNode))
  notes.push(...parseNoteReferences(pNode))
  fields.push(...parseFieldNodes(pNode))

  for (const c of childrenOf(pNode)) {
    const tn = tagNameOf(c)
    if (tn === "w:pPr") continue

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

    if (tn === "w:r") {
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

export function parseRunNode(
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

export function parseRunPropsFromOrdered(rPr: OrderedXmlNode | undefined): { props: Partial<RunStyle>; charStyleId?: string } {
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

export function parseParaPropsFromOrdered(pPr: OrderedXmlNode | undefined): {
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

export function buildParagraphFromRuns(
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

export function parseDrawingNode(drawingNode: OrderedXmlNode, hyperlinkRels: Map<string, string>): ImageSpec[] {
  const images: ImageSpec[] = []
  
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
      width: Number.isFinite(cxRaw) ? cxRaw / 914400 * 72 : undefined,
      height: Number.isFinite(cyRaw) ? cyRaw / 914400 * 72 : undefined,
      description: typeof descr === "string" ? descr : undefined,
      title: typeof title === "string" ? title : undefined
    })
  }
  
  return images
}

export function parsePictNode(pictNode: OrderedXmlNode): ImageSpec[] {
  const images: ImageSpec[] = []
  
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

export function parseBookmarkNodes(pNode: OrderedXmlNode): BookmarkSpec[] {
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

export function parseFieldNodes(pNode: OrderedXmlNode): FieldSpec[] {
  const fields: FieldSpec[] = []
  
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
          // Start collecting result
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

export function parseCommentNodes(pNode: OrderedXmlNode): CommentSpec[] {
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

export function parseNoteReferences(pNode: OrderedXmlNode): NoteSpec[] {
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

export function parseMathNodes(pNode: OrderedXmlNode): MathSpec[] {
  const maths: MathSpec[] = []
  
  for (const c of childrenOf(pNode)) {
    if (tagNameOf(c) === "m:oMath" || tagNameOf(c) === "m:oMathPara") {
      maths.push({ type: "math", omml: JSON.stringify(c) })
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
