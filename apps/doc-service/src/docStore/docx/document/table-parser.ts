/**
 * 表格解析器
 * 从 docxGenerator.ts 迁移的表格解析逻辑
 * 
 * 核心功能：
 * - 解析表格属性（布局、边框、样式ID）
 * - 解析单元格样式（底纹、边框、对齐）
 * - 处理单元格合并（垂直、水平）
 * - 保留样式继承逻辑（合并单元格继承顶部样式）
 */

import type {
  OrderedXmlNode,
  DocxParagraph,
  CellStyle,
  BorderSpec,
  TableBordersSpec,
  CellBordersSpec,
  StyleMap,
  RunStyle
} from "../types.js"
import {
  childOf,
  childrenNamed,
  attrsOf,
  attrOf,
  normalizeColor,
  normalizeBorderColor,
  childrenOf,
  tagNameOf
} from "../core/utils.js"

/**
 * 解析边框规格
 * @param n - 边框节点
 * @returns 边框规格或 undefined
 */
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

/**
 * 解析表格边框
 * @param bordersNode - 表格边框节点
 * @returns 表格边框规格或 undefined
 */
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

/**
 * 解析单元格边框
 * @param bordersNode - 单元格边框节点
 * @returns 单元格边框规格或 undefined
 */
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

/**
 * 表格解析器类
 * 负责解析 OOXML 表格节点并转换为 DocxParagraph 格式
 */
export class TableParser {
  /**
   * 解析表格节点（静态方法）
   * 
   * @param tblNode - 表格节点（w:tbl）
   * @param styles - 样式映射表
   * @param docDefaultRun - 文档默认运行样式
   * @param docDefaultPara - 文档默认段落样式
   * @param hyperlinkRels - 超链接关系映射
   * @param parseParagraphNode - 段落解析函数（用于解析单元格内的段落）
   * @returns 表格的 DocxParagraph 表示
   * 
   * 关键特性：
   * - 保留 tableStyleId 解析
   * - 保留单元格合并的样式继承逻辑（合并单元格继承顶部样式）
   * - 支持 gridBefore 和 gridAfter
   * - 支持垂直和水平合并
   */
  static parseTableNode(
    tblNode: OrderedXmlNode,
    styles: StyleMap,
    docDefaultRun: Partial<RunStyle>,
    docDefaultPara: Pick<DocxParagraph, "alignment">,
    hyperlinkRels: Map<string, string>,
    parseParagraphNode: (
      pNode: OrderedXmlNode,
      styles: StyleMap,
      docDefaultRun: Partial<RunStyle>,
      docDefaultPara: Pick<DocxParagraph, "alignment">,
      hyperlinkRels: Map<string, string>
    ) => DocxParagraph
  ): DocxParagraph {
    // 提取表格节点的子节点（处理 { "w:tbl": [...] } 格式）
    const tblChildren = childrenOf(tblNode)
    
    // 从子节点中查找各个部分
    const tblPr = tblChildren.find(child => tagNameOf(child) === "w:tblPr")
    const tblGrid = tblChildren.find(child => tagNameOf(child) === "w:tblGrid")
    const trs = tblChildren.filter(child => tagNameOf(child) === "w:tr")
    
    // 解析表格属性
    const tblLayoutNode = tblPr ? childOf(tblPr, "w:tblLayout") : undefined
    const layoutTypeRaw = tblLayoutNode ? String(attrOf(attrsOf(tblLayoutNode), "w:type") ?? "").trim().toLowerCase() : ""
    const tableLayout: "fixed" | "autofit" | undefined = layoutTypeRaw === "fixed" ? "fixed" : layoutTypeRaw === "autofit" ? "autofit" : undefined

    // 解析表格边框
    const tblBordersNode = tblPr ? childOf(tblPr, "w:tblBorders") : undefined
    const tableBorders = parseTableBordersFromNode(tblBordersNode)
    
    // 解析表格样式引用（关键：保留 tableStyleId）
    const tblStyleNode = tblPr ? childOf(tblPr, "w:tblStyle") : undefined
    const tableStyleId = tblStyleNode ? String(attrOf(attrsOf(tblStyleNode), "w:val") ?? "") : undefined

    // 解析表格网格列
    const tableGridColsRaw = tblGrid
      ? childrenNamed(tblGrid, "w:gridCol").map((col) => Number.parseInt(String(attrOf(attrsOf(col), "w:w") ?? ""), 10))
      : []
    const tableGridCols = tableGridColsRaw.length ? tableGridColsRaw.filter((n) => Number.isFinite(n) && n > 0) : undefined

    const colCountFromGrid = tableGridCols?.length

    // 计算最大列数
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

    // 跟踪垂直合并的单元格（关键：用于样式继承）
    const mergeTracker = new Map<string, CellStyle>()

    // 解析每一行
    for (const tr of trs) {
      const rowTexts = new Array(maxCols).fill("")
      const rowStyles = new Array(maxCols).fill(undefined).map(() => ({} as CellStyle))

      // 解析行属性
      const trPr = childOf(tr, "w:trPr")
      const gridBeforeNode = trPr ? childOf(trPr, "w:gridBefore") : undefined
      const gridAfterNode = trPr ? childOf(trPr, "w:gridAfter") : undefined
      const gridBefore = Number.parseInt(String(attrOf(attrsOf(gridBeforeNode ?? {}), "w:val") ?? ""), 10)
      const gridAfter = Number.parseInt(String(attrOf(attrsOf(gridAfterNode ?? {}), "w:val") ?? ""), 10)

      // 标记已占用的列（由垂直合并的单元格占用）
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

      // 处理 gridBefore（行首空列）
      const beforeCols = Number.isFinite(gridBefore) && gridBefore > 0 ? gridBefore : 0
      for (let c = 0; c < Math.min(maxCols, beforeCols); c += 1) {
        occupied[c] = true
        rowStyles[c].skip = true
      }

      let colCursor = beforeCols

      // 解析每个单元格
      for (const tc of childrenNamed(tr, "w:tc")) {
        // 跳过已占用的列
        while (colCursor < maxCols && occupied[colCursor]) colCursor += 1
        if (colCursor >= maxCols) break

        // 解析单元格属性
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

        // 处理垂直合并（关键：样式继承逻辑）
        if (vMerge != null && vMergeVal !== "restart") {
          const top = mergeTracker.get(mergeKey)
          if (top) {
            top.rowSpan = (top.rowSpan ?? 1) + 1
            // 关键：继承顶部单元格的样式到合并的单元格
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

        // 清除合并跟踪（新的合并开始）
        mergeTracker.delete(mergeKey)

        // 解析单元格内的段落
        const paras = childrenNamed(tc, "w:p").map((p) =>
          parseParagraphNode(p, styles, docDefaultRun, docDefaultPara, hyperlinkRels)
        )

        const cellText = paras.map((p) => p.text).join("\n")
        rowTexts[colStart] = cellText

        // 提取第一个运行的样式
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

        // 开始新的垂直合并
        if (vMergeVal === "restart") {
          style.rowSpan = 1
          mergeTracker.set(mergeKey, style)
        }

        rowStyles[colStart] = style
        for (let c = colStart + 1; c < colEnd; c += 1) rowStyles[c] = { skip: true }

        colCursor = colEnd
      }

      // 处理 gridAfter（行尾空列）
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

    // 返回表格的 DocxParagraph 表示
    return {
      text: "",
      isTable: true,
      tableData: rows,
      tableCellStyles: cellStyles,
      tableLayout,
      tableGridCols,
      tableBorders,
      tableStyleId // 关键：保留表格样式引用
    }
  }
}
