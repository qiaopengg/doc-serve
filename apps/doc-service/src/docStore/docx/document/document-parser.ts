/**
 * 主文档解析器
 * 从 docxGenerator.ts 迁移的文档解析逻辑
 * 
 * 核心功能：
 * - 解析完整的 DOCX 文档结构
 * - 整合所有子解析器（段落、表格、样式等）
 * - 处理文档遍历和分节属性
 * - 支持编号格式解析
 */

import { XMLParser } from "fast-xml-parser"
import type {
  DocxParagraph,
  StyleMap,
  RunStyle,
  OrderedXmlNode,
  SectionPropertiesSpec
} from "../types.js"
import {
  tagNameOf,
  childOf,
  childrenOf
} from "../core/utils.js"
import { parseParagraphNode } from "./paragraph-parser.js"
import { TableParser } from "./table-parser.js"
import { SectionParser } from "./section-parser.js"
import { StyleParser } from "../styles/style-parser.js"
import { parseNumbering } from "../styles/numbering-parser.js"
import { RelationshipParser } from "../core/relationship-parser.js"
import { readZipEntry } from "../core/zip-reader.js"

/**
 * 从 DOCX Buffer 提取段落列表
 * 
 * 这是主要的文档解析入口函数，负责：
 * 1. 读取 DOCX 文件中的各个 XML 部分（styles.xml, document.xml, relationships, numbering.xml 等）
 * 2. 解析样式、编号、关系等辅助信息
 * 3. 遍历文档主体，解析段落和表格
 * 4. 处理分节属性（页面设置、页边距等）
 * 5. 补充编号格式信息（列表项的格式和文本）
 * 
 * 解析策略：
 * - 从后向前遍历文档主体，以正确处理分节属性的继承关系
 * - 使用 preserveOrder 模式解析 XML，保持元素顺序
 * - 整合多个子解析器（段落、表格、样式、编号等）
 * - 为未识别的元素创建占位符，保持文档结构完整性
 * 
 * 支持的元素类型：
 * - w:p - 段落（包括文本、图片、书签、域代码等）
 * - w:tbl - 表格（包括单元格样式、合并、边框等）
 * - w:sectPr - 分节属性（页面设置、页眉页脚等）
 * 
 * @param docxBuffer - DOCX 文件的 Buffer
 * @returns 解析后的段落数组（包括表格，表格也表示为特殊的段落）
 *          如果文件为空或解析失败，返回空数组
 * 
 * @example
 * ```typescript
 * const docxBuffer = fs.readFileSync('document.docx')
 * const paragraphs = await extractParagraphsFromDocx(docxBuffer)
 * 
 * for (const para of paragraphs) {
 *   if (para.isTable) {
 *     console.log('Table:', para.tableData)
 *   } else {
 *     console.log('Paragraph:', para.text)
 *   }
 * }
 * ```
 */
export async function extractParagraphsFromDocx(docxBuffer: Buffer): Promise<DocxParagraph[]> {
  if (!docxBuffer || docxBuffer.length === 0) return []

  // 1. 并行读取所有需要的 XML 文件
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

  // 2. 解析样式定义
  const styleParser = new StyleParser()
  const { styles, docDefaultRun, docDefaultPara } = stylesXmlBuf
    ? styleParser.parseStyles(stylesXmlBuf.toString("utf-8"))
    : { styles: new Map(), docDefaultRun: {}, docDefaultPara: { alignment: undefined } }

  // 3. 解析关系映射（用于超链接、图片等）
  const relationshipParser = new RelationshipParser()
  const hyperlinkRels = relsBuf ? relationshipParser.parseRelationships(relsBuf.toString("utf-8")) : new Map()
  
  // 4. 解析编号定义
  const numberingDefs = numberingXmlBuf ? parseNumbering(numberingXmlBuf.toString("utf-8")) : new Map()

  // 5. 解析文档主体 XML
  const parser = new XMLParser({ ignoreAttributes: false, preserveOrder: true })
  const ordered: any[] = parser.parse(documentXmlBuf.toString("utf-8"))

  const docNode = ordered.find((n) => tagNameOf(n) === "w:document")
  if (!docNode) return []
  const bodyNode = childOf(docNode, "w:body")
  if (!bodyNode) return []

  const bodyChildren = childrenOf(bodyNode)

  // 6. 遍历文档主体，从后向前处理（用于正确处理分节属性）
  const outReversed: DocxParagraph[] = []
  let currentSectPr: OrderedXmlNode | undefined

  for (let i = bodyChildren.length - 1; i >= 0; i -= 1) {
    const c = bodyChildren[i] as OrderedXmlNode
    const tn = tagNameOf(c)
    
    // 处理分节属性节点
    if (tn === "w:sectPr") {
      currentSectPr = c
      continue
    }
    
    // 处理段落
    if (tn === "w:p") {
      // 检查段落末尾的分节属性
      const pPr = childOf(c, "w:pPr")
      const s = pPr ? childOf(pPr, "w:sectPr") : undefined
      if (s) currentSectPr = s
      
      // 解析段落
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
      
      // 添加分节属性
      const sectionParser = new SectionParser()
      const sectionProperties = sectionParser.parseSectionPropertiesFromSectPr(currentSectPr)
      if (sectionProperties) para.sectionProperties = sectionProperties
      
      outReversed.push(para)
      continue
    }
    
    // 处理表格
    if (tn === "w:tbl") {
      const tbl = TableParser.parseTableNode(c, styles, docDefaultRun, docDefaultPara, hyperlinkRels, parseParagraphNode)
      
      // 添加分节属性
      const sectionParser = new SectionParser()
      const sectionProperties = sectionParser.parseSectionPropertiesFromSectPr(currentSectPr)
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

  // 7. 反转数组（因为我们是从后向前处理的）
  outReversed.reverse()
  return outReversed
}
