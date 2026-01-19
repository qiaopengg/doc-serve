/**
 * DOCX 流式切片生成
 */

import { XMLParser } from "fast-xml-parser"
import type { DocxParagraph, OrderedXmlNode, FlattenedElement } from "./types.js"
import {
  tagNameOf,
  childOf,
  childrenOf,
  childrenNamed
} from "./core/utils.js"
import { buildXml } from "./core/xml-parser.js"
import { readZipEntry, replaceZipEntry } from "./core/zip-reader.js"

export function flattenParagraphsForStreaming(paragraphs: DocxParagraph[]): FlattenedElement[] {
  const flattened: FlattenedElement[] = []
  let tableIndex = 0
  
  for (const para of paragraphs) {
    if (para.isTable && para.tableData) {
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
      flattened.push({
        type: "paragraph",
        paragraph: para
      })
    }
  }
  
  return flattened
}

export async function streamDocxSlices(
  sourceDocxBuffer: Buffer,
  bodyElementCount: number
): Promise<Buffer> {
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
    
    if (tn === "w:tbl") {
      const tblPr = childOf(n, "w:tblPr")
      const tblGrid = childOf(n, "w:tblGrid")
      const allRows = childrenNamed(n, "w:tr")
      
      const remainingCount = bodyElementCount - included
      const rowsToInclude = Math.min(remainingCount, allRows.length)
      
      if (rowsToInclude > 0) {
        const partialTable = cloneNode(n)
        const tableKey = "w:tbl"
        const tableChildren: OrderedXmlNode[] = []
        
        if (tblPr) tableChildren.push(cloneNode(tblPr))
        if (tblGrid) tableChildren.push(cloneNode(tblGrid))
        
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
