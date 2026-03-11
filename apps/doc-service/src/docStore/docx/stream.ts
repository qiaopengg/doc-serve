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
  
  for (const para of paragraphs) {
    // 表格和段落都作为一个完整单元
    flattened.push({
      type: para.isTable ? "paragraph" : "paragraph",
      paragraph: para
    })
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
    
    // 表格作为一个完整单元
    if (tn === "w:tbl") {
      newBodyChildren.push(cloneNode(n))
      lastIncludedBodyIndex = i
      included += 1
      continue
    }
    
    // 段落作为一个单元
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
