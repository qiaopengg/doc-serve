import type { DocxParagraph } from "./docxGenerator.js"
import { extractParagraphsFromDocx } from "./docxGenerator.js"
import type { DocumentMetadata, HeaderFooterContent, Comment, Note } from "./docxMetadata.js"
import { parseCoreProperties, parseAppProperties, parseCustomProperties, parseHeaderFooter, parseComments, parseNotes } from "./docxMetadata.js"
import { createRequire } from "node:module"
import type { Entry, ZipFile } from "yauzl"
import type { Readable } from "node:stream"

/**
 * 完整的文档结构
 */
export interface FullDocxDocument {
  // 主体内容
  paragraphs: DocxParagraph[]
  
  // 元数据
  metadata?: DocumentMetadata
  
  // 页眉页脚
  headers?: Map<string, HeaderFooterContent>
  footers?: Map<string, HeaderFooterContent>
  
  // 注释
  comments?: Map<string, Comment>
  
  // 脚注和尾注
  footnotes?: Map<string, Note>
  endnotes?: Map<string, Note>
  
  // 图片数据（relationshipId -> base64 data）
  images?: Map<string, string>
  
  // 其他资源
  resources?: {
    styles?: string
    numbering?: string
    settings?: string
    fontTable?: string
    theme?: string
  }
}

async function readZipEntry(docxBuffer: Buffer, entryName: string): Promise<Buffer | undefined> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")

  return await new Promise((resolve, reject) => {
    yauzl.fromBuffer(docxBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
      if (err || !zipfile) {
        reject(err ?? new Error("zip_open_failed"))
        return
      }

      let done = false

      const finish = (result: Buffer | undefined) => {
        if (done) return
        done = true
        zipfile.close()
        resolve(result)
      }

      zipfile.readEntry()
      zipfile.on("entry", (entry: Entry) => {
        if (entry.fileName !== entryName) {
          zipfile.readEntry()
          return
        }
        zipfile.openReadStream(entry, (err2: Error | null, stream?: Readable) => {
          if (err2 || !stream) {
            reject(err2 ?? new Error("zip_entry_stream_failed"))
            return
          }
          const chunks: Buffer[] = []
          stream.on("data", (c: Buffer) => chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c)))
          stream.on("end", () => finish(Buffer.concat(chunks)))
          stream.on("error", reject)
        })
      })
      zipfile.on("end", () => finish(undefined))
      zipfile.on("error", reject)
    })
  })
}

async function listZipEntries(docxBuffer: Buffer): Promise<string[]> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")

  return await new Promise((resolve, reject) => {
    yauzl.fromBuffer(docxBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
      if (err || !zipfile) {
        reject(err ?? new Error("zip_open_failed"))
        return
      }

      const entries: string[] = []

      zipfile.readEntry()
      zipfile.on("entry", (entry: Entry) => {
        entries.push(entry.fileName)
        zipfile.readEntry()
      })
      zipfile.on("end", () => {
        zipfile.close()
        resolve(entries)
      })
      zipfile.on("error", reject)
    })
  })
}

/**
 * 完整解析 docx 文档，包含所有元素
 */
export async function parseFullDocx(docxBuffer: Buffer): Promise<FullDocxDocument> {
  if (!docxBuffer || docxBuffer.length === 0) {
    return { paragraphs: [] }
  }

  // 列出所有文件
  const entries = await listZipEntries(docxBuffer)
  
  // 解析主体内容
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  
  // 解析元数据
  const [coreXml, appXml, customXml] = await Promise.all([
    readZipEntry(docxBuffer, "docProps/core.xml"),
    readZipEntry(docxBuffer, "docProps/app.xml"),
    readZipEntry(docxBuffer, "docProps/custom.xml")
  ])
  
  const metadata: DocumentMetadata = {
    core: coreXml ? parseCoreProperties(coreXml.toString("utf-8")) : undefined,
    app: appXml ? parseAppProperties(appXml.toString("utf-8")) : undefined,
    custom: customXml ? parseCustomProperties(customXml.toString("utf-8")) : undefined
  }
  
  // 解析页眉页脚
  const headers = new Map<string, HeaderFooterContent>()
  const footers = new Map<string, HeaderFooterContent>()
  
  for (const entry of entries) {
    if (entry.startsWith("word/header") && entry.endsWith(".xml")) {
      const xml = await readZipEntry(docxBuffer, entry)
      if (xml) {
        const content = parseHeaderFooter(xml.toString("utf-8"))
        content.id = entry
        headers.set(entry, content)
      }
    } else if (entry.startsWith("word/footer") && entry.endsWith(".xml")) {
      const xml = await readZipEntry(docxBuffer, entry)
      if (xml) {
        const content = parseHeaderFooter(xml.toString("utf-8"))
        content.id = entry
        footers.set(entry, content)
      }
    }
  }
  
  // 解析注释
  const commentsXml = await readZipEntry(docxBuffer, "word/comments.xml")
  const comments = commentsXml ? parseComments(commentsXml.toString("utf-8")) : undefined
  
  // 解析脚注和尾注
  const footnotesXml = await readZipEntry(docxBuffer, "word/footnotes.xml")
  const endnotesXml = await readZipEntry(docxBuffer, "word/endnotes.xml")
  const footnotes = footnotesXml ? parseNotes(footnotesXml.toString("utf-8"), "footnote") : undefined
  const endnotes = endnotesXml ? parseNotes(endnotesXml.toString("utf-8"), "endnote") : undefined
  
  // 解析图片（提取 base64）
  const images = new Map<string, string>()
  for (const entry of entries) {
    if (entry.startsWith("word/media/") && /\.(png|jpg|jpeg|gif|bmp|svg)$/i.test(entry)) {
      const imgBuf = await readZipEntry(docxBuffer, entry)
      if (imgBuf) {
        const base64 = imgBuf.toString("base64")
        const ext = entry.split(".").pop()?.toLowerCase() || "png"
        const mimeType = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : `image/${ext}`
        images.set(entry, `data:${mimeType};base64,${base64}`)
      }
    }
  }
  
  // 解析其他资源
  const [stylesXml, numberingXml, settingsXml, fontTableXml, themeXml] = await Promise.all([
    readZipEntry(docxBuffer, "word/styles.xml"),
    readZipEntry(docxBuffer, "word/numbering.xml"),
    readZipEntry(docxBuffer, "word/settings.xml"),
    readZipEntry(docxBuffer, "word/fontTable.xml"),
    readZipEntry(docxBuffer, "word/theme/theme1.xml")
  ])
  
  const resources = {
    styles: stylesXml?.toString("utf-8"),
    numbering: numberingXml?.toString("utf-8"),
    settings: settingsXml?.toString("utf-8"),
    fontTable: fontTableXml?.toString("utf-8"),
    theme: themeXml?.toString("utf-8")
  }
  
  return {
    paragraphs,
    metadata: Object.keys(metadata.core || {}).length || Object.keys(metadata.app || {}).length ? metadata : undefined,
    headers: headers.size > 0 ? headers : undefined,
    footers: footers.size > 0 ? footers : undefined,
    comments: comments && comments.size > 0 ? comments : undefined,
    footnotes: footnotes && footnotes.size > 0 ? footnotes : undefined,
    endnotes: endnotes && endnotes.size > 0 ? endnotes : undefined,
    images: images.size > 0 ? images : undefined,
    resources: Object.values(resources).some(v => v) ? resources : undefined
  }
}

/**
 * 获取文档统计信息
 */
export interface DocxStatistics {
  paragraphCount: number
  tableCount: number
  imageCount: number
  wordCount: number
  characterCount: number
  hasNumbering: boolean
  hasComments: boolean
  hasFootnotes: boolean
  hasEndnotes: boolean
  hasHeaders: boolean
  hasFooters: boolean
  styleIds: string[]
}

export function getDocxStatistics(doc: FullDocxDocument): DocxStatistics {
  let tableCount = 0
  let imageCount = 0
  let wordCount = 0
  let characterCount = 0
  let hasNumbering = false
  const styleIds = new Set<string>()
  
  for (const para of doc.paragraphs) {
    if (para.isTable) {
      tableCount++
    }
    
    if (para.images && para.images.length > 0) {
      imageCount += para.images.length
    }
    
    if (para.numbering) {
      hasNumbering = true
    }
    
    if (para.styleId) {
      styleIds.add(para.styleId)
    }
    
    const text = para.text || ""
    characterCount += text.length
    wordCount += text.split(/\s+/).filter(w => w.length > 0).length
  }
  
  return {
    paragraphCount: doc.paragraphs.length,
    tableCount,
    imageCount: imageCount + (doc.images?.size || 0),
    wordCount,
    characterCount,
    hasNumbering,
    hasComments: (doc.comments?.size || 0) > 0,
    hasFootnotes: (doc.footnotes?.size || 0) > 0,
    hasEndnotes: (doc.endnotes?.size || 0) > 0,
    hasHeaders: (doc.headers?.size || 0) > 0,
    hasFooters: (doc.footers?.size || 0) > 0,
    styleIds: Array.from(styleIds)
  }
}
