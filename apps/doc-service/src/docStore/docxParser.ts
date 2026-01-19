import type { DocxParagraph } from "./docxGenerator.js"
import type { DocumentMetadata, HeaderFooterContent, Comment, Note } from "./docxMetadata.js"
import { readZipEntry, listZipEntries } from "./core/zipReader.js"
import { parseCoreProperties, parseAppProperties, parseCustomProperties, parseHeaderFooter, parseComments, parseNotes } from "./docxMetadata.js"

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


/**
 * 解析选项
 */
export interface DocxParseOptions {
  includeMetadata?: boolean       // 是否解析元数据（core.xml, app.xml, custom.xml）
  includeHeadersFooters?: boolean // 是否解析页眉页脚
  includeComments?: boolean       // 是否解析注释
  includeNotes?: boolean          // 是否解析脚注和尾注
  includeImages?: boolean         // 是否提取图片 base64
  includeResources?: boolean      // 是否保留资源 XML（styles, numbering, settings等）
}

/**
 * 统一的 docx 解析入口
 * 根据选项灵活控制解析深度
 */
export async function parseDocx(
  docxBuffer: Buffer,
  options: DocxParseOptions = {}
): Promise<FullDocxDocument> {
  if (!docxBuffer || docxBuffer.length === 0) {
    return { paragraphs: [] }
  }

  // 导入内部解析函数（避免循环依赖）
  const { extractParagraphsFromDocx } = await import("./docxGenerator.js")

  // 解析主体内容（段落和表格）
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)

  const result: FullDocxDocument = { paragraphs }

  // 根据选项解析额外内容
  if (options.includeMetadata) {
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

    if (Object.keys(metadata.core || {}).length || Object.keys(metadata.app || {}).length) {
      result.metadata = metadata
    }
  }

  if (options.includeHeadersFooters) {
    const entries = await listZipEntries(docxBuffer)
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

    if (headers.size > 0) result.headers = headers
    if (footers.size > 0) result.footers = footers
  }

  if (options.includeComments) {
    const commentsXml = await readZipEntry(docxBuffer, "word/comments.xml")
    if (commentsXml) {
      const comments = parseComments(commentsXml.toString("utf-8"))
      if (comments && comments.size > 0) {
        result.comments = comments
      }
    }
  }

  if (options.includeNotes) {
    const [footnotesXml, endnotesXml] = await Promise.all([
      readZipEntry(docxBuffer, "word/footnotes.xml"),
      readZipEntry(docxBuffer, "word/endnotes.xml")
    ])

    if (footnotesXml) {
      const footnotes = parseNotes(footnotesXml.toString("utf-8"), "footnote")
      if (footnotes && footnotes.size > 0) {
        result.footnotes = footnotes
      }
    }

    if (endnotesXml) {
      const endnotes = parseNotes(endnotesXml.toString("utf-8"), "endnote")
      if (endnotes && endnotes.size > 0) {
        result.endnotes = endnotes
      }
    }
  }

  if (options.includeImages) {
    const entries = await listZipEntries(docxBuffer)
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

    if (images.size > 0) result.images = images
  }

  if (options.includeResources) {
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

    if (Object.values(resources).some(v => v)) {
      result.resources = resources
    }
  }

  return result
}

/**
 * 快捷方法：只解析段落（最快）
 */
export async function parseParagraphs(docxBuffer: Buffer): Promise<DocxParagraph[]> {
  const doc = await parseDocx(docxBuffer, {})
  return doc.paragraphs
}

/**
 * 快捷方法：解析完整文档（包含所有元数据）
 */
export async function parseWithMetadata(docxBuffer: Buffer): Promise<FullDocxDocument> {
  return await parseDocx(docxBuffer, {
    includeMetadata: true,
    includeHeadersFooters: true,
    includeComments: true,
    includeNotes: true,
    includeImages: true,
    includeResources: true
  })
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
    wordCount += text.split(/\s+/).filter((w: string) => w.length > 0).length
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
