/**
 * DOCX 文档解析
 */

import type { DocxParagraph } from './types.js'
import type { DocumentMetadata, HeaderFooterContent, Comment, Note } from './metadata/metadata.js'
import { extractParagraphsFromDocx } from './document/document-parser.js'
import { readZipEntry, listZipEntries } from './core/zip-reader.js'
import {
  parseCoreProperties,
  parseAppProperties,
  parseCustomProperties,
  parseHeaderFooter,
  parseComments,
  parseNotes
} from './metadata/metadata.js'

export interface DocxDocument {
  paragraphs: DocxParagraph[]
  metadata?: DocumentMetadata
  headers?: Map<string, HeaderFooterContent>
  footers?: Map<string, HeaderFooterContent>
  comments?: Map<string, Comment>
  footnotes?: Map<string, Note>
  endnotes?: Map<string, Note>
  images?: Map<string, string>
  resources?: {
    styles?: string
    numbering?: string
    settings?: string
    fontTable?: string
    theme?: string
  }
}

export interface DocxParseOptions {
  includeMetadata?: boolean
  includeHeadersFooters?: boolean
  includeComments?: boolean
  includeNotes?: boolean
  includeImages?: boolean
  includeResources?: boolean
}

export async function parseDocxDocument(
  docxBuffer: Buffer,
  options: DocxParseOptions = {}
): Promise<DocxDocument> {
  if (!docxBuffer || docxBuffer.length === 0) {
    return { paragraphs: [] }
  }

  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  const result: DocxDocument = { paragraphs }

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
