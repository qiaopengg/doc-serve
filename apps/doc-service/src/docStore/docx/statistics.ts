/**
 * DOCX 文档统计信息
 */

import type { DocxDocument } from './parse.js'

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

export function getDocumentStatistics(doc: DocxDocument): DocxStatistics {
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
