import type { Router } from "../http/router.js"
import type { DocStore } from "../docStore/types.js"
import { pipeDocxChunksToResponse } from "@wps/doc-core"
import { URL } from "node:url"
import { createDocxFromSourceDocxSlice } from "../docStore/docxGenerator.js"
import { parseDocx, getDocxStatistics } from "../docStore/docxParser.js"
import type { SectionPropertiesSpec } from "../docStore/docxGenerator.js"
import type { Readable } from "node:stream"

function randomIntInclusive(min: number, max: number): number {
  const a = Math.ceil(Math.min(min, max))
  const b = Math.floor(Math.max(min, max))
  return a + Math.floor(Math.random() * (b - a + 1))
}

function getQueryParam(req: import("node:http").IncomingMessage, key: string): string | undefined {
  const url = new URL(req.url ?? "/", "http://localhost")
  const val = url.searchParams.get(key)
  return val == null || val === "" ? undefined : val
}

function twipsToPoints(v: unknown): number | undefined {
  const n = typeof v === "number" ? v : Number(v)
  if (!Number.isFinite(n)) return undefined
  return n / 20
}

function sectionPropertiesToPageSetupPoints(section: SectionPropertiesSpec | undefined): Record<string, any> | undefined {
  if (!section) return undefined

  const page = section.page ?? {}
  const size = page.size ?? {}
  const margin = page.margin ?? {}
  const column = section.column ?? {}

  const out: Record<string, any> = {}

  const pageWidth = twipsToPoints(size.width)
  const pageHeight = twipsToPoints(size.height)
  if (pageWidth != null) out.pageWidth = pageWidth
  if (pageHeight != null) out.pageHeight = pageHeight
  if (size.orientation === "landscape" || size.orientation === "portrait") out.orientation = size.orientation

  const topMargin = twipsToPoints(margin.top)
  const rightMargin = twipsToPoints(margin.right)
  const bottomMargin = twipsToPoints(margin.bottom)
  const leftMargin = twipsToPoints(margin.left)
  const headerDistance = twipsToPoints(margin.header)
  const footerDistance = twipsToPoints(margin.footer)
  const gutter = twipsToPoints(margin.gutter)

  if (topMargin != null) out.topMargin = topMargin
  if (rightMargin != null) out.rightMargin = rightMargin
  if (bottomMargin != null) out.bottomMargin = bottomMargin
  if (leftMargin != null) out.leftMargin = leftMargin
  if (headerDistance != null) out.headerDistance = headerDistance
  if (footerDistance != null) out.footerDistance = footerDistance
  if (gutter != null) out.gutter = gutter

  if (typeof column.count === "number" && Number.isFinite(column.count)) out.columnsCount = column.count
  if (typeof column.space === "number" && Number.isFinite(column.space)) out.columnsSpace = twipsToPoints(column.space)

  return Object.keys(out).length ? out : undefined
}

export function registerDocRoutes(router: Router, deps: { docStore: DocStore }): void {
  async function readableToBuffer(stream: Readable): Promise<Buffer> {
    const chunks: Buffer[] = []
    for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
    return Buffer.concat(chunks)
  }

  // 接口：按完整 docx 块流式推送（支持逐行流式输出表格）
  router.add("GET", "/api/v1/docs/stream-docx", async ({ req, res }) => {
    const delayMs = randomIntInclusive(100, 200)
    
    res.statusCode = 200
    res.setHeader("Content-Type", "application/x-wps-docx-chunks")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.setHeader("X-WPS-Stream-Mode", "docx-chunks")
    res.setHeader("X-WPS-Delay-Ms", String(delayMs))
    
    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)

    const docId = getQueryParam(req, "docId") || "mock.docx"
    const { stream } = await deps.docStore.openStream(docId)
    const docxBuffer = await readableToBuffer(stream)
    
    // 使用统一解析器（只解析段落，不解析元数据）
    const doc = await parseDocx(docxBuffer, {})
    const allParagraphs = doc.paragraphs

    const firstSection = allParagraphs.find((p: any) => p.sectionProperties)?.sectionProperties
    const pageSetupPoints = sectionPropertiesToPageSetupPoints(firstSection)
    res.setHeader("X-WPS-PageSetup-Unit", "point")
    res.setHeader("X-WPS-PageSetup-Source-Unit", "twip")
    res.setHeader("X-WPS-PageSetup", encodeURIComponent(JSON.stringify(pageSetupPoints ?? {})))

    if (typeof res.flushHeaders === "function") res.flushHeaders()

    async function* generateDocxChunks(): AsyncGenerator<Buffer> {
      // 计算总的流式单元数（段落 + 表格行）
      let totalUnits = 0
      for (const para of allParagraphs) {
        if (para.isTable && para.tableData) {
          totalUnits += para.tableData.length
        } else {
          totalUnits += 1
        }
      }
      
      // 逐个单元生成 docx 块
      for (let i = 0; i < totalUnits; i += 1) {
        yield await createDocxFromSourceDocxSlice(docxBuffer, i + 1)
      }
    }

    await pipeDocxChunksToResponse(req, res, generateDocxChunks(), { delayMs })
  })

  // 接口：使用完整解析器的流式推送（返回格式与 stream-docx 一样，但支持逐行流式输出表格）
  router.add("GET", "/api/v1/docs/better-stream-docx", async ({ req, res }) => {
    const docId = getQueryParam(req, "docId") || "mock.docx"
    const delayMs = randomIntInclusive(100, 200)
    
    res.statusCode = 200
    res.setHeader("Content-Type", "application/x-wps-docx-chunks")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.setHeader("X-WPS-Stream-Mode", "docx-chunks")
    res.setHeader("X-WPS-Delay-Ms", String(delayMs))
    
    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)

    const { stream } = await deps.docStore.openStream(docId)
    const docxBuffer = await readableToBuffer(stream)
    
    // 使用统一解析器（包含元数据）
    const fullDoc = await parseDocx(docxBuffer, {
      includeMetadata: true,
      includeHeadersFooters: true,
      includeComments: true,
      includeNotes: true
    })
    const allParagraphs = fullDoc.paragraphs

    // 提取页面设置信息
    const firstSection = allParagraphs.find((p: any) => p.sectionProperties)?.sectionProperties
    const pageSetupPoints = sectionPropertiesToPageSetupPoints(firstSection)
    
    // 添加增强的元数据到响应头
    res.setHeader("X-WPS-PageSetup-Unit", "point")
    res.setHeader("X-WPS-PageSetup-Source-Unit", "twip")
    res.setHeader("X-WPS-PageSetup", encodeURIComponent(JSON.stringify(pageSetupPoints ?? {})))
    
    // 添加文档统计信息
    const stats = getDocxStatistics(fullDoc)
    res.setHeader("X-WPS-Doc-Stats", encodeURIComponent(JSON.stringify({
      paragraphCount: stats.paragraphCount,
      tableCount: stats.tableCount,
      imageCount: stats.imageCount,
      wordCount: stats.wordCount,
      hasNumbering: stats.hasNumbering,
      hasComments: stats.hasComments,
      hasHeaders: stats.hasHeaders,
      hasFooters: stats.hasFooters
    })))

    if (typeof res.flushHeaders === "function") res.flushHeaders()

    async function* generateEnhancedDocxChunks(): AsyncGenerator<Buffer> {
      // 计算总的流式单元数（段落 + 表格行）
      let totalUnits = 0
      for (const para of allParagraphs) {
        if (para.isTable && para.tableData) {
          totalUnits += para.tableData.length
        } else {
          totalUnits += 1
        }
      }
      
      // 逐个单元生成 docx 块（支持逐行输出表格）
      for (let i = 0; i < totalUnits; i += 1) {
        yield await createDocxFromSourceDocxSlice(docxBuffer, i + 1)
      }
    }

    await pipeDocxChunksToResponse(req, res, generateEnhancedDocxChunks(), { delayMs })
  })
}
