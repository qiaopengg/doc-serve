import type { Router } from "../http/router.js"
import type { DocStore } from "../docStore/types.js"
import { HttpError, pipeDocxChunksToResponse } from "@wps/doc-core"
import { URL } from "node:url"
import { 
  streamDocxSlices,
  parseDocxDocument,
  getDocumentStatistics,
  type SectionPropertiesSpec,
  type DocxDocument,
  type DocxParseOptions
} from "../docStore/docx/index.js"
import type { Readable } from "node:stream"
import type { IncomingMessage, ServerResponse } from "node:http"

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

type StreamRouteConfig = {
  path: string
  parseOptions: DocxParseOptions
  includeStatsHeader: boolean
}

const DEFAULT_DOC_CANDIDATES = ["test.docx", "text.docx", "text.dock", "mock.docx"] as const

async function readableToBuffer(stream: Readable): Promise<Buffer> {
  const chunks: Buffer[] = []
  for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
  return Buffer.concat(chunks)
}

function setCommonStreamHeaders(req: IncomingMessage, res: ServerResponse, delayMs: number): void {
  res.statusCode = 200
  res.setHeader("Content-Type", "application/x-wps-docx-chunks")
  res.setHeader("X-Content-Type-Options", "nosniff")
  res.setHeader("Cache-Control", "no-store")
  res.setHeader("X-WPS-Stream-Mode", "docx-chunks")
  res.setHeader("X-WPS-Delay-Ms", String(delayMs))
  if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)
}

function setPageSetupHeaders(res: ServerResponse, doc: DocxDocument): void {
  const firstSection = doc.paragraphs.find((p) => p.sectionProperties)?.sectionProperties
  const pageSetupPoints = sectionPropertiesToPageSetupPoints(firstSection)
  res.setHeader("X-WPS-PageSetup-Unit", "point")
  res.setHeader("X-WPS-PageSetup-Source-Unit", "twip")
  res.setHeader("X-WPS-PageSetup", encodeURIComponent(JSON.stringify(pageSetupPoints ?? {})))
}

function setStatsHeader(res: ServerResponse, doc: DocxDocument): void {
  const stats = getDocumentStatistics(doc)
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
}

async function streamDocxByConfig(
  req: IncomingMessage,
  res: ServerResponse,
  deps: { docStore: DocStore },
  config: StreamRouteConfig
): Promise<void> {
  const delayMs = randomIntInclusive(100, 200)
  setCommonStreamHeaders(req, res, delayMs)

  const requestedDocId = getQueryParam(req, "docId")
  const candidateDocIds = requestedDocId ? [requestedDocId] : [...DEFAULT_DOC_CANDIDATES]
  let selectedFilename = ""
  let selectedDocxBuffer: Buffer | undefined
  let parsedDoc: DocxDocument | undefined
  let lastErr: unknown

  for (const candidateDocId of candidateDocIds) {
    try {
      const { meta, stream } = await deps.docStore.openStream(candidateDocId)
      const docxBuffer = await readableToBuffer(stream)
      const parsed = await parseDocxDocument(docxBuffer, config.parseOptions)
      selectedFilename = meta.filename
      selectedDocxBuffer = docxBuffer
      parsedDoc = parsed
      break
    } catch (err) {
      lastErr = err
      if (requestedDocId) {
        throw err
      }
      if (err instanceof HttpError && err.statusCode === 404) {
        continue
      }
    }
  }

  if (!selectedDocxBuffer || !parsedDoc) {
    if (lastErr instanceof HttpError && lastErr.statusCode === 404) {
      throw lastErr
    }
    throw new HttpError(404, "document_not_found")
  }

  const resolvedDocxBuffer = selectedDocxBuffer
  const resolvedParsedDoc = parsedDoc
  res.setHeader("X-WPS-Filename", selectedFilename)
  setPageSetupHeaders(res, resolvedParsedDoc)
  if (config.includeStatsHeader) {
    setStatsHeader(res, resolvedParsedDoc)
  }

  if (typeof res.flushHeaders === "function") res.flushHeaders()

  async function* generateDocxChunks(): AsyncGenerator<Buffer> {
    for (let i = 1; i <= resolvedParsedDoc.paragraphs.length; i += 1) {
      yield await streamDocxSlices(resolvedDocxBuffer, i)
    }
  }

  await pipeDocxChunksToResponse(req, res, generateDocxChunks(), { delayMs })
}

export function registerDocRoutes(router: Router, deps: { docStore: DocStore }): void {
  const configs: StreamRouteConfig[] = [
    {
      path: "/api/v1/docs/stream-docx",
      parseOptions: {},
      includeStatsHeader: false
    },
    {
      path: "/api/v1/docs/better-stream-docx",
      parseOptions: {
        includeMetadata: true,
        includeHeadersFooters: true,
        includeComments: true,
        includeNotes: true
      },
      includeStatsHeader: true
    }
  ]

  for (const config of configs) {
    router.add("GET", config.path, async ({ req, res }) => {
      await streamDocxByConfig(req, res, deps, config)
    })
  }
}
