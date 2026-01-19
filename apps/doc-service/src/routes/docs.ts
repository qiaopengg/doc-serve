import type { Router } from "../http/router.js"
import type { DocStore } from "../docStore/types.js"
import { HttpError, pipeAsyncIterableToResponseAsNdjson, pipeReadableToResponse, pipeReadableToResponseAsFramedChunks, pipeDocxChunksToResponse } from "@wps/doc-core"
import { URL } from "node:url"
import { createDocxFromSourceDocxSlice, extractParagraphsFromDocx } from "../docStore/docxGenerator.js"
import { parseFullDocx, getDocxStatistics } from "../docStore/docxFullParser.js"
import type { SectionPropertiesSpec } from "../docStore/docxGenerator.js"
import type { Readable } from "node:stream"

const DEFAULT_DOC_ID = "test.docx"

function randomIntInclusive(min: number, max: number): number {
  const a = Math.ceil(Math.min(min, max))
  const b = Math.floor(Math.max(min, max))
  return a + Math.floor(Math.random() * (b - a + 1))
}

function getRequestOrigin(req: import("node:http").IncomingMessage): string {
  const host = String(req.headers.host ?? "")
  const proto = String(req.headers["x-forwarded-proto"] ?? "http").split(",")[0]?.trim() || "http"
  if (host) return `${proto}://${host}`
  return "http://127.0.0.1"
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
  async function openFirstAvailable(docIds: string[]): Promise<Awaited<ReturnType<DocStore["openStream"]>> & { chosenDocId: string }> {
    let lastErr: unknown
    for (const docId of docIds) {
      try {
        const out = await deps.docStore.openStream(docId)
        return { ...out, chosenDocId: docId }
      } catch (err) {
        lastErr = err
        if (!(err instanceof HttpError && err.statusCode === 404)) {
          throw err
        }
      }
    }
    throw lastErr
  }

  async function sendDownload(req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse, docId: string) {
    const { meta, stream } = await deps.docStore.openStream(docId)

    res.statusCode = 200
    res.setHeader("Content-Type", meta.contentType)
    res.setHeader("Content-Length", String(meta.contentLength))
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeURIComponent(meta.filename)}`
    )
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")

    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)
    if (typeof res.flushHeaders === "function") res.flushHeaders()
    await pipeReadableToResponse(req, res, stream)
  }

  async function sendFrames(
    req: import("node:http").IncomingMessage,
    res: import("node:http").ServerResponse,
    docId: string,
    downloadPath: string
  ) {
    const origin = getRequestOrigin(req)
    const { meta, stream } = await deps.docStore.openStream(docId)
    stream.destroy()
    const simulateError = String(req.headers["x-wps-simulate-error"] ?? "") === "1"

    const delayMs = randomIntInclusive(100, 200)
    res.statusCode = 200
    res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.setHeader("X-WPS-Stream-Mode", "ndjson-frames")
    res.setHeader("X-WPS-Delay-Ms", String(delayMs))
    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)
    if (typeof res.flushHeaders === "function") res.flushHeaders()

    async function* frames(): AsyncGenerator<unknown> {
      let seq = 1
      const ts = () => Date.now()

      try {
        const runs = [
          { text: `标题：${docId}`, bold: true, fontSize: 16, headingLevel: 1, newParagraph: true },
          { text: "这是预览内容，用于测试 WPS 实时写入与逐字效果。", fontSize: 11, newParagraph: true },
          { text: "第二段：样式与分页最终以 docx 为准。", fontSize: 11, newParagraph: true }
        ]

        for (const run of runs) {
          yield { docId, seq, type: "preview.runs", ts: ts(), payload: { runs: [run] } }
          seq += 1
          if (simulateError) {
            throw new Error("simulated_error")
          }
        }

        yield {
          docId,
          seq,
          type: "final.docx.url",
          ts: ts(),
          payload: {
            url: `${origin}${downloadPath}`,
            fileName: meta.filename,
            expiresAt: ts() + 10 * 60_000
          }
        }
        seq += 1

        yield { docId, seq, type: "control.done", ts: ts(), payload: { reason: "completed" } }
      } catch (err) {
        const message = err instanceof Error ? err.message : "unknown_error"
        yield {
          docId,
          seq,
          type: "control.error",
          ts: ts(),
          payload: { code: "INTERNAL_ERROR", message, retryable: false }
        }
      }
    }

    await pipeAsyncIterableToResponseAsNdjson(req, res, frames(), { delayMs })
  }

  async function sendDefaultFrames(req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse, downloadPath: string) {
    const { chosenDocId } = await openFirstAvailable([DEFAULT_DOC_ID, "text.docx", "text.dock", "text.doc"])
    await sendFrames(req, res, chosenDocId, downloadPath)
  }

  async function sendDefaultDownload(req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse) {
    const { chosenDocId } = await openFirstAvailable([DEFAULT_DOC_ID, "text.docx", "text.dock", "text.doc"])
    await sendDownload(req, res, chosenDocId)
  }

  router.add("GET", "/api/v1/docs/stream", async ({ req, res }) => {
    let meta: { filename: string }
    let stream: import("node:stream").Readable
    try {
      ;({ meta, stream } = await deps.docStore.openStream("text.doc"))
    } catch (err) {
      if (err instanceof HttpError && err.statusCode === 404) {
        ;({ meta, stream } = await deps.docStore.openStream("text.docx"))
      } else {
        throw err
      }
    }

    res.setHeader("Content-Type", "application/x-wps-framed-chunks")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.setHeader("X-WPS-Stream-Mode", "framed-chunks")
    res.setHeader("X-WPS-Filename", meta.filename)
    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)

    const chunkSize = 1024
    res.setHeader("X-WPS-Chunk-Size", String(chunkSize))
    const delayMs = randomIntInclusive(100, 200)
    res.setHeader("X-WPS-Delay-Ms", String(delayMs))

    res.statusCode = 200
    if (typeof res.flushHeaders === "function") res.flushHeaders()
    await pipeReadableToResponseAsFramedChunks(req, res, stream, { chunkSize, delayMs })
  })

  // 新接口：按完整 docx 块流式推送
  router.add("GET", "/api/v1/docs/stream-docx", async ({ req, res }) => {
    const delayMs = randomIntInclusive(100, 200)
    
    res.statusCode = 200
    res.setHeader("Content-Type", "application/x-wps-docx-chunks")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.setHeader("X-WPS-Stream-Mode", "docx-chunks")
    res.setHeader("X-WPS-Delay-Ms", String(delayMs))
    
    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)

    async function readableToBuffer(stream: Readable): Promise<Buffer> {
      const chunks: Buffer[] = []
      for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
      return Buffer.concat(chunks)
    }

    const docId = getQueryParam(req, "docId") || "mock.docx"
    const { stream } = await deps.docStore.openStream(docId)
    const docxBuffer = await readableToBuffer(stream)
    const allParagraphs = await extractParagraphsFromDocx(docxBuffer)

    const firstSection = allParagraphs.find((p) => p.sectionProperties)?.sectionProperties
    const pageSetupPoints = sectionPropertiesToPageSetupPoints(firstSection)
    res.setHeader("X-WPS-PageSetup-Unit", "point")
    res.setHeader("X-WPS-PageSetup-Source-Unit", "twip")
    res.setHeader("X-WPS-PageSetup", encodeURIComponent(JSON.stringify(pageSetupPoints ?? {})))

    if (typeof res.flushHeaders === "function") res.flushHeaders()

    async function* generateDocxChunks(): AsyncGenerator<Buffer> {
      for (let i = 0; i < allParagraphs.length; i += 1) {
        yield await createDocxFromSourceDocxSlice(docxBuffer, i + 1)
      }
    }

    await pipeDocxChunksToResponse(req, res, generateDocxChunks(), { delayMs })
  })

  router.add("GET", "/api/v1/docs/frames-docx", async ({ req, res }) => {
    const origin = getRequestOrigin(req)
    const docId = getQueryParam(req, "docId") || "mock.docx"

    const delayMs = randomIntInclusive(30, 80)
    res.statusCode = 200
    res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.setHeader("X-WPS-Stream-Mode", "ndjson-frames")
    res.setHeader("X-WPS-Delay-Ms", String(delayMs))
    if (typeof req.socket?.setNoDelay === "function") req.socket.setNoDelay(true)
    if (typeof res.flushHeaders === "function") res.flushHeaders()

    async function readableToBuffer(stream: Readable): Promise<Buffer> {
      const chunks: Buffer[] = []
      for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
      return Buffer.concat(chunks)
    }

    async function* frames(): AsyncGenerator<unknown> {
      let seq = 1
      const ts = () => Date.now()

      const { meta, stream } = await deps.docStore.openStream(docId)
      const docxBuffer = await readableToBuffer(stream)
      const paragraphs = await extractParagraphsFromDocx(docxBuffer)
      const firstSection = paragraphs.find((p) => p.sectionProperties)?.sectionProperties
      const pageSetupPoints = sectionPropertiesToPageSetupPoints(firstSection)

      yield {
        docId,
        seq,
        type: "preview.pageSetup",
        ts: ts(),
        payload: {
          unit: "point",
          pageSetup: pageSetupPoints ?? {},
          sourceSectionProperties: firstSection ?? {},
          sourceUnit: "twip"
        }
      }
      seq += 1

      for (let i = 0; i < paragraphs.length; i += 1) {
        yield {
          docId,
          seq,
          type: "preview.paragraphs",
          ts: ts(),
          payload: { unit: "twip", paragraphs: [paragraphs[i]!] }
        }
        seq += 1
      }

      yield {
        docId,
        seq,
        type: "final.docx.url",
        ts: ts(),
        payload: {
          url: `${origin}/api/v1/docs/${encodeURIComponent(docId)}/download`,
          fileName: meta.filename,
          expiresAt: ts() + 10 * 60_000
        }
      }
      seq += 1

      yield { docId, seq, type: "control.done", ts: ts(), payload: { reason: "completed" } }
    }

    await pipeAsyncIterableToResponseAsNdjson(req, res, frames(), { delayMs })
  })

  router.add("GET", "/api/v1/docs/parsed", async ({ req, res }) => {
    const docId = getQueryParam(req, "docId") || "mock.docx"
    const takeRaw = getQueryParam(req, "take")
    const take = takeRaw ? Number.parseInt(takeRaw, 10) : undefined

    const { stream } = await deps.docStore.openStream(docId)
    const chunks: Buffer[] = []
    for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
    const docxBuffer = Buffer.concat(chunks)

    let paragraphs = await extractParagraphsFromDocx(docxBuffer)
    if (take != null && Number.isFinite(take) && take >= 0) {
      paragraphs = paragraphs.slice(0, take)
    }

    res.statusCode = 200
    res.setHeader("Content-Type", "application/json; charset=utf-8")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.end(JSON.stringify({ docId, count: paragraphs.length, paragraphs }))
  })

  router.add("GET", "/api/v1/docs/:docId/download", async ({ req, res, params }) => {
    await sendDownload(req, res, params.docId)
  })

  router.add("GET", "/api/v1/docs/:docId/frames", async ({ req, res, params }) => {
    const docId = params.docId
    await sendFrames(req, res, docId, `/api/v1/docs/${encodeURIComponent(docId)}/download`)
  })

  router.add("GET", "/api/v1/docs/frames", async ({ req, res }) => {
    const docId = getQueryParam(req, "docId")
    if (docId) {
      await sendFrames(req, res, docId, `/api/v1/docs/${encodeURIComponent(docId)}/download`)
      return
    }
    await sendDefaultFrames(req, res, "/api/v1/docs/download")
  })

  router.add("GET", "/api/v1/docs/download", async ({ req, res }) => {
    await sendDefaultDownload(req, res)
  })

  router.add("GET", "/docs/frames", async ({ req, res }) => {
    const docId = getQueryParam(req, "docId")
    if (docId) {
      await sendFrames(req, res, docId, `/api/v1/docs/${encodeURIComponent(docId)}/download`)
      return
    }
    await sendDefaultFrames(req, res, "/docs/download")
  })

  router.add("GET", "/docs/download", async ({ req, res }) => {
    await sendDefaultDownload(req, res)
  })

  // 新接口：完整解析文档（包含所有元素）
  router.add("GET", "/api/v1/docs/full-parse", async ({ req, res }) => {
    const docId = getQueryParam(req, "docId") || "mock.docx"
    
    const { stream } = await deps.docStore.openStream(docId)
    const chunks: Buffer[] = []
    for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
    const docxBuffer = Buffer.concat(chunks)
    
    const fullDoc = await parseFullDocx(docxBuffer)
    const stats = getDocxStatistics(fullDoc)
    
    res.statusCode = 200
    res.setHeader("Content-Type", "application/json; charset=utf-8")
    res.setHeader("X-Content-Type-Options", "nosniff")
    res.setHeader("Cache-Control", "no-store")
    res.end(JSON.stringify({ 
      docId, 
      statistics: stats,
      document: fullDoc 
    }, null, 2))
  })
}
