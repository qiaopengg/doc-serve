import type { ServerResponse } from "node:http"

export type ContentDisposition = "inline" | "attachment"

export type DocHeadersInput = {
  contentType?: string
  disposition: ContentDisposition
  filename?: string
  contentLength?: number
}

function toAsciiSafeFilename(filename: string): string {
  return filename.replaceAll(/[^\w.\- ]/g, "_")
}

export function setDocHeaders(res: ServerResponse, input: DocHeadersInput): void {
  const contentType = input.contentType ?? "application/octet-stream"
  res.setHeader("Content-Type", contentType)
  res.setHeader("X-Content-Type-Options", "nosniff")

  if (typeof input.contentLength === "number" && Number.isFinite(input.contentLength)) {
    res.setHeader("Content-Length", String(input.contentLength))
  }

  if (input.filename) {
    const safe = toAsciiSafeFilename(input.filename)
    res.setHeader("Content-Disposition", `${input.disposition}; filename="${safe}"`)
  } else {
    res.setHeader("Content-Disposition", input.disposition)
  }
}

