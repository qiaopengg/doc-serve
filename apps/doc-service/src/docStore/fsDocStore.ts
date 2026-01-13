import { createReadStream } from "node:fs"
import { stat } from "node:fs/promises"
import { basename, join } from "node:path"
import { HttpError } from "@wps/doc-core"
import type { DocStore, DocStream } from "./types.js"

function guessContentType(filename: string): string {
  if (filename.toLowerCase().endsWith(".docx")) {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  }
  if (filename.toLowerCase().endsWith(".doc")) {
    return "application/msword"
  }
  return "application/octet-stream"
}

export class FsDocStore implements DocStore {
  readonly docsDir: string

  constructor(docsDir: string) {
    this.docsDir = docsDir
  }

  async openStream(docId: string): Promise<DocStream> {
    const normalized = docId.trim()
    if (!normalized) {
      throw new HttpError(400, "invalid_doc_id")
    }

    const rawFilename = normalized.includes(".") ? normalized : `${normalized}.docx`
    const safeFilename = basename(rawFilename)
    if (safeFilename !== rawFilename) {
      throw new HttpError(400, "invalid_doc_id")
    }

    const filePath = join(this.docsDir, safeFilename)

    let fileStat: { size: number }
    try {
      fileStat = await stat(filePath)
    } catch {
      throw new HttpError(404, "document_not_found")
    }

    const stream = createReadStream(filePath, { highWaterMark: 64 * 1024 })
    return {
      meta: {
        docId: normalized,
        filename: safeFilename,
        contentType: guessContentType(safeFilename),
        contentLength: fileStat.size
      },
      stream
    }
  }
}
