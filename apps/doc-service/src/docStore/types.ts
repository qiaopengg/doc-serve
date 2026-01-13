import type { Readable } from "node:stream"

export type DocMeta = {
  docId: string
  filename: string
  contentType: string
  contentLength?: number
}

export type DocStream = {
  meta: DocMeta
  stream: Readable
}

export interface DocStore {
  openStream(docId: string): Promise<DocStream>
}

