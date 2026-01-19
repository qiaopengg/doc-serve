import { createRequire } from "node:module"
import type { Entry, ZipFile } from "yauzl"
import type { Readable } from "node:stream"

/**
 * 读取 ZIP 文件中的指定条目
 */
export async function readZipEntry(zipBuffer: Buffer, entryName: string): Promise<Buffer | undefined> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")

  return await new Promise((resolve, reject) => {
    yauzl.fromBuffer(zipBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
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

/**
 * 列出 ZIP 文件中的所有条目
 */
export async function listZipEntries(zipBuffer: Buffer): Promise<string[]> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")

  return await new Promise((resolve, reject) => {
    yauzl.fromBuffer(zipBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
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
 * 读取 Readable 流的所有内容
 */
async function readAllReadable(stream: Readable): Promise<Buffer> {
  const chunks: Buffer[] = []
  for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
  return Buffer.concat(chunks)
}

/**
 * 替换 ZIP 文件中的指定条目
 */
export async function replaceZipEntry(zipBuffer: Buffer, entryName: string, replacement: Buffer): Promise<Buffer> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")
  const yazl = require("yazl") as typeof import("yazl")

  const zipOut = new yazl.ZipFile()
  const outPromise = readAllReadable(zipOut.outputStream as unknown as Readable)

  await new Promise<void>((resolve, reject) => {
    yauzl.fromBuffer(zipBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
      if (err || !zipfile) {
        reject(err ?? new Error("zip_open_failed"))
        return
      }

      zipfile.readEntry()
      zipfile.on("entry", (entry: Entry) => {
        if (entry.fileName.endsWith("/")) {
          zipOut.addEmptyDirectory(entry.fileName)
          zipfile.readEntry()
          return
        }
        zipfile.openReadStream(entry, async (err2: Error | null, stream?: Readable) => {
          if (err2 || !stream) {
            reject(err2 ?? new Error("zip_entry_stream_failed"))
            return
          }

          try {
            const data = entry.fileName === entryName ? replacement : await readAllReadable(stream)
            zipOut.addBuffer(data, entry.fileName)
            zipfile.readEntry()
          } catch (e) {
            reject(e)
          }
        })
      })

      zipfile.on("end", () => {
        resolve()
      })
      zipfile.on("error", reject)
    })
  })

  zipOut.end()
  return await outPromise
}
