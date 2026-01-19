import test from "node:test"
import assert from "node:assert/strict"
import { mkdtemp, writeFile, readFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { createRequire } from "node:module"
import type { Readable } from "node:stream"
import type { Entry, ZipFile } from "yauzl"
import { buildServer } from "../server.js"
import { FsDocStore } from "../docStore/fsDocStore.js"

async function readZipEntryFromBuffer(docxBuffer: Buffer, entryName: string): Promise<Buffer | undefined> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")

  return await new Promise<Buffer | undefined>((resolve, reject) => {
    yauzl.fromBuffer(docxBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
      if (err || !zipfile) {
        reject(err ?? new Error("zip_open_failed"))
        return
      }

      const finish = (buf: Buffer | undefined) => {
        try {
          zipfile.close()
        } catch (e) {
          void e
        }
        resolve(buf)
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

async function startServer(docsDir: string): Promise<{ baseUrl: string; close: () => Promise<void> }> {
  const server = buildServer({ docStore: new FsDocStore(docsDir) })
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve))
  const address = server.address()
  assert.ok(address && typeof address === "object")

  return {
    baseUrl: `http://127.0.0.1:${address.port}`,
    close: async () => {
      await new Promise<void>((resolve, reject) => {
        server.close((err: unknown) => (err ? reject(err) : resolve()))
      })
    }
  }
}

test("stream returns framed chunks for text.doc", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const content = Buffer.from("fake-doc-bytes-" + "n".repeat(96))
  await writeFile(join(docsDir, "text.doc"), content)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/stream`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    assert.equal(res.headers.get("access-control-allow-origin"), "*")
    assert.equal(res.headers.get("x-wps-stream-mode"), "framed-chunks")
    assert.equal(res.headers.get("x-wps-chunk-size"), "1024")
    const delayMs = Number(res.headers.get("x-wps-delay-ms"))
    assert.ok(Number.isFinite(delayMs))
    assert.ok(delayMs >= 100 && delayMs <= 200)
    const framed = Buffer.from(await res.arrayBuffer())
    let offset = 0
    const parts: Buffer[] = []
    while (offset + 4 <= framed.length) {
      const len = framed.readUInt32BE(offset)
      offset += 4
      if (len === 0) break
      parts.push(framed.subarray(offset, offset + len))
      offset += len
    }
    assert.deepEqual(Buffer.concat(parts), content)
  } finally {
    await app.close()
  }
})

test("stream-docx returns full docx chunks and keeps merged table structure", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const mock = await readFile(new URL("../../src/mock.docx", import.meta.url))
  await writeFile(join(docsDir, "mock.docx"), mock)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/stream-docx?docId=mock.docx`, {
      headers: { Origin: "http://example.com" }
    })
    assert.equal(res.status, 200)
    assert.equal(res.headers.get("x-wps-stream-mode"), "docx-chunks")
    assert.equal(res.headers.get("x-wps-pagesetup-unit"), "point")
    assert.equal(res.headers.get("x-wps-pagesetup-source-unit"), "twip")
    const pageSetupRaw = res.headers.get("x-wps-pagesetup")
    assert.ok(pageSetupRaw)
    const pageSetup = JSON.parse(decodeURIComponent(pageSetupRaw)) as Record<string, unknown>
    assert.ok(Math.abs(Number(pageSetup.pageWidth) - 11906 / 20) < 1e-6)
    assert.ok(Math.abs(Number(pageSetup.leftMargin) - 1440 / 20) < 1e-6)

    const framed = Buffer.from(await res.arrayBuffer())
    let offset = 0
    const parts: Buffer[] = []
    while (offset + 4 <= framed.length) {
      const len = framed.readUInt32BE(offset)
      offset += 4
      if (len === 0) break
      parts.push(framed.subarray(offset, offset + len))
      offset += len
    }

    assert.ok(parts.length >= 1)
    for (const p of parts) {
      assert.equal(p[0], 0x50)
      assert.equal(p[1], 0x4b)
    }

    const withTable = await (async () => {
      for (const p of parts) {
        const docXmlBuf = await readZipEntryFromBuffer(p, "word/document.xml")
        if (!docXmlBuf) continue
        const xml = docXmlBuf.toString("utf-8")
        if (xml.includes("<w:tbl")) return xml
      }
      return undefined
    })()

    assert.ok(withTable)
    assert.ok(withTable.includes('w:fill="808080"'))
    assert.ok(withTable.includes("<w:gridSpan"))
    assert.ok(withTable.includes('w:val="4"'))
  } finally {
    await app.close()
  }
})

test("frames-docx streams preview.pageSetup + preview.paragraphs with increasing seq", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const mock = await readFile(new URL("../../src/mock.docx", import.meta.url))
  await writeFile(join(docsDir, "mock.docx"), mock)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/frames-docx?docId=mock.docx`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    assert.equal(res.headers.get("access-control-allow-origin"), "*")
    assert.equal(res.headers.get("x-wps-stream-mode"), "ndjson-frames")

    const frames = (await res.text())
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { docId: string; seq: number; type: string; payload: any })

    assert.ok(frames.length >= 3)
    for (let i = 0; i < frames.length; i += 1) {
      assert.equal(frames[i]?.seq, i + 1)
    }

    assert.equal(frames[0]?.type, "preview.pageSetup")
    assert.equal(frames[0]?.payload?.unit, "point")
    assert.equal(frames[0]?.payload?.sourceUnit, "twip")
    assert.equal(frames[0]?.payload?.sourceSectionProperties?.page?.size?.width, 11906)
    assert.equal(frames[0]?.payload?.sourceSectionProperties?.page?.margin?.left, 1440)

    const firstParas = frames.find((f) => f.type === "preview.paragraphs")
    assert.ok(firstParas?.payload?.paragraphs?.length === 1)
    assert.equal(firstParas?.payload?.unit, "twip")

    const final = frames.find((f) => f.type === "final.docx.url")
    assert.ok(final?.payload?.url)
  } finally {
    await app.close()
  }
})

test("frames returns preview.runs + final.docx.url + done with increasing seq", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const content = Buffer.from("fake-docx-bytes-" + "x".repeat(96))
  await writeFile(join(docsDir, "text.docx"), content)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/text/frames`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    assert.equal(res.headers.get("access-control-allow-origin"), "*")
    assert.equal(res.headers.get("x-wps-stream-mode"), "ndjson-frames")
    const delayMs = Number(res.headers.get("x-wps-delay-ms"))
    assert.ok(Number.isFinite(delayMs))
    assert.ok(delayMs >= 100 && delayMs <= 200)

    const body = await res.text()
    const frames = body
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { docId: string; seq: number; type: string; payload: any })

    assert.ok(frames.length >= 3)
    assert.equal(frames[0]?.docId, "text")
    for (let i = 0; i < frames.length; i += 1) {
      assert.equal(frames[i]?.seq, i + 1)
    }

    const last = frames[frames.length - 1]
    assert.equal(last?.type, "control.done")

    const final = frames.find((f) => f.type === "final.docx.url")
    assert.ok(final?.payload?.url)
    const downloadRes = await fetch(final.payload.url)
    assert.equal(downloadRes.status, 200)
    const downloaded = Buffer.from(await downloadRes.arrayBuffer())
    assert.deepEqual(downloaded, content)
  } finally {
    await app.close()
  }
})

test("frames alias supports query docId", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  await writeFile(join(docsDir, "text.docx"), Buffer.from("fake-docx-bytes"))

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/frames?docId=text`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    const lines = (await res.text())
      .split("\n")
      .map((l) => l.trim())
      .filter(Boolean)
    assert.ok(lines.length >= 2)
    const first = JSON.parse(lines[0]!) as { docId: string; seq: number }
    assert.equal(first.docId, "text")
    assert.equal(first.seq, 1)
  } finally {
    await app.close()
  }
})

test("frames default uses test.docx and final url points to /docs/download", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const content = Buffer.from("fake-test-docx-bytes-" + "y".repeat(96))
  await writeFile(join(docsDir, "test.docx"), content)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/frames`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    const frames = (await res.text())
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { docId: string; seq: number; type: string; payload: any })

    assert.ok(frames.length >= 3)
    assert.equal(frames[0]?.docId, "test.docx")
    for (let i = 0; i < frames.length; i += 1) {
      assert.equal(frames[i]?.seq, i + 1)
    }

    const final = frames.find((f) => f.type === "final.docx.url")
    assert.ok(final?.payload?.url)
    assert.match(String(final.payload.url), /\/api\/v1\/docs\/download$/)
    const downloadRes = await fetch(final.payload.url)
    assert.equal(downloadRes.status, 200)
    const downloaded = Buffer.from(await downloadRes.arrayBuffer())
    assert.deepEqual(downloaded, content)
  } finally {
    await app.close()
  }
})

test("frames default falls back to text.docx when test.docx missing", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const content = Buffer.from("fake-text-docx-bytes-" + "q".repeat(96))
  await writeFile(join(docsDir, "text.docx"), content)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/frames`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    const frames = (await res.text())
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { docId: string; type: string; payload: any })

    assert.ok(frames.length >= 3)
    assert.equal(frames[0]?.docId, "text.docx")

    const final = frames.find((f) => f.type === "final.docx.url")
    assert.ok(final?.payload?.url)
    const downloadRes = await fetch(final.payload.url)
    assert.equal(downloadRes.status, 200)
    const downloaded = Buffer.from(await downloadRes.arrayBuffer())
    assert.deepEqual(downloaded, content)
  } finally {
    await app.close()
  }
})

test("frames default falls back to text.dock before text.doc", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const dockContent = Buffer.from("fake-text-dock-bytes-" + "d".repeat(96))
  const docContent = Buffer.from("fake-text-doc-bytes-" + "c".repeat(96))
  await writeFile(join(docsDir, "text.dock"), dockContent)
  await writeFile(join(docsDir, "text.doc"), docContent)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/frames`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    const frames = (await res.text())
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { docId: string; type: string; payload: any })

    assert.ok(frames.length >= 3)
    assert.equal(frames[0]?.docId, "text.dock")

    const final = frames.find((f) => f.type === "final.docx.url")
    assert.ok(final?.payload?.url)
    const downloadRes = await fetch(final.payload.url)
    assert.equal(downloadRes.status, 200)
    const downloaded = Buffer.from(await downloadRes.arrayBuffer())
    assert.deepEqual(downloaded, dockContent)
  } finally {
    await app.close()
  }
})

test("/docs/frames works without /api/v1 prefix", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const content = Buffer.from("fake-test-docx-bytes-" + "z".repeat(96))
  await writeFile(join(docsDir, "test.docx"), content)

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/docs/frames`, { headers: { Origin: "http://example.com" } })
    assert.equal(res.status, 200)
    const frames = (await res.text())
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { type: string; payload: any })
    const final = frames.find((f) => f.type === "final.docx.url")
    assert.ok(final?.payload?.url)
    assert.match(String(final.payload.url), /\/docs\/download$/)
    const downloadRes = await fetch(final.payload.url)
    assert.equal(downloadRes.status, 200)
    const downloaded = Buffer.from(await downloadRes.arrayBuffer())
    assert.deepEqual(downloaded, content)
  } finally {
    await app.close()
  }
})

test("frames can return control.error when generation fails", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  await writeFile(join(docsDir, "text.docx"), Buffer.from("fake-docx-bytes"))

  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/text/frames`, {
      headers: { Origin: "http://example.com", "x-wps-simulate-error": "1" }
    })
    assert.equal(res.status, 200)
    const frames = (await res.text())
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { type: string; seq: number })

    assert.ok(frames.length >= 2)
    for (let i = 0; i < frames.length; i += 1) {
      assert.equal(frames[i]?.seq, i + 1)
    }
    assert.equal(frames[frames.length - 1]?.type, "control.error")
  } finally {
    await app.close()
  }
})

test("missing doc returns 404", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/stream`)
    assert.equal(res.status, 404)
    const body = await res.json()
    assert.deepEqual(body, { error: "document_not_found" })
  } finally {
    await app.close()
  }
})

test("cors preflight returns 204", async () => {
  const docsDir = await mkdtemp(join(tmpdir(), "wps-docs-"))
  const app = await startServer(docsDir)
  try {
    const res = await fetch(`${app.baseUrl}/api/v1/docs/stream`, {
      method: "OPTIONS",
      headers: {
        Origin: "http://example.com",
        "Access-Control-Request-Method": "GET",
        "Access-Control-Request-Headers": "x-test"
      }
    })
    assert.equal(res.status, 204)
    assert.equal(res.headers.get("access-control-allow-origin"), "*")
    assert.match(res.headers.get("access-control-allow-headers") ?? "", /x-test/i)
  } finally {
    await app.close()
  }
})
