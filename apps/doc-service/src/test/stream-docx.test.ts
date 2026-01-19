import { test } from "node:test"
import assert from "node:assert"
import { readFile } from "node:fs/promises"
import { createRequire } from "node:module"
import type { Readable } from "node:stream"
import type { Entry, ZipFile } from "yauzl"
import { createDocxFromParagraphs, createDocxFromSourceDocxSlice, extractParagraphsFromDocx } from "../docStore/docxGenerator.js"

function findTag(xml: string, tagName: string): string | undefined {
  const re = new RegExp(`<${tagName}\\b[\\s\\S]*?(?:\\/?>|<\\/${tagName}>)`, "i")
  const m = xml.match(re)
  return m?.[0]
}

function lastSectPrXml(xml: string): string | undefined {
  const matches = [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)]
  return matches.length ? matches[matches.length - 1]![0] : undefined
}

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

async function replaceZipEntryFromBuffer(docxBuffer: Buffer, entryName: string, replacement: Buffer): Promise<Buffer> {
  const require = createRequire(import.meta.url)
  const yauzl = require("yauzl") as typeof import("yauzl")
  const yazl = require("yazl") as typeof import("yazl")

  const zipOut = new yazl.ZipFile()
  const outPromise = new Promise<Buffer>((resolve, reject) => {
    const chunks: Buffer[] = []
    ;(zipOut.outputStream as unknown as Readable).on("data", (c: Buffer) => chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c)))
    ;(zipOut.outputStream as unknown as Readable).on("end", () => resolve(Buffer.concat(chunks)))
    ;(zipOut.outputStream as unknown as Readable).on("error", reject)
  })

  const writing = new Promise<void>((resolve, reject) => {
    yauzl.fromBuffer(docxBuffer, { lazyEntries: true }, (err: Error | null, zipfile?: ZipFile) => {
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
        zipfile.openReadStream(entry, (err2: Error | null, stream?: Readable) => {
          if (err2 || !stream) {
            reject(err2 ?? new Error("zip_entry_stream_failed"))
            return
          }
          const chunks: Buffer[] = []
          stream.on("data", (c: Buffer) => chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c)))
          stream.on("end", () => {
            const data = entry.fileName === entryName ? replacement : Buffer.concat(chunks)
            zipOut.addBuffer(data, entry.fileName)
            zipfile.readEntry()
          })
          stream.on("error", reject)
        })
      })

      zipfile.on("end", resolve)
      zipfile.on("error", reject)
    })
  })

  await writing
  zipOut.end()
  return await outPromise
}

function firstTag(xml: string, tagName: string): string | undefined {
  const re = new RegExp(`<${tagName}\\b[^>]*(?:\\/|)>`, "i")
  const m = xml.match(re)
  return m?.[0]
}

function attrValue(tag: string | undefined, name: string): string | undefined {
  if (!tag) return undefined
  const re = new RegExp(`${name}="([^"]+)"`)
  const m = tag.match(re)
  return m?.[1]
}

test("extractParagraphsFromDocx returns paragraphs", async () => {
  const docxBuffer = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  assert.ok(Array.isArray(paragraphs))
  assert.ok(paragraphs.length > 0)
  assert.ok(paragraphs[0].text)
})

test("extractParagraphsFromDocx preserves key styles in mock.docx", async () => {
  const docxBuffer = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  const title = paragraphs.find((p) => p.text.includes("AI 应用"))!
  assert.ok(title)
  assert.strictEqual(title.alignment, "center")
  assert.strictEqual(Boolean(title.bold), true)

  const table = paragraphs.find((p) => p.isTable)!
  assert.ok(table)
  assert.ok(Array.isArray(table.tableData))
  assert.ok(table.tableData!.length >= 2)
  assert.ok(table.tableCellStyles?.[0]?.[0]?.fill)
  assert.ok(table.tableCellStyles?.[0]?.[0]?.gridSpan)
})

test("extractParagraphsFromDocx keeps justify alignment", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const pXml = `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>JUSTIFY_TEST</w:t></w:r></w:p>`
  const sectPrIndex = srcXml.lastIndexOf("<w:sectPr")
  const insertAt = sectPrIndex >= 0 ? sectPrIndex : srcXml.lastIndexOf("</w:body>")
  assert.ok(insertAt > 0)
  const modifiedXml = srcXml.slice(0, insertAt) + pXml + srcXml.slice(insertAt)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(modifiedXml, "utf-8"))
  const paragraphs = await extractParagraphsFromDocx(modified)
  const p = paragraphs.find((x) => x.text.includes("JUSTIFY_TEST"))
  assert.ok(p)
  assert.strictEqual(p.alignment, "justify")
})

test("extractParagraphsFromDocx reads font size from w:szCs", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const pXml = `<w:p><w:r><w:rPr><w:szCs w:val="36"/></w:rPr><w:t>SZCS_TEST</w:t></w:r></w:p>`
  const sectPrIndex = srcXml.lastIndexOf("<w:sectPr")
  const insertAt = sectPrIndex >= 0 ? sectPrIndex : srcXml.lastIndexOf("</w:body>")
  assert.ok(insertAt > 0)
  const modifiedXml = srcXml.slice(0, insertAt) + pXml + srcXml.slice(insertAt)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(modifiedXml, "utf-8"))
  const paragraphs = await extractParagraphsFromDocx(modified)
  const p = paragraphs.find((x) => x.text.includes("SZCS_TEST"))
  assert.ok(p)
  assert.ok(p.runs && p.runs.length === 1)
  assert.strictEqual(p.runs[0]!.fontSize, 18)
})

test("extractParagraphsFromDocx reads underline none as explicit false", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const pXml = `<w:p><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>UNDERLINE_NONE_TEST</w:t></w:r></w:p>`
  const sectPrIndex = srcXml.lastIndexOf("<w:sectPr")
  const insertAt = sectPrIndex >= 0 ? sectPrIndex : srcXml.lastIndexOf("</w:body>")
  assert.ok(insertAt > 0)
  const modifiedXml = srcXml.slice(0, insertAt) + pXml + srcXml.slice(insertAt)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(modifiedXml, "utf-8"))
  const paragraphs = await extractParagraphsFromDocx(modified)
  const p = paragraphs.find((x) => x.text.includes("UNDERLINE_NONE_TEST"))
  assert.ok(p)
  assert.ok(p.runs && p.runs.length === 1)
  assert.strictEqual(p.runs[0]!.underline, false)
})

test("extractParagraphsFromDocx keeps on/off semantics for w:b", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const pXml = `<w:p><w:pPr><w:rPr><w:b w:val="0"/></w:rPr></w:pPr><w:r><w:t>NO_B_NODE_IN_RUN</w:t></w:r></w:p>`
  const sectPrIndex = srcXml.lastIndexOf("<w:sectPr")
  const insertAt = sectPrIndex >= 0 ? sectPrIndex : srcXml.lastIndexOf("</w:body>")
  assert.ok(insertAt > 0)
  const modifiedXml = srcXml.slice(0, insertAt) + pXml + srcXml.slice(insertAt)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(modifiedXml, "utf-8"))
  const paragraphs = await extractParagraphsFromDocx(modified)
  const p = paragraphs.find((x) => x.text.includes("NO_B_NODE_IN_RUN"))
  assert.ok(p)
  assert.ok(p.runs && p.runs.length === 1)
  assert.strictEqual(p.runs[0]!.bold, false)
})

test("extractParagraphsFromDocx treats <w:b/> as true", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const pXml = `<w:p><w:pPr><w:rPr><w:b w:val="0"/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>B_EMPTY_TAG_TRUE</w:t></w:r></w:p>`
  const sectPrIndex = srcXml.lastIndexOf("<w:sectPr")
  const insertAt = sectPrIndex >= 0 ? sectPrIndex : srcXml.lastIndexOf("</w:body>")
  assert.ok(insertAt > 0)
  const modifiedXml = srcXml.slice(0, insertAt) + pXml + srcXml.slice(insertAt)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(modifiedXml, "utf-8"))
  const paragraphs = await extractParagraphsFromDocx(modified)
  const p = paragraphs.find((x) => x.text.includes("B_EMPTY_TAG_TRUE"))
  assert.ok(p)
  assert.ok(p.runs && p.runs.length === 1)
  assert.strictEqual(p.runs[0]!.bold, true)
})

test("createDocxFromParagraphs keeps justify alignment in document.xml", async () => {
  const buffer = await createDocxFromParagraphs([{ text: "x", alignment: "justify" }])
  const docXmlBuf = await readZipEntryFromBuffer(buffer, "word/document.xml")
  assert.ok(docXmlBuf)
  const docXml = docXmlBuf.toString("utf-8")
  assert.ok(/<w:jc\b[^>]*w:val="(both|justify)"/.test(docXml))
})

test("extractParagraphsFromDocx restores merged table layout in mock.docx", async () => {
  const docxBuffer = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  const table = paragraphs.find((p) => p.isTable)!
  assert.ok(table)

  assert.strictEqual(table.tableData?.length, 4)
  assert.strictEqual(table.tableData?.[0]?.length, 4)
  assert.deepStrictEqual(table.tableGridCols, [2310, 2310, 2311, 2311])

  const row0 = table.tableCellStyles?.[0]
  const row1 = table.tableCellStyles?.[1]
  const row2 = table.tableCellStyles?.[2]
  const row3 = table.tableCellStyles?.[3]
  assert.ok(row0 && row1 && row2 && row3)

  assert.strictEqual(row0[0]?.gridSpan, 4)
  assert.strictEqual(row0[0]?.fill, "808080")
  assert.strictEqual(Boolean(row0[1]?.skip), true)
  assert.strictEqual(Boolean(row0[2]?.skip), true)
  assert.strictEqual(Boolean(row0[3]?.skip), true)

  for (let i = 0; i < 4; i += 1) {
    assert.strictEqual(row1[i]?.fill, "808080")
    assert.strictEqual(Boolean(row1[i]?.skip), false)
    assert.strictEqual(row1[i]?.gridSpan, undefined)
  }

  for (let i = 0; i < 4; i += 1) {
    assert.strictEqual(row2[i]?.fill, undefined)
    assert.strictEqual(Boolean(row2[i]?.skip), false)
    assert.strictEqual(row2[i]?.gridSpan, undefined)
    assert.strictEqual(row3[i]?.fill, undefined)
    assert.strictEqual(Boolean(row3[i]?.skip), false)
    assert.strictEqual(row3[i]?.gridSpan, undefined)
  }

  const regenerated = await createDocxFromParagraphs([table])
  const docXmlBuf = await readZipEntryFromBuffer(regenerated, "word/document.xml")
  assert.ok(docXmlBuf)
  const docXml = docXmlBuf.toString("utf-8")

  assert.ok(docXml.includes('w:fill="808080"'))
  assert.ok(docXml.includes("<w:gridSpan"))
  assert.ok(docXml.includes('w:val="4"'))

  const trMatch = docXml.match(/<w:tr\b[\s\S]*?<\/w:tr>/)
  assert.ok(trMatch)
  const firstTrXml = trMatch[0]
  const tcCount = [...firstTrXml.matchAll(/<w:tc\b/g)].length
  assert.strictEqual(tcCount, 1)
})

test("extractParagraphsFromDocx preserves table borders from mock.docx", async () => {
  const docxBuffer = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  const table = paragraphs.find((p) => p.isTable)!
  assert.ok(table)

  assert.ok(table.tableBorders)
  assert.strictEqual(table.tableBorders?.top?.style, "single")
  assert.strictEqual(table.tableBorders?.top?.size, 4)
  assert.strictEqual(table.tableBorders?.top?.color, "auto")

  const regenerated = await createDocxFromParagraphs([table])
  const docXmlBuf = await readZipEntryFromBuffer(regenerated, "word/document.xml")
  assert.ok(docXmlBuf)
  const docXml = docXmlBuf.toString("utf-8")
  assert.ok(docXml.includes("<w:tblBorders"))
  assert.ok(docXml.includes('w:val="single"'))
})

test("createDocxFromParagraphs generates valid docx buffer", async () => {
  const paragraphs = [
    { text: "标题", bold: true, fontSize: 16, headingLevel: 1 as const },
    { text: "正文内容" }
  ]
  
  const buffer = await createDocxFromParagraphs(paragraphs)
  assert.ok(Buffer.isBuffer(buffer))
  assert.ok(buffer.length > 0)
  
  // 验证是否是 ZIP 格式（docx 是 ZIP 压缩的 XML）
  assert.strictEqual(buffer[0], 0x50) // 'P'
  assert.strictEqual(buffer[1], 0x4B) // 'K'
})

test("createDocxFromParagraphs writes paragraph spacing into document.xml", async () => {
  const buffer = await createDocxFromParagraphs([
    { text: "a", spacing: { before: 240, after: 120, line: 360, lineRule: "auto" } }
  ])

  const docXmlBuf = await readZipEntryFromBuffer(buffer, "word/document.xml")
  assert.ok(docXmlBuf)
  const docXml = docXmlBuf.toString("utf-8")

  assert.ok(docXml.includes("w:spacing"))
  assert.ok(docXml.includes('w:before="240"'))
  assert.ok(docXml.includes('w:after="120"'))
})

test("createDocxFromParagraphs applies section properties page margins and size", async () => {
  const buffer = await createDocxFromParagraphs([
    {
      text: "x",
      sectionProperties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720 }
        }
      }
    }
  ])
  const docXmlBuf = await readZipEntryFromBuffer(buffer, "word/document.xml")
  assert.ok(docXmlBuf)
  const docXml = docXmlBuf.toString("utf-8")
  const pgSz = findTag(docXml, "w:pgSz")
  const pgMar = findTag(docXml, "w:pgMar")
  assert.ok(pgSz && pgMar)
  assert.ok(pgSz.includes('w:w="12240"'))
  assert.ok(pgSz.includes('w:h="15840"'))
  assert.ok(pgMar.includes('w:top="1440"'))
  assert.ok(pgMar.includes('w:bottom="1440"'))
  assert.ok(pgMar.includes('w:left="1440"'))
  assert.ok(pgMar.includes('w:right="1440"'))
  assert.ok(pgMar.includes('w:header="720"'))
  assert.ok(pgMar.includes('w:footer="720"'))
})

test("extractParagraphsFromDocx keeps section properties for createDocxFromParagraphs", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")
  const srcSectPr = lastSectPrXml(srcXml)
  assert.ok(srcSectPr)
  const srcPgMar = findTag(srcSectPr, "w:pgMar")
  const srcPgSz = findTag(srcSectPr, "w:pgSz")
  assert.ok(srcPgMar && srcPgSz)

  const paragraphs = await extractParagraphsFromDocx(source)
  assert.ok(paragraphs.length > 0)

  const regenerated = await createDocxFromParagraphs(paragraphs.slice(0, 1))
  const outDocXmlBuf = await readZipEntryFromBuffer(regenerated, "word/document.xml")
  assert.ok(outDocXmlBuf)
  const outXml = outDocXmlBuf.toString("utf-8")
  const outSectPr = lastSectPrXml(outXml)
  assert.ok(outSectPr)
  const outPgMar = findTag(outSectPr, "w:pgMar")
  const outPgSz = findTag(outSectPr, "w:pgSz")
  assert.ok(outPgMar && outPgSz)

  assert.strictEqual(attrValue(outPgMar, "w:top"), attrValue(srcPgMar, "w:top"))
  assert.strictEqual(attrValue(outPgMar, "w:bottom"), attrValue(srcPgMar, "w:bottom"))
  assert.strictEqual(attrValue(outPgMar, "w:left"), attrValue(srcPgMar, "w:left"))
  assert.strictEqual(attrValue(outPgMar, "w:right"), attrValue(srcPgMar, "w:right"))
  assert.strictEqual(attrValue(outPgMar, "w:header"), attrValue(srcPgMar, "w:header"))
  assert.strictEqual(attrValue(outPgMar, "w:footer"), attrValue(srcPgMar, "w:footer"))

  assert.strictEqual(attrValue(outPgSz, "w:w"), attrValue(srcPgSz, "w:w"))
  assert.strictEqual(attrValue(outPgSz, "w:h"), attrValue(srcPgSz, "w:h"))

  const srcOrient = attrValue(srcPgSz, "w:orient")
  const outOrient = attrValue(outPgSz, "w:orient")
  if (srcOrient) assert.strictEqual(outOrient, srcOrient)
})

test("extractParagraphsFromDocx maps section properties per position for multi-section documents", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const sectPrXml = lastSectPrXml(srcXml)
  assert.ok(sectPrXml)
  const srcPgMar = findTag(sectPrXml, "w:pgMar")
  assert.ok(srcPgMar)

  const existingLeft = attrValue(srcPgMar, "w:left") ?? "1440"
  const leftA = existingLeft === "1111" ? "1222" : "1111"
  const leftB = existingLeft === "2222" ? "2333" : "2222"

  const sectPrA = sectPrXml.replace(/w:left="[^"]*"/, `w:left="${leftA}"`)
  const sectPrB = sectPrXml.replace(/w:left="[^"]*"/, `w:left="${leftB}"`)

  const sectPrIndex = srcXml.lastIndexOf(sectPrXml)
  assert.ok(sectPrIndex >= 0)
  let xmlWithTwoSections = srcXml.slice(0, sectPrIndex) + sectPrB + srcXml.slice(sectPrIndex + sectPrXml.length)

  const pPrCloses = [...xmlWithTwoSections.matchAll(/<\/w:pPr>/g)].map((m) => m.index ?? -1).filter((n) => n >= 0)
  assert.ok(pPrCloses.length >= 2)
  const insertAt = pPrCloses[Math.floor(pPrCloses.length / 2)]!
  xmlWithTwoSections =
    xmlWithTwoSections.slice(0, insertAt) + sectPrA + xmlWithTwoSections.slice(insertAt)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(xmlWithTwoSections, "utf-8"))
  const paragraphs = await extractParagraphsFromDocx(modified)
  assert.ok(paragraphs.length >= 3)

  const first = paragraphs[0]!
  const last = paragraphs[paragraphs.length - 1]!
  assert.strictEqual(String(first.sectionProperties?.page?.margin?.left ?? ""), leftA)
  assert.strictEqual(String(last.sectionProperties?.page?.margin?.left ?? ""), leftB)

  const outA = await createDocxFromParagraphs([first])
  const outADocXmlBuf = await readZipEntryFromBuffer(outA, "word/document.xml")
  assert.ok(outADocXmlBuf)
  const outAXml = outADocXmlBuf.toString("utf-8")
  const outASectPr = lastSectPrXml(outAXml)
  assert.ok(outASectPr)
  const outAPgMar = findTag(outASectPr, "w:pgMar")
  assert.ok(outAPgMar)
  assert.strictEqual(attrValue(outAPgMar, "w:left"), leftA)

  const outB = await createDocxFromParagraphs([last])
  const outBDocXmlBuf = await readZipEntryFromBuffer(outB, "word/document.xml")
  assert.ok(outBDocXmlBuf)
  const outBXml = outBDocXmlBuf.toString("utf-8")
  const outBSectPr = lastSectPrXml(outBXml)
  assert.ok(outBSectPr)
  const outBPgMar = findTag(outBSectPr, "w:pgMar")
  assert.ok(outBPgMar)
  assert.strictEqual(attrValue(outBPgMar, "w:left"), leftB)
})

test("createDocxFromParagraphs with multiple paragraphs", async () => {
  const docxBuffer = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const paragraphs = await extractParagraphsFromDocx(docxBuffer)
  
  for (let i = 1; i <= paragraphs.length; i++) {
    const subset = paragraphs.slice(0, i)
    const buffer = await createDocxFromParagraphs(subset)
    
    assert.ok(Buffer.isBuffer(buffer))
    assert.ok(buffer.length > 0)
    
    // 每个都应该是有效的 ZIP/docx
    assert.strictEqual(buffer[0], 0x50)
    assert.strictEqual(buffer[1], 0x4B)
  }
})

test("createDocxFromSourceDocxSlice generates valid docx buffers", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const paragraphs = await extractParagraphsFromDocx(source)
  assert.ok(paragraphs.length > 0)

  const buf1 = await createDocxFromSourceDocxSlice(source, 1)
  assert.ok(buf1.length > 0)
  assert.strictEqual(buf1[0], 0x50)
  assert.strictEqual(buf1[1], 0x4b)

  const bufAll = await createDocxFromSourceDocxSlice(source, paragraphs.length)
  assert.ok(bufAll.length > 0)
  assert.strictEqual(bufAll[0], 0x50)
  assert.strictEqual(bufAll[1], 0x4b)
})

test("createDocxFromSourceDocxSlice keeps section properties when sectPr is stored in paragraph pPr", async () => {
  const source = await readFile(new URL("../../src/mock.docx", import.meta.url))
  const srcDocXmlBuf = await readZipEntryFromBuffer(source, "word/document.xml")
  assert.ok(srcDocXmlBuf)
  const srcXml = srcDocXmlBuf.toString("utf-8")

  const matches = [...srcXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)]
  assert.ok(matches.length >= 1)
  const last = matches[matches.length - 1]!
  const sectPrXml = last[0]
  const sectPrIndex = last.index ?? -1
  assert.ok(sectPrIndex >= 0)

  let withoutBodySectPr = srcXml.slice(0, sectPrIndex) + srcXml.slice(sectPrIndex + sectPrXml.length)
  const lastPprClose = withoutBodySectPr.lastIndexOf("</w:pPr>")
  assert.ok(lastPprClose > 0)
  withoutBodySectPr = withoutBodySectPr.slice(0, lastPprClose) + sectPrXml + withoutBodySectPr.slice(lastPprClose)

  const modified = await replaceZipEntryFromBuffer(source, "word/document.xml", Buffer.from(withoutBodySectPr, "utf-8"))
  const chunk = await createDocxFromSourceDocxSlice(modified, 1)
  const chunkDocXmlBuf = await readZipEntryFromBuffer(chunk, "word/document.xml")
  assert.ok(chunkDocXmlBuf)
  const chunkXml = chunkDocXmlBuf.toString("utf-8")

  assert.ok(chunkXml.includes("w:sectPr"))

  const srcPgMar = firstTag(withoutBodySectPr, "w:pgMar")
  const chunkPgMar = firstTag(chunkXml, "w:pgMar")
  assert.ok(srcPgMar)
  assert.ok(chunkPgMar)
  assert.strictEqual(attrValue(chunkPgMar, "w:left"), attrValue(srcPgMar, "w:left"))
  assert.strictEqual(attrValue(chunkPgMar, "w:right"), attrValue(srcPgMar, "w:right"))
})
