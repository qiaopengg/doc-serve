import { test } from "node:test"
import assert from "node:assert"
import { createDocxFromParagraphs, extractParagraphsFromDocx } from "../docStore/docxGenerator.js"

test("extractParagraphsFromDocx returns paragraphs", async () => {
  const paragraphs = await extractParagraphsFromDocx(Buffer.alloc(0))
  assert.ok(Array.isArray(paragraphs))
  assert.ok(paragraphs.length > 0)
  assert.ok(paragraphs[0].text)
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

test("createDocxFromParagraphs with multiple paragraphs", async () => {
  const paragraphs = await extractParagraphsFromDocx(Buffer.alloc(0))
  
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
