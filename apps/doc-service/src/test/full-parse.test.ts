import { readFile } from "node:fs/promises"
import { test } from "node:test"
import assert from "node:assert"
import { parseFullDocx, getDocxStatistics } from "../docStore/docxFullParser.js"
import { join } from "node:path"

test("完整解析 docx 文档", async () => {
  // 尝试读取测试文档
  let docxBuffer: Buffer
  try {
    docxBuffer = await readFile(join(process.cwd(), "src", "mock.docx"))
  } catch {
    console.log("跳过测试：未找到 mock.docx")
    return
  }

  const fullDoc = await parseFullDocx(docxBuffer)
  
  // 验证基本结构
  assert.ok(fullDoc.paragraphs, "应该有段落数组")
  assert.ok(Array.isArray(fullDoc.paragraphs), "段落应该是数组")
  
  console.log("解析结果：")
  console.log(`- 段落数量: ${fullDoc.paragraphs.length}`)
  console.log(`- 元数据: ${fullDoc.metadata ? "有" : "无"}`)
  console.log(`- 页眉: ${fullDoc.headers?.size || 0} 个`)
  console.log(`- 页脚: ${fullDoc.footers?.size || 0} 个`)
  console.log(`- 注释: ${fullDoc.comments?.size || 0} 个`)
  console.log(`- 脚注: ${fullDoc.footnotes?.size || 0} 个`)
  console.log(`- 尾注: ${fullDoc.endnotes?.size || 0} 个`)
  console.log(`- 图片: ${fullDoc.images?.size || 0} 个`)
  
  // 获取统计信息
  const stats = getDocxStatistics(fullDoc)
  console.log("\n文档统计：")
  console.log(`- 段落: ${stats.paragraphCount}`)
  console.log(`- 表格: ${stats.tableCount}`)
  console.log(`- 图片: ${stats.imageCount}`)
  console.log(`- 字数: ${stats.wordCount}`)
  console.log(`- 字符数: ${stats.characterCount}`)
  console.log(`- 有编号: ${stats.hasNumbering}`)
  console.log(`- 有注释: ${stats.hasComments}`)
  console.log(`- 样式ID: ${stats.styleIds.join(", ")}`)
  
  // 检查扩展属性
  let hasExtendedFeatures = false
  for (const para of fullDoc.paragraphs) {
    if (para.images?.length) {
      console.log(`\n段落包含 ${para.images.length} 个图片`)
      hasExtendedFeatures = true
    }
    if (para.numbering) {
      console.log(`\n段落有编号: numId=${para.numbering.numId}, level=${para.numbering.level}, format=${para.numbering.format}`)
      hasExtendedFeatures = true
    }
    if (para.bookmarks?.length) {
      console.log(`\n段落包含 ${para.bookmarks.length} 个书签`)
      hasExtendedFeatures = true
    }
    if (para.fields?.length) {
      console.log(`\n段落包含 ${para.fields.length} 个域代码`)
      for (const field of para.fields) {
        console.log(`  - 类型: ${field.fieldType}, 代码: ${field.code}`)
      }
      hasExtendedFeatures = true
    }
    if (para.notes?.length) {
      console.log(`\n段落包含 ${para.notes.length} 个脚注/尾注引用`)
      hasExtendedFeatures = true
    }
    if (para.comments?.length) {
      console.log(`\n段落包含 ${para.comments.length} 个注释标记`)
      hasExtendedFeatures = true
    }
  }
  
  if (!hasExtendedFeatures) {
    console.log("\n注意：文档不包含扩展功能（图片、编号、书签等）")
  }
  
  // 显示元数据
  if (fullDoc.metadata?.core) {
    console.log("\n核心元数据：")
    console.log(`- 标题: ${fullDoc.metadata.core.title || "无"}`)
    console.log(`- 作者: ${fullDoc.metadata.core.creator || "无"}`)
    console.log(`- 创建时间: ${fullDoc.metadata.core.created || "无"}`)
    console.log(`- 修改时间: ${fullDoc.metadata.core.modified || "无"}`)
  }
  
  if (fullDoc.metadata?.app) {
    console.log("\n应用程序属性：")
    console.log(`- 应用: ${fullDoc.metadata.app.application || "无"}`)
    console.log(`- 页数: ${fullDoc.metadata.app.pages || "无"}`)
    console.log(`- 字数: ${fullDoc.metadata.app.words || "无"}`)
  }
})

test("解析包含表格的文档", async () => {
  let docxBuffer: Buffer
  try {
    docxBuffer = await readFile(join(process.cwd(), "src", "mock.docx"))
  } catch {
    console.log("跳过测试：未找到 mock.docx")
    return
  }

  const fullDoc = await parseFullDocx(docxBuffer)
  const stats = getDocxStatistics(fullDoc)
  
  console.log(`\n表格测试 - 找到 ${stats.tableCount} 个表格`)
  
  for (const para of fullDoc.paragraphs) {
    if (para.isTable && para.tableData) {
      console.log(`\n表格: ${para.tableData.length} 行 x ${para.tableData[0]?.length || 0} 列`)
      if (para.tableCellStyles) {
        console.log("包含单元格样式信息")
      }
      if (para.tableBorders) {
        console.log("包含表格边框信息")
      }
    }
  }
})
