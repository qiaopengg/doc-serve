#!/usr/bin/env node

/**
 * 测试表格完整返回功能
 * 验证表格作为一个完整单元返回
 */

import http from 'node:http'
import { exec } from 'node:child_process'
import { writeFileSync, unlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const API_URL = 'http://localhost:3002/api/v1/docs/stream-docx?docId=mock.docx'

console.log('🧪 测试表格完整返回功能...\n')

const req = http.get(API_URL, (res) => {
  console.log(`✅ 状态码: ${res.statusCode}`)
  console.log(`✅ Content-Type: ${res.headers['content-type']}`)
  console.log(`✅ Stream Mode: ${res.headers['x-wps-stream-mode']}`)
  
  const pageSetup = res.headers['x-wps-pagesetup']
  if (pageSetup) {
    const decoded = JSON.parse(decodeURIComponent(pageSetup))
    console.log(`✅ 页面设置: ${decoded.pageWidth}x${decoded.pageHeight} pt`)
  }
  
  console.log('\n📦 接收 chunks...\n')
  
  let buffer = Buffer.alloc(0)
  let chunkCount = 0
  let tableChunkCount = 0
  const pendingChecks = []
  
  res.on('data', (chunk) => {
    buffer = Buffer.concat([buffer, chunk])
    
    // 解析帧格式: [4字节长度][DOCX数据]
    while (buffer.length >= 4) {
      const chunkLength = buffer.readUInt32BE(0)
      
      if (chunkLength === 0) {
        console.log('\n🏁 收到结束标记')
        
        // 等待所有检查完成
        Promise.all(pendingChecks).then(() => {
          console.log(`\n✅ 流式传输完成`)
          console.log(`📊 统计:`)
          console.log(`  - 总 chunks: ${chunkCount}`)
          console.log(`  - 表格 chunks: ${tableChunkCount}`)
          console.log(`  - 段落 chunks: ${chunkCount - tableChunkCount}`)
          console.log('\n✅ 测试完成!')
          process.exit(0)
        })
        return
      }
      
      if (buffer.length < 4 + chunkLength) break
      
      chunkCount++
      const docxData = buffer.slice(4, 4 + chunkLength)
      const currentChunkNum = chunkCount
      
      // 异步检查是否包含表格
      const checkPromise = new Promise((resolve) => {
        const tmpFile = join(tmpdir(), `chunk${currentChunkNum}.docx`)
        writeFileSync(tmpFile, docxData)
        
        exec(`unzip -p "${tmpFile}" word/document.xml 2>/dev/null | grep -o "<w:tbl" | wc -l`, (err, stdout) => {
          const tableCount = parseInt(stdout.trim()) || 0
          
          if (tableCount > 0) {
            tableChunkCount++
            console.log(`  Chunk #${currentChunkNum}: 📊 表格 (${tableCount}个表格, ${docxData.length} bytes)`)
          } else {
            console.log(`  Chunk #${currentChunkNum}: 📄 段落 (${docxData.length} bytes)`)
          }
          
          try { unlinkSync(tmpFile) } catch {}
          resolve()
        })
      })
      
      pendingChecks.push(checkPromise)
      buffer = buffer.slice(4 + chunkLength)
    }
  })
  
  res.on('end', () => {
    if (pendingChecks.length === 0) {
      console.log(`\n✅ 流式传输完成`)
      console.log(`📊 总 chunks: ${chunkCount}`)
      console.log('\n✅ 测试完成!')
    }
  })
})

req.on('error', (err) => {
  console.error('❌ 请求失败:', err.message)
  console.error('💡 请确保服务已启动: npm start')
  process.exit(1)
})
