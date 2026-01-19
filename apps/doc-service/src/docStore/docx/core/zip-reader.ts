/**
 * ZIP 文件读取器
 * 用于读取 DOCX 文件中的所有 XML 文件
 */

import { createRequire } from 'node:module'
import type { Entry, ZipFile } from 'yauzl'
import type { Readable } from 'node:stream'

export interface ZipEntry {
  name: string
  content: Buffer
}

export class ZipReader {
  // 安全配置常量
  private readonly MAX_ENTRY_SIZE = 100 * 1024 * 1024 // 100MB
  private readonly MAX_TOTAL_SIZE = 500 * 1024 * 1024 // 500MB
  private readonly MAX_COMPRESSION_RATIO = 100 // 最大压缩比
  
  /**
   * 验证路径安全性（防止路径遍历攻击）
   */
  private isUnsafePath(path: string): boolean {
    // 检查路径遍历
    if (path.includes('..') || path.startsWith('/') || path.includes('\\')) {
      return true
    }
    
    // 检查绝对路径（Windows）
    if (path.match(/^[a-zA-Z]:/)) {
      return true
    }
    
    // 检查特殊字符
    if (path.includes('\0') || path.includes('\r') || path.includes('\n')) {
      return true
    }
    
    return false
  }
  
  /**
   * 读取 ZIP 文件中的单个条目
   */
  async readEntry(buffer: Buffer, entryName: string): Promise<Buffer | undefined> {
    // 安全检查：路径遍历
    if (this.isUnsafePath(entryName)) {
      throw new Error(`Unsafe path detected: ${entryName}`)
    }
    
    const require = createRequire(import.meta.url)
    const yauzl = require('yauzl') as typeof import('yauzl')

    return new Promise((resolve, reject) => {
      yauzl.fromBuffer(buffer, { lazyEntries: true }, (err, zipfile) => {
        if (err || !zipfile) {
          reject(err ?? new Error('Failed to open ZIP'))
          return
        }

        let found = false

        zipfile.readEntry()
        zipfile.on('entry', (entry: Entry) => {
          if (entry.fileName !== entryName) {
            zipfile.readEntry()
            return
          }

          found = true
          zipfile.openReadStream(entry, (err2, stream) => {
            if (err2 || !stream) {
              reject(err2 ?? new Error('Failed to open stream'))
              return
            }

            const chunks: Buffer[] = []
            let totalSize = 0
            
            stream.on('data', (chunk) => {
              totalSize += chunk.length
              
              // 安全检查：单个文件大小
              if (totalSize > this.MAX_ENTRY_SIZE) {
                stream.destroy()
                zipfile.close()
                reject(new Error(`Entry size exceeds limit: ${totalSize} bytes (max: ${this.MAX_ENTRY_SIZE})`))
                return
              }
              
              // 安全检查：压缩比（防止 ZIP 炸弹）
              if (entry.compressedSize > 0) {
                const compressionRatio = totalSize / entry.compressedSize
                if (compressionRatio > this.MAX_COMPRESSION_RATIO) {
                  stream.destroy()
                  zipfile.close()
                  reject(new Error(`Compression ratio too high: ${compressionRatio.toFixed(2)} (max: ${this.MAX_COMPRESSION_RATIO}) - possible ZIP bomb`))
                  return
                }
              }
              
              chunks.push(Buffer.from(chunk))
            })
            
            stream.on('end', () => {
              zipfile.close()
              resolve(Buffer.concat(chunks))
            })
            stream.on('error', (err) => {
              stream.destroy()
              zipfile.close()
              reject(err)
            })
          })
        })

        zipfile.on('end', () => {
          if (!found) {
            zipfile.close()
            resolve(undefined)
          }
        })

        zipfile.on('error', reject)
      })
    })
  }

  /**
   * 列出 ZIP 文件中的所有条目
   */
  async listEntries(buffer: Buffer): Promise<string[]> {
    const require = createRequire(import.meta.url)
    const yauzl = require('yauzl') as typeof import('yauzl')

    return new Promise((resolve, reject) => {
      yauzl.fromBuffer(buffer, { lazyEntries: true }, (err, zipfile) => {
        if (err || !zipfile) {
          reject(err ?? new Error('Failed to open ZIP'))
          return
        }

        const entries: string[] = []

        zipfile.readEntry()
        zipfile.on('entry', (entry: Entry) => {
          entries.push(entry.fileName)
          zipfile.readEntry()
        })

        zipfile.on('end', () => {
          zipfile.close()
          resolve(entries)
        })

        zipfile.on('error', reject)
      })
    })
  }

  /**
   * 读取所有条目
   */
  async readAllEntries(buffer: Buffer): Promise<Map<string, Buffer>> {
    const require = createRequire(import.meta.url)
    const yauzl = require('yauzl') as typeof import('yauzl')

    return new Promise((resolve, reject) => {
      yauzl.fromBuffer(buffer, { lazyEntries: true }, (err, zipfile) => {
        if (err || !zipfile) {
          reject(err ?? new Error('Failed to open ZIP'))
          return
        }

        const entries = new Map<string, Buffer>()
        const pending: Promise<void>[] = []
        let totalUncompressedSize = 0

        zipfile.readEntry()
        zipfile.on('entry', (entry: Entry) => {
          if (entry.fileName.endsWith('/')) {
            zipfile.readEntry()
            return
          }
          
          // 安全检查：路径遍历
          if (this.isUnsafePath(entry.fileName)) {
            reject(new Error(`Unsafe path detected: ${entry.fileName}`))
            return
          }

          const promise = new Promise<void>((resolveEntry, rejectEntry) => {
            zipfile.openReadStream(entry, (err2, stream) => {
              if (err2 || !stream) {
                rejectEntry(err2 ?? new Error('Failed to open stream'))
                return
              }

              const chunks: Buffer[] = []
              let entrySize = 0
              
              stream.on('data', (chunk) => {
                entrySize += chunk.length
                totalUncompressedSize += chunk.length
                
                // 安全检查：单个文件大小
                if (entrySize > this.MAX_ENTRY_SIZE) {
                  stream.destroy()
                  rejectEntry(new Error(`Entry ${entry.fileName} size exceeds limit: ${entrySize} bytes`))
                  return
                }
                
                // 安全检查：总大小
                if (totalUncompressedSize > this.MAX_TOTAL_SIZE) {
                  stream.destroy()
                  rejectEntry(new Error(`Total uncompressed size exceeds limit: ${totalUncompressedSize} bytes`))
                  return
                }
                
                // 安全检查：压缩比
                if (entry.compressedSize > 0) {
                  const compressionRatio = entrySize / entry.compressedSize
                  if (compressionRatio > this.MAX_COMPRESSION_RATIO) {
                    stream.destroy()
                    rejectEntry(new Error(`Entry ${entry.fileName} compression ratio too high: ${compressionRatio.toFixed(2)}`))
                    return
                  }
                }
                
                chunks.push(Buffer.from(chunk))
              })
              
              stream.on('end', () => {
                entries.set(entry.fileName, Buffer.concat(chunks))
                resolveEntry()
              })
              stream.on('error', (err) => {
                stream.destroy()
                rejectEntry(err)
              })
            })
          })

          pending.push(promise)
          zipfile.readEntry()
        })

        zipfile.on('end', async () => {
          try {
            await Promise.all(pending)
            zipfile.close()
            resolve(entries)
          } catch (e) {
            zipfile.close()
            reject(e)
          }
        })

        zipfile.on('error', reject)
      })
    })
  }
}

export const zipReader = new ZipReader()

// 便捷函数（兼容旧 API）
export async function readZipEntry(zipBuffer: Buffer, entryName: string): Promise<Buffer | undefined> {
  return zipReader.readEntry(zipBuffer, entryName)
}

export async function listZipEntries(zipBuffer: Buffer): Promise<string[]> {
  return zipReader.listEntries(zipBuffer)
}

export async function replaceZipEntry(zipBuffer: Buffer, entryName: string, replacement: Buffer): Promise<Buffer> {
  const require = createRequire(import.meta.url)
  const yazl = require('yazl') as typeof import('yazl')
  
  const zipOut = new yazl.ZipFile()
  const allEntries = await zipReader.readAllEntries(zipBuffer)
  
  for (const [name, content] of allEntries) {
    const data = name === entryName ? replacement : content
    zipOut.addBuffer(data, name)
  }
  
  zipOut.end()
  
  const chunks: Buffer[] = []
  for await (const chunk of zipOut.outputStream as any) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
  }
  return Buffer.concat(chunks)
}
