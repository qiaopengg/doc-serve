import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { PictureRun } from '../types.js'

export class PictureParser {
  /**
   * 解析文档中的所有图片
   */
  parse(xml: string): Map<string, any> {
    const pictures = new Map<string, any>()
    const doc = xmlParser.parse(xml)
    
    this.extractPictures(doc, pictures)
    return pictures
  }

  /**
   * 解析图片运行
   */
  parsePictureRun(pictNode: XmlNode): PictureRun | null {
    if (!pictNode) return null

    // 从 v:shape 或 v:imagedata 中提取图片信息
    const imageData = pictNode['v:imagedata']
    if (!imageData) return null

    const relationshipId = imageData['@_r:id'] || imageData['@_o:relid']
    if (!relationshipId) return null

    const picture: PictureRun = {
      type: 'picture',
      relationshipId
    }

    // 提取尺寸（如果有）
    const shape = pictNode
    if (shape) {
      const style = shape['@_style']
      if (style) {
        const widthMatch = style.match(/width:([^;]+)/)
        const heightMatch = style.match(/height:([^;]+)/)
        
        if (widthMatch) {
          picture.width = this.parseLength(widthMatch[1])
        }
        if (heightMatch) {
          picture.height = this.parseLength(heightMatch[1])
        }
      }
    }

    // 提取描述
    const alt = imageData['@_o:title'] || shape?.['@_alt']
    if (alt) {
      picture.description = alt
    }

    return picture
  }

  /**
   * 从 pict 元素中提取图片
   */
  parsePict(pictNode: XmlNode): any | null {
    if (!pictNode) return null

    const shape = pictNode['v:shape']
    if (!shape) return null

    const imageData = shape['v:imagedata']
    if (!imageData) return null

    const relationshipId = imageData['@_r:id'] || imageData['@_o:relid']
    if (!relationshipId) return null

    return {
      type: 'vml-picture',
      relationshipId,
      shape: {
        id: shape['@_id'],
        style: shape['@_style'],
        alt: shape['@_alt']
      }
    }
  }

  /**
   * 递归提取所有图片
   */
  private extractPictures(node: any, pictures: Map<string, any>): void {
    if (!node || typeof node !== 'object') return

    // 处理 w:pict（旧版图片）
    if (node['w:pict']) {
      const pictArray = Array.isArray(node['w:pict']) 
        ? node['w:pict'] 
        : [node['w:pict']]
      
      for (const pict of pictArray) {
        const picture = this.parsePict(pict)
        if (picture && picture.relationshipId) {
          pictures.set(picture.relationshipId, picture)
        }
      }
    }

    // 处理 v:shape（VML 图形）
    if (node['v:shape']) {
      const shapeArray = Array.isArray(node['v:shape'])
        ? node['v:shape']
        : [node['v:shape']]
      
      for (const shape of shapeArray) {
        const picture = this.parsePictureRun(shape)
        if (picture && picture.relationshipId) {
          pictures.set(picture.relationshipId, picture)
        }
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractPictures(node[key], pictures)
      }
    }
  }

  /**
   * 解析长度值（pt, px, in 等）
   */
  private parseLength(value: string): number | undefined {
    if (!value) return undefined
    
    const match = value.match(/^([\d.]+)(pt|px|in|cm|mm)?$/)
    if (!match) return undefined
    
    const num = parseFloat(match[1])
    const unit = match[2] || 'pt'
    
    // 转换为 EMU (English Metric Unit)
    switch (unit) {
      case 'pt':
        return num * 12700 // 1 pt = 12700 EMU
      case 'px':
        return num * 9525 // 1 px ≈ 9525 EMU (at 96 DPI)
      case 'in':
        return num * 914400 // 1 inch = 914400 EMU
      case 'cm':
        return num * 360000 // 1 cm = 360000 EMU
      case 'mm':
        return num * 36000 // 1 mm = 36000 EMU
      default:
        return num
    }
  }
}

export const pictureParser = new PictureParser()
