import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { DrawingRun, DrawingAnchor, EMUValue } from '../types.js'

export class DrawingParser {
  /**
   * 解析文档中的所有图形元素
   */
  parse(xml: string): Map<string, any> {
    const drawings = new Map<string, any>()
    const doc = xmlParser.parse(xml)
    
    this.extractDrawings(doc, drawings)
    return drawings
  }

  /**
   * 从段落运行中提取图形
   */
  parseDrawingRun(drawingNode: XmlNode): DrawingRun | null {
    if (!drawingNode) return null

    const inline = drawingNode['wp:inline']
    const anchor = drawingNode['wp:anchor']

    if (inline) {
      return this.parseInlineDrawing(inline)
    } else if (anchor) {
      return this.parseAnchorDrawing(anchor)
    }

    return null
  }

  /**
   * 解析内联图形（嵌入文本流）
   */
  private parseInlineDrawing(inlineNode: XmlNode): DrawingRun {
    const extent = inlineNode['wp:extent']
    const docPr = inlineNode['wp:docPr']
    const graphic = inlineNode['a:graphic']

    const drawing: DrawingRun = {
      type: 'drawing',
      drawingId: docPr?.['@_id'] || '',
      inline: true
    }

    // 提取图形数据
    if (graphic) {
      const graphicData = graphic['a:graphicData']
      if (graphicData) {
        const pic = graphicData['pic:pic']
        if (pic) {
          Object.assign(drawing, this.parsePicture(pic))
        }
      }
    }

    return drawing
  }

  /**
   * 解析浮动图形（锚定定位）
   */
  private parseAnchorDrawing(anchorNode: XmlNode): DrawingRun {
    const docPr = anchorNode['wp:docPr']
    const graphic = anchorNode['a:graphic']

    const anchor: DrawingAnchor = {
      distT: this.parseEMU(anchorNode['@_distT']),
      distB: this.parseEMU(anchorNode['@_distB']),
      distL: this.parseEMU(anchorNode['@_distL']),
      distR: this.parseEMU(anchorNode['@_distR']),
      simplePos: anchorNode['@_simplePos'] === '1',
      relativeHeight: parseInt(anchorNode['@_relativeHeight'] || '0'),
      behindDoc: anchorNode['@_behindDoc'] === '1',
      locked: anchorNode['@_locked'] === '1',
      layoutInCell: anchorNode['@_layoutInCell'] === '1',
      allowOverlap: anchorNode['@_allowOverlap'] === '1'
    }

    const drawing: DrawingRun = {
      type: 'drawing',
      drawingId: docPr?.['@_id'] || '',
      inline: false,
      anchor
    }

    // 提取图形数据
    if (graphic) {
      const graphicData = graphic['a:graphicData']
      if (graphicData) {
        const pic = graphicData['pic:pic']
        if (pic) {
          Object.assign(drawing, this.parsePicture(pic))
        }
      }
    }

    return drawing
  }

  /**
   * 解析图片元素
   */
  private parsePicture(picNode: XmlNode): any {
    const nvPicPr = picNode['pic:nvPicPr']
    const blipFill = picNode['pic:blipFill']
    const spPr = picNode['pic:spPr']

    const picture: any = {}

    // 非视觉属性（名称、描述等）
    if (nvPicPr) {
      const cNvPr = nvPicPr['pic:cNvPr']
      if (cNvPr) {
        picture.name = cNvPr['@_name']
        picture.description = cNvPr['@_descr']
        picture.title = cNvPr['@_title']
      }
    }

    // 图片填充（关系ID）
    if (blipFill) {
      const blip = blipFill['a:blip']
      if (blip) {
        picture.relationshipId = blip['@_r:embed'] || blip['@_r:link']
      }
    }

    // 形状属性（大小、变换等）
    if (spPr) {
      const xfrm = spPr['a:xfrm']
      if (xfrm) {
        const ext = xfrm['a:ext']
        if (ext) {
          picture.width = this.parseEMU(ext['@_cx'])
          picture.height = this.parseEMU(ext['@_cy'])
        }

        const off = xfrm['a:off']
        if (off) {
          picture.x = this.parseEMU(off['@_x'])
          picture.y = this.parseEMU(off['@_y'])
        }

        if (xfrm['@_rot']) {
          picture.rotation = parseInt(xfrm['@_rot']) / 60000 // 转换为度
        }
      }

      // 效果
      const effectLst = spPr['a:effectLst']
      if (effectLst) {
        picture.effects = this.parseEffects(effectLst)
      }
    }

    return picture
  }

  /**
   * 解析图形效果
   */
  private parseEffects(effectLst: XmlNode): any {
    const effects: any = {}

    if (effectLst['a:outerShdw']) {
      const shadow = effectLst['a:outerShdw']
      effects.shadow = {
        blurRadius: this.parseEMU(shadow['@_blurRad']),
        distance: this.parseEMU(shadow['@_dist']),
        direction: parseInt(shadow['@_dir'] || '0') / 60000,
        alignment: shadow['@_algn']
      }
    }

    if (effectLst['a:glow']) {
      const glow = effectLst['a:glow']
      effects.glow = {
        radius: this.parseEMU(glow['@_rad'])
      }
    }

    if (effectLst['a:reflection']) {
      effects.reflection = true
    }

    return effects
  }

  /**
   * 递归提取所有图形
   */
  private extractDrawings(node: any, drawings: Map<string, any>): void {
    if (!node || typeof node !== 'object') return

    if (node['w:drawing']) {
      const drawingArray = Array.isArray(node['w:drawing']) 
        ? node['w:drawing'] 
        : [node['w:drawing']]

      for (const drawing of drawingArray) {
        const parsed = this.parseDrawingRun(drawing)
        if (parsed) {
          drawings.set(parsed.drawingId, parsed)
        }
      }
    }

    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractDrawings(node[key], drawings)
      }
    }
  }

  /**
   * 解析 EMU 值（English Metric Unit）
   */
  private parseEMU(value: string | undefined): EMUValue | undefined {
    if (!value) return undefined
    const num = parseInt(value)
    return isNaN(num) ? undefined : num
  }
}

export const drawingParser = new DrawingParser()
