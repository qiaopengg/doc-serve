import { xmlParser, type XmlNode } from '../core/xml-parser.js'
import type { DocumentSettings } from '../types.js'

export class SettingsParser {
  /**
   * 解析文档设置（settings.xml）
   */
  parse(xml: string): Partial<DocumentSettings> {
    const doc = xmlParser.parse(xml)
    const settings = doc['w:settings']
    
    if (!settings) return {}

    const result: Partial<DocumentSettings> = {}

    // 视图和缩放
    if (settings['w:view']) {
      result.view = settings['w:view']['@_w:val'] as any
    }
    if (settings['w:zoom']) {
      result.zoom = parseInt(settings['w:zoom']['@_w:percent'] || '100')
    }

    // 修订追踪
    if (settings['w:trackRevisions']) {
      result.trackRevisions = true
    }
    if (settings['w:doNotTrackMoves']) {
      result.doNotTrackMoves = true
    }
    if (settings['w:doNotTrackFormatting']) {
      result.doNotTrackFormatting = true
    }

    // 默认制表位
    if (settings['w:defaultTabStop']) {
      result.defaultTabStop = parseInt(settings['w:defaultTabStop']['@_w:val'] || '720')
    }

    // 字符间距控制
    if (settings['w:characterSpacingControl']) {
      result.characterSpacingControl = settings['w:characterSpacingControl']['@_w:val'] as any
    }

    // 页眉页脚
    if (settings['w:evenAndOddHeaders']) {
      result.evenAndOddHeaders = true
    }

    // 书籍折叠打印
    if (settings['w:bookFoldPrinting']) {
      result.bookFoldPrinting = true
    }
    if (settings['w:bookFoldRevPrinting']) {
      result.bookFoldRevPrinting = true
    }
    if (settings['w:bookFoldPrintingSheets']) {
      result.bookFoldPrintingSheets = parseInt(
        settings['w:bookFoldPrintingSheets']['@_w:val'] || '1'
      )
    }

    // 网格设置
    if (settings['w:drawingGridHorizontalSpacing']) {
      result.drawingGridHorizontalSpacing = parseInt(
        settings['w:drawingGridHorizontalSpacing']['@_w:val'] || '0'
      )
    }
    if (settings['w:drawingGridVerticalSpacing']) {
      result.drawingGridVerticalSpacing = parseInt(
        settings['w:drawingGridVerticalSpacing']['@_w:val'] || '0'
      )
    }

    // 其他设置
    if (settings['w:doNotShadeFormData']) {
      result.doNotShadeFormData = true
    }
    if (settings['w:noPunctuationKerning']) {
      result.noPunctuationKerning = true
    }
    if (settings['w:printTwoOnOne']) {
      result.printTwoOnOne = true
    }
    if (settings['w:strictFirstAndLastChars']) {
      result.strictFirstAndLastChars = true
    }

    // 禁则字符
    if (settings['w:noLineBreaksAfter']) {
      result.noLineBreaksAfter = settings['w:noLineBreaksAfter']['@_w:val']
    }
    if (settings['w:noLineBreaksBefore']) {
      result.noLineBreaksBefore = settings['w:noLineBreaksBefore']['@_w:val']
    }

    // 预览图片
    if (settings['w:savePreviewPicture']) {
      result.savePreviewPicture = true
    }

    // XML 相关
    if (settings['w:doNotValidateAgainstSchema']) {
      result.doNotValidateAgainstSchema = true
    }
    if (settings['w:saveInvalidXml']) {
      result.saveInvalidXml = true
    }
    if (settings['w:ignoreMixedContent']) {
      result.ignoreMixedContent = true
    }
    if (settings['w:alwaysShowPlaceholderText']) {
      result.alwaysShowPlaceholderText = true
    }
    if (settings['w:doNotDemarcateInvalidXml']) {
      result.doNotDemarcateInvalidXml = true
    }
    if (settings['w:saveXmlDataOnly']) {
      result.saveXmlDataOnly = true
    }

    // 更新域
    if (settings['w:updateFields']) {
      result.updateFields = true
    }

    // 脚注和尾注属性
    if (settings['w:footnotePr']) {
      result.footnotePr = this.parseFootnoteProperties(settings['w:footnotePr'])
    }
    if (settings['w:endnotePr']) {
      result.endnotePr = this.parseEndnoteProperties(settings['w:endnotePr'])
    }

    // 兼容性设置
    if (settings['w:compat']) {
      result.compat = this.parseCompatibilitySettings(settings['w:compat'])
    }

    // 修订标识符
    if (settings['w:rsids']) {
      result.rsids = this.parseRevisionIdentifiers(settings['w:rsids'])
    }

    // 数学属性
    if (settings['m:mathPr']) {
      result.mathPr = this.parseMathProperties(settings['m:mathPr'])
    }

    // 附加模板
    if (settings['w:attachedTemplate']) {
      result.attachedTemplate = settings['w:attachedTemplate']['@_r:id']
    }

    // 链接样式
    if (settings['w:linkStyles']) {
      result.linkStyles = true
    }

    // 文档类型
    if (settings['w:documentType']) {
      result.documentType = settings['w:documentType']['@_w:val'] as any
    }

    // 主题字体语言
    if (settings['w:themeFontLang']) {
      result.themeFontLang = {
        val: settings['w:themeFontLang']['@_w:val'],
        eastAsia: settings['w:themeFontLang']['@_w:eastAsia'],
        bidi: settings['w:themeFontLang']['@_w:bidi']
      }
    }

    // 颜色方案映射
    if (settings['w:clrSchemeMapping']) {
      result.clrSchemeMapping = this.parseColorSchemeMapping(settings['w:clrSchemeMapping'])
    }

    // 其他布尔设置
    if (settings['w:doNotIncludeSubdocsInStats']) {
      result.doNotIncludeSubdocsInStats = true
    }
    if (settings['w:doNotAutoCompressPictures']) {
      result.doNotAutoCompressPictures = true
    }
    if (settings['w:forceUpgrade']) {
      result.forceUpgrade = true
    }
    if (settings['w:doNotEmbedSmartTags']) {
      result.doNotEmbedSmartTags = true
    }

    // 小数和列表分隔符
    if (settings['w:decimalSymbol']) {
      result.decimalSymbol = settings['w:decimalSymbol']['@_w:val']
    }
    if (settings['w:listSeparator']) {
      result.listSeparator = settings['w:listSeparator']['@_w:val']
    }

    return result
  }

  /**
   * 解析脚注属性
   */
  private parseFootnoteProperties(node: XmlNode): any {
    return {
      position: node['w:pos']?.['@_w:val'],
      numberingFormat: node['w:numFmt']?.['@_w:val'],
      numberingStart: node['w:numStart'] ? parseInt(node['w:numStart']['@_w:val']) : undefined,
      numberingRestart: node['w:numRestart']?.['@_w:val']
    }
  }

  /**
   * 解析尾注属性
   */
  private parseEndnoteProperties(node: XmlNode): any {
    return {
      position: node['w:pos']?.['@_w:val'],
      numberingFormat: node['w:numFmt']?.['@_w:val'],
      numberingStart: node['w:numStart'] ? parseInt(node['w:numStart']['@_w:val']) : undefined,
      numberingRestart: node['w:numRestart']?.['@_w:val']
    }
  }

  /**
   * 解析兼容性设置
   */
  private parseCompatibilitySettings(node: XmlNode): any {
    const compat: any = {}
    
    for (const key in node) {
      if (key.startsWith('w:')) {
        const settingName = key.substring(2)
        const value = node[key]['@_w:val']
        compat[settingName] = value === '1' || value === 'true' || value
      }
    }
    
    return compat
  }

  /**
   * 解析修订标识符
   */
  private parseRevisionIdentifiers(node: XmlNode): any {
    const rsids: string[] = []
    
    if (node['w:rsid']) {
      const rsidArray = Array.isArray(node['w:rsid']) ? node['w:rsid'] : [node['w:rsid']]
      for (const rsid of rsidArray) {
        if (rsid['@_w:val']) {
          rsids.push(rsid['@_w:val'])
        }
      }
    }
    
    return {
      rsidRoot: node['w:rsidRoot']?.['@_w:val'],
      rsids
    }
  }

  /**
   * 解析数学属性
   */
  private parseMathProperties(node: XmlNode): any {
    return {
      mathFont: node['m:mathFont']?.['@_m:val'],
      brkBin: node['m:brkBin']?.['@_m:val'],
      brkBinSub: node['m:brkBinSub']?.['@_m:val'],
      defJc: node['m:defJc']?.['@_m:val'],
      dispDef: node['m:dispDef']?.['@_m:val'] === '1',
      smallFrac: node['m:smallFrac']?.['@_m:val'] === '1',
      wrapRight: node['m:wrapRight']?.['@_m:val'] === '1'
    }
  }

  /**
   * 解析颜色方案映射
   */
  private parseColorSchemeMapping(node: XmlNode): any {
    return {
      bg1: node['@_w:bg1'],
      t1: node['@_w:t1'],
      bg2: node['@_w:bg2'],
      t2: node['@_w:t2'],
      accent1: node['@_w:accent1'],
      accent2: node['@_w:accent2'],
      accent3: node['@_w:accent3'],
      accent4: node['@_w:accent4'],
      accent5: node['@_w:accent5'],
      accent6: node['@_w:accent6'],
      hyperlink: node['@_w:hyperlink'],
      followedHyperlink: node['@_w:followedHyperlink']
    }
  }
}

export const settingsParser = new SettingsParser()
