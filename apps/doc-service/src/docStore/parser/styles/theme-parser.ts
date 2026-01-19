import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface Theme {
  name?: string
  colorScheme?: ColorScheme
  fontScheme?: FontScheme
  formatScheme?: any
}

export interface ColorScheme {
  name?: string
  colors?: {
    dk1?: string  // Dark 1
    lt1?: string  // Light 1
    dk2?: string  // Dark 2
    lt2?: string  // Light 2
    accent1?: string
    accent2?: string
    accent3?: string
    accent4?: string
    accent5?: string
    accent6?: string
    hlink?: string  // Hyperlink
    folHlink?: string  // Followed Hyperlink
  }
}

export interface FontScheme {
  name?: string
  majorFont?: FontCollection
  minorFont?: FontCollection
}

export interface FontCollection {
  latin?: string
  ea?: string  // East Asian
  cs?: string  // Complex Script
}

export class ThemeParser {
  /**
   * 解析主题文件
   */
  parse(xml: string): Theme | null {
    const doc = xmlParser.parse(xml)
    
    const theme = doc['a:theme']
    if (!theme) return null

    const result: Theme = {
      name: theme['@_name']
    }

    // 解析颜色方案
    const themeElements = theme['a:themeElements']
    if (themeElements) {
      if (themeElements['a:clrScheme']) {
        result.colorScheme = this.parseColorScheme(themeElements['a:clrScheme'])
      }

      if (themeElements['a:fontScheme']) {
        result.fontScheme = this.parseFontScheme(themeElements['a:fontScheme'])
      }

      if (themeElements['a:fmtScheme']) {
        result.formatScheme = themeElements['a:fmtScheme']
      }
    }

    return result
  }

  /**
   * 解析颜色方案
   */
  private parseColorScheme(clrScheme: XmlNode): ColorScheme {
    const scheme: ColorScheme = {
      name: clrScheme['@_name'],
      colors: {}
    }

    // 深色 1
    if (clrScheme['a:dk1']) {
      scheme.colors!.dk1 = this.extractColor(clrScheme['a:dk1'])
    }

    // 浅色 1
    if (clrScheme['a:lt1']) {
      scheme.colors!.lt1 = this.extractColor(clrScheme['a:lt1'])
    }

    // 深色 2
    if (clrScheme['a:dk2']) {
      scheme.colors!.dk2 = this.extractColor(clrScheme['a:dk2'])
    }

    // 浅色 2
    if (clrScheme['a:lt2']) {
      scheme.colors!.lt2 = this.extractColor(clrScheme['a:lt2'])
    }

    // 强调色
    for (let i = 1; i <= 6; i++) {
      const accentKey = `a:accent${i}`
      if (clrScheme[accentKey]) {
        const colorKey = `accent${i}` as 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
        if (scheme.colors) {
          scheme.colors[colorKey] = this.extractColor(clrScheme[accentKey])
        }
      }
    }

    // 超链接
    if (clrScheme['a:hlink']) {
      scheme.colors!.hlink = this.extractColor(clrScheme['a:hlink'])
    }

    // 已访问超链接
    if (clrScheme['a:folHlink']) {
      scheme.colors!.folHlink = this.extractColor(clrScheme['a:folHlink'])
    }

    return scheme
  }

  /**
   * 解析字体方案
   */
  private parseFontScheme(fontScheme: XmlNode): FontScheme {
    const scheme: FontScheme = {
      name: fontScheme['@_name']
    }

    // 主要字体
    if (fontScheme['a:majorFont']) {
      scheme.majorFont = this.parseFontCollection(fontScheme['a:majorFont'])
    }

    // 次要字体
    if (fontScheme['a:minorFont']) {
      scheme.minorFont = this.parseFontCollection(fontScheme['a:minorFont'])
    }

    return scheme
  }

  /**
   * 解析字体集合
   */
  private parseFontCollection(fontCollection: XmlNode): FontCollection {
    const collection: FontCollection = {}

    // 拉丁字体
    if (fontCollection['a:latin']) {
      collection.latin = fontCollection['a:latin']['@_typeface']
    }

    // 东亚字体
    if (fontCollection['a:ea']) {
      collection.ea = fontCollection['a:ea']['@_typeface']
    }

    // 复杂脚本字体
    if (fontCollection['a:cs']) {
      collection.cs = fontCollection['a:cs']['@_typeface']
    }

    return collection
  }

  /**
   * 提取颜色值
   */
  private extractColor(colorNode: XmlNode): string {
    // sRGB 颜色
    if (colorNode['a:srgbClr']) {
      return colorNode['a:srgbClr']['@_val'] || ''
    }

    // 系统颜色
    if (colorNode['a:sysClr']) {
      return colorNode['a:sysClr']['@_lastClr'] || colorNode['a:sysClr']['@_val'] || ''
    }

    // 方案颜色
    if (colorNode['a:schemeClr']) {
      return colorNode['a:schemeClr']['@_val'] || ''
    }

    return ''
  }
}

export const themeParser = new ThemeParser()
