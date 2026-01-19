/**
 * DOCX 完整解析器 - 主入口
 * 目标：100% 覆盖 OOXML 规范
 */

import { zipReader } from './core/zip-reader.js'
import { xmlParser } from './core/xml-parser.js'
import { RelationshipParser } from './core/relationship-parser.js'
import { ParagraphParser } from './document/paragraph-parser.js'
import { RunParser } from './document/run-parser.js'
import { TableParser } from './document/table-parser.js'
import { SectionParser } from './document/section-parser.js'
import { StyleParser } from './styles/style-parser.js'
import { NumberingParser } from './styles/numbering-parser.js'
import { ThemeParser } from './styles/theme-parser.js'
import { DrawingParser } from './content/drawing-parser.js'
import { PictureParser } from './content/picture-parser.js'
import { ObjectParser } from './content/object-parser.js'
import { MathParser } from './content/math-parser.js'
import { ChartParser } from './content/chart-parser.js'
import { SmartArtParser } from './content/smartart-parser.js'
import { FieldParser } from './interactive/field-parser.js'
import { SdtParser } from './interactive/sdt-parser.js'
import { HyperlinkParser } from './interactive/hyperlink-parser.js'
import { BookmarkParser } from './interactive/bookmark-parser.js'
import { CommentParser } from './collaboration/comment-parser.js'
import { RevisionParser } from './collaboration/revision-parser.js'
import { ProtectionParser } from './collaboration/protection-parser.js'
import { CorePropsParser } from './metadata/core-props-parser.js'
import { AppPropsParser } from './metadata/app-props-parser.js'
import { CustomPropsParser } from './metadata/custom-props-parser.js'
import { SettingsParser } from './metadata/settings-parser.js'

export interface CompleteDocxDocument {
  // 文档主体
  body: {
    paragraphs: any[]
    tables: any[]
    sections: any[]
    contentControls: any[]
  }
  
  // 样式系统
  styles: {
    definitions: any[]
    defaults: any
    numbering: any[]
    theme: any
  }
  
  // 内容元素
  content: {
    drawings: Map<string, any>
    pictures: Map<string, any>
    objects: Map<string, any>
    math: any[]
    charts: Map<string, any>
    smartArts: Map<string, any>
  }
  
  // 交互元素
  interactive: {
    fields: any[]
    contentControls: any[]
    hyperlinks: Map<string, string>
    bookmarks: Map<string, any>
  }
  
  // 协作元素
  collaboration: {
    comments: Map<string, any>
    revisions: any[]
    protection: any
  }
  
  // 元数据
  metadata: {
    core: any
    app: any
    custom: any
  }
  
  // 文档设置
  settings: any
  
  // 页眉页脚
  headers: Map<string, any>
  footers: Map<string, any>
  
  // 脚注尾注
  footnotes: Map<string, any>
  endnotes: Map<string, any>
  
  // 关系映射
  relationships: Map<string, any>
  
  // 原始文件
  rawFiles: Map<string, Buffer>
}

export class CompleteDocxParser {
  private relationshipParser: RelationshipParser
  private paragraphParser: ParagraphParser
  private runParser: RunParser
  private tableParser: TableParser
  private sectionParser: SectionParser
  private styleParser: StyleParser
  private numberingParser: NumberingParser
  private themeParser: ThemeParser
  private drawingParser: DrawingParser
  private pictureParser: PictureParser
  private objectParser: ObjectParser
  private mathParser: MathParser
  private chartParser: ChartParser
  private smartArtParser: SmartArtParser
  private fieldParser: FieldParser
  private sdtParser: SdtParser
  private hyperlinkParser: HyperlinkParser
  private bookmarkParser: BookmarkParser
  private commentParser: CommentParser
  private revisionParser: RevisionParser
  private protectionParser: ProtectionParser
  private corePropsParser: CorePropsParser
  private appPropsParser: AppPropsParser
  private customPropsParser: CustomPropsParser
  private settingsParser: SettingsParser

  constructor() {
    // 初始化所有解析器
    this.relationshipParser = new RelationshipParser()
    this.paragraphParser = new ParagraphParser()
    this.runParser = new RunParser()
    this.tableParser = new TableParser()
    this.sectionParser = new SectionParser()
    this.styleParser = new StyleParser()
    this.numberingParser = new NumberingParser()
    this.themeParser = new ThemeParser()
    this.drawingParser = new DrawingParser()
    this.pictureParser = new PictureParser()
    this.objectParser = new ObjectParser()
    this.mathParser = new MathParser()
    this.chartParser = new ChartParser()
    this.smartArtParser = new SmartArtParser()
    this.fieldParser = new FieldParser()
    this.sdtParser = new SdtParser()
    this.hyperlinkParser = new HyperlinkParser()
    this.bookmarkParser = new BookmarkParser()
    this.commentParser = new CommentParser()
    this.revisionParser = new RevisionParser()
    this.protectionParser = new ProtectionParser()
    this.corePropsParser = new CorePropsParser()
    this.appPropsParser = new AppPropsParser()
    this.customPropsParser = new CustomPropsParser()
    this.settingsParser = new SettingsParser()
  }

  /**
   * 解析完整的 DOCX 文档
   */
  async parse(buffer: Buffer): Promise<CompleteDocxDocument> {
    // 1. 读取所有文件
    const rawFiles = await zipReader.readAllEntries(buffer)
    
    // 2. 解析关系
    const relationships = await this.parseRelationships(rawFiles)
    
    // 3. 解析样式系统
    const styles = await this.parseStyles(rawFiles, relationships)
    
    // 4. 解析元数据
    const metadata = await this.parseMetadata(rawFiles)
    
    // 5. 解析设置
    const settings = await this.parseSettings(rawFiles)
    
    // 6. 解析协作元素
    const collaboration = await this.parseCollaboration(rawFiles, relationships)
    
    // 7. 解析文档主体
    const body = await this.parseBody(rawFiles, relationships, styles, collaboration)
    
    // 8. 解析内容元素
    const content = await this.parseContent(rawFiles, relationships)
    
    // 9. 解析交互元素
    const interactive = await this.parseInteractive(rawFiles, relationships)
    
    // 10. 解析页眉页脚
    const { headers, footers } = await this.parseHeadersFooters(rawFiles, relationships)
    
    // 11. 解析脚注尾注
    const { footnotes, endnotes } = await this.parseNotes(rawFiles, relationships)
    
    return {
      body,
      styles,
      content,
      interactive,
      collaboration,
      metadata,
      settings,
      headers,
      footers,
      footnotes,
      endnotes,
      relationships,
      rawFiles
    }
  }

  private async parseRelationships(files: Map<string, Buffer>): Promise<Map<string, any>> {
    const relationships = new Map<string, any>()
    
    for (const [path, content] of files.entries()) {
      if (path.endsWith('.rels')) {
        const xml = content.toString('utf-8')
        const rels = this.relationshipParser.parse(xml)
        relationships.set(path, rels)
      }
    }
    
    return relationships
  }

  private async parseStyles(files: Map<string, Buffer>, relationships: Map<string, any>): Promise<any> {
    const stylesXml = files.get('word/styles.xml')
    const numberingXml = files.get('word/numbering.xml')
    const themeXml = files.get('word/theme/theme1.xml')
    
    return {
      definitions: stylesXml ? this.styleParser.parse(stylesXml.toString('utf-8')) : [],
      defaults: stylesXml ? this.styleParser.parseDefaults(stylesXml.toString('utf-8')) : {},
      numbering: numberingXml ? this.numberingParser.parse(numberingXml.toString('utf-8')) : [],
      theme: themeXml ? this.themeParser.parse(themeXml.toString('utf-8')) : null
    }
  }

  private async parseMetadata(files: Map<string, Buffer>): Promise<any> {
    const coreXml = files.get('docProps/core.xml')
    const appXml = files.get('docProps/app.xml')
    const customXml = files.get('docProps/custom.xml')
    
    return {
      core: coreXml ? this.corePropsParser.parse(coreXml.toString('utf-8')) : null,
      app: appXml ? this.appPropsParser.parse(appXml.toString('utf-8')) : null,
      custom: customXml ? this.customPropsParser.parse(customXml.toString('utf-8')) : null
    }
  }

  private async parseSettings(files: Map<string, Buffer>): Promise<any> {
    const settingsXml = files.get('word/settings.xml')
    return settingsXml ? this.settingsParser.parse(settingsXml.toString('utf-8')) : null
  }

  private async parseCollaboration(files: Map<string, Buffer>, relationships: Map<string, any>): Promise<any> {
    const commentsXml = files.get('word/comments.xml')
    const documentXml = files.get('word/document.xml')
    
    return {
      comments: commentsXml ? this.commentParser.parse(commentsXml.toString('utf-8')) : new Map(),
      revisions: documentXml ? this.revisionParser.parse(documentXml.toString('utf-8')) : [],
      protection: documentXml ? this.protectionParser.parse(documentXml.toString('utf-8')) : null
    }
  }

  private async parseBody(
    files: Map<string, Buffer>,
    relationships: Map<string, any>,
    styles: any,
    collaboration: any
  ): Promise<any> {
    const documentXml = files.get('word/document.xml')
    if (!documentXml) {
      return { paragraphs: [], tables: [], sections: [], contentControls: [] }
    }
    
    const xml = documentXml.toString('utf-8')
    const doc = xmlParser.parse(xml)
    const body = doc['w:document']?.['w:body']
    
    if (!body) {
      return { paragraphs: [], tables: [], sections: [], contentControls: [] }
    }
    
    const paragraphs: any[] = []
    const tables: any[] = []
    const sections: any[] = []
    const contentControls: any[] = []
    
    // 按顺序处理所有子元素，保持文档结构
    const bodyChildren = this.getBodyChildren(body)
    
    for (const child of bodyChildren) {
      const [tagName, node] = child
      
      switch (tagName) {
        case 'w:p':
          // 解析段落
          const paragraph = this.paragraphParser.parseParagraph(node)
          if (paragraph) {
            paragraphs.push(paragraph)
          }
          
          // 检查段落末尾的分节属性
          if (node['w:pPr']?.['w:sectPr']) {
            const section = this.sectionParser.parseSectionProperties(node['w:pPr']['w:sectPr'])
            if (section) {
              sections.push(section)
            }
          }
          break
          
        case 'w:tbl':
          // 解析表格
          const table = this.tableParser.parseTable(node)
          if (table) {
            tables.push(table)
          }
          break
          
        case 'w:sdt':
          // 解析内容控件
          const sdt = this.sdtParser.parseContentControl(node)
          if (sdt) {
            contentControls.push(sdt)
          }
          break
          
        case 'w:bookmarkStart':
        case 'w:bookmarkEnd':
        case 'w:commentRangeStart':
        case 'w:commentRangeEnd':
        case 'w:moveFromRangeStart':
        case 'w:moveFromRangeEnd':
        case 'w:moveToRangeStart':
        case 'w:moveToRangeEnd':
        case 'w:permStart':
        case 'w:permEnd':
        case 'w:proofErr':
        case 'w:customXml':
        case 'w:smartTag':
          // 这些元素已在其他解析器中处理
          break
      }
    }
    
    // 解析文档末尾的分节属性
    if (body['w:sectPr']) {
      const section = this.sectionParser.parseSectionProperties(body['w:sectPr'])
      if (section) {
        sections.push(section)
      }
    }
    
    return {
      paragraphs,
      tables,
      sections,
      contentControls
    }
  }
  
  /**
   * 获取 body 的所有子元素（按顺序）
   */
  private getBodyChildren(body: any): Array<[string, any]> {
    const children: Array<[string, any]> = []
    
    // Word 文档主体可能包含的元素类型
    const elementTypes = [
      'w:p',           // 段落
      'w:tbl',         // 表格
      'w:sdt',         // 内容控件
      'w:bookmarkStart',
      'w:bookmarkEnd',
      'w:commentRangeStart',
      'w:commentRangeEnd',
      'w:moveFromRangeStart',
      'w:moveFromRangeEnd',
      'w:moveToRangeStart',
      'w:moveToRangeEnd',
      'w:permStart',
      'w:permEnd',
      'w:proofErr',
      'w:customXml',
      'w:smartTag'
    ]
    
    for (const type of elementTypes) {
      if (body[type]) {
        const nodes = Array.isArray(body[type]) ? body[type] : [body[type]]
        for (const node of nodes) {
          children.push([type, node])
        }
      }
    }
    
    // 注意：这里简化了顺序处理，实际应该按照 XML 中的原始顺序
    // 但由于 XML 解析器已经将元素分组，我们只能按类型处理
    
    return children
  }

  private async parseContent(files: Map<string, Buffer>, relationships: Map<string, any>): Promise<any> {
    const documentXml = files.get('word/document.xml')
    
    const drawings = new Map()
    const pictures = new Map()
    const objects = new Map()
    const math: any[] = []
    const charts = new Map()
    const smartArts = new Map()
    
    if (documentXml) {
      const xml = documentXml.toString('utf-8')
      
      // 解析图形
      const parsedDrawings = this.drawingParser.parse(xml)
      for (const [id, drawing] of parsedDrawings) {
        drawings.set(id, drawing)
      }
      
      // 解析图片
      const parsedPictures = this.pictureParser.parse(xml)
      for (const [id, picture] of parsedPictures) {
        pictures.set(id, picture)
      }
      
      // 解析对象
      const parsedObjects = this.objectParser.parse(xml)
      for (const [id, obj] of parsedObjects) {
        objects.set(id, obj)
      }
      
      // 解析数学公式
      const parsedMath = this.mathParser.parse(xml)
      math.push(...parsedMath)
      
      // 解析图表
      const parsedCharts = this.chartParser.parse(xml)
      for (const [id, chart] of parsedCharts) {
        charts.set(id, chart)
      }
      
      // 解析 SmartArt
      const parsedSmartArts = this.smartArtParser.parse(xml)
      for (const [id, smartArt] of parsedSmartArts) {
        smartArts.set(id, smartArt)
      }
    }
    
    // 解析独立的图表文件
    for (const [path, content] of files.entries()) {
      if (path.startsWith('word/charts/') && path.endsWith('.xml')) {
        const chartXml = content.toString('utf-8')
        const chart = this.chartParser.parseChart({}, chartXml)
        if (chart) {
          charts.set(path, chart)
        }
      }
    }
    
    return {
      drawings,
      pictures,
      objects,
      math,
      charts,
      smartArts
    }
  }

  private async parseInteractive(files: Map<string, Buffer>, relationships: Map<string, any>): Promise<any> {
    const documentXml = files.get('word/document.xml')
    
    let fields: any[] = []
    let contentControls: any[] = []
    let hyperlinks = new Map()
    let bookmarks = new Map()
    
    if (documentXml) {
      const xml = documentXml.toString('utf-8')
      
      // 解析域代码
      fields = this.fieldParser.parse(xml)
      
      // 解析内容控件
      contentControls = this.sdtParser.parse(xml)
      
      // 解析超链接
      hyperlinks = this.hyperlinkParser.parse(xml)
      
      // 解析书签
      bookmarks = this.bookmarkParser.parse(xml)
    }
    
    return {
      fields,
      contentControls,
      hyperlinks,
      bookmarks
    }
  }

  private async parseHeadersFooters(files: Map<string, Buffer>, relationships: Map<string, any>): Promise<any> {
    const headers = new Map()
    const footers = new Map()
    
    for (const [path, content] of files.entries()) {
      if (path.startsWith('word/header') && path.endsWith('.xml')) {
        const xml = content.toString('utf-8')
        const doc = xmlParser.parse(xml)
        const hdr = doc['w:hdr']
        
        if (hdr) {
          const paragraphs: any[] = []
          if (hdr['w:p']) {
            const pArray = Array.isArray(hdr['w:p']) ? hdr['w:p'] : [hdr['w:p']]
            for (const p of pArray) {
              const paragraph = this.paragraphParser.parseParagraph(p)
              if (paragraph) paragraphs.push(paragraph)
            }
          }
          headers.set(path, { paragraphs, raw: xml })
        }
      } else if (path.startsWith('word/footer') && path.endsWith('.xml')) {
        const xml = content.toString('utf-8')
        const doc = xmlParser.parse(xml)
        const ftr = doc['w:ftr']
        
        if (ftr) {
          const paragraphs: any[] = []
          if (ftr['w:p']) {
            const pArray = Array.isArray(ftr['w:p']) ? ftr['w:p'] : [ftr['w:p']]
            for (const p of pArray) {
              const paragraph = this.paragraphParser.parseParagraph(p)
              if (paragraph) paragraphs.push(paragraph)
            }
          }
          footers.set(path, { paragraphs, raw: xml })
        }
      }
    }
    
    return { headers, footers }
  }

  private async parseNotes(files: Map<string, Buffer>, relationships: Map<string, any>): Promise<any> {
    const footnotesXml = files.get('word/footnotes.xml')
    const endnotesXml = files.get('word/endnotes.xml')
    
    const footnotes = new Map()
    const endnotes = new Map()
    
    // 解析脚注
    if (footnotesXml) {
      const xml = footnotesXml.toString('utf-8')
      const doc = xmlParser.parse(xml)
      const footnotesRoot = doc['w:footnotes']
      
      if (footnotesRoot && footnotesRoot['w:footnote']) {
        const footnoteArray = Array.isArray(footnotesRoot['w:footnote'])
          ? footnotesRoot['w:footnote']
          : [footnotesRoot['w:footnote']]
        
        for (const fn of footnoteArray) {
          const id = fn['@_w:id']
          if (id) {
            const paragraphs: any[] = []
            if (fn['w:p']) {
              const pArray = Array.isArray(fn['w:p']) ? fn['w:p'] : [fn['w:p']]
              for (const p of pArray) {
                const paragraph = this.paragraphParser.parseParagraph(p)
                if (paragraph) paragraphs.push(paragraph)
              }
            }
            footnotes.set(id, { id, type: fn['@_w:type'], paragraphs })
          }
        }
      }
    }
    
    // 解析尾注
    if (endnotesXml) {
      const xml = endnotesXml.toString('utf-8')
      const doc = xmlParser.parse(xml)
      const endnotesRoot = doc['w:endnotes']
      
      if (endnotesRoot && endnotesRoot['w:endnote']) {
        const endnoteArray = Array.isArray(endnotesRoot['w:endnote'])
          ? endnotesRoot['w:endnote']
          : [endnotesRoot['w:endnote']]
        
        for (const en of endnoteArray) {
          const id = en['@_w:id']
          if (id) {
            const paragraphs: any[] = []
            if (en['w:p']) {
              const pArray = Array.isArray(en['w:p']) ? en['w:p'] : [en['w:p']]
              for (const p of pArray) {
                const paragraph = this.paragraphParser.parseParagraph(p)
                if (paragraph) paragraphs.push(paragraph)
              }
            }
            endnotes.set(id, { id, type: en['@_w:type'], paragraphs })
          }
        }
      }
    }
    
    return {
      footnotes,
      endnotes
    }
  }
}

// 导出单例
export const completeDocxParser = new CompleteDocxParser()

// 便捷函数
export async function parseCompleteDocx(buffer: Buffer): Promise<CompleteDocxDocument> {
  return completeDocxParser.parse(buffer)
}
