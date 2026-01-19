/**
 * DOCX 解析器 - 统一入口
 */

// ============================================================================
// 核心解析功能
// ============================================================================

export { parseDocxDocument, type DocxDocument, type DocxParseOptions } from './parse.js'
export { streamDocxSlices, flattenParagraphsForStreaming } from './stream.js'
export { getDocumentStatistics, type DocxStatistics } from './statistics.js'

// ============================================================================
// 工具函数
// ============================================================================

export { zipReader, readZipEntry, listZipEntries, replaceZipEntry } from './core/zip-reader.js'
export { buildXml } from './core/xml-parser.js'

// ============================================================================
// 类型定义
// ============================================================================

export type {
  // 元数据类型
  CoreProperties,
  AppProperties,
  CustomProperties,
  DocumentMetadata,
  HeaderFooterContent,
  Comment,
  Note
} from './metadata/metadata.js'

export type {
  // 段落类型
  DocxParagraph,
  RunStyle,
  ParagraphSpacing,
  SectionPropertiesSpec,
  
  // 表格类型
  CellStyle,
  BorderSpec,
  TableBordersSpec,
  CellBordersSpec,
  
  // 扩展类型
  ImageSpec,
  NumberingSpec,
  BookmarkSpec,
  FieldSpec,
  
  // 流式类型
  FlattenedElement,
  
  // 内部类型
  StyleMap,
  OrderedXmlNode
} from './types.js'

// ============================================================================
// 工具函数导出
// ============================================================================

export {
  asArray,
  parseBooleanOnOff,
  normalizeColor,
  normalizeBorderColor,
  normalizeAlignment,
  tagNameOf,
  attrsOf,
  attrOf,
  childrenOf,
  childOf,
  childrenNamed,
  textFromOrdered,
  mergeDefined,
  detectHeadingLevel
} from './core/utils.js'
