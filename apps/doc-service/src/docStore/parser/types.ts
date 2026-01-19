/**
 * 完整的 DOCX 解析类型定义
 * 覆盖 OOXML 规范的所有主要元素
 */

// ============= 基础类型 =============

export type TwipValue = number // 1/20 point
export type PointValue = number
export type EMUValue = number // English Metric Unit (1/914400 inch)
export type PercentValue = number
export type HexColor = string // 6位十六进制颜色

// ============= 文本运行（Run）=============

export interface TextRun {
  text: string
  properties?: RunProperties
}

export interface RunProperties {
  // 基础样式
  bold?: boolean
  italic?: boolean
  underline?: UnderlineType
  strike?: boolean
  doubleStrike?: boolean
  
  // 字体
  fonts?: FontInfo
  fontSize?: number // 半点
  fontSizeCs?: number // 复杂脚本字号
  
  // 颜色和效果
  color?: HexColor
  highlight?: HighlightColor
  shading?: Shading
  border?: Border
  
  // 位置和间距
  position?: number // 升高/降低
  spacing?: number // 字符间距
  scale?: number // 字符缩放百分比
  kern?: number // 字距调整
  
  // 特殊格式
  verticalAlign?: 'baseline' | 'superscript' | 'subscript'
  smallCaps?: boolean
  allCaps?: boolean
  hidden?: boolean
  webHidden?: boolean
  
  // 效果
  emboss?: boolean
  imprint?: boolean
  outline?: boolean
  shadow?: boolean
  vanish?: boolean
  specVanish?: boolean
  
  // 东亚排版
  eastAsianLayout?: EastAsianLayout
  emphasis?: EmphasisMark
  
  // 语言
  lang?: LanguageInfo
  
  // 样式引用
  styleId?: string
  
  // 其他
  fitText?: { width: TwipValue; id: number }
  effect?: TextEffect
  rtl?: boolean
}

export type UnderlineType = 
  | 'single' | 'double' | 'thick' | 'dotted' | 'dottedHeavy'
  | 'dash' | 'dashedHeavy' | 'dashLong' | 'dashLongHeavy'
  | 'dotDash' | 'dashDotHeavy' | 'dotDotDash' | 'dashDotDotHeavy'
  | 'wave' | 'wavyHeavy' | 'wavyDouble' | 'words' | 'none'

export type HighlightColor = 
  | 'black' | 'blue' | 'cyan' | 'green' | 'magenta' | 'red' | 'yellow' | 'white'
  | 'darkBlue' | 'darkCyan' | 'darkGreen' | 'darkMagenta' | 'darkRed' | 'darkYellow' | 'darkGray' | 'lightGray'

export type TextEffect = 
  | 'blinkBackground' | 'lights' | 'antsBlack' | 'antsRed' | 'shimmer' | 'sparkle' | 'none'

export interface FontInfo {
  ascii?: string
  hAnsi?: string
  eastAsia?: string
  cs?: string // Complex Script
  hint?: 'default' | 'eastAsia' | 'cs'
}

export interface Shading {
  fill?: HexColor
  color?: HexColor
  pattern?: ShadingPattern
}

export type ShadingPattern = 
  | 'clear' | 'solid' | 'horzStripe' | 'vertStripe' | 'reverseDiagStripe'
  | 'diagStripe' | 'horzCross' | 'diagCross' | 'thinHorzStripe' | 'thinVertStripe'
  | 'thinReverseDiagStripe' | 'thinDiagStripe' | 'thinHorzCross' | 'thinDiagCross'
  | 'pct5' | 'pct10' | 'pct12' | 'pct15' | 'pct20' | 'pct25' | 'pct30' | 'pct35'
  | 'pct37' | 'pct40' | 'pct45' | 'pct50' | 'pct55' | 'pct60' | 'pct62' | 'pct65'
  | 'pct70' | 'pct75' | 'pct80' | 'pct85' | 'pct87' | 'pct90' | 'pct95'

export interface Border {
  style?: BorderStyle
  color?: HexColor
  size?: number // 1/8 point
  space?: number // points
  shadow?: boolean
  frame?: boolean
}

export type BorderStyle = 
  | 'none' | 'single' | 'thick' | 'double' | 'dotted' | 'dashed' | 'dotDash'
  | 'dotDotDash' | 'triple' | 'thinThickSmallGap' | 'thickThinSmallGap'
  | 'thinThickThinSmallGap' | 'thinThickMediumGap' | 'thickThinMediumGap'
  | 'thinThickThinMediumGap' | 'thinThickLargeGap' | 'thickThinLargeGap'
  | 'thinThickThinLargeGap' | 'wave' | 'doubleWave' | 'dashSmallGap' | 'dashDotStroked'
  | 'threeDEmboss' | 'threeDEngrave' | 'outset' | 'inset'

export interface EastAsianLayout {
  id?: number
  combine?: boolean
  combineBrackets?: 'none' | 'round' | 'square' | 'angle' | 'curly'
  vert?: boolean
  vertCompress?: boolean
}

export type EmphasisMark = 'none' | 'dot' | 'comma' | 'circle' | 'underDot'

export interface LanguageInfo {
  val?: string // BCP 47 language tag
  eastAsia?: string
  bidi?: string
}

// ============= 段落（Paragraph）=============

export interface Paragraph {
  properties?: ParagraphProperties
  runs: (TextRun | SpecialRun)[]
  id?: string // w14:paraId
}

export interface ParagraphProperties {
  // 样式
  styleId?: string
  
  // 对齐
  alignment?: 'left' | 'center' | 'right' | 'both' | 'distribute' | 'start' | 'end'
  
  // 缩进
  indentation?: Indentation
  
  // 间距
  spacing?: Spacing
  
  // 边框和底纹
  borders?: ParagraphBorders
  shading?: Shading
  
  // 制表位
  tabs?: Tab[]
  
  // 分页控制
  keepNext?: boolean
  keepLines?: boolean
  pageBreakBefore?: boolean
  widowControl?: boolean
  
  // 编号
  numbering?: NumberingReference
  
  // 框架（定位）
  framePr?: FrameProperties
  
  // 文本方向和对齐
  textDirection?: 'lrTb' | 'tbRl' | 'btLr' | 'lrTbV' | 'tbRlV' | 'tbLrV'
  textAlignment?: 'top' | 'center' | 'baseline' | 'bottom' | 'auto'
  bidi?: boolean
  
  // 行高和网格
  snapToGrid?: boolean
  contextualSpacing?: boolean
  mirrorIndents?: boolean
  
  // 禁则处理
  suppressLineNumbers?: boolean
  suppressAutoHyphens?: boolean
  kinsoku?: boolean
  wordWrap?: boolean
  overflowPunct?: boolean
  topLinePunct?: boolean
  autoSpaceDE?: boolean
  autoSpaceDN?: boolean
  
  // 大纲级别
  outlineLevel?: number
  
  // 分节属性（段落末尾）
  sectionProperties?: SectionProperties
  
  // 修订
  revisionId?: string
  revisionProperties?: ParagraphPropertiesChange
}

export interface Indentation {
  left?: TwipValue
  right?: TwipValue
  start?: TwipValue
  end?: TwipValue
  firstLine?: TwipValue
  hanging?: TwipValue
}

export interface Spacing {
  before?: TwipValue
  after?: TwipValue
  line?: TwipValue
  lineRule?: 'auto' | 'exact' | 'atLeast'
  beforeAutoSpacing?: boolean
  afterAutoSpacing?: boolean
}

export interface ParagraphBorders {
  top?: Border
  bottom?: Border
  left?: Border
  right?: Border
  between?: Border
  bar?: Border
}

export interface Tab {
  position: TwipValue
  alignment?: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'num' | 'start' | 'end'
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot'
}

export interface NumberingReference {
  numId: number
  ilvl: number
}

export interface FrameProperties {
  dropCap?: 'none' | 'drop' | 'margin'
  lines?: number
  width?: TwipValue
  height?: TwipValue
  vAnchor?: 'text' | 'margin' | 'page'
  hAnchor?: 'text' | 'margin' | 'page'
  x?: TwipValue
  xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside'
  y?: TwipValue
  yAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline'
  hRule?: 'auto' | 'exact' | 'atLeast'
  wrap?: boolean
  hSpace?: TwipValue
  vSpace?: TwipValue
}

// ============= 特殊运行（SpecialRun）=============

export type SpecialRun = 
  | BreakRun
  | TabRun
  | FieldRun
  | DrawingRun
  | PictureRun
  | ObjectRun
  | FootnoteRun
  | EndnoteRun
  | CommentRangeRun
  | BookmarkRun

export interface BreakRun {
  type: 'break'
  breakType?: 'page' | 'column' | 'textWrapping'
  clear?: 'none' | 'left' | 'right' | 'all'
}

export interface TabRun {
  type: 'tab'
}

export interface FieldRun {
  type: 'field'
  fieldType: 'begin' | 'separate' | 'end'
  fieldCode?: string
  dirty?: boolean
  locked?: boolean
}

export interface DrawingRun {
  type: 'drawing'
  drawingId: string
  inline?: boolean
  anchor?: DrawingAnchor
}

export interface DrawingAnchor {
  distT?: EMUValue
  distB?: EMUValue
  distL?: EMUValue
  distR?: EMUValue
  simplePos?: boolean
  relativeHeight?: number
  behindDoc?: boolean
  locked?: boolean
  layoutInCell?: boolean
  allowOverlap?: boolean
}

export interface PictureRun {
  type: 'picture'
  relationshipId: string
  width?: EMUValue
  height?: EMUValue
  description?: string
}

export interface ObjectRun {
  type: 'object'
  objectId: string
  progId?: string
}

export interface FootnoteRun {
  type: 'footnote'
  id: string
}

export interface EndnoteRun {
  type: 'endnote'
  id: string
}

export interface CommentRangeRun {
  type: 'commentRange'
  commentId: string
  rangeType: 'start' | 'end'
}

export interface BookmarkRun {
  type: 'bookmark'
  bookmarkId: string
  bookmarkName?: string
  rangeType: 'start' | 'end'
}

// ============= 分节属性（SectionProperties）=============

export interface SectionProperties {
  // 页面设置
  pageSize?: PageSize
  pageMargin?: PageMargin
  pageOrientation?: 'portrait' | 'landscape'
  
  // 页眉页脚
  headerReference?: HeaderFooterReference[]
  footerReference?: HeaderFooterReference[]
  
  // 页码
  pageNumberType?: PageNumberType
  
  // 分栏
  columns?: ColumnProperties
  
  // 行号
  lineNumberType?: LineNumberType
  
  // 分节类型
  sectionType?: 'nextPage' | 'nextColumn' | 'continuous' | 'evenPage' | 'oddPage'
  
  // 文本方向
  textDirection?: 'lrTb' | 'tbRl' | 'btLr' | 'lrTbV' | 'tbRlV' | 'tbLrV'
  
  // 垂直对齐
  verticalAlign?: 'top' | 'center' | 'both' | 'bottom'
  
  // 边框
  borders?: SectionBorders
  
  // 其他
  titlePage?: boolean
  rtlGutter?: boolean
  formProtection?: boolean
  bidi?: boolean
}

export interface PageSize {
  width: TwipValue
  height: TwipValue
  code?: number
}

export interface PageMargin {
  top: TwipValue
  right: TwipValue
  bottom: TwipValue
  left: TwipValue
  header: TwipValue
  footer: TwipValue
  gutter: TwipValue
}

export interface HeaderFooterReference {
  type: 'default' | 'first' | 'even'
  relationshipId: string
}

export interface PageNumberType {
  format?: 'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter'
  start?: number
  chapStyle?: number
  chapSep?: 'hyphen' | 'period' | 'colon' | 'emDash' | 'enDash'
}

export interface ColumnProperties {
  count?: number
  space?: TwipValue
  equalWidth?: boolean
  separator?: boolean
  columns?: Column[]
}

export interface Column {
  width: TwipValue
  space: TwipValue
}

export interface LineNumberType {
  countBy?: number
  start?: number
  restart?: 'newPage' | 'newSection' | 'continuous'
  distance?: TwipValue
}

export interface SectionBorders {
  top?: Border
  bottom?: Border
  left?: Border
  right?: Border
  offsetFrom?: 'text' | 'page'
  zOrder?: 'front' | 'back'
  display?: 'allPages' | 'firstPage' | 'notFirstPage'
}

// ============= 修订属性（RevisionProperties）=============

export interface ParagraphPropertiesChange {
  author?: string
  date?: Date
  id?: string
  properties?: ParagraphProperties
}

export interface RunPropertiesChange {
  author?: string
  date?: Date
  id?: string
  properties?: RunProperties
}

// ============= 表格（Table）=============

export interface Table {
  properties?: TableProperties
  grid?: TableGrid
  rows: TableRow[]
}

export interface TableProperties {
  styleId?: string
  width?: TableWidth
  alignment?: 'left' | 'center' | 'right' | 'start' | 'end'
  indent?: TwipValue
  borders?: TableBorders
  shading?: Shading
  layout?: 'fixed' | 'autofit'
  cellSpacing?: TwipValue
  cellMargin?: TableCellMargin
  look?: TableLook
  caption?: string
  description?: string
  bidiVisual?: boolean
  overlap?: 'never' | 'overlap'
}

export interface TableWidth {
  type: 'auto' | 'dxa' | 'pct' | 'nil'
  value: number
}

export interface TableBorders {
  top?: Border
  bottom?: Border
  left?: Border
  right?: Border
  insideH?: Border
  insideV?: Border
}

export interface TableCellMargin {
  top?: TwipValue
  bottom?: TwipValue
  left?: TwipValue
  right?: TwipValue
  start?: TwipValue
  end?: TwipValue
}

export interface TableLook {
  firstRow?: boolean
  lastRow?: boolean
  firstColumn?: boolean
  lastColumn?: boolean
  noHBand?: boolean
  noVBand?: boolean
}

export interface TableGrid {
  columns: TableGridColumn[]
}

export interface TableGridColumn {
  width: TwipValue
}

export interface TableRow {
  properties?: TableRowProperties
  cells: TableCell[]
}

export interface TableRowProperties {
  cantSplit?: boolean
  height?: TableRowHeight
  header?: boolean
  gridBefore?: number
  gridAfter?: number
  hidden?: boolean
}

export interface TableRowHeight {
  value: TwipValue
  rule?: 'auto' | 'exact' | 'atLeast'
}

export interface TableCell {
  properties?: TableCellProperties
  content: Paragraph[]
}

export interface TableCellProperties {
  width?: TableWidth
  gridSpan?: number
  vMerge?: 'restart' | 'continue'
  borders?: TableCellBorders
  shading?: Shading
  margins?: TableCellMargin
  verticalAlign?: 'top' | 'center' | 'bottom'
  textDirection?: 'lrTb' | 'tbRl' | 'btLr' | 'lrTbV' | 'tbRlV' | 'tbLrV'
  fitText?: boolean
  noWrap?: boolean
  hidden?: boolean
}

export interface TableCellBorders {
  top?: Border
  bottom?: Border
  left?: Border
  right?: Border
  insideH?: Border
  insideV?: Border
  tl2br?: Border
  tr2bl?: Border
}

// ============= 样式（Styles）=============

export interface Style {
  styleId: string
  type: 'paragraph' | 'character' | 'table' | 'numbering'
  name?: string
  basedOn?: string
  next?: string
  link?: string
  autoRedefine?: boolean
  hidden?: boolean
  uiPriority?: number
  semiHidden?: boolean
  unhideWhenUsed?: boolean
  qFormat?: boolean
  locked?: boolean
  personal?: boolean
  personalCompose?: boolean
  personalReply?: boolean
  rsid?: string
  paragraphProperties?: ParagraphProperties
  runProperties?: RunProperties
  tableProperties?: TableProperties
  tableRowProperties?: TableRowProperties
  tableCellProperties?: TableCellProperties
}

// ============= 编号（Numbering）=============

export interface NumberingDefinition {
  abstractNumId: number
  levels: NumberingLevel[]
}

export interface NumberingLevel {
  level: number
  start?: number
  format?: NumberingFormat
  text?: string
  alignment?: 'left' | 'center' | 'right'
  paragraphProperties?: ParagraphProperties
  runProperties?: RunProperties
  restart?: number
  suffix?: 'tab' | 'space' | 'nothing'
  isLegal?: boolean
}

export type NumberingFormat = 
  | 'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter'
  | 'ordinal' | 'cardinalText' | 'ordinalText' | 'hex' | 'chicago'
  | 'ideographDigital' | 'japaneseCounting' | 'aiueo' | 'iroha'
  | 'decimalFullWidth' | 'decimalHalfWidth' | 'japaneseLegal'
  | 'japaneseDigitalTenThousand' | 'decimalEnclosedCircle'
  | 'decimalFullWidth2' | 'aiueoFullWidth' | 'irohaFullWidth'
  | 'decimalZero' | 'bullet' | 'ganada' | 'chosung'
  | 'decimalEnclosedFullstop' | 'decimalEnclosedParen'
  | 'decimalEnclosedCircleChinese' | 'ideographEnclosedCircle'
  | 'ideographTraditional' | 'ideographZodiac' | 'ideographZodiacTraditional'
  | 'taiwaneseCounting' | 'ideographLegalTraditional' | 'taiwaneseCountingThousand'
  | 'taiwaneseDigital' | 'chineseCounting' | 'chineseLegalSimplified'
  | 'chineseCountingThousand' | 'koreanDigital' | 'koreanCounting'
  | 'koreanLegal' | 'koreanDigital2' | 'vietnameseCounting'
  | 'russianLower' | 'russianUpper' | 'none' | 'numberInDash'
  | 'hebrew1' | 'hebrew2' | 'arabicAlpha' | 'arabicAbjad'
  | 'hindiVowels' | 'hindiConsonants' | 'hindiNumbers' | 'hindiCounting'
  | 'thaiLetters' | 'thaiNumbers' | 'thaiCounting'

// ============= 完整文档（Document）=============

export interface Document {
  body: DocumentBody
  background?: DocumentBackground
  settings?: DocumentSettings
}

export interface DocumentBody {
  sections: Section[]
}

export interface Section {
  properties?: SectionProperties
  content: (Paragraph | Table)[]
}

export interface DocumentBackground {
  color?: HexColor
  themeColor?: string
  themeShade?: string
  themeTint?: string
}

export interface DocumentSettings {
  zoom?: number
  view?: 'none' | 'print' | 'outline' | 'masterPages' | 'normal' | 'web'
  trackRevisions?: boolean
  doNotTrackMoves?: boolean
  doNotTrackFormatting?: boolean
  defaultTabStop?: TwipValue
  characterSpacingControl?: 'doNotCompress' | 'compressPunctuation' | 'compressPunctuationAndJapaneseKana'
  evenAndOddHeaders?: boolean
  bookFoldPrinting?: boolean
  bookFoldRevPrinting?: boolean
  bookFoldPrintingSheets?: number
  drawingGridHorizontalSpacing?: TwipValue
  drawingGridVerticalSpacing?: TwipValue
  displayHorizontalDrawingGridEvery?: number
  displayVerticalDrawingGridEvery?: number
  doNotUseMarginsForDrawingGridOrigin?: boolean
  doNotShadeFormData?: boolean
  noPunctuationKerning?: boolean
  printTwoOnOne?: boolean
  strictFirstAndLastChars?: boolean
  noLineBreaksAfter?: string
  noLineBreaksBefore?: string
  savePreviewPicture?: boolean
  doNotValidateAgainstSchema?: boolean
  saveInvalidXml?: boolean
  ignoreMixedContent?: boolean
  alwaysShowPlaceholderText?: boolean
  doNotDemarcateInvalidXml?: boolean
  saveXmlDataOnly?: boolean
  useXSLTWhenSaving?: boolean
  saveThroughXslt?: string
  showXMLTags?: boolean
  alwaysMergeEmptyNamespace?: boolean
  updateFields?: boolean
  footnotePr?: FootnoteProperties
  endnotePr?: EndnoteProperties
  compat?: CompatibilitySettings
  rsids?: RevisionIdentifiers
  mathPr?: MathProperties
  attachedTemplate?: string
  linkStyles?: boolean
  stylePaneFormatFilter?: string
  stylePaneSortMethod?: string
  documentType?: 'notSpecified' | 'letter' | 'eMail'
  mailMerge?: MailMergeSettings
  revisionView?: RevisionView
  formsDesign?: boolean
  attachedSchema?: string[]
  themeFontLang?: ThemeFontLanguage
  clrSchemeMapping?: ColorSchemeMapping
  doNotIncludeSubdocsInStats?: boolean
  doNotAutoCompressPictures?: boolean
  forceUpgrade?: boolean
  captions?: CaptionSettings[]
  readModeInkLockDown?: ReadModeInkLockDown
  smartTagType?: SmartTagType[]
  shapeDefaults?: ShapeDefaults
  doNotEmbedSmartTags?: boolean
  decimalSymbol?: string
  listSeparator?: string
}

export interface FootnoteProperties {
  position?: 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd'
  numberingFormat?: NumberingFormat
  numberingStart?: number
  numberingRestart?: 'continuous' | 'eachSect' | 'eachPage'
}

export interface EndnoteProperties {
  position?: 'sectEnd' | 'docEnd'
  numberingFormat?: NumberingFormat
  numberingStart?: number
  numberingRestart?: 'continuous' | 'eachSect'
}

export interface CompatibilitySettings {
  [key: string]: boolean | string | number
}

export interface RevisionIdentifiers {
  rsidRoot?: string
  rsids?: string[]
}

export interface MathProperties {
  mathFont?: string
  brkBin?: 'before' | 'after' | 'repeat'
  brkBinSub?: 'minusMinus' | 'minusPlus' | 'plusMinus' | 'plusPlus'
  defJc?: 'left' | 'right' | 'center' | 'centerGroup'
  dispDef?: boolean
  interSp?: TwipValue
  intraSp?: TwipValue
  lMargin?: TwipValue
  rMargin?: TwipValue
  postSp?: TwipValue
  preSp?: TwipValue
  smallFrac?: boolean
  wrapIndent?: TwipValue
  wrapRight?: boolean
}

export interface MailMergeSettings {
  mainDocumentType?: string
  linkToQuery?: boolean
  dataType?: string
  connectString?: string
  query?: string
  dataSourceReference?: string
  headerSourceReference?: string
  doNotSuppressBlankLines?: boolean
  destination?: string
  addressFieldName?: string
  mailSubject?: string
  mailAsAttachment?: boolean
  viewMergedData?: boolean
  activeRecord?: number
  checkErrors?: number
}

export interface RevisionView {
  markup?: 'none' | 'simple' | 'all'
  formatting?: boolean
  inkAnnotations?: boolean
  insDel?: boolean
}

export interface ThemeFontLanguage {
  val?: string
  eastAsia?: string
  bidi?: string
}

export interface ColorSchemeMapping {
  bg1?: string
  t1?: string
  bg2?: string
  t2?: string
  accent1?: string
  accent2?: string
  accent3?: string
  accent4?: string
  accent5?: string
  accent6?: string
  hyperlink?: string
  followedHyperlink?: string
}

export interface CaptionSettings {
  name?: string
  pos?: 'above' | 'below' | 'left' | 'right'
  chapNum?: boolean
  heading?: number
  noLabel?: boolean
  numFmt?: NumberingFormat
  sep?: 'hyphen' | 'period' | 'colon' | 'emDash' | 'enDash'
}

export interface ReadModeInkLockDown {
  actualPg?: boolean
  w?: number
  h?: number
  fontSz?: number
}

export interface SmartTagType {
  namespaceuri?: string
  name?: string
  url?: string
}

export interface ShapeDefaults {
  [key: string]: any
}
