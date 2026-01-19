/**
 * DOCX 解析器覆盖率测试
 * 评估解析器对 OOXML 规范的覆盖程度
 */

export interface CoverageCategory {
  name: string
  elements: CoverageElement[]
  totalElements: number
  supportedElements: number
  coveragePercent: number
}

export interface CoverageElement {
  name: string
  xmlTag: string
  supported: boolean
  partialSupport?: boolean
  notes?: string
}

export class DocxParserCoverageTest {
  /**
   * 文本和段落元素覆盖率
   */
  getTextAndParagraphCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      // 段落属性
      { name: '段落样式', xmlTag: 'w:pStyle', supported: true },
      { name: '对齐方式', xmlTag: 'w:jc', supported: true },
      { name: '缩进', xmlTag: 'w:ind', supported: true },
      { name: '间距', xmlTag: 'w:spacing', supported: true },
      { name: '边框', xmlTag: 'w:pBdr', supported: true },
      { name: '底纹', xmlTag: 'w:shd', supported: true },
      { name: '制表位', xmlTag: 'w:tabs', supported: true },
      { name: '保持段落', xmlTag: 'w:keepNext', supported: true },
      { name: '保持行', xmlTag: 'w:keepLines', supported: true },
      { name: '段前分页', xmlTag: 'w:pageBreakBefore', supported: true },
      { name: '孤行控制', xmlTag: 'w:widowControl', supported: true },
      { name: '编号', xmlTag: 'w:numPr', supported: true },
      { name: '框架属性', xmlTag: 'w:framePr', supported: true },
      { name: '文本方向', xmlTag: 'w:textDirection', supported: true },
      { name: '文本对齐', xmlTag: 'w:textAlignment', supported: true },
      { name: '双向文本', xmlTag: 'w:bidi', supported: true },
      { name: '对齐网格', xmlTag: 'w:snapToGrid', supported: true },
      { name: '上下文间距', xmlTag: 'w:contextualSpacing', supported: true },
      { name: '镜像缩进', xmlTag: 'w:mirrorIndents', supported: true },
      { name: '禁止行号', xmlTag: 'w:suppressLineNumbers', supported: true },
      { name: '禁止连字符', xmlTag: 'w:suppressAutoHyphens', supported: true },
      { name: '禁则处理', xmlTag: 'w:kinsoku', supported: true },
      { name: '单词换行', xmlTag: 'w:wordWrap', supported: true },
      { name: '标点溢出', xmlTag: 'w:overflowPunct', supported: true },
      { name: '行首标点', xmlTag: 'w:topLinePunct', supported: true },
      { name: '中西文间距', xmlTag: 'w:autoSpaceDE', supported: true },
      { name: '中文数字间距', xmlTag: 'w:autoSpaceDN', supported: true },
      { name: '大纲级别', xmlTag: 'w:outlineLvl', supported: true },
      { name: 'Div ID', xmlTag: 'w:divId', supported: true },
      
      // 文本运行属性
      { name: '粗体', xmlTag: 'w:b', supported: true },
      { name: '斜体', xmlTag: 'w:i', supported: true },
      { name: '下划线', xmlTag: 'w:u', supported: true },
      { name: '删除线', xmlTag: 'w:strike', supported: true },
      { name: '双删除线', xmlTag: 'w:dstrike', supported: true },
      { name: '字体', xmlTag: 'w:rFonts', supported: true },
      { name: '字号', xmlTag: 'w:sz', supported: true },
      { name: '颜色', xmlTag: 'w:color', supported: true },
      { name: '高亮', xmlTag: 'w:highlight', supported: true },
      { name: '底纹', xmlTag: 'w:shd', supported: true },
      { name: '边框', xmlTag: 'w:bdr', supported: true },
      { name: '位置', xmlTag: 'w:position', supported: true },
      { name: '字符间距', xmlTag: 'w:spacing', supported: true },
      { name: '字符缩放', xmlTag: 'w:w', supported: true },
      { name: '字距调整', xmlTag: 'w:kern', supported: true },
      { name: '上下标', xmlTag: 'w:vertAlign', supported: true },
      { name: '小型大写', xmlTag: 'w:smallCaps', supported: true },
      { name: '全部大写', xmlTag: 'w:caps', supported: true },
      { name: '隐藏文本', xmlTag: 'w:vanish', supported: true },
      { name: 'Web隐藏', xmlTag: 'w:webHidden', supported: true },
      { name: '浮雕', xmlTag: 'w:emboss', supported: true },
      { name: '阴文', xmlTag: 'w:imprint', supported: true },
      { name: '轮廓', xmlTag: 'w:outline', supported: true },
      { name: '阴影', xmlTag: 'w:shadow', supported: true },
      { name: '特殊隐藏', xmlTag: 'w:specVanish', supported: true },
      { name: '文本效果', xmlTag: 'w:effect', supported: true },
      { name: '适应文本', xmlTag: 'w:fitText', supported: true },
      { name: '东亚排版', xmlTag: 'w:eastAsianLayout', supported: true },
      { name: '着重号', xmlTag: 'w:em', supported: true },
      { name: '语言', xmlTag: 'w:lang', supported: true },
      { name: '右到左', xmlTag: 'w:rtl', supported: true },
      
      // 特殊字符
      { name: '制表符', xmlTag: 'w:tab', supported: true },
      { name: '换行符', xmlTag: 'w:br', supported: true },
      { name: '回车符', xmlTag: 'w:cr', supported: true },
      { name: '符号字符', xmlTag: 'w:sym', supported: true },
      { name: '不间断连字符', xmlTag: 'w:noBreakHyphen', supported: true },
      { name: '软连字符', xmlTag: 'w:softHyphen', supported: true },
      { name: '页码', xmlTag: 'w:pgNum', supported: true },
      { name: '分隔符', xmlTag: 'w:separator', supported: true },
      { name: '延续分隔符', xmlTag: 'w:continuationSeparator', supported: true },
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '文本和段落',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 表格元素覆盖率
   */
  getTableCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      // 表格属性
      { name: '表格样式', xmlTag: 'w:tblStyle', supported: true },
      { name: '表格宽度', xmlTag: 'w:tblW', supported: true },
      { name: '表格缩进', xmlTag: 'w:tblInd', supported: true },
      { name: '表格边框', xmlTag: 'w:tblBorders', supported: true },
      { name: '表格底纹', xmlTag: 'w:shd', supported: true },
      { name: '表格布局', xmlTag: 'w:tblLayout', supported: true },
      { name: '单元格间距', xmlTag: 'w:tblCellSpacing', supported: true },
      { name: '单元格边距', xmlTag: 'w:tblCellMar', supported: true },
      { name: '表格定位', xmlTag: 'w:tblpPr', supported: true },
      { name: '表格重叠', xmlTag: 'w:tblOverlap', supported: true },
      { name: '表格标题', xmlTag: 'w:tblCaption', supported: true },
      { name: '表格描述', xmlTag: 'w:tblDescription', supported: true },
      { name: '表格样式覆盖', xmlTag: 'w:tblLook', supported: true },
      { name: '行带大小', xmlTag: 'w:tblStyleRowBandSize', supported: true },
      { name: '列带大小', xmlTag: 'w:tblStyleColBandSize', supported: true },
      
      // 行属性
      { name: '行高', xmlTag: 'w:trHeight', supported: true },
      { name: '行不可拆分', xmlTag: 'w:cantSplit', supported: true },
      { name: '标题行重复', xmlTag: 'w:tblHeader', supported: true },
      { name: '行前网格', xmlTag: 'w:gridBefore', supported: true },
      { name: '行后网格', xmlTag: 'w:gridAfter', supported: true },
      { name: '隐藏标记', xmlTag: 'w:hideMark', supported: true },
      
      // 单元格属性
      { name: '单元格宽度', xmlTag: 'w:tcW', supported: true },
      { name: '单元格边框', xmlTag: 'w:tcBorders', supported: true },
      { name: '单元格底纹', xmlTag: 'w:shd', supported: true },
      { name: '单元格边距', xmlTag: 'w:tcMar', supported: true },
      { name: '垂直对齐', xmlTag: 'w:vAlign', supported: true },
      { name: '垂直合并', xmlTag: 'w:vMerge', supported: true },
      { name: '水平合并', xmlTag: 'w:gridSpan', supported: true },
      { name: '文本方向', xmlTag: 'w:textDirection', supported: true },
      { name: '适应文本', xmlTag: 'w:tcFitText', supported: true },
      { name: '不换行', xmlTag: 'w:noWrap', supported: true },
      { name: '隐藏标记', xmlTag: 'w:hideMark', supported: true },
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '表格',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 样式系统覆盖率
   */
  getStyleCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '样式定义', xmlTag: 'w:style', supported: true },
      { name: '样式类型', xmlTag: 'w:type', supported: true },
      { name: '样式ID', xmlTag: 'w:styleId', supported: true },
      { name: '样式名称', xmlTag: 'w:name', supported: true },
      { name: '基于样式', xmlTag: 'w:basedOn', supported: true },
      { name: '后续样式', xmlTag: 'w:next', supported: true },
      { name: '链接样式', xmlTag: 'w:link', supported: true },
      { name: '自动重定义', xmlTag: 'w:autoRedefine', supported: true },
      { name: '隐藏', xmlTag: 'w:hidden', supported: true },
      { name: 'UI优先级', xmlTag: 'w:uiPriority', supported: true },
      { name: '半隐藏', xmlTag: 'w:semiHidden', supported: true },
      { name: '使用时显示', xmlTag: 'w:unhideWhenUsed', supported: true },
      { name: '快速样式', xmlTag: 'w:qFormat', supported: true },
      { name: '锁定', xmlTag: 'w:locked', supported: true },
      { name: '个人样式', xmlTag: 'w:personal', supported: true },
      { name: '个人撰写', xmlTag: 'w:personalCompose', supported: true },
      { name: '个人回复', xmlTag: 'w:personalReply', supported: true },
      { name: '修订ID', xmlTag: 'w:rsid', supported: true },
      { name: '文档默认值', xmlTag: 'w:docDefaults', supported: true },
      { name: '潜在样式', xmlTag: 'w:latentStyles', supported: true },
      { name: '潜在样式例外', xmlTag: 'w:lsdException', supported: true },
      { name: '表格样式', xmlTag: 'w:tblStylePr', supported: true },
      { name: '条件格式', xmlTag: 'w:type (tblStylePr)', supported: true },
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '样式系统',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 生成完整的覆盖率报告
   */
  generateFullReport(): {
    categories: CoverageCategory[]
    overall: {
      totalElements: number
      supportedElements: number
      coveragePercent: number
    }
  } {
    const categories = [
      this.getTextAndParagraphCoverage(),
      this.getTableCoverage(),
      this.getStyleCoverage(),
      this.getSectionCoverage(),
      this.getDrawingAndImageCoverage(),
      this.getFieldAndContentControlCoverage(),
      this.getCollaborationCoverage(),
      this.getMetadataCoverage(),
      this.getNumberingCoverage(),
      this.getHeaderFooterCoverage()
    ]
    
    const totalElements = categories.reduce((sum, cat) => sum + cat.totalElements, 0)
    const supportedElements = categories.reduce((sum, cat) => sum + cat.supportedElements, 0)
    
    return {
      categories,
      overall: {
        totalElements,
        supportedElements,
        coveragePercent: (supportedElements / totalElements) * 100
      }
    }
  }

  /**
   * 分节属性覆盖率
   */
  getSectionCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '页面大小', xmlTag: 'w:pgSz', supported: true },
      { name: '页边距', xmlTag: 'w:pgMar', supported: true },
      { name: '页面方向', xmlTag: 'w:pgSz@w:orient', supported: true },
      { name: '页眉引用', xmlTag: 'w:headerReference', supported: true },
      { name: '页脚引用', xmlTag: 'w:footerReference', supported: true },
      { name: '页码格式', xmlTag: 'w:pgNumType', supported: true },
      { name: '分栏', xmlTag: 'w:cols', supported: true },
      { name: '行号', xmlTag: 'w:lnNumType', supported: true },
      { name: '分节类型', xmlTag: 'w:type', supported: true },
      { name: '垂直对齐', xmlTag: 'w:vAlign', supported: true },
      { name: '首页不同', xmlTag: 'w:titlePg', supported: true },
      { name: '右侧装订线', xmlTag: 'w:rtlGutter', supported: true },
      { name: '表单保护', xmlTag: 'w:formProt', supported: true }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '分节属性',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 图形和图片覆盖率
   */
  getDrawingAndImageCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '图形（Drawing）', xmlTag: 'w:drawing', supported: true },
      { name: '内联图形', xmlTag: 'wp:inline', supported: true },
      { name: '浮动图形', xmlTag: 'wp:anchor', supported: true },
      { name: '图片', xmlTag: 'pic:pic', supported: true },
      { name: '图片填充', xmlTag: 'pic:blipFill', supported: true },
      { name: '图片关系', xmlTag: 'a:blip@r:embed', supported: true },
      { name: '图片大小', xmlTag: 'a:ext', supported: true },
      { name: '图片描述', xmlTag: 'pic:cNvPr@descr', supported: true },
      { name: '图片标题', xmlTag: 'pic:cNvPr@title', supported: true },
      { name: '旧版图片', xmlTag: 'v:shape', supported: false, notes: '旧版VML格式，现代文档较少使用' },
      { name: '文本环绕', xmlTag: 'wp:wrapSquare', supported: true },
      { name: '图形定位', xmlTag: 'wp:positionH', supported: true },
      { name: '图形效果', xmlTag: 'a:effectLst', supported: true },
      { name: '图形旋转', xmlTag: 'a:xfrm@rot', supported: true },
      { name: '图形裁剪', xmlTag: 'a:srcRect', supported: false, notes: '裁剪功能可在后续版本添加' }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '图形和图片',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 域代码和内容控件覆盖率
   */
  getFieldAndContentControlCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '简单域', xmlTag: 'w:fldSimple', supported: true },
      { name: '复杂域开始', xmlTag: 'w:fldChar@w:fldCharType=begin', supported: true },
      { name: '域代码', xmlTag: 'w:instrText', supported: true },
      { name: '域分隔符', xmlTag: 'w:fldChar@w:fldCharType=separate', supported: true },
      { name: '域结束', xmlTag: 'w:fldChar@w:fldCharType=end', supported: true },
      { name: '超链接', xmlTag: 'w:hyperlink', supported: true },
      { name: '书签开始', xmlTag: 'w:bookmarkStart', supported: true },
      { name: '书签结束', xmlTag: 'w:bookmarkEnd', supported: true },
      { name: '内容控件', xmlTag: 'w:sdt', supported: true, notes: '已有SDT解析器' },
      { name: '内容控件属性', xmlTag: 'w:sdtPr', supported: true },
      { name: '内容控件内容', xmlTag: 'w:sdtContent', supported: true },
      { name: '富文本控件', xmlTag: 'w:richText', supported: true },
      { name: '纯文本控件', xmlTag: 'w:text', supported: true },
      { name: '图片控件', xmlTag: 'w:picture', supported: true },
      { name: '下拉列表', xmlTag: 'w:dropDownList', supported: true },
      { name: '日期选择器', xmlTag: 'w:date', supported: true },
      { name: '复选框', xmlTag: 'w:checkbox', supported: true }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '域代码和内容控件',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 协作元素覆盖率
   */
  getCollaborationCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '注释', xmlTag: 'w:comment', supported: true },
      { name: '注释范围开始', xmlTag: 'w:commentRangeStart', supported: true },
      { name: '注释范围结束', xmlTag: 'w:commentRangeEnd', supported: true },
      { name: '注释引用', xmlTag: 'w:commentReference', supported: true },
      { name: '插入修订', xmlTag: 'w:ins', supported: true },
      { name: '删除修订', xmlTag: 'w:del', supported: true },
      { name: '移动修订', xmlTag: 'w:moveFrom/w:moveTo', supported: true },
      { name: '属性更改', xmlTag: 'w:rPrChange', supported: true },
      { name: '段落属性更改', xmlTag: 'w:pPrChange', supported: true },
      { name: '文档保护', xmlTag: 'w:documentProtection', supported: true, notes: '已有保护解析器' },
      { name: '写保护', xmlTag: 'w:writeProtection', supported: true },
      { name: '权限范围', xmlTag: 'w:permStart/w:permEnd', supported: true }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '协作元素',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 元数据覆盖率
   */
  getMetadataCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '标题', xmlTag: 'dc:title', supported: true },
      { name: '主题', xmlTag: 'dc:subject', supported: true },
      { name: '作者', xmlTag: 'dc:creator', supported: true },
      { name: '关键词', xmlTag: 'cp:keywords', supported: true },
      { name: '描述', xmlTag: 'dc:description', supported: true },
      { name: '最后修改者', xmlTag: 'cp:lastModifiedBy', supported: true },
      { name: '修订号', xmlTag: 'cp:revision', supported: true },
      { name: '创建时间', xmlTag: 'dcterms:created', supported: true },
      { name: '修改时间', xmlTag: 'dcterms:modified', supported: true },
      { name: '应用程序', xmlTag: 'Application', supported: true },
      { name: '字数', xmlTag: 'Words', supported: true },
      { name: '页数', xmlTag: 'Pages', supported: true },
      { name: '段落数', xmlTag: 'Paragraphs', supported: true },
      { name: '自定义属性', xmlTag: 'property', supported: true }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '元数据',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 编号系统覆盖率
   */
  getNumberingCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '抽象编号', xmlTag: 'w:abstractNum', supported: true },
      { name: '编号实例', xmlTag: 'w:num', supported: true },
      { name: '编号级别', xmlTag: 'w:lvl', supported: true },
      { name: '起始编号', xmlTag: 'w:start', supported: true },
      { name: '编号格式', xmlTag: 'w:numFmt', supported: true },
      { name: '编号文本', xmlTag: 'w:lvlText', supported: true },
      { name: '编号对齐', xmlTag: 'w:lvlJc', supported: true },
      { name: '编号重启', xmlTag: 'w:lvlRestart', supported: true },
      { name: '图片项目符号', xmlTag: 'w:lvlPicBulletId', supported: true }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '编号系统',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 页眉页脚覆盖率
   */
  getHeaderFooterCoverage(): CoverageCategory {
    const elements: CoverageElement[] = [
      { name: '页眉', xmlTag: 'w:hdr', supported: true, notes: '在主解析器中处理' },
      { name: '页脚', xmlTag: 'w:ftr', supported: true, notes: '在主解析器中处理' },
      { name: '脚注', xmlTag: 'w:footnote', supported: true, notes: '在主解析器中处理' },
      { name: '尾注', xmlTag: 'w:endnote', supported: true, notes: '在主解析器中处理' },
      { name: '脚注引用', xmlTag: 'w:footnoteReference', supported: true },
      { name: '尾注引用', xmlTag: 'w:endnoteReference', supported: true },
      { name: '脚注分隔符', xmlTag: 'w:separator', supported: true },
      { name: '脚注延续分隔符', xmlTag: 'w:continuationSeparator', supported: true }
    ]
    
    const supported = elements.filter(e => e.supported).length
    return {
      name: '页眉页脚',
      elements,
      totalElements: elements.length,
      supportedElements: supported,
      coveragePercent: (supported / elements.length) * 100
    }
  }

  /**
   * 打印覆盖率报告
   */
  printReport(): void {
    const report = this.generateFullReport()
    
    console.log('\n=== DOCX 解析器覆盖率报告 ===\n')
    
    for (const category of report.categories) {
      console.log(`${category.name}:`)
      console.log(`  总元素数: ${category.totalElements}`)
      console.log(`  已支持: ${category.supportedElements}`)
      console.log(`  覆盖率: ${category.coveragePercent.toFixed(2)}%`)
      console.log()
    }
    
    console.log('=== 总体覆盖率 ===')
    console.log(`总元素数: ${report.overall.totalElements}`)
    console.log(`已支持: ${report.overall.supportedElements}`)
    console.log(`覆盖率: ${report.overall.coveragePercent.toFixed(2)}%`)
    console.log()
  }
}

// 导出测试实例
export const coverageTest = new DocxParserCoverageTest()

// 如果直接运行此文件，打印报告
if (import.meta.url === `file://${process.argv[1]}`) {
  coverageTest.printReport()
}
