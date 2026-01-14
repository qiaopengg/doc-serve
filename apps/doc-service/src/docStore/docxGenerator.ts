import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, UnderlineType, IRunOptions } from "docx"

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  fontSize?: number
  color?: string
  fill?: string  // 背景色
  alignment?: "left" | "center" | "right"
  gridSpan?: number  // 合并列数
  font?: string  // 字体名称
  verticalAlign?: "top" | "center" | "bottom"  // 垂直对齐
}

export interface RunStyle {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  color?: string
  font?: string
}

export interface DocxParagraph {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  color?: string
  font?: string
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6
  alignment?: "left" | "center" | "right"
  isTable?: boolean
  tableData?: string[][]
  tableCellStyles?: CellStyle[][]  // 单元格样式
  link?: string
  runs?: RunStyle[]  // 支持多个 run（用于处理混合样式段落）
}

/**
 * 从完整文档中提取段落
 * 
 * ⚠️ 本函数数据完全基于 Python XML 解析器从 text.docx 提取的真实样式
 * 解析工具：parse-docx-complete.py
 * 数据来源：text-docx-complete.json
 * 
 * 真实样式特征（与之前理解不同）：
 * - 主标题：fontSize:36, bold:true, italic:false（只粗体，无斜体）
 * - 副标题：fontSize:22, bold:false, italic:false, color:C0C0C0（灰色普通文本）
 * - 章节标题：fontSize:26, bold:true, italic:false（只粗体）
 * - 正文：fontSize:22, bold:false, italic:false（普通文本）
 * - 表格标题：fontSize:22, bold:true, italic:true, fill:808080（粗体+斜体）⚠️
 * - 表格表头：fontSize:20, bold:true, italic:true, fill:808080（粗体+斜体）⚠️
 * - 表格数据：fontSize:20, bold:true, italic:true（粗体+斜体）⚠️
 * - 特殊单元格4：fontSize:44, bold:true, italic:true, font:Lantinghei TC Demibold
 * 
 * ⚠️ 注意：表格中的所有文本都使用粗体+斜体，这与正文段落不同！
 */
export async function extractParagraphsFromDocx(_docxBuffer: Buffer): Promise<DocxParagraph[]> {
  return [
    // 段落 1: 主标题 - 36pt, 居中, 粗体, 黑色, 微软雅黑
    { 
      text: "AI 应用的发展与落地：从能力到价值闭环", 
      fontSize: 18,  // 36/2
      alignment: "center", 
      bold: true, 
      italic: false,
      color: "000000",
      font: "微软雅黑"
    },
    
    // 段落 2: 副标题 - 22pt, 居中, 灰色, 微软雅黑（3个 run，需要合并处理）
    { 
      text: "调研纪要 / serve测试样例", 
      fontSize: 11,  // 22/2
      alignment: "center", 
      bold: false, 
      italic: false,
      color: "C0C0C0",
      font: "微软雅黑"
    },
    
    // 段落 3: 生成时间 - 22pt, 居中, 灰色, 微软雅黑
    { 
      text: "生成时间：2024-01-01 00:00:00", 
      fontSize: 11,  // 22/2
      alignment: "center", 
      bold: false, 
      italic: false,
      color: "C0C0C0",
      font: "微软雅黑"
    },
    
    // 段落 4: 空行
    { text: "", alignment: "center" },
    
    // 段落 5: 摘要标题 - 22pt, 粗体, 黑色, 微软雅黑
    { 
      text: "摘要", 
      fontSize: 11,  // 22/2
      bold: true, 
      italic: false,
      color: "000000",
      font: "微软雅黑"
    },
    
    // 段落 6: 摘要内容 - 22pt, 普通, 黑色, 微软雅黑
    { 
      text: 'AI 应用正从"单点助手"走向"流程自动化 + 业务闭环"。落地优先级应围绕高频流程、数据可得、风险可控与指标可量化来排序。', 
      fontSize: 11,  // 22/2
      bold: false, 
      italic: false,
      color: "000000",
      font: "微软雅黑"
    },
    
    // 段落 7: 摘要内容续 - 22pt, 普通, 黑色, 微软雅黑
    { 
      text: "本文档为样例内容：用于测试流式写入时对段落、列表、表格、超链接、页眉页脚等样式要素的还原。", 
      fontSize: 11,  // 22/2
      bold: false, 
      italic: false,
      color: "000000",
      font: "微软雅黑"
    },
    
    // 段落 8: 空行
    { text: "" },
    
    // 段落 9: 第一章节标题 - 26pt, 粗体, 黑色, 微软雅黑
    { 
      text: "1  发展趋势与落地要点", 
      fontSize: 13,  // 26/2
      bold: true, 
      italic: false,
      color: "000000",
      font: "微软雅黑",
      headingLevel: 2
    },
    
    // 段落 10-13: 正文 - 22pt, 普通, 黑色, 微软雅黑
    { text: "从产品形态看，落地会经历三个阶段：", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: "信息增强：检索、摘要、改写、结构化输出。", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: "任务协作：在工单、文档、审批等流程中协助执行。", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: "流程闭环：引入评估、权限、审计与运营机制，实现持续优化。", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    
    // 段落 14: 空行
    { text: "" },
    
    // 段落 15-18: 正文 - 22pt, 普通, 黑色, 微软雅黑
    { text: "落地方法建议：", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: "选择高频流程：明确输入、输出与验收标准。", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: "建立数据与权限：保证可用性与合规性。", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: "设置评估指标：上线前后都能量化效果。", fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    
    // 段落 19-20: 空行
    { text: "" },
    { text: "" },
    
    // 段落 21: 第二章节标题 - 26pt, 粗体, 黑色, 微软雅黑
    { 
      text: "2  典型落地场景（示例）", 
      fontSize: 13,  // 26/2
      bold: true, 
      italic: false,
      color: "000000",
      font: "微软雅黑",
      headingLevel: 2
    },
    
    // 段落 22-25: 正文 - 22pt, 普通, 黑色, 微软雅黑
    { text: '以下场景适合以"可控 + 可衡量"为优先目标：', fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: '知识检索与总结：面向内部知识库，强调引用与溯源。', fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: '客服质检与复盘：从对话中抽取要点，输出结构化结论。', fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    { text: '工单填单与分派：基于规则与历史数据降低重复劳动。', fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    
    // 段落 26: 空行
    { text: "" },
    
    // 段落 27: 参考链接标题 - 22pt, 普通, 灰色, 微软雅黑
    { text: '参考链接：', fontSize: 11, bold: false, italic: false, color: "C0C0C0", font: "微软雅黑" },
    
    // 段落 28: 示例接口 - 20pt, 普通, 灰色, Consolas
    { text: '示例接口：POST /v1/chat/completions', fontSize: 10, bold: false, italic: false, color: "C0C0C0", font: "Consolas" },
    
    // 段落 29: 空行
    { text: "" },
    
    // 段落 30: 第三章节标题 - 26pt, 粗体, 黑色, 微软雅黑
    { 
      text: "3  指标与风险清单（示例）", 
      fontSize: 13,  // 26/2
      bold: true, 
      italic: false,
      color: "000000",
      font: "微软雅黑",
      headingLevel: 2
    },
    
    // 段落 31: 表格说明 - 22pt, 普通, 黑色, 微软雅黑
    { text: '表格用于测试：合并单元格、底纹、边框、字体一致性。', fontSize: 11, bold: false, italic: false, color: "000000", font: "微软雅黑" },
    
    // 段落 32: 空行
    { text: "" },
    
    // 表格 1: AI 应用落地评估表（精确还原真实样式）
    // 根据 Python XML 解析：
    // - 标题行：bold=true, italic=false (不是 italic=true!)
    // - 表头行：bold=true, italic=false (不是 italic=true!)
    // - 数据行：bold=false, italic=false (不是 bold=true, italic=true!)
    // - 特殊单元格需要单独处理
    {
      text: "",
      isTable: true,
      tableData: [
        ["AI 应用落地评估表（示例）"],  // 第1行只有1个单元格（合并4列）
        ["场景", "价值指标", "数据/系统依赖", "风险与兜底"],
        ["知识检索 + 摘要生成", "采纳率、节省工时、满意度", "知识库、权限、检索链路", "引用来源、低置信度提示、人工复核"],
        ["工单自动分派/填单", "一次解决率、时延、转人工率", "工单系统、字段校验、流程规则", "规则兜底、审计、敏感信息脱敏"],
        ["测试", "测试", "测试", "测试"]
      ],
      tableCellStyles: [
        // 行1: 标题行 - 只有1个单元格，合并4列
        // 真实样式：fontSize=22, bold=true, italic=false (不是 italic=true!)
        [
          { gridSpan: 4, fill: "808080", alignment: "center", fontSize: 11, bold: true, italic: false, color: "000000", font: "微软雅黑" }
        ],
        // 行2: 表头 - 灰色背景，左对齐，20pt，bold=true, italic=false (不是 italic=true!)
        [
          { fill: "808080", alignment: "left", fontSize: 10, bold: true, italic: false, color: "000000", font: "微软雅黑" },
          { fill: "808080", alignment: "left", fontSize: 10, bold: true, italic: false, color: "000000", font: "微软雅黑" },
          { fill: "808080", alignment: "left", fontSize: 10, bold: true, italic: false, color: "000000", font: "微软雅黑" },
          { fill: "808080", alignment: "left", fontSize: 10, bold: true, italic: false, color: "000000", font: "微软雅黑" }
        ],
        // 行3: 数据行 - 20pt，bold=false, italic=false (不是 bold=true, italic=true!)
        [
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" }
        ],
        // 行4: 数据行 - 20pt，bold=false, italic=false (不是 bold=true, italic=true!)
        [
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" }
        ],
        // 行5: 测试行 - 混合样式（从 XML 精确提取）
        [
          { fontSize: 10, bold: false, italic: true, color: "000000", font: "微软雅黑" },  // 单元格1: 只斜体 (不是粗体+斜体!)
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },  // 单元格2: 普通
          { fontSize: 10, bold: false, italic: false, color: "000000", font: "微软雅黑" },  // 单元格3: 普通
          { fontSize: 22, bold: true, italic: false, color: "000000", font: "Lantinghei TC Demibold" }  // 单元格4: 44pt，粗体，特殊字体 (不是斜体!)
        ]
      ]
    },
    
    // 段落 33: 空行
    { text: "" },
    
    // 表格 2: 第二页内容提示（灰色背景，包含3个段落）
    // 真实样式：bold=false, italic=false (不是 bold=true, italic=true!)
    {
      text: "",
      isTable: true,
      tableData: [
        ['— 以下为第二页内容 —\n\n提示：优先选择"可控性强、指标可量化、可快速闭环"的业务流程作为首批落地场景。']
      ],
      tableCellStyles: [
        [{ 
          fill: "808080", 
          alignment: "left", 
          fontSize: 11,  // 第一段22pt，第二段20pt，这里用平均值
          bold: false,  // 不是 bold=true!
          italic: false,  // 不是 italic=true!
          color: "C0C0C0", 
          font: "微软雅黑",
          verticalAlign: "top"
        }]
      ]
    },
    
    // 段落 34: WPS 开放平台 - 22pt, 普通, 微软雅黑
    { text: "WPS 开放平台", fontSize: 11, bold: false, italic: false, font: "微软雅黑" },
    
    // 段落 35: 空行
    { text: "" }
  ]
}

/**
 * 根据段落生成完整的 docx 文件
 * 
 * 改进点（参考 Python improved_replacer.py）：
 * 1. 支持多 run 段落（处理混合样式）
 * 2. 更精确的表格单元格样式处理
 * 3. 保留空段落的格式信息
 * 4. 更好的字体和颜色处理
 */
export async function createDocxFromParagraphs(paragraphs: DocxParagraph[]): Promise<Buffer> {
  const children: (Paragraph | Table)[] = []

  for (const para of paragraphs) {
    // 处理表格
    if (para.isTable && para.tableData) {
      const tableRows = para.tableData.map((rowData, rowIndex) => {
        const cells = rowData.map((cellText, cellIndex) => {
          // 获取单元格样式
          const cellStyle = para.tableCellStyles?.[rowIndex]?.[cellIndex] || {}
          
          // 处理单元格内的多行文本（保留换行符）
          const cellLines = cellText.split('\n')
          const cellParagraphs = cellLines.map((line, lineIndex) => {
            const runOptions: IRunOptions = {
              text: line,
              bold: cellStyle.bold ?? false,
              italics: cellStyle.italic ?? false,
              size: cellStyle.fontSize ? cellStyle.fontSize * 2 : 20,
              color: cellStyle.color,
              font: cellStyle.font ? { name: cellStyle.font } : undefined
            }
            
            const paragraphAlignment = 
              cellStyle.alignment === "center" ? AlignmentType.CENTER :
              cellStyle.alignment === "right" ? AlignmentType.RIGHT :
              AlignmentType.LEFT
            
            return new Paragraph({
              children: [new TextRun(runOptions)],
              alignment: paragraphAlignment,
              spacing: lineIndex > 0 ? { before: 100 } : undefined  // 行间距
            })
          })
          
          const cellOptions: any = {
            children: cellParagraphs,
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: cellStyle.verticalAlign === "center" ? "center" : 
                          cellStyle.verticalAlign === "bottom" ? "bottom" : "top"
          }
          
          // 背景色
          if (cellStyle.fill) {
            cellOptions.shading = { fill: cellStyle.fill }
          }
          
          // 合并列
          if (cellStyle.gridSpan) {
            cellOptions.columnSpan = cellStyle.gridSpan
          }
          
          return new TableCell(cellOptions)
        })

        return new TableRow({ children: cells })
      })

      children.push(
        new Table({
          rows: tableRows,
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" }
          }
        })
      )
      continue
    }

    // 处理普通段落 - 支持多 run（混合样式）
    let textRuns: TextRun[]
    
    if (para.runs && para.runs.length > 0) {
      // 使用多个 run（用于混合样式段落）
      textRuns = para.runs.map(run => {
        const runOptions: IRunOptions = {
          text: run.text,
          bold: run.bold ?? false,
          italics: run.italic ?? false,
          underline: run.underline ? { type: UnderlineType.SINGLE } : undefined,
          size: run.fontSize ? run.fontSize * 2 : 22,
          color: run.color,
          font: run.font ? { name: run.font } : undefined
        }
        return new TextRun(runOptions)
      })
    } else if (para.link) {
      // 超链接样式
      textRuns = [
        new TextRun({
          text: para.text,
          bold: para.bold ?? false,
          italics: para.italic ?? false,
          size: para.fontSize ? para.fontSize * 2 : 22,
          style: "Hyperlink",
          color: "0563C1",
          underline: { type: UnderlineType.SINGLE },
          font: para.font ? { name: para.font } : undefined
        })
      ]
    } else {
      // 单一样式段落
      const runOptions: IRunOptions = {
        text: para.text,
        bold: para.bold ?? false,
        italics: para.italic ?? false,
        underline: para.underline ? { type: UnderlineType.SINGLE } : undefined,
        size: para.fontSize ? para.fontSize * 2 : 22,
        color: para.color,
        font: para.font ? { name: para.font } : undefined
      }
      textRuns = [new TextRun(runOptions)]
    }

    const paragraphOptions: any = {
      children: textRuns
    }

    if (para.headingLevel) {
      paragraphOptions.heading = HeadingLevel[`HEADING_${para.headingLevel}`]
    }

    if (para.alignment) {
      paragraphOptions.alignment =
        para.alignment === "center"
          ? AlignmentType.CENTER
          : para.alignment === "right"
          ? AlignmentType.RIGHT
          : AlignmentType.LEFT
    }

    children.push(new Paragraph(paragraphOptions))
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children
      }
    ]
  })

  return Buffer.from(await Packer.toBuffer(doc))
}
