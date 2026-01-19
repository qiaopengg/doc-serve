/**
 * 图表解析器
 * 解析 Excel 图表（Chart）
 */

import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface Chart {
  id: string
  type: ChartType
  title?: string
  data: ChartData
  style?: ChartStyle
  legend?: ChartLegend
  axes?: ChartAxes
}

export type ChartType = 
  | 'bar' | 'barStacked' | 'barPercentStacked'
  | 'column' | 'columnStacked' | 'columnPercentStacked'
  | 'line' | 'lineMarkers' | 'lineStacked' | 'linePercentStacked'
  | 'pie' | 'pie3D' | 'doughnut'
  | 'area' | 'areaStacked' | 'areaPercentStacked'
  | 'scatter' | 'scatterLines' | 'scatterSmooth'
  | 'bubble'
  | 'radar' | 'radarMarkers' | 'radarFilled'
  | 'stock'
  | 'surface' | 'surface3D'

export interface ChartData {
  series: ChartSeries[]
  categories?: string[]
}

export interface ChartSeries {
  name: string
  values: number[]
  color?: string
  marker?: ChartMarker
}

export interface ChartMarker {
  symbol?: 'circle' | 'square' | 'diamond' | 'triangle' | 'x' | 'star' | 'dash' | 'dot'
  size?: number
  fill?: string
  line?: string
}

export interface ChartStyle {
  colorScheme?: string
  font?: string
  fontSize?: number
}

export interface ChartLegend {
  position?: 'top' | 'bottom' | 'left' | 'right' | 'topRight'
  overlay?: boolean
}

export interface ChartAxes {
  category?: ChartAxis
  value?: ChartAxis
}

export interface ChartAxis {
  title?: string
  min?: number
  max?: number
  majorUnit?: number
  minorUnit?: number
  format?: string
}

export class ChartParser {
  private idCounter = 0
  
  /**
   * 解析文档中的所有图表
   */
  parse(xml: string): Map<string, Chart> {
    const charts = new Map<string, Chart>()
    const doc = xmlParser.parse(xml)
    
    this.extractCharts(doc, charts)
    return charts
  }
  
  /**
   * 解析单个图表
   */
  parseChart(chartNode: XmlNode, chartXml?: string): Chart | null {
    if (!chartNode) return null
    
    // 如果有独立的图表 XML，解析它
    if (chartXml) {
      return this.parseChartXml(chartXml)
    }
    
    // 否则从嵌入的图表节点解析
    const chart: Chart = {
      id: this.generateId(),
      type: 'bar', // 默认类型
      data: { series: [] }
    }
    
    // 解析图表类型
    chart.type = this.detectChartType(chartNode)
    
    // 解析图表标题
    chart.title = this.parseChartTitle(chartNode)
    
    // 解析图表数据
    chart.data = this.parseChartData(chartNode)
    
    // 解析图表样式
    chart.style = this.parseChartStyle(chartNode)
    
    // 解析图例
    chart.legend = this.parseChartLegend(chartNode)
    
    // 解析坐标轴
    chart.axes = this.parseChartAxes(chartNode)
    
    return chart
  }
  
  /**
   * 解析独立的图表 XML 文件
   */
  private parseChartXml(xml: string): Chart | null {
    try {
      const doc = xmlParser.parse(xml)
      const chartSpace = doc['c:chartSpace']
      if (!chartSpace) return null
      
      const chart = chartSpace['c:chart']
      if (!chart) return null
      
      const result: Chart = {
        id: this.generateId(),
        type: 'bar',
        data: { series: [] }
      }
      
      // 解析图表标题
      if (chart['c:title']) {
        result.title = this.extractTextFromTitle(chart['c:title'])
      }
      
      // 解析绘图区域
      const plotArea = chart['c:plotArea']
      if (plotArea) {
        // 检测图表类型并解析数据
        result.type = this.detectChartTypeFromPlotArea(plotArea)
        result.data = this.parseChartDataFromPlotArea(plotArea, result.type)
        
        // 解析坐标轴
        result.axes = this.parseAxesFromPlotArea(plotArea)
      }
      
      // 解析图例
      if (chart['c:legend']) {
        result.legend = this.parseLegendNode(chart['c:legend'])
      }
      
      return result
    } catch (error) {
      console.error('Failed to parse chart XML:', error)
      return null
    }
  }
  
  /**
   * 检测图表类型
   */
  private detectChartType(chartNode: XmlNode): ChartType {
    // 从节点属性或子节点推断图表类型
    // 这是简化实现，实际需要根据 OOXML 规范解析
    return 'bar'
  }
  
  /**
   * 从绘图区域检测图表类型
   */
  private detectChartTypeFromPlotArea(plotArea: XmlNode): ChartType {
    // 检查各种图表类型节点
    if (plotArea['c:barChart']) return 'bar'
    if (plotArea['c:bar3DChart']) return 'bar'
    if (plotArea['c:lineChart']) return 'line'
    if (plotArea['c:line3DChart']) return 'line'
    if (plotArea['c:pieChart']) return 'pie'
    if (plotArea['c:pie3DChart']) return 'pie3D'
    if (plotArea['c:doughnutChart']) return 'doughnut'
    if (plotArea['c:areaChart']) return 'area'
    if (plotArea['c:area3DChart']) return 'area'
    if (plotArea['c:scatterChart']) return 'scatter'
    if (plotArea['c:bubbleChart']) return 'bubble'
    if (plotArea['c:radarChart']) return 'radar'
    if (plotArea['c:stockChart']) return 'stock'
    if (plotArea['c:surfaceChart']) return 'surface'
    if (plotArea['c:surface3DChart']) return 'surface3D'
    
    return 'bar' // 默认
  }
  
  /**
   * 解析图表标题
   */
  private parseChartTitle(chartNode: XmlNode): string | undefined {
    const title = chartNode['c:title']
    if (!title) return undefined
    
    return this.extractTextFromTitle(title)
  }
  
  /**
   * 从标题节点提取文本
   */
  private extractTextFromTitle(titleNode: XmlNode): string {
    const tx = titleNode['c:tx']
    if (!tx) return ''
    
    const rich = tx['c:rich']
    if (rich) {
      return this.extractTextFromRich(rich)
    }
    
    const strRef = tx['c:strRef']
    if (strRef) {
      const strCache = strRef['c:strCache']
      if (strCache) {
        const pt = strCache['c:pt']
        if (pt) {
          const v = pt['c:v']
          return v ? String(v) : ''
        }
      }
    }
    
    return ''
  }
  
  /**
   * 从富文本节点提取文本
   */
  private extractTextFromRich(richNode: XmlNode): string {
    const texts: string[] = []
    
    const paragraphs = xmlParser.getChildren(richNode, 'a:p')
    for (const p of paragraphs) {
      const runs = xmlParser.getChildren(p, 'a:r')
      for (const r of runs) {
        const t = r['a:t']
        if (t) {
          texts.push(String(t))
        }
      }
    }
    
    return texts.join(' ')
  }
  
  /**
   * 解析图表数据
   */
  private parseChartData(chartNode: XmlNode): ChartData {
    // 简化实现
    return {
      series: [],
      categories: []
    }
  }
  
  /**
   * 从绘图区域解析图表数据
   */
  private parseChartDataFromPlotArea(plotArea: XmlNode, chartType: ChartType): ChartData {
    const data: ChartData = {
      series: [],
      categories: []
    }
    
    // 根据图表类型选择正确的节点
    const chartTypeNode = this.getChartTypeNode(plotArea, chartType)
    if (!chartTypeNode) return data
    
    // 解析系列
    const serNodes = xmlParser.getChildren(chartTypeNode, 'c:ser')
    for (const ser of serNodes) {
      const series = this.parseSeries(ser)
      if (series) {
        data.series.push(series)
      }
    }
    
    // 解析分类（从第一个系列获取）
    if (serNodes.length > 0) {
      data.categories = this.parseCategories(serNodes[0])
    }
    
    return data
  }
  
  /**
   * 获取图表类型节点
   */
  private getChartTypeNode(plotArea: XmlNode, chartType: ChartType): XmlNode | undefined {
    const typeMap: Record<string, string> = {
      'bar': 'c:barChart',
      'column': 'c:barChart',
      'line': 'c:lineChart',
      'pie': 'c:pieChart',
      'pie3D': 'c:pie3DChart',
      'area': 'c:areaChart',
      'scatter': 'c:scatterChart',
      'bubble': 'c:bubbleChart',
      'radar': 'c:radarChart',
      'stock': 'c:stockChart',
      'surface': 'c:surfaceChart',
    }
    
    const nodeName = typeMap[chartType] || 'c:barChart'
    return plotArea[nodeName]
  }
  
  /**
   * 解析系列
   */
  private parseSeries(serNode: XmlNode): ChartSeries | null {
    const series: ChartSeries = {
      name: '',
      values: []
    }
    
    // 解析系列名称
    const tx = serNode['c:tx']
    if (tx) {
      series.name = this.extractSeriesName(tx)
    }
    
    // 解析系列值
    const val = serNode['c:val'] || serNode['c:yVal']
    if (val) {
      series.values = this.extractSeriesValues(val)
    }
    
    // 解析标记
    const marker = serNode['c:marker']
    if (marker) {
      series.marker = this.parseMarker(marker)
    }
    
    return series
  }
  
  /**
   * 提取系列名称
   */
  private extractSeriesName(txNode: XmlNode): string {
    const strRef = txNode['c:strRef']
    if (strRef) {
      const strCache = strRef['c:strCache']
      if (strCache) {
        const pt = strCache['c:pt']
        if (pt) {
          const v = pt['c:v']
          return v ? String(v) : ''
        }
      }
    }
    
    const v = txNode['c:v']
    return v ? String(v) : ''
  }
  
  /**
   * 提取系列值
   */
  private extractSeriesValues(valNode: XmlNode): number[] {
    const values: number[] = []
    
    const numRef = valNode['c:numRef']
    if (numRef) {
      const numCache = numRef['c:numCache']
      if (numCache) {
        const pts = xmlParser.getChildren(numCache, 'c:pt')
        for (const pt of pts) {
          const v = pt['c:v']
          if (v) {
            const num = parseFloat(String(v))
            if (!isNaN(num)) {
              values.push(num)
            }
          }
        }
      }
    }
    
    return values
  }
  
  /**
   * 解析分类
   */
  private parseCategories(serNode: XmlNode): string[] {
    const categories: string[] = []
    
    const cat = serNode['c:cat'] || serNode['c:xVal']
    if (!cat) return categories
    
    const strRef = cat['c:strRef']
    if (strRef) {
      const strCache = strRef['c:strCache']
      if (strCache) {
        const pts = xmlParser.getChildren(strCache, 'c:pt')
        for (const pt of pts) {
          const v = pt['c:v']
          if (v) {
            categories.push(String(v))
          }
        }
      }
    }
    
    return categories
  }
  
  /**
   * 解析标记
   */
  private parseMarker(markerNode: XmlNode): ChartMarker {
    const marker: ChartMarker = {}
    
    const symbol = markerNode['c:symbol']
    if (symbol) {
      marker.symbol = xmlParser.getAttr(symbol, 'val') as any
    }
    
    const size = markerNode['c:size']
    if (size) {
      marker.size = xmlParser.parseInt(xmlParser.getAttr(size, 'val'))
    }
    
    return marker
  }
  
  /**
   * 解析图表样式
   */
  private parseChartStyle(chartNode: XmlNode): ChartStyle | undefined {
    // 简化实现
    return undefined
  }
  
  /**
   * 解析图例
   */
  private parseChartLegend(chartNode: XmlNode): ChartLegend | undefined {
    const legend = chartNode['c:legend']
    if (!legend) return undefined
    
    return this.parseLegendNode(legend)
  }
  
  /**
   * 解析图例节点
   */
  private parseLegendNode(legendNode: XmlNode): ChartLegend {
    const legend: ChartLegend = {}
    
    const legendPos = legendNode['c:legendPos']
    if (legendPos) {
      const val = xmlParser.getAttr(legendPos, 'val')
      legend.position = val as any
    }
    
    const overlay = legendNode['c:overlay']
    if (overlay) {
      legend.overlay = xmlParser.getAttr(overlay, 'val') === '1'
    }
    
    return legend
  }
  
  /**
   * 解析坐标轴
   */
  private parseChartAxes(chartNode: XmlNode): ChartAxes | undefined {
    // 简化实现
    return undefined
  }
  
  /**
   * 从绘图区域解析坐标轴
   */
  private parseAxesFromPlotArea(plotArea: XmlNode): ChartAxes {
    const axes: ChartAxes = {}
    
    // 解析分类轴
    const catAx = plotArea['c:catAx']
    if (catAx) {
      axes.category = this.parseAxis(catAx)
    }
    
    // 解析数值轴
    const valAx = plotArea['c:valAx']
    if (valAx) {
      axes.value = this.parseAxis(valAx)
    }
    
    return axes
  }
  
  /**
   * 解析坐标轴
   */
  private parseAxis(axisNode: XmlNode): ChartAxis {
    const axis: ChartAxis = {}
    
    // 解析标题
    const title = axisNode['c:title']
    if (title) {
      axis.title = this.extractTextFromTitle(title)
    }
    
    // 解析缩放
    const scaling = axisNode['c:scaling']
    if (scaling) {
      const min = scaling['c:min']
      if (min) {
        axis.min = parseFloat(xmlParser.getAttr(min, 'val') || '0')
      }
      
      const max = scaling['c:max']
      if (max) {
        axis.max = parseFloat(xmlParser.getAttr(max, 'val') || '0')
      }
    }
    
    return axis
  }
  
  /**
   * 递归提取所有图表
   */
  private extractCharts(node: any, charts: Map<string, Chart>): void {
    if (!node || typeof node !== 'object') return
    
    // 处理图表节点
    if (node['c:chart']) {
      const chartArray = Array.isArray(node['c:chart'])
        ? node['c:chart']
        : [node['c:chart']]
      
      for (const chart of chartArray) {
        const parsed = this.parseChart(chart)
        if (parsed) {
          charts.set(parsed.id, parsed)
        }
      }
    }
    
    // 递归处理子节点
    for (const key in node) {
      if (typeof node[key] === 'object') {
        this.extractCharts(node[key], charts)
      }
    }
  }
  
  /**
   * 生成唯一 ID
   */
  private generateId(): string {
    return `chart_${++this.idCounter}`
  }
}

export const chartParser = new ChartParser()
