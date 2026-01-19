import { XMLParser } from "fast-xml-parser"

/**
 * 文档核心属性（core.xml）
 */
export interface CoreProperties {
  title?: string
  subject?: string
  creator?: string
  keywords?: string
  description?: string
  lastModifiedBy?: string
  revision?: string
  created?: Date
  modified?: Date
  category?: string
  contentStatus?: string
}

/**
 * 应用程序属性（app.xml）
 */
export interface AppProperties {
  application?: string
  appVersion?: string
  totalTime?: number
  pages?: number
  words?: number
  characters?: number
  paragraphs?: number
  lines?: number
  company?: string
  manager?: string
}

/**
 * 自定义属性
 */
export interface CustomProperties {
  [key: string]: string | number | boolean | Date
}

/**
 * 完整的文档元数据
 */
export interface DocumentMetadata {
  core?: CoreProperties
  app?: AppProperties
  custom?: CustomProperties
}

function asArray<T>(v: T | T[] | undefined | null): T[] {
  if (!v) return []
  return Array.isArray(v) ? v : [v]
}

/**
 * 解析 core.xml（核心属性）
 */
export function parseCoreProperties(xml: string): CoreProperties {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  const coreProps = obj?.["cp:coreProperties"] || obj?.["coreProperties"]
  
  if (!coreProps) return {}
  
  const parseDate = (v: unknown): Date | undefined => {
    if (!v) return undefined
    const d = new Date(String(v))
    return isNaN(d.getTime()) ? undefined : d
  }
  
  return {
    title: coreProps["dc:title"] ? String(coreProps["dc:title"]) : undefined,
    subject: coreProps["dc:subject"] ? String(coreProps["dc:subject"]) : undefined,
    creator: coreProps["dc:creator"] ? String(coreProps["dc:creator"]) : undefined,
    keywords: coreProps["cp:keywords"] ? String(coreProps["cp:keywords"]) : undefined,
    description: coreProps["dc:description"] ? String(coreProps["dc:description"]) : undefined,
    lastModifiedBy: coreProps["cp:lastModifiedBy"] ? String(coreProps["cp:lastModifiedBy"]) : undefined,
    revision: coreProps["cp:revision"] ? String(coreProps["cp:revision"]) : undefined,
    created: parseDate(coreProps["dcterms:created"]),
    modified: parseDate(coreProps["dcterms:modified"]),
    category: coreProps["cp:category"] ? String(coreProps["cp:category"]) : undefined,
    contentStatus: coreProps["cp:contentStatus"] ? String(coreProps["cp:contentStatus"]) : undefined
  }
}

/**
 * 解析 app.xml（应用程序属性）
 */
export function parseAppProperties(xml: string): AppProperties {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  const props = obj?.["Properties"] || obj?.["ap:Properties"]
  
  if (!props) return {}
  
  const parseNum = (v: unknown): number | undefined => {
    const n = Number(v)
    return isNaN(n) ? undefined : n
  }
  
  return {
    application: props["Application"] ? String(props["Application"]) : undefined,
    appVersion: props["AppVersion"] ? String(props["AppVersion"]) : undefined,
    totalTime: parseNum(props["TotalTime"]),
    pages: parseNum(props["Pages"]),
    words: parseNum(props["Words"]),
    characters: parseNum(props["Characters"]),
    paragraphs: parseNum(props["Paragraphs"]),
    lines: parseNum(props["Lines"]),
    company: props["Company"] ? String(props["Company"]) : undefined,
    manager: props["Manager"] ? String(props["Manager"]) : undefined
  }
}

/**
 * 解析 custom.xml（自定义属性）
 */
export function parseCustomProperties(xml: string): CustomProperties {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  const props = obj?.["Properties"]
  
  if (!props) return {}
  
  const custom: CustomProperties = {}
  
  for (const prop of asArray<any>(props?.["property"])) {
    const name = prop?.["@_name"]
    if (!name || typeof name !== "string") continue
    
    // 尝试解析不同类型的值
    const lpwstr = prop?.["vt:lpwstr"]
    const i4 = prop?.["vt:i4"]
    const bool = prop?.["vt:bool"]
    const filetime = prop?.["vt:filetime"]
    
    if (lpwstr !== undefined) {
      custom[name] = String(lpwstr)
    } else if (i4 !== undefined) {
      custom[name] = Number(i4)
    } else if (bool !== undefined) {
      custom[name] = String(bool).toLowerCase() === "true"
    } else if (filetime !== undefined) {
      const d = new Date(String(filetime))
      custom[name] = isNaN(d.getTime()) ? String(filetime) : d
    }
  }
  
  return custom
}

/**
 * 解析页眉内容
 */
export interface HeaderFooterContent {
  id: string
  paragraphs: string[]
}

export function parseHeaderFooter(xml: string): HeaderFooterContent {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  
  const hdr = obj?.["w:hdr"] || obj?.["w:ftr"]
  if (!hdr) return { id: "", paragraphs: [] }
  
  const paragraphs: string[] = []
  
  for (const p of asArray<any>(hdr?.["w:p"])) {
    let text = ""
    for (const r of asArray<any>(p?.["w:r"])) {
      const t = r?.["w:t"]
      if (t) text += String(t)
    }
    if (text) paragraphs.push(text)
  }
  
  return { id: "", paragraphs }
}

/**
 * 解析注释内容
 */
export interface Comment {
  id: string
  author?: string
  date?: Date
  initials?: string
  content: string
}

export function parseComments(xml: string): Map<string, Comment> {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  
  const comments = new Map<string, Comment>()
  const commentsRoot = obj?.["w:comments"]
  
  for (const comment of asArray<any>(commentsRoot?.["w:comment"])) {
    const id = String(comment?.["@_w:id"] ?? "")
    if (!id) continue
    
    const author = comment?.["@_w:author"]
    const date = comment?.["@_w:date"]
    const initials = comment?.["@_w:initials"]
    
    let content = ""
    for (const p of asArray<any>(comment?.["w:p"])) {
      for (const r of asArray<any>(p?.["w:r"])) {
        const t = r?.["w:t"]
        if (t) content += String(t)
      }
    }
    
    comments.set(id, {
      id,
      author: typeof author === "string" ? author : undefined,
      date: date ? new Date(String(date)) : undefined,
      initials: typeof initials === "string" ? initials : undefined,
      content
    })
  }
  
  return comments
}

/**
 * 解析脚注/尾注内容
 */
export interface Note {
  id: string
  type: "footnote" | "endnote"
  content: string
}

export function parseNotes(xml: string, type: "footnote" | "endnote"): Map<string, Note> {
  const parser = new XMLParser({ ignoreAttributes: false })
  const obj: any = parser.parse(xml)
  
  const notes = new Map<string, Note>()
  const notesRoot = type === "footnote" ? obj?.["w:footnotes"] : obj?.["w:endnotes"]
  const noteTag = type === "footnote" ? "w:footnote" : "w:endnote"
  
  for (const note of asArray<any>(notesRoot?.[noteTag])) {
    const id = String(note?.["@_w:id"] ?? "")
    if (!id) continue
    
    let content = ""
    for (const p of asArray<any>(note?.["w:p"])) {
      for (const r of asArray<any>(p?.["w:r"])) {
        const t = r?.["w:t"]
        if (t) content += String(t)
      }
    }
    
    notes.set(id, { id, type, content })
  }
  
  return notes
}
