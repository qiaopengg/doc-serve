/**
 * 通用工具函数模块
 * 
 * 本模块提供从 docxGenerator.ts 迁移的工具函数集合，用于处理 OOXML 文档解析中的常见操作。
 * 包括：
 * - 数组和类型转换工具
 * - 颜色标准化工具
 * - 对齐方式标准化
 * - XML 节点操作工具
 * - 对象合并工具
 * - 样式检测工具
 * 
 * @module parser/core/utils
 */

import type { OrderedXmlNode } from "../types.js"

// ============= 数组和类型转换工具 =============

/**
 * 将值转换为数组
 * 
 * 这是一个类型安全的工具函数，用于确保返回值始终是数组。
 * 常用于处理 XML 解析结果，因为单个元素和多个元素的表示方式不同。
 * 
 * @template T - 数组元素的类型
 * @param v - 可能是单个值、数组或 undefined/null
 * @returns 始终返回数组：
 *          - 如果输入是 null/undefined，返回空数组
 *          - 如果输入是数组，直接返回
 *          - 如果输入是单个值，包装成数组返回
 * 
 * @example
 * ```typescript
 * asArray(undefined)        // []
 * asArray(null)             // []
 * asArray('hello')          // ['hello']
 * asArray(['a', 'b'])       // ['a', 'b']
 * ```
 */
export function asArray<T>(v: T | T[] | undefined | null): T[] {
  if (!v) return []
  return Array.isArray(v) ? v : [v]
}

/**
 * 解析 OOXML 布尔值（on/off 格式）
 * 
 * OOXML 规范中的布尔值可以用多种方式表示：
 * - "true"/"false"
 * - "on"/"off"
 * - "1"/"0"
 * - 属性存在但无值（默认为 true）
 * 
 * @param v - 布尔值或字符串，可以是任何类型
 * @param defaultWhenMissingVal - 当值为 null/undefined 时的默认值
 *                                如果为 true，则缺失值被视为 true
 *                                如果未提供，则缺失值返回 undefined
 * @returns 解析后的布尔值，或 undefined（如果无法解析）
 * 
 * @example
 * ```typescript
 * parseBooleanOnOff("on")              // true
 * parseBooleanOnOff("off")             // false
 * parseBooleanOnOff("1")               // true
 * parseBooleanOnOff("0")               // false
 * parseBooleanOnOff(null)              // undefined
 * parseBooleanOnOff(null, true)        // true
 * parseBooleanOnOff("false")           // false
 * parseBooleanOnOff("none")            // false
 * ```
 */
export function parseBooleanOnOff(v: unknown, defaultWhenMissingVal?: true): boolean | undefined {
  if (v == null) return defaultWhenMissingVal ? true : undefined
  const s = String(v).trim().toLowerCase()
  if (s === "0" || s === "false" || s === "off" || s === "none") return false
  return true
}

// ============= 颜色标准化工具 =============

/**
 * 标准化颜色值（支持 6 位和 8 位十六进制）
 * 
 * OOXML 文档中的颜色可以用多种格式表示：
 * - 6 位十六进制（RGB）：如 "FF0000"
 * - 8 位十六进制（ARGB）：如 "80FF0000"（前两位是 alpha 通道）
 * - "auto" 关键字（自动颜色）
 * 
 * 此函数将所有有效颜色标准化为 6 位十六进制格式（大写，无 # 前缀）。
 * 对于 8 位颜色，会去掉 alpha 通道，只保留 RGB 部分。
 * 
 * @param v - 颜色值，可以是任何类型
 * @returns 标准化的 6 位十六进制颜色（大写，无 # 前缀），
 *          如果输入无效或为 "auto"，返回 undefined
 * 
 * @example
 * ```typescript
 * normalizeColor("FF0000")      // "FF0000"
 * normalizeColor("#ff0000")     // "FF0000"
 * normalizeColor("80FF0000")    // "FF0000" (去掉 alpha)
 * normalizeColor("auto")        // undefined
 * normalizeColor("")            // undefined
 * normalizeColor(null)          // undefined
 * ```
 */
export function normalizeColor(v: unknown): string | undefined {
  if (!v) return undefined
  const s = String(v).trim()
  if (!s || s.toLowerCase() === "auto") return undefined
  const hex = s.replace(/^#/, "").toUpperCase()
  // 支持 6 位和 8 位十六进制颜色
  if (/^[0-9A-F]{6}$/.test(hex)) return hex
  if (/^[0-9A-F]{8}$/.test(hex)) return hex.substring(2, 8) // 去掉 alpha 通道（前2位）
  return undefined
}

/**
 * 标准化边框颜色值
 * 
 * 边框颜色的处理与普通颜色略有不同：
 * - 支持 "auto" 关键字（保留为 "auto"）
 * - 只支持 6 位十六进制格式（不支持 8 位）
 * 
 * @param v - 颜色值，可以是任何类型
 * @returns 标准化的颜色值：
 *          - "auto" 如果输入是 "auto"
 *          - 6 位十六进制颜色（大写，无 # 前缀）
 *          - undefined 如果输入无效
 * 
 * @example
 * ```typescript
 * normalizeBorderColor("FF0000")    // "FF0000"
 * normalizeBorderColor("#ff0000")   // "FF0000"
 * normalizeBorderColor("auto")      // "auto"
 * normalizeBorderColor("AUTO")      // "auto"
 * normalizeBorderColor("")          // undefined
 * normalizeBorderColor("invalid")   // undefined
 * ```
 */
export function normalizeBorderColor(v: unknown): string | undefined {
  if (!v) return undefined
  const s = String(v).trim()
  if (!s) return undefined
  if (s.toLowerCase() === "auto") return "auto"
  const hex = s.replace(/^#/, "").toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(hex)) return undefined
  return hex
}

// ============= 对齐方式标准化 =============

/**
 * 标准化对齐方式
 * 
 * OOXML 中的对齐方式有多种表示方法，此函数将它们统一为标准格式。
 * 支持的输入格式：
 * - "left" / "start" → "left"
 * - "center" → "center"
 * - "right" / "end" → "right"
 * - "both" / "justify" → "justify"
 * 
 * @param v - 对齐方式值，可以是任何类型
 * @returns 标准化的对齐方式，如果输入无效则返回 undefined
 * 
 * @example
 * ```typescript
 * normalizeAlignment("left")      // "left"
 * normalizeAlignment("start")     // "left"
 * normalizeAlignment("center")    // "center"
 * normalizeAlignment("right")     // "right"
 * normalizeAlignment("end")       // "right"
 * normalizeAlignment("both")      // "justify"
 * normalizeAlignment("justify")   // "justify"
 * normalizeAlignment("")          // undefined
 * normalizeAlignment("invalid")   // undefined
 * ```
 */
export function normalizeAlignment(v: unknown): "left" | "center" | "right" | "justify" | undefined {
  const s = String(v ?? "").trim().toLowerCase()
  if (!s) return undefined
  if (s === "center") return "center"
  if (s === "right" || s === "end") return "right"
  if (s === "left" || s === "start") return "left"
  if (s === "both" || s === "justify") return "justify"
  return undefined
}

// ============= XML 节点操作工具 =============

/**
 * 获取 XML 节点的标签名
 * 
 * 在 fast-xml-parser 的 preserveOrder 模式下，每个节点是一个对象，
 * 其键名就是标签名。此函数提取该标签名。
 * 
 * 特殊处理：
 * - 跳过 ":@" 键（属性对象）
 * - 跳过 "#text" 键（除非它是唯一的键）
 * 
 * @param node - 有序 XML 节点对象
 * @returns 标签名，如果是纯文本节点则返回 "#text"，如果无法确定则返回 undefined
 * 
 * @example
 * ```typescript
 * tagNameOf({ "w:p": [...] })              // "w:p"
 * tagNameOf({ "#text": "hello" })          // "#text"
 * tagNameOf({ "w:r": [...], ":@": {...} }) // "w:r"
 * ```
 */
export function tagNameOf(node: OrderedXmlNode): string | undefined {
  const keys = Object.keys(node)
  for (const k of keys) {
    if (k === ":@" || k === "#text") continue
    return k
  }
  if (keys.includes("#text")) return "#text"
  return undefined
}

/**
 * 获取 XML 节点的属性对象
 * 
 * 在 fast-xml-parser 的 preserveOrder 模式下，节点的属性存储在 ":@" 键中。
 * 此函数安全地提取属性对象。
 * 
 * @param node - 有序 XML 节点对象
 * @returns 属性对象（键值对），如果没有属性则返回空对象
 * 
 * @example
 * ```typescript
 * attrsOf({ "w:p": [...], ":@": { "@_id": "1" } })  // { "@_id": "1" }
 * attrsOf({ "w:p": [...] })                         // {}
 * ```
 */
export function attrsOf(node: OrderedXmlNode): Record<string, any> {
  const attrs = node[":@"]
  if (!attrs || typeof attrs !== "object") return {}
  return attrs
}

/**
 * 获取节点的指定属性值
 * 
 * fast-xml-parser 会在属性名前添加 "@_" 前缀，此函数会尝试两种格式：
 * 1. 直接使用属性名
 * 2. 使用 "@_" 前缀的属性名
 * 
 * @param attrs - 属性对象（通常来自 attrsOf()）
 * @param name - 属性名（不带前缀）
 * @returns 属性值，如果不存在则返回 undefined
 * 
 * @example
 * ```typescript
 * const attrs = { "@_id": "1", "val": "test" }
 * attrOf(attrs, "id")    // "1"
 * attrOf(attrs, "val")   // "test"
 * attrOf(attrs, "foo")   // undefined
 * ```
 */
export function attrOf(attrs: Record<string, any>, name: string): any {
  if (name in attrs) return attrs[name]
  const alt = `@_${name}`
  if (alt in attrs) return attrs[alt]
  return undefined
}

/**
 * 获取 XML 节点的所有子节点
 * 
 * 在 fast-xml-parser 的 preserveOrder 模式下，子节点存储在标签名对应的数组中。
 * 此函数提取该数组。
 * 
 * @param node - 有序 XML 节点对象
 * @returns 子节点数组，如果没有子节点或节点是文本节点则返回空数组
 * 
 * @example
 * ```typescript
 * childrenOf({ "w:p": [{ "w:r": [...] }, { "w:r": [...] }] })  // [{ "w:r": [...] }, { "w:r": [...] }]
 * childrenOf({ "#text": "hello" })                             // []
 * ```
 */
export function childrenOf(node: OrderedXmlNode): OrderedXmlNode[] {
  const tn = tagNameOf(node)
  if (!tn || tn === "#text") return []
  const v = node[tn]
  return Array.isArray(v) ? v : []
}

/**
 * 获取节点的第一个指定名称的子节点
 * 
 * 在子节点列表中查找第一个匹配指定标签名的节点。
 * 
 * @param node - 有序 XML 节点对象
 * @param name - 要查找的子节点标签名（如 "w:r", "w:t" 等）
 * @returns 第一个匹配的子节点，如果未找到则返回 undefined
 * 
 * @example
 * ```typescript
 * const para = { "w:p": [{ "w:pPr": [...] }, { "w:r": [...] }] }
 * childOf(para, "w:pPr")  // { "w:pPr": [...] }
 * childOf(para, "w:r")    // { "w:r": [...] }
 * childOf(para, "w:tbl")  // undefined
 * ```
 */
export function childOf(node: OrderedXmlNode, name: string): OrderedXmlNode | undefined {
  for (const c of childrenOf(node)) {
    if (tagNameOf(c) === name) return c
  }
  return undefined
}

/**
 * 获取节点的所有指定名称的子节点
 * 
 * 在子节点列表中查找所有匹配指定标签名的节点。
 * 
 * @param node - 有序 XML 节点对象
 * @param name - 要查找的子节点标签名（如 "w:r", "w:t" 等）
 * @returns 所有匹配的子节点数组，如果未找到则返回空数组
 * 
 * @example
 * ```typescript
 * const para = { "w:p": [{ "w:r": [...] }, { "w:r": [...] }, { "w:pPr": [...] }] }
 * childrenNamed(para, "w:r")    // [{ "w:r": [...] }, { "w:r": [...] }]
 * childrenNamed(para, "w:pPr")  // [{ "w:pPr": [...] }]
 * childrenNamed(para, "w:tbl")  // []
 * ```
 */
export function childrenNamed(node: OrderedXmlNode, name: string): OrderedXmlNode[] {
  return childrenOf(node).filter((c) => tagNameOf(c) === name)
}

/**
 * 从有序 XML 节点递归提取文本内容
 * 
 * 此函数递归遍历 XML 节点树，提取所有文本内容。
 * 特殊处理 OOXML 的特殊元素：
 * - w:t - 文本节点
 * - w:tab - 制表符（转换为 \t）
 * - w:br / w:cr - 换行符（转换为 \n）
 * - #text - 纯文本节点
 * 
 * @param node - 有序 XML 节点对象
 * @returns 提取的文本内容，多个文本片段会连接在一起
 * 
 * @example
 * ```typescript
 * textFromOrdered({ "#text": "hello" })                    // "hello"
 * textFromOrdered({ "w:t": [{ "#text": "world" }] })       // "world"
 * textFromOrdered({ "w:tab": [] })                         // "\t"
 * textFromOrdered({ "w:br": [] })                          // "\n"
 * ```
 */
export function textFromOrdered(node: OrderedXmlNode): string {
  const tn = tagNameOf(node)
  if (tn === "#text") return String((node as any)["#text"] ?? "")

  if (tn === "w:t") {
    const children = childrenOf(node)
    if (children.length === 0) return ""
    return children.map(textFromOrdered).join("")
  }

  if (tn === "w:tab") return "\t"
  if (tn === "w:br" || tn === "w:cr") return "\n"

  return childrenOf(node).map(textFromOrdered).join("")
}

// ============= 对象合并工具 =============

/**
 * 合并多个对象，只保留已定义的值
 * 
 * 此函数用于合并多个部分对象，创建一个完整的对象。
 * 与 Object.assign 不同，此函数会跳过 undefined 值，避免覆盖已有的有效值。
 * 
 * 合并规则：
 * - 跳过 null 或 undefined 的对象
 * - 跳过值为 undefined 的属性
 * - 后面的对象会覆盖前面对象的同名属性（如果值不是 undefined）
 * 
 * @template T - 对象类型，必须是记录类型
 * @param parts - 要合并的对象数组，可以包含 undefined
 * @returns 合并后的对象，包含所有已定义的属性
 * 
 * @example
 * ```typescript
 * mergeDefined({ a: 1 }, { b: 2 })                    // { a: 1, b: 2 }
 * mergeDefined({ a: 1 }, { a: 2, b: 3 })              // { a: 2, b: 3 }
 * mergeDefined({ a: 1 }, undefined, { b: 2 })         // { a: 1, b: 2 }
 * mergeDefined({ a: 1 }, { a: undefined, b: 2 })      // { a: 1, b: 2 }
 * ```
 */
export function mergeDefined<T extends Record<string, any>>(...parts: Array<T | undefined>): T {
  const out: any = {}
  for (const p of parts) {
    if (!p) continue
    for (const [k, v] of Object.entries(p)) {
      if (v !== undefined) out[k] = v
    }
  }
  return out
}

// ============= 样式检测工具 =============

/**
 * 从样式 ID 或样式名称检测标题级别
 * 
 * OOXML 文档中的标题样式通常包含 "Heading" 或 "heading" 字样，后跟数字 1-6。
 * 此函数尝试从样式标识符中提取标题级别。
 * 
 * 支持的格式：
 * - "Heading1", "Heading 1", "heading_1"
 * - "Heading2", "Heading 2", "heading_2"
 * - 等等（1-6 级）
 * 
 * @param styleId - 样式 ID（如 "Heading1"）
 * @param styleName - 样式名称（如 "Heading 1"）
 * @returns 标题级别（1-6），如果不是标题样式则返回 undefined
 * 
 * @example
 * ```typescript
 * detectHeadingLevel("Heading1", undefined)           // 1
 * detectHeadingLevel("Heading 2", undefined)          // 2
 * detectHeadingLevel("heading_3", undefined)          // 3
 * detectHeadingLevel(undefined, "Heading 4")          // 4
 * detectHeadingLevel("Normal", undefined)             // undefined
 * detectHeadingLevel("Heading7", undefined)           // undefined (超出范围)
 * ```
 */
export function detectHeadingLevel(styleId: string | undefined, styleName: string | undefined): 1 | 2 | 3 | 4 | 5 | 6 | undefined {
  const s = String(styleId ?? styleName ?? "").toLowerCase()
  if (!s) return undefined
  const m = s.match(/heading\s*([1-6])/i) || s.match(/heading_([1-6])/i) || s.match(/heading([1-6])/i)
  if (!m) return undefined
  const n = Number.parseInt(m[1]!, 10)
  if (n >= 1 && n <= 6) return n as any
  return undefined
}
