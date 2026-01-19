import { XMLBuilder, XMLParser } from "fast-xml-parser"

export interface ParseOptions {
  ignoreAttributes?: boolean
  preserveOrder?: boolean
  allowBooleanAttributes?: boolean
}

export interface BuildOptions {
  ignoreAttributes?: boolean
  preserveOrder?: boolean
  format?: boolean
}

/**
 * 解析 XML 字符串
 */
export function parseXml(xml: string, options: ParseOptions = {}): any {
  const parser = new XMLParser({
    ignoreAttributes: options.ignoreAttributes ?? false,
    preserveOrder: options.preserveOrder ?? false,
    allowBooleanAttributes: options.allowBooleanAttributes ?? false
  })
  return parser.parse(xml)
}

/**
 * 构建 XML 字符串
 */
export function buildXml(obj: any, options: BuildOptions = {}): string {
  const builder = new XMLBuilder({
    ignoreAttributes: options.ignoreAttributes ?? false,
    preserveOrder: options.preserveOrder ?? false,
    format: options.format ?? false
  })
  return builder.build(obj)
}
