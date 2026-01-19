import { xmlParser, type XmlNode } from '../core/xml-parser.js'

export interface CustomProperties {
  [key: string]: string | number | boolean | Date
}

export class CustomPropsParser {
  parse(xml: string): CustomProperties {
    const doc = xmlParser.parse(xml)
    const props = doc['Properties']
    if (!props) return {}
    
    const custom: CustomProperties = {}
    
    const propertyNodes = xmlParser.getChildren(props, 'property')
    for (const prop of propertyNodes) {
      const name = xmlParser.getAttr(prop, 'name')
      if (!name) continue
      
      // 尝试解析不同类型的值
      const lpwstr = prop['vt:lpwstr']
      const i4 = prop['vt:i4']
      const bool = prop['vt:bool']
      const filetime = prop['vt:filetime']
      
      if (lpwstr !== undefined) {
        custom[name] = String(lpwstr)
      } else if (i4 !== undefined) {
        custom[name] = Number(i4)
      } else if (bool !== undefined) {
        custom[name] = String(bool).toLowerCase() === 'true'
      } else if (filetime !== undefined) {
        const d = new Date(String(filetime))
        custom[name] = isNaN(d.getTime()) ? String(filetime) : d
      }
    }
    
    return custom
  }
}
