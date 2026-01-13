import type { IncomingMessage, ServerResponse } from "node:http"
import { URL } from "node:url"
import { HttpError } from "@wps/doc-core"

export type Handler = (ctx: {
  req: IncomingMessage
  res: ServerResponse
  params: Record<string, string>
}) => Promise<void> | void

type Route = {
  method: string
  pattern: RegExp
  paramNames: string[]
  handler: Handler
}

function compilePath(path: string): { pattern: RegExp; paramNames: string[] } {
  const paramNames: string[] = []
  const pattern = path
    .split("/")
    .filter(Boolean)
    .map((seg) => {
      if (seg.startsWith(":")) {
        paramNames.push(seg.slice(1))
        return "([^/]+)"
      }
      return seg.replaceAll(/[.*+?^${}()|[\]\\]/g, "\\$&")
    })
    .join("\\/")

  return { pattern: new RegExp(`^\\/${pattern}\\/?$`), paramNames }
}

export class Router {
  private readonly routes: Route[] = []

  add(method: string, path: string, handler: Handler): void {
    const { pattern, paramNames } = compilePath(path)
    this.routes.push({ method: method.toUpperCase(), pattern, paramNames, handler })
  }

  async handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const method = (req.method ?? "GET").toUpperCase()
    const url = new URL(req.url ?? "/", "http://localhost")
    const pathname = url.pathname

    for (const route of this.routes) {
      if (route.method !== method) continue
      const match = route.pattern.exec(pathname)
      if (!match) continue

      const params: Record<string, string> = {}
      for (let i = 0; i < route.paramNames.length; i += 1) {
        params[route.paramNames[i]] = decodeURIComponent(match[i + 1] ?? "")
      }

      await route.handler({ req, res, params })
      return
    }

    throw new HttpError(404, "not_found")
  }
}
