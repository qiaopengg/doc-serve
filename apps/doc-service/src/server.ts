import { createServer } from "node:http"
import type { Server } from "node:http"
import { HttpError } from "@wps/doc-core"
import { Router } from "./http/router.js"
import { registerDocRoutes } from "./routes/docs.js"
import type { DocStore } from "./docStore/types.js"

export type BuildServerDeps = {
  docStore: DocStore
  corsOrigin?: string
}

function sendJson(res: import("node:http").ServerResponse, statusCode: number, body: unknown): void {
  const json = JSON.stringify(body)
  res.statusCode = statusCode
  res.setHeader("Content-Type", "application/json; charset=utf-8")
  res.setHeader("Content-Length", Buffer.byteLength(json))
  res.end(json)
}

export function buildServer(deps: BuildServerDeps): Server {
  const router = new Router()
  registerDocRoutes(router, { docStore: deps.docStore })

  return createServer(async (req, res) => {
    res.setHeader("Access-Control-Allow-Origin", deps.corsOrigin ?? "*")
    res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS")
    res.setHeader(
      "Access-Control-Allow-Headers",
      String(req.headers["access-control-request-headers"] ?? "Content-Type")
    )
    res.setHeader(
      "Access-Control-Expose-Headers",
      "X-WPS-Stream-Mode,X-WPS-Filename,X-WPS-Chunk-Size,X-WPS-Delay-Ms"
    )
    res.setHeader("Access-Control-Max-Age", "600")

    if ((req.method ?? "").toUpperCase() === "OPTIONS") {
      res.statusCode = 204
      res.end()
      return
    }

    try {
      await router.handle(req, res)
    } catch (err) {
      if (res.headersSent) {
        res.destroy()
        return
      }

      if (err instanceof HttpError) {
        if (err.statusCode === 404 && err.message === "not_found") {
          sendJson(res, err.statusCode, { error: err.message, method: (req.method ?? "GET").toUpperCase(), path: req.url ?? "/" })
          return
        }
        sendJson(res, err.statusCode, { error: err.message })
        return
      }

      sendJson(res, 500, { error: "internal_error" })
    }
  })
}
