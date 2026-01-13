import { dirname, resolve } from "node:path"
import { fileURLToPath } from "node:url"

export type AppConfig = {
  port: number
  host: string
  docsDir: string
  corsOrigin: string
}

export function loadConfig(env: NodeJS.ProcessEnv): AppConfig {
  const port = Number(env.PORT ?? "3000")
  const host = env.HOST ?? "0.0.0.0"
  const moduleDir = dirname(fileURLToPath(import.meta.url))
  const docsDir = env.DOCS_DIR ? resolve(env.DOCS_DIR) : resolve(moduleDir)
  const corsOrigin = env.CORS_ORIGIN ?? "*"

  return { port, host, docsDir, corsOrigin }
}
