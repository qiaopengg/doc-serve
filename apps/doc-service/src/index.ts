import { loadConfig } from "./config.js"
import { FsDocStore } from "./docStore/fsDocStore.js"
import { buildServer } from "./server.js"

function isAddrInUseError(err: unknown): boolean {
  return Boolean(err && typeof err === "object" && "code" in err && (err as any).code === "EADDRINUSE")
}

async function listenOnce(
  server: import("node:http").Server,
  host: string,
  port: number
): Promise<import("node:net").AddressInfo> {
  return await new Promise((resolve, reject) => {
    const onError = (err: unknown) => {
      server.off("listening", onListening)
      reject(err)
    }
    const onListening = () => {
      server.off("error", onError)
      const addr = server.address()
      if (!addr || typeof addr === "string") {
        reject(new Error("unexpected_server_address"))
        return
      }
      resolve(addr)
    }
    server.once("error", onError)
    server.once("listening", onListening)
    server.listen(port, host)
  })
}

async function listenWithFallback(server: import("node:http").Server, host: string, startPort: number) {
  const maxAttempts = 20
  for (let i = 0; i <= maxAttempts; i += 1) {
    const port = startPort + i
    try {
      return await listenOnce(server, host, port)
    } catch (err) {
      if (isAddrInUseError(err) && i < maxAttempts) continue
      throw err
    }
  }
  throw new Error("no_available_port_found")
}

function formatListeningUrl(addr: import("node:net").AddressInfo): string {
  const host = addr.address === "0.0.0.0" || addr.address === "::" ? "127.0.0.1" : addr.address
  return `http://${host}:${addr.port}`
}

async function main() {
  const config = loadConfig(process.env)
  const docStore = new FsDocStore(config.docsDir)

  const server = buildServer({ docStore, corsOrigin: config.corsOrigin })
  const addr = await listenWithFallback(server, config.host, config.port)
  console.log(`doc-service listening on ${formatListeningUrl(addr)}`)
}

main().catch((err) => {
  console.error(err)
  process.exitCode = 1
})
