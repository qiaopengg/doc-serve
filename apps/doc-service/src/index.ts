import { loadConfig } from "./config.js"
import { FsDocStore } from "./docStore/fsDocStore.js"
import { buildServer } from "./server.js"

const config = loadConfig(process.env)
const docStore = new FsDocStore(config.docsDir)

const server = buildServer({ docStore, corsOrigin: config.corsOrigin })
server.listen(config.port, config.host, () => {
  const addr = server.address()
  if (addr && typeof addr === "object") {
    console.log(`doc-service listening on http://${addr.address}:${addr.port}`)
  }
})
