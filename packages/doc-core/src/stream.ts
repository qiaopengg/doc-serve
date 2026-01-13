import type { IncomingMessage, ServerResponse } from "node:http"
import { once } from "node:events"
import { pipeline } from "node:stream/promises"
import { setTimeout as sleep } from "node:timers/promises"
import type { Readable } from "node:stream"

export async function pipeReadableToResponse(
  req: IncomingMessage,
  res: ServerResponse,
  readable: Readable
): Promise<void> {
  const abort = () => {
    readable.destroy()
  }

  req.on("aborted", abort)
  res.on("close", abort)

  try {
    await pipeline(readable, res)
  } finally {
    req.off("aborted", abort)
    res.off("close", abort)
  }
}

export async function pipeReadableToResponseInChunks(
  req: IncomingMessage,
  res: ServerResponse,
  readable: Readable,
  options?: { chunkSize?: number }
): Promise<void> {
  const chunkSize = Math.max(1024, Math.min(options?.chunkSize ?? 64 * 1024, 1024 * 1024))

  const abort = () => {
    readable.destroy()
  }

  req.on("aborted", abort)
  res.on("close", abort)

  try {
    for await (const chunk of readable) {
      const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array)
      for (let offset = 0; offset < buf.length; offset += chunkSize) {
        const piece = buf.subarray(offset, Math.min(offset + chunkSize, buf.length))
        if (!res.write(piece)) {
          await once(res, "drain")
        }
      }
    }
    res.end()
  } finally {
    req.off("aborted", abort)
    res.off("close", abort)
  }
}

export async function pipeReadableToResponseAsFramedChunks(
  req: IncomingMessage,
  res: ServerResponse,
  readable: Readable,
  options?: { chunkSize?: number; delayMs?: number }
): Promise<void> {
  const chunkSize = Math.max(1024, Math.min(options?.chunkSize ?? 64 * 1024, 1024 * 1024))
  const delayMs = Math.max(0, Math.min(options?.delayMs ?? 0, 60_000))

  const abort = () => {
    readable.destroy()
  }

  req.on("aborted", abort)
  res.on("close", abort)

  try {
    let didWriteAny = false
    for await (const chunk of readable) {
      const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array)
      for (let offset = 0; offset < buf.length; offset += chunkSize) {
        const piece = buf.subarray(offset, Math.min(offset + chunkSize, buf.length))
        if (didWriteAny && delayMs > 0) {
          await sleep(delayMs)
        }
        const header = Buffer.alloc(4)
        header.writeUInt32BE(piece.length, 0)
        if (!res.write(header)) {
          await once(res, "drain")
        }
        if (!res.write(piece)) {
          await once(res, "drain")
        }
        didWriteAny = true
      }
    }
    res.end(Buffer.alloc(4))
  } finally {
    req.off("aborted", abort)
    res.off("close", abort)
  }
}

export async function pipeAsyncIterableToResponseAsNdjson(
  req: IncomingMessage,
  res: ServerResponse,
  iterable: AsyncIterable<unknown>,
  options?: { delayMs?: number }
): Promise<void> {
  const delayMs = Math.max(0, Math.min(options?.delayMs ?? 0, 60_000))

  let aborted = false
  const abort = () => {
    aborted = true
  }

  req.on("aborted", abort)
  res.on("close", abort)

  try {
    let didWriteAny = false
    for await (const frame of iterable) {
      if (aborted) break

      if (didWriteAny && delayMs > 0) {
        await sleep(delayMs)
      }

      const line = JSON.stringify(frame) + "\n"
      if (!res.write(line)) {
        await once(res, "drain")
      }
      didWriteAny = true
    }
    res.end()
  } finally {
    req.off("aborted", abort)
    res.off("close", abort)
  }
}
