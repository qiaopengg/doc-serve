# WPS 端文档“流式写入 + 最终保真”服务端协议说明（text.docx 作为数据源）

## 0. 这份设计要解决什么问题（1 分钟版）

本次改造的前提变了：服务端“流式推送给前端”的内容来源不再是模型输出文本，而是一个已经存在的 `text.docx`（或 `text` 对应的 `text.docx`）文档。

但前端的功能诉求不变：

- 实时写入 WPS（边接收边写入，逐字/逐块打字效果）
- 接收与写入解耦（队列 + 背压，避免写入慢导致接收阻塞）
- 最终 100% 保真（最终以 docx 文件为准）

核心矛盾也不变：

- `.docx` 是 Zip 容器，二进制分片在“未完整”时无法可靠解析成可写入的文档内容

因此仍采用“双通道”，但语义调整为“Docx 回放（Replay）”：

- 通道 A（Preview Stream）：服务端从 `text.docx` 中提取“可增量应用”的内容（推荐 `preview.runs`），按 `seq` 回放给前端；前端实时写入并做打字效果
- 通道 B（Final Docx）：最终仍给 `text.docx` 的下载 URL（或二进制）；前端用最终 docx 覆盖预览，保证 100% 保真

## 0.1 服务端需要提供的能力（Checklist）

必须支持：

- [ ] Docx 回放预览流：持续输出 `preview.runs`（最低可降级为 `preview.text`）
- [ ] 最终 docx：输出 `final.docx.url`（推荐）或 `final.docx.binary`
- [ ] 顺序字段：`docId` + 严格递增 `seq`
- [ ] 结束/错误信号：`control.done` / `control.error`

强烈建议支持：

- [ ] 启动元信息帧：`control.start`（告知 source doc、能力与策略）
- [ ] 心跳：`control.heartbeat`（避免中间层超时断开）
- [ ] 限流/背压：服务端控制推送节奏；或支持客户端 `ack`（WebSocket 最适合）
- [ ] 断线续传：客户端带 `lastSeq` 续传（可选）

## 1. 总体方案：Docx Replay 双通道

### 1.1 数据源定义

- `sourceDoc`：服务端已有的 `text.docx` 文档（或 `docId=text` 映射到 `text.docx`）
- `docId`：一次回放任务的唯一 ID（用于断线续传、幂等去重、日志关联）

建议：

- `docId` 不要直接复用文件名（避免多次回放冲突）；可以用 `text@<timestamp>` 这类可读 ID
- 预览流中通过 `payload.source` 明确指向 `text.docx`，做到“内容来源可追溯”

### 1.2 服务端职责（关键变化）

服务端需要把 `text.docx` 转成可增量应用的数据，再按 `seq` 推送：

- 解析 `word/document.xml`，按段落/文本 run 顺序提取文本与最小样式意图
- 可选解析 `word/styles.xml`/主题/编号，补全 `headingLevel/styleName` 等映射
- 将提取结果切分为若干帧（每帧一个或多个 run），持续推送

前端不需要理解 docx 结构，只消费 `preview.runs` 并写入 WPS。

## 2. 传输方式与接口（推荐 NDJSON）

承载层任选：WebSocket / SSE / HTTP Fetch Stream。以下用 HTTP NDJSON 举例。

### 2.1 预览回放流（Preview Frames）

**GET** `/api/v2/docs/:docId/frames`

其中 `:docId` 是回放任务标识，服务端内部绑定其 `sourceDoc=text.docx`。

Query（可选）：

- `source=text`：指定源文档（默认 `text`，映射到 `text.docx`）
- `fromSeq=number`：断线续传，从 `fromSeq + 1` 开始补发

Response（关键）：

- `Content-Type: application/x-ndjson; charset=utf-8`
- `Cache-Control: no-store`

Body：每行一个 JSON 帧（NDJSON）。

### 2.2 最终 docx 下载（Final Docx）

**GET** `/api/v2/docs/:docId/download`

语义：下载 `sourceDoc` 对应的最终 docx（即 `text.docx`），供前端最终替换。

## 3. 核心协议：统一帧格式

每一帧都必须包含：

```json
{
  "docId": "string",
  "seq": 1,
  "type": "control.start | preview.runs | preview.text | final.docx.url | control.done | control.error | control.heartbeat",
  "ts": 1700000000000,
  "payload": {}
}
```

约束：

- `seq` 从 1 开始严格递增
- 客户端必须用 `docId + seq` 幂等去重
- 服务端允许重发（网络重试/续传常见）

## 4. 帧类型定义

### 4.1 control.start（建议）

第一帧建议为 `control.start`，用于让前端建立上下文：

```json
{
  "type": "control.start",
  "payload": {
    "source": {
      "docKey": "text",
      "fileName": "text.docx",
      "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    },
    "capabilities": {
      "preview": ["preview.runs", "preview.text"],
      "final": ["final.docx.url"]
    }
  }
}
```

### 4.2 preview.runs（推荐）

服务端从 `text.docx` 提取出的增量 runs。前端收到后立刻写入 WPS，并对 `text` 做逐字效果：

```json
{
  "type": "preview.runs",
  "payload": {
    "runs": [
      {
        "text": "标题一",
        "newParagraph": true,
        "bold": true,
        "fontSize": 16,
        "headingLevel": 1
      }
    ],
    "source": { "docKey": "text", "paraIndex": 0, "runIndex": 0 }
  }
}
```

字段建议（可按需扩展）：

- `text`：字符串（必须）
- `newParagraph`：是否在 run 末尾追加换行
- 字体：`bold / italic / underline / fontName / fontSize / color`
- 段落：`paragraph: { alignment, firstLineIndent, leftIndent, rightIndent, spaceBefore, spaceAfter, lineSpacing }`
- 样式：`headingLevel`（1-9）或 `styleName`
- `payload.source`：可选，便于调试与定位（源 doc + 段落/run 游标）

降级策略：

- 若服务端短期无法从 docx 可靠提取样式：仍输出 `preview.runs`，但只填 `text/newParagraph`，其它字段留空
- 若连 runs 切分都无法做：输出 `preview.text`（见下）

### 4.3 preview.text（最低保底）

```json
{
  "type": "preview.text",
  "payload": { "text": "从 docx 提取到的纯文本片段...\n" }
}
```

### 4.4 final.docx.url（必须）

最终 docx 用于 100% 保真替换（这里的“最终”就是 `text.docx` 本体）：

```json
{
  "type": "final.docx.url",
  "payload": {
    "url": "https://xxx/api/v2/docs/<docId>/download",
    "fileName": "text.docx",
    "expiresAt": 1700000000000
  }
}
```

### 4.5 control.done / control.error / control.heartbeat（必须/建议）

结束：

```json
{ "type": "control.done", "payload": { "reason": "completed" } }
```

错误：

```json
{
  "type": "control.error",
  "payload": { "code": "DOCX_PARSE_ERROR", "message": "xxx", "retryable": false }
}
```

心跳：

```json
{ "type": "control.heartbeat", "payload": { "serverTime": 1700000000000 } }
```

## 5. 顺序、背压与续传

### 5.1 顺序模型（必须）

- 服务端按 `seq` 递增推送
- 客户端严格按 `seq` 写入，乱序必须缓存等待或丢弃并触发续传

### 5.2 背压（建议）

无 ack 的最低要求：

- 服务端控制推送节奏（例如每帧间隔 20~100ms），避免瞬间堆积导致前端写入队列爆炸

有 ack 的推荐要求（适合 WebSocket）：

客户端：

```json
{ "type": "client.ack", "payload": { "docId": "...", "ackSeq": 123 } }
```

服务端：

- 基于 `ackSeq` 限流
- 支持断线重连后从 `ackSeq + 1` 续传

### 5.3 断线续传（建议）

客户端重连携带：

```json
{ "type": "client.resume", "payload": { "docId": "...", "lastSeq": 120 } }
```

服务端从 `lastSeq + 1` 开始补发，直到 done/error。

## 6. 一次完整交互示例（Docx Replay）

1. 客户端发起回放：`GET /api/v2/docs/<docId>/frames?source=text`
2. 服务端推送（seq 1..N）：
   - `control.start`
   - 多个 `preview.runs`（从 `text.docx` 提取并分片回放）
   - `final.docx.url`（指向 `text.docx` 下载）
   - `control.done`
3. 客户端行为：
   - `preview.runs`：立刻写入 WPS（逐字效果）
   - `final.docx.url`：下载并在 done 后整体替换，保证最终保真
